VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBatchRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Receipt"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   14670
   Begin VB.Frame Frame8 
      BackColor       =   &H00DEDEDE&
      Caption         =   "Bank Balance:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1185
      Index           =   0
      Left            =   9180
      TabIndex        =   107
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtAvailableBankBal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "0.00"
         Top             =   855
         Width           =   1125
      End
      Begin VB.TextBox txtRetentions1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "0.00"
         Top             =   525
         Width           =   1125
      End
      Begin VB.TextBox txtBankBal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "0.00"
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avail.Bank Balance£"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   113
         Top             =   855
         Width           =   1380
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retentions  £"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   112
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Balance  £"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   111
         Top             =   195
         Width           =   1050
      End
   End
   Begin VB.PictureBox picPaySelected 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   4455
      ScaleHeight     =   1650
      ScaleWidth      =   4125
      TabIndex        =   96
      Top             =   1710
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H80000000&
         Caption         =   "OK"
         Height          =   375
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1170
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000000&
         Caption         =   "C&lose"
         Height          =   380
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1170
         Width           =   1400
      End
      Begin VB.TextBox txtReveiptDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1170
         TabIndex        =   99
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtReference1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   101
         Top             =   720
         Width           =   2850
      End
      Begin VB.CommandButton cmdClosePic1 
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
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   45
         Width           =   300
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   300
         Left            =   2520
         TabIndex        =   104
         Top             =   360
         Width           =   225
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "397;529"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   100
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   97
         Top             =   405
         Width           =   915
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      Height          =   195
      Left            =   135
      TabIndex        =   95
      Top             =   1305
      Width           =   195
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   135
      TabIndex        =   94
      Top             =   5175
      Width           =   4785
      Begin VB.CommandButton cmdSPayAll 
         Caption         =   "Pay in &Full"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1470
      End
      Begin VB.CommandButton cmdClearSel 
         Caption         =   "&Clear Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdPaymentDiscard 
         Caption         =   "&Clear All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.TextBox txtPostingDateGrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   11925
      TabIndex        =   88
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtReferenceGrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7830
      TabIndex        =   87
      Top             =   7200
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.TextBox txtPaymentDateGrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6390
      TabIndex        =   86
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAmountGrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4950
      TabIndex        =   85
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbAmountTypeGrid 
      Height          =   315
      Left            =   3510
      TabIndex        =   9
      Top             =   7155
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbFund 
      Height          =   315
      Left            =   2025
      TabIndex        =   8
      Top             =   7155
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cmbTenantID 
      Height          =   315
      Left            =   585
      TabIndex        =   7
      Top             =   7155
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtPostingDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   13455
      TabIndex        =   13
      Top             =   1665
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame grpUploadReceipts 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Upload Receipts"
      Height          =   750
      Index           =   0
      Left            =   11205
      TabIndex        =   72
      Top             =   8910
      Width           =   1725
      Begin VB.CommandButton cmdUploadReceipts 
         Caption         =   "&Upload Receipts"
         Height          =   375
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   225
         Width           =   1560
      End
   End
   Begin VB.TextBox txtRefInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12285
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1665
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtRptDt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   11085
      TabIndex        =   11
      Top             =   1665
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picDmdLeaseList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3135
      Left            =   11895
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   55
      Top             =   2385
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdDmdGridUnitLookup 
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
         Left            =   6080
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   20
         Width           =   255
      End
      Begin VB.TextBox txtDmdTenantSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   61
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtDmdTenantSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1650
         TabIndex        =   62
         Top             =   300
         Width           =   2370
      End
      Begin VB.TextBox txtDmdTenantSearchUnitName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4080
         TabIndex        =   63
         Top             =   300
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   56
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin MSForms.ComboBox ComboBox1 
            Height          =   315
            Left            =   480
            TabIndex        =   60
            Top             =   0
            Width           =   2415
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4260;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   8
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411"
         End
         Begin MSForms.ComboBox ComboBox2 
            Height          =   315
            Left            =   3675
            TabIndex        =   59
            Top             =   0
            Width           =   2295
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4048;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   465
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            Height          =   195
            Index           =   6
            Left            =   3000
            TabIndex        =   57
            Top             =   0
            Width           =   645
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   64
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4392
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   68
         Top             =   75
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Name"
         Height          =   195
         Index           =   8
         Left            =   1650
         TabIndex        =   67
         Top             =   75
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant ID"
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   66
         Top             =   75
         Width           =   690
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   45
         Top             =   70
         Width           =   6015
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   240
      Index           =   1
      Left            =   15345
      TabIndex        =   26
      Text            =   "this control is using to use tab in the grid"
      Top             =   7350
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   240
      Index           =   0
      Left            =   15345
      TabIndex        =   23
      Text            =   "this control is using to use tab in the grid"
      Top             =   6870
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Book Receipts:"
      Height          =   750
      Index           =   5
      Left            =   4410
      TabIndex        =   44
      Top             =   8910
      Width           =   1770
      Begin VB.CommandButton cmdGPayment 
         Caption         =   "&Book Receipts Now"
         Height          =   375
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   225
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdSPClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9630
      Width           =   1400
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ReceiptSelection"
      Height          =   735
      Left            =   135
      TabIndex        =   43
      Top             =   8910
      Width           =   4200
      Begin VB.CommandButton cmdClearSelection 
         Caption         =   "Remo&ve  Selection"
         Height          =   375
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   1575
      End
      Begin VB.CommandButton CmdDeleteAll 
         Caption         =   "R&emove All"
         Height          =   375
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton cmdSavePayment 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Generate Payment later"
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.TextBox txtSPayment 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   9765
      TabIndex        =   10
      Top             =   1665
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Receipt Method"
      Enabled         =   0   'False
      Height          =   1065
      Left            =   14865
      TabIndex        =   25
      Top             =   5595
      Visible         =   0   'False
      Width           =   2055
      Begin VB.OptionButton optBR_Cheque 
         Caption         =   "Cheque"
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optBR_Bank 
         Caption         =   "Bank"
         Height          =   300
         Left            =   360
         TabIndex        =   20
         Top             =   580
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dont delete the frame. this is in use."
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   2505
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
      Height          =   3600
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6350
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAllocation 
      Height          =   2655
      Left            =   90
      TabIndex        =   6
      Top             =   6165
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   405
      TabIndex        =   73
      Top             =   6210
      Visible         =   0   'False
      Width           =   14370
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting date"
         Height          =   195
         Index           =   9
         Left            =   11295
         TabIndex        =   89
         Top             =   45
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting date"
         Height          =   195
         Index           =   8
         Left            =   11295
         TabIndex        =   84
         Top             =   45
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting date"
         Height          =   195
         Index           =   7
         Left            =   10170
         TabIndex        =   81
         Top             =   90
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   80
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FundCode"
         Height          =   195
         Index           =   2
         Left            =   2070
         TabIndex        =   79
         Top             =   90
         Width           =   735
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount type"
         Height          =   195
         Index           =   3
         Left            =   3435
         TabIndex        =   78
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "amount"
         Height          =   195
         Index           =   4
         Left            =   4950
         TabIndex        =   77
         Top             =   90
         Width           =   525
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
         Height          =   195
         Index           =   5
         Left            =   6390
         TabIndex        =   76
         Top             =   90
         Width           =   975
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   6
         Left            =   7695
         TabIndex        =   75
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblAllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LesseeID"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   74
         Top             =   90
         Width           =   660
      End
   End
   Begin VB.Label lblBC 
      BackStyle       =   0  'Transparent
      Caption         =   "BC"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3285
      TabIndex        =   106
      Top             =   240
      Width           =   780
   End
   Begin MSForms.CommandButton cmdAll 
      Height          =   330
      Left            =   6300
      TabIndex        =   105
      Top             =   585
      Width           =   480
      Caption         =   "All"
      Size            =   "847;582"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Receipts:"
      Height          =   195
      Index           =   2
      Left            =   225
      TabIndex        =   93
      Top             =   1035
      Width           =   1050
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipts on Account/Refunds"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   92
      Top             =   5895
      Width           =   2070
   End
   Begin MSForms.TextBox txtGrandTotal 
      Height          =   300
      Left            =   9585
      TabIndex        =   91
      Top             =   9270
      Width           =   1200
      VariousPropertyBits=   679495711
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2117;529"
      Value           =   "0.00"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label lblbb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total:"
      Height          =   195
      Left            =   8460
      TabIndex        =   90
      Top             =   9315
      Width           =   840
   End
   Begin MSForms.TextBox txtAlloctotal 
      Height          =   300
      Left            =   9585
      TabIndex        =   83
      Top             =   8910
      Width           =   1200
      VariousPropertyBits=   679495711
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2117;529"
      Value           =   "0.00"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label lblbEE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipts on account:"
      Height          =   195
      Left            =   7515
      TabIndex        =   82
      Top             =   9000
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   195
      Index           =   13
      Left            =   13260
      TabIndex        =   71
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   195
      Index           =   12
      Left            =   12135
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
      Height          =   195
      Index           =   11
      Left            =   10950
      TabIndex        =   69
      Top             =   1320
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSForms.CommandButton cmdDmdTenantLookup 
      Height          =   240
      Left            =   5970
      TabIndex        =   0
      Top             =   630
      Width           =   255
      Caption         =   """"
      Size            =   "450;432"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtReference_ 
      Height          =   300
      Left            =   7200
      TabIndex        =   52
      Top             =   1035
      Visible         =   0   'False
      Width           =   2700
      VariousPropertyBits=   679495707
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "4762;529"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref:"
      Height          =   195
      Index           =   5
      Left            =   15300
      TabIndex        =   51
      Top             =   945
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipts:"
      Height          =   195
      Left            =   8865
      TabIndex        =   50
      Top             =   5355
      Width           =   1035
   End
   Begin MSForms.TextBox txtGrossTotal 
      Height          =   300
      Left            =   9945
      TabIndex        =   49
      Top             =   5310
      Width           =   1380
      VariousPropertyBits=   679495711
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2434;529"
      Value           =   "0.00"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDate"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7470
      TabIndex        =   48
      Top             =   225
      Width           =   1845
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "lblBank"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4155
      TabIndex        =   47
      Top             =   240
      Width           =   3120
   End
   Begin VB.Label lblProperty 
      BackStyle       =   0  'Transparent
      Caption         =   "lblProperty"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   46
      Top             =   645
      Width           =   2505
   End
   Begin VB.Label lblClient 
      BackStyle       =   0  'Transparent
      Caption         =   "lblClient"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   45
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant:"
      Height          =   195
      Index           =   1
      Left            =   2745
      TabIndex        =   42
      Top             =   645
      Width           =   525
   End
   Begin MSForms.ComboBox cboTenant_ 
      Height          =   300
      Left            =   8865
      TabIndex        =   21
      Top             =   990
      Visible         =   0   'False
      Width           =   3360
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5927;529"
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
      Object.Width           =   "2116"
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   41
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   40
      Top             =   645
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   195
      Index           =   2
      Left            =   1095
      TabIndex        =   39
      Top             =   1320
      Width           =   495
   End
   Begin MSForms.TextBox txtChqNo 
      Height          =   300
      Left            =   30860
      TabIndex        =   22
      Top             =   195
      Width           =   1860
      VariousPropertyBits=   679495707
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3281;529"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batch No.:"
      Height          =   195
      Index           =   4
      Left            =   30080
      TabIndex        =   38
      Top             =   240
      Width           =   690
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   195
      Index           =   3
      Left            =   6885
      TabIndex        =   37
      Top             =   225
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt £"
      Height          =   195
      Index           =   10
      Left            =   9975
      TabIndex        =   35
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Amt. £"
      Height          =   195
      Index           =   9
      Left            =   8895
      TabIndex        =   34
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      Height          =   195
      Index           =   8
      Left            =   7905
      TabIndex        =   33
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   195
      Index           =   7
      Left            =   5985
      TabIndex        =   32
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   195
      Index           =   6
      Left            =   5985
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Index           =   5
      Left            =   5025
      TabIndex        =   30
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   195
      Index           =   4
      Left            =   3800
      TabIndex        =   29
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   28
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   27
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      Height          =   195
      Index           =   0
      Left            =   2745
      TabIndex        =   24
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Width           =   14415
   End
   Begin MSForms.TextBox txtTenantName 
      Height          =   285
      Left            =   3345
      TabIndex        =   54
      Top             =   600
      Width           =   2895
      VariousPropertyBits=   746604575
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "5106;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   1110
      Left            =   120
      Top             =   75
      Width           =   9015
   End
End
Attribute VB_Name = "frmBatchRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bBRPreForm As Boolean
Dim szaTntID() As String
Dim iaPI_RowNo() As Integer, iaPC_RowNo() As Integer, iCurRow  As Integer
Dim iLeft As Integer, iTop As Integer, szSuppCnt As String
Public bSavedPayment As Boolean
Dim iFlxSPayCol As Integer
'Dim BoolManualMode As Boolean
'Added by Asif. Issue:0000534. Date: 10 Apr 2015
Public BankStatementFile As String
Dim strTransactionID As String
Dim frmLockingDialogisActive As Boolean
Dim UserSessionID As String
Dim colTransactionIDOther As String 'this variable shall hold all the locked transaction number which i s locked by other screen

Private Sub cboTenant_Click()
   RefreshGridSupp
End Sub

Private Sub RefreshGridSupp()
   Dim iRow As Integer

   For iRow = 1 To flxSPayment.Rows - 1
      flxSPayment.RowHeight(iRow) = 240
   Next iRow

   If cboTenant_.Column(0) = "ALL" Then Exit Sub

   For iRow = flxSPayment.Rows - 1 To 1 Step -1
   'Below line has been modified by anol 10 Apr 2015
   'Grid was not showing the correct value after search
   'issue 445
   
      If flxSPayment.TextMatrix(iRow, 2) <> cboTenant_.Column(0) Then
         flxSPayment.RowHeight(iRow) = 0
         cmdAll.Tag = 1
      End If
      If flxSPayment.RowHeight(iRow) = 240 Then
            flxSPayment.row = iRow
            flxSPayment.col = 11
        End If
   Next iRow
End Sub

Private Sub ConfigFlxDmdLeaseList()
   Dim szHeader As String

   flxDmdLeaseList.Clear
   flxDmdLeaseList.Cols = 4
   flxDmdLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Tenant ID|<Tenant Name|<Unit Name"
   flxDmdLeaseList.FormatString = szHeader$
   flxDmdLeaseList.ColWidth(0) = 240 'Label20(9).Left - flxDmdLeaseList.Left    '240        Solid column
   flxDmdLeaseList.ColWidth(1) = Label20(8).Left - Label20(9).Left - 20    '1400       'Tenant ID
   flxDmdLeaseList.ColWidth(2) = Label20(7).Left - Label20(8).Left - 20    'Tenant Name
   flxDmdLeaseList.ColWidth(3) = flxDmdLeaseList.Left + flxDmdLeaseList.Width - Label20(7).Left - 300 'Unit Name
   flxDmdLeaseList.Rows = 3
End Sub

Public Function PopulateDmdTenantLookup(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Me.MousePointer = vbHourglass
   
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer

   ConfigFlxDmdLeaseList
   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
         szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units, Property " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = Property.PropertyID AND " & _
            "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units, Property " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = Property.PropertyID AND " & _
            "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND " & _
            "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   flxDmdLeaseList.TextMatrix(1, 1) = "ALL"
   flxDmdLeaseList.TextMatrix(1, 2) = "All Lessees"
   cboTenant_.text = "All Lessees"

   iRow = 2

   While Not adoRst.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRst!Name
      flxDmdLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber

      iRow = iRow + 1
      adoRst.MoveNext

      If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   Wend
   adoRst.Close
   Set adoRst = Nothing

   txtTenantName.text = "ALL / All Lessees"

   Me.MousePointer = vbArrow
End Function


'Resolved by BOSL
'Issue No: 0000445.
'The function generates the expression of matching string pattern by using SQL LIKE operation and
'uses the in-built Filter function of the ADODB recordset to filter the records that match with the
'expression and finally bind the filtered records to the grid.
'Modified By: Asif. 09 Aug 2014
Private Function FilterTenantsList() As String
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim tempstr As String
   
   Dim Filter As String
   
   If Len(txtDmdTenantSearchID.text) > 0 Then
      txtDmdTenantSearchName.text = ""
      txtDmdTenantSearchUnitName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchID.text), "'", "''")
      Filter = " SageAccountNumber LIKE '%" + tempstr + "*'"
      
   End If
  'Issue No: 0000445. note 933 Wild card searching has been implemented by anol 23 Feb 2015
   If Len(txtDmdTenantSearchName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchUnitName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchName.text), "'", "''")
      Filter = " Name LIKE '%" + tempstr + "*'"
   End If

   If Len(txtDmdTenantSearchUnitName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchUnitName.text), "'", "''")
      Filter = " UnitNumber LIKE '%" + tempstr + "*'"
   End If
   
   
   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
         szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units, Property " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = Property.PropertyID AND " & _
            "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
           "From Tenants, LeaseDetails, Units, Property " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
            "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
            "LeaseDetails.Status = True AND " & _
            "Units.PropertyID = Property.PropertyID AND " & _
            "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND " & _
            "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
          "ORDER BY Tenants.SageAccountNumber;"
   End If
   
   'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
 
   adoRst.Filter = Filter
        
   flxDmdLeaseList.Clear
   'Resolved by BOSL
   ' Issue 530/992 On selection of lessees in batch receipts user not able to go back to full list.
   'Modified by anol 23 Mar 2015
   flxDmdLeaseList.Rows = adoRst.RecordCount + 2
   
   flxDmdLeaseList.TextMatrix(1, 1) = "ALL"
   flxDmdLeaseList.TextMatrix(1, 2) = "All Lessees"

   Dim iRow As Integer
   iRow = 2

   While Not adoRst.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRst!Name
      flxDmdLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber

      iRow = iRow + 1
      adoRst.MoveNext

'      If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   Wend
   
   adoRst.Close
   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing

End Function

Private Sub chkAll_Click()
     Dim i As Integer
    If chkAll.Value = 1 Then
        For i = 1 To flxSPayment.Rows - 1
           If flxSPayment.TextMatrix(i, 0) = "" And flxSPayment.TextMatrix(i, 1) <> "" Then
                flxSPayment.TextMatrix(i, 0) = "X"
           End If
        Next i
    Else
        For i = 1 To flxSPayment.Rows - 1
           'If flxSPayment.TextMatrix(i, 0) = "" And flxSPayment.TextMatrix(i, 1) <> "" Then
                flxSPayment.TextMatrix(i, 0) = ""
           'End If
        Next i
   End If
End Sub

Private Sub cmbAmountTypeGrid_GotFocus()
    flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub cmbAmountTypeGrid_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      FocusControl flxAllocation
      cmbAmountTypeGrid.text = ""
      cmbAmountTypeGrid.Visible = False
   End If
    If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
        
        
            If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
                    MsgBox "Please select a lessee ID", vbInformation, "Sorry"
                    flxAllocation.col = 2
                    Exit Sub
                    If Trim(flxAllocation.TextMatrix(flxAllocation.row, 3)) = "" Then
                            MsgBox "Please select a fund code", vbInformation, "Sorry"
                            flxAllocation.col = 3
                            Exit Sub
                    End If
            End If
            flxAllocation.TextMatrix(flxAllocation.row, 4) = RetAmountTypeID(cmbAmountTypeGrid.text)
             If cmbAmountTypeGrid.text <> "" And RetAmountTypeID(cmbAmountTypeGrid.text) = "" Then
                    MsgBox "Please select valid amount type to proceed", vbInformation, "Sorry"
                    cmbAmountTypeGrid.text = ""
                    If cmbAmountTypeGrid.Visible = True Then
                        FocusControl cmbAmountTypeGrid
                    End If

                     Exit Sub
             Else
                flxAllocation.CellForeColor = vbBlack
             End If
            If flxAllocation.TextMatrix(flxAllocation.row, 9) = "B" Then
                flxAllocation.col = 8
            Else
                If cmbAmountTypeGrid.text <> "" Then
                    flxAllocation.col = 5
                End If
            End If
       
        cmbAmountTypeGrid.Visible = False
        flxAllocation_DblClick
    End If
End Sub
Private Function RetAmountTypeID(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode AND SecondaryCode.Value='" & Code & "' " & _
             "ORDER BY SecondaryCode.Value ;", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        RetAmountTypeID = rsAmountType.Fields("C").Value
        Exit Function
   End If
   rsAmountType.Close
   adoConn.Close
    
End Function
Private Function RetAmountTypeText(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode AND SecondaryCode.Code='" & Code & "' " & _
             "ORDER BY SecondaryCode.Value ;", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        RetAmountTypeText = rsAmountType.Fields("V").Value
        Exit Function
   End If
   rsAmountType.Close
   adoConn.Close
    
End Function

Private Sub cmbAmountTypeGrid_KeyUp(KeyCode As Integer, Shift As Integer)
     Call FindComboString(cmbAmountTypeGrid, KeyCode)
End Sub

Private Sub cmbAmountTypeGrid_LostFocus()
     flxAllocation.ScrollBars = flexScrollBarVertical
     If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
                    MsgBox "Please select a lessee ID", vbInformation, "Sorry"
                    flxAllocation.col = 2
                    Exit Sub
                    If Trim(flxAllocation.TextMatrix(flxAllocation.row, 3)) = "" Then
                            MsgBox "Please select a fund code", vbInformation, "Sorry"
                            flxAllocation.col = 3
                            Exit Sub
                    End If
            End If
            flxAllocation.TextMatrix(flxAllocation.row, 4) = RetAmountTypeID(cmbAmountTypeGrid.text)
             If cmbAmountTypeGrid.text <> "" And RetAmountTypeID(cmbAmountTypeGrid.text) = "" Then
                    MsgBox "Please select valid amount type to proceed", vbInformation, "Sorry"
                    cmbAmountTypeGrid.text = ""
                    If cmbAmountTypeGrid.Visible = True Then
                        FocusControl cmbAmountTypeGrid
                    End If

                     Exit Sub
             Else
                flxAllocation.CellForeColor = vbBlack
             End If
            If flxAllocation.TextMatrix(flxAllocation.row, 9) = "B" Then
                flxAllocation.col = 8
            Else
                If cmbAmountTypeGrid.text <> "" Then
                    flxAllocation.col = 5
                End If
            End If
       
        cmbAmountTypeGrid.Visible = False
        flxAllocation_DblClick
End Sub

Private Sub cmbFund_GotFocus()
    flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub cmbFund_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        FocusControl flxAllocation
        cmbFund.text = ""
        cmbFund.Visible = False
    End If
    If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
        If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
                MsgBox "Please select a lessee ID", vbInformation, "Sorry"
                flxAllocation.col = 2
                Exit Sub
        End If
        flxAllocation.TextMatrix(flxAllocation.row, 3) = cmbFund.text
        If cmbFund.text <> "" And isvalidFund(cmbFund.text) = False Then
                MsgBox "Please select a valid fund code to proceed", vbInformation, "Please select a valid fund code"
                cmbFund.text = ""
                FocusControl cmbFund
                Exit Sub
        Else
                flxAllocation.CellForeColor = vbBlack
        End If
        
        'cmbFund.Visible = False
        FocusControl flxAllocation
        flxAllocation_DblClick
    End If
End Sub

Private Sub cmbFund_KeyUp(KeyCode As Integer, Shift As Integer)
        Call FindComboString(cmbFund, KeyCode)
End Sub

Private Sub cmbFund_LostFocus()
    flxAllocation.ScrollBars = flexScrollBarVertical
    flxAllocation.col = 4
    If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
            MsgBox "Please select a lessee ID", vbInformation, "Sorry"
            flxAllocation.col = 2
            Exit Sub
    End If
    flxAllocation.TextMatrix(flxAllocation.row, 3) = cmbFund.text
    If cmbFund.text <> "" And isvalidFund(cmbFund.text) = False Then
            MsgBox "Please select a valid fund code to proceed", vbInformation, "Please select a valid fund code"
            cmbFund.text = ""
            FocusControl cmbFund
            Exit Sub
    Else
            flxAllocation.CellForeColor = vbBlack
    End If
    
    FocusControl flxAllocation
    flxAllocation_DblClick
End Sub

Private Sub cmbTenantID_GotFocus()
     flxAllocation.ScrollBars = flexScrollBarNone
End Sub



Private Sub cmbTenantID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
          FocusControl flxAllocation
          cmbTenantID.text = ""
          cmbTenantID.Visible = False
    End If
    If (KeyAscii = 13 Or KeyAscii = 9) And flxAllocation.row <= flxAllocation.Rows - 1 Then
            If cmbTenantID.text <> "" And ValidtenantID(cmbTenantID) = False Then
                   MsgBox "Please select valid lessee ID to proceed", vbInformation, "Sorry"
                   cmbTenantID.text = ""
                   cmbTenantID.SetFocus
                   Exit Sub
             Else
                   flxAllocation.CellForeColor = vbBlack
                   flxAllocation.TextMatrix(flxAllocation.row, 2) = cmbTenantID.text
            End If
            'flxAllocation.col = 3
            If cmbTenantID.text <> "" Then
                    If flxAllocation.TextMatrix(flxAllocation.row, 9) = "" Then
                       flxAllocation.TextMatrix(flxAllocation.row, 9) = "M"
                    End If
                    'adding sequential number to the grid
                Dim i As Integer
                For i = 1 To flxAllocation.Rows - 1
                    flxAllocation.TextMatrix(i, 1) = i
                Next i
            End If
            FocusControl flxAllocation
            cmbTenantID.Visible = False
            flxAllocation_DblClick
    End If
End Sub

Private Sub cmbTenantID_KeyUp(KeyCode As Integer, Shift As Integer)
    Call FindComboString(cmbTenantID, KeyCode)
End Sub

Private Sub cmbTenantID_LostFocus()
            flxAllocation.ScrollBars = flexScrollBarVertical
            If cmbTenantID.text <> "" And ValidtenantID(cmbTenantID) = False Then
                   MsgBox "Please select valid lessee ID to proceed", vbInformation, "Sorry"
                   cmbTenantID.text = ""
                   cmbTenantID.SetFocus
                   Exit Sub
            Else
                   flxAllocation.CellForeColor = vbBlack
                   flxAllocation.TextMatrix(flxAllocation.row, 2) = cmbTenantID.text
            End If
            
            If cmbTenantID.text <> "" Then
                If flxAllocation.TextMatrix(flxAllocation.row, 9) = "" Then
                   flxAllocation.TextMatrix(flxAllocation.row, 9) = "M"
                End If
                'adding sequential number to the grid
                Dim i As Integer
                For i = 1 To flxAllocation.Rows - 1
                    flxAllocation.TextMatrix(i, 1) = i
                Next i
            End If
            FocusControl flxAllocation
            cmbTenantID.Visible = False
            flxAllocation.col = 3
           ' flxAllocation_DblClick
    
End Sub

Private Sub cmdAll_Click()
    Dim iRow As Integer
    txtTenantName.text = ""
    cmdAll.Tag = 0
    For iRow = 1 To flxSPayment.Rows - 1
      flxSPayment.RowHeight(iRow) = 240
    Next iRow
    picDmdLeaseList.Visible = False
    FocusControl cmdDmdTenantLookup
End Sub

Private Sub cmdCancel_Click()
    picPaySelected.Visible = False
End Sub

Private Sub cmdClearSel_Click()
On Error GoTo Err
        'added by anol 13 Apr 2015
      txtSPayment.text = ""
      txtSPayment.Visible = False
      txtRptDt.text = ""
      txtRptDt.Visible = False
      txtRefInput.text = ""
      txtRefInput.Visible = False
      txtPostingDate.text = ""
      txtPostingDate.Visible = False
      Dim i As Integer
      For i = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(i, 0) = "X" Then
            flxSPayment.TextMatrix(i, 0) = ""
            flxSPayment.TextMatrix(i, 11) = "0.00"
            flxSPayment.TextMatrix(i, 22) = ""
            If frmBRPreForm.chkMultiple.Value = 1 Then
                flxSPayment.TextMatrix(i, 24) = ""
                flxSPayment.TextMatrix(i, 25) = ""
            End If
            
        End If
      Next i
      flxSPayment.col = 11
      FocusControl flxSPayment
      SumUpTotal
      chkAll.Value = 0
      Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Function isvalidFund(ID As String) As Boolean
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND where FundCode='" & ID & "';", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsFund.EOF Then
        isvalidFund = True
        Exit Function
   End If
   rsFund.Close
   adoConn.Close
    
End Function
Private Function RetFundID(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND where FundCode='" & Code & "';", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsFund.EOF Then
        RetFundID = rsFund.Fields("FundID").Value
        Exit Function
   End If
   rsFund.Close
   adoConn.Close
    
End Function
Private Function RetFundCode(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND where FundID=" & Code & ";", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsFund.EOF Then
        RetFundCode = rsFund.Fields("FundCode").Value
        Exit Function
   End If
   rsFund.Close
   adoConn.Close
    
End Function
Private Function RetAmountType(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND SecondaryCode.Value='" & Code & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        RetAmountType = rsAmountType.Fields("C").Value
        Exit Function
   End If
   rsAmountType.Close
   adoConn.Close
    
End Function
Private Function RetAmountTypeDes(Code As String) As String
    Dim adoConn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND SecondaryCode.Code='" & Code & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        RetAmountTypeDes = rsAmountType.Fields("v").Value
        Exit Function
   End If
   rsAmountType.Close
   adoConn.Close
    
End Function
Private Sub LoadFund()
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND;", adoConn, adOpenStatic, adLockReadOnly
    cmbFund.Clear
    While Not rsFund.EOF
        cmbFund.AddItem rsFund.Fields("FundCode").Value
        rsFund.MoveNext
   Wend
   rsFund.Close
   adoConn.Close
    
End Sub

Private Sub cmdClearSelection_Click()
On Error Resume Next
            cmbTenantID.Visible = False
            cmbFund.Visible = False
            cmbAmountTypeGrid.Visible = False
            txtAmountGrid.Visible = False
            txtPaymentDateGrid.Visible = False
            txtReferenceGrid.Visible = False
            txtPostingDateGrid.Visible = False
            cmbTenantID.text = ""
            cmbFund.text = ""
            cmbAmountTypeGrid.text = ""
            txtAmountGrid.text = ""
            txtPaymentDateGrid.text = ""
            txtReferenceGrid.text = ""
            txtPostingDateGrid.text = ""
      Dim i, k As Integer
      For i = 1 To flxAllocation.Rows - 1
            If flxAllocation.TextMatrix(i, 0) = "X" Then
                If flxAllocation.TextMatrix(i, 9) = "M" Then
'                    If i > 1 Then
'                        flxAllocation.RowHeight(i) = 0
'                        flxAllocation.TextMatrix(i, 9) = "D"
'                    End If
                    If i = 1 And flxAllocation.Rows = 2 Then
                        For k = 0 To flxAllocation.Cols - 1
                            flxAllocation.TextMatrix(1, k) = ""
                        Next k
                   Else
                     flxAllocation.RowHeight(i) = 0
                        flxAllocation.TextMatrix(i, 9) = "D"
                    End If
                End If
                
            End If
      Next i
      FocusControl flxAllocation
      Dim dblAmount As Double
       For i = 1 To flxAllocation.Rows - 1
             If flxAllocation.TextMatrix(i, 9) <> "D" Then
                   dblAmount = dblAmount + Val(flxAllocation.TextMatrix(i, 5))
            End If
         Next i
            
     txtAlloctotal.text = Format(dblAmount, "0.00")
End Sub

Private Sub cmdClosePic1_Click()
    picPaySelected.Visible = False
End Sub

Private Sub CmdDeleteAll_Click()
    If MsgBox("Do you wish to delete all Receipts on account?", vbYesNo, "Delete?") = vbNo Then Exit Sub
'    flxAllocation.Clear
    ConfigureFlxallocation
    txtAlloctotal.text = "0.00"
'    'after deleiting all you must need to refresh upload receipt form
'
'    frmUploadReceipts.flxSPayment.Clear
'    frmUploadReceipts.flxBankTransactions.Clear
'    Dim i, k As Integer
'      For i = 1 To flxAllocation.Rows - 1
'            If flxAllocation.TextMatrix(i, 9) = "M" Then
'
'                    flxAllocation.RemoveItem (i)
'
'            End If
'      Next i
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
   
     If frmBRPreForm.chkMultiple.Value = 1 Then
     If Trim(txtReveiptDate1.text) = "" Then
        MsgBox "Please enter a valid receipt date", vbInformation, "warning"
        Exit Sub
     End If
    For i = 1 To flxSPayment.Rows - 1
        If Val(flxSPayment.TextMatrix(i, 10)) > 0 And flxSPayment.TextMatrix(i, 0) = "X" And flxSPayment.TextMatrix(i, 2) <> flxSPayment.TextMatrix(i, 26) Then
            flxSPayment.TextMatrix(i, 11) = flxSPayment.TextMatrix(i, 10)
            flxSPayment.TextMatrix(i, 22) = txtReveiptDate1.text
            flxSPayment.TextMatrix(i, 25) = lblPostingDate.ToolTipText
            flxSPayment.TextMatrix(i, 24) = txtReference1.text
            flxSPayment.TextMatrix(i, 0) = ""
         
         End If
    Next i
    End If
    picPaySelected.Visible = False
    SumUpTotal
    chkAll.Value = 0
    

End Sub

Private Sub cmdSaveReceipttonAccount_Click()
'    If ifanybankreceipt = True And IsFormLoaded(frmUploadReceipts.Name) = True Then
'        If InStr(1, frmUploadReceipts.txtInputFile.text, ".csv") = 0 Then
'            MsgBox "Please select the csv file on upload receipt", vbInformation, "Sorry"
'        End If
'    End If
    If flxAllocation.Rows = 2 Then
        If flxAllocation.TextMatrix(1, 1) = "" Then
            'ShowMsgInTaskBar "No transaction to save", "N", "P"
        Exit Sub
        End If
    End If
    Dim i As Integer
    '  Generate the next Transaction Id from the tlbReceipt
   Dim lRpt_ID As Long, lRptT_ID, lSlNumber As Long, szSQL As String
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   adoConn.BeginTrans
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbReceipt"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRpt_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
   Dim strFundcode As String
   
   Dim lngtranscount  As Double
    For i = 1 To flxAllocation.Rows - 1
        strFundcode = RetFundID(flxAllocation.TextMatrix(i, 3))
        If flxAllocation.TextMatrix(i, 5) = "" Or flxAllocation.TextMatrix(i, 2) = "" Or _
            flxAllocation.TextMatrix(i, 3) = "" Or flxAllocation.TextMatrix(i, 4) = "" Or Val(flxAllocation.TextMatrix(i, 5)) = 0 Or _
                strFundcode = "" Or flxAllocation.TextMatrix(i, 9) = "D" Then GoTo Nextrec
            
        lRpt_ID = lRpt_ID + 1
        lngtranscount = lngtranscount + 1
        szSQL = "SELECT * FROM tlbReceipt;"
      With rstSet
         .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
         If flxAllocation.TextMatrix(i, 5) > 0 Then
            lSlNumber = SlNumber("SA", "tlbReceipt", adoConn)
         Else
            lSlNumber = SlNumber("SRR", "tlbReceipt", adoConn)
         End If

         .AddNew
         !TransactionID = lRpt_ID
         !szTransactionID = !TransactionID
         If flxAllocation.TextMatrix(i, 5) > 0 Then
                !Type = 4 'CByte(sdoSA)                   'tlbTransactionType.TYPE_ID (4) = Sales Receipt on Account
                !ref = "SA" & Format(Now, "yymmddhhmmss")
                !Details = "Receipt on Account"
         Else
                !Type = 23
                !ref = "SRR" & Format(Now, "yymmddhhmmss")
                !Details = "Sales Receipt Refund"
         End If
         !SageAccountNumber = flxAllocation.TextMatrix(i, 2) 'tenant ID
         !unitid = GetUnitIDbyTenantID(flxAllocation.TextMatrix(i, 2), adoConn)
         !RDate = flxAllocation.TextMatrix(i, 6)
         !dDate = flxAllocation.TextMatrix(i, 6)
         
        
         !amount = Abs(flxAllocation.TextMatrix(i, 5)) 'Amount
         !OSAmount = Abs(RoundingNumber(!amount, 2))           'amount to be allocated
         !ReceiptView = True
            'issue 973 clientID was not writing 2021-07-27
         !ClientID = frmBRPreForm.txtClient.Tag
         !BankCode = frmBRPreForm.cmbBankAc.Column(0)
         !nominalCode = !BankCode
         !ExtRef = txtChqNo.text 'flxAllocation.TextMatrix(i, 7) 'Reference
         !RptAmtType = flxAllocation.TextMatrix(i, 4) 'amount type,right now this is saving description, it should save the code
         !SlNumber = lSlNumber
         !fundID = strFundcode 'fund code
         !postingDate = Format(flxAllocation.TextMatrix(i, 8), "dd mmmm yyyy")

         .Update
         .Close
      End With

'     Saving the split(s) of the header
      szSQL = "SELECT * FROM tlbReceiptSplit;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With rstSet
         .AddNew
         .Fields.Item("TransactionID").Value = UniqueID()
         .Fields.Item("RptHeader").Value = lRpt_ID
         .Fields.Item("FundID").Value = strFundcode 'fund code
         .Fields.Item("Amount").Value = Abs(flxAllocation.TextMatrix(i, 5)) 'Amount
         .Fields.Item("OSAmount").Value = Abs(RoundingNumber(!amount, 2))
         .Fields.Item("SplitID").Value = 1
         .Fields.Item("DueDate").Value = flxAllocation.TextMatrix(i, 6)
         If flxAllocation.TextMatrix(i, 5) > 0 Then
            .Fields.Item("Description").Value = "Receipt on Account"
         Else
            .Fields.Item("Description").Value = "Sales Receipt Refund"
         End If
         .Update
      End With
      rstSet.Close
Nextrec:
   Next i
   adoConn.CommitTrans
   adoConn.Close
'   If lngtranscount > 0 Then
'        ShowMsgInTaskBar "Total: " & lngtranscount & " Transactions has been saved successfully", "Y", "P"
'    Else
'        ShowMsgInTaskBar "No transactions was saved.", "Y", "P"
'    End If
End Sub

Private Sub Command1_Click()
   ' SaveCSV "C:\Myfile.csv", frmUploadReceipts.flxBankTransactions
   
End Sub

Private Sub flxAllocation_Click()
        cmbTenantID.Visible = False
        cmbFund.Visible = False
        cmbAmountTypeGrid.Visible = False
        txtAmountGrid.Visible = False
        txtPaymentDateGrid.Visible = False
        txtReferenceGrid.Visible = False
        txtPostingDateGrid.Visible = False
        If flxAllocation.col < 2 Then
           If flxAllocation.TextMatrix(flxAllocation.row, 0) = "" And flxAllocation.TextMatrix(flxAllocation.row, 1) <> "" And flxAllocation.TextMatrix(flxAllocation.row, 9) <> "B" Then
              flxAllocation.TextMatrix(flxAllocation.row, 0) = "X"
           Else
              flxAllocation.TextMatrix(flxAllocation.row, 0) = ""
           End If
        End If
End Sub

Private Sub flxAllocation_DblClick()
    Dim i As Integer
    If flxAllocation.row <= flxAllocation.Rows - 1 Then
            cmbTenantID.Visible = False
            cmbFund.Visible = False
            cmbAmountTypeGrid.Visible = False
            txtAmountGrid.Visible = False
            txtPaymentDateGrid.Visible = False
            txtReferenceGrid.Visible = False
            txtPostingDateGrid.Visible = False
            If flxAllocation.col = 0 Or flxAllocation.col = 1 Then
                flxAllocation.col = 2
            End If
            If flxAllocation.col = 2 Then
                'szHeader$ = "|<No|<LesseID|<FundCode|<Amount type|<Amount|<Payment Date|>Reference|>Posting date"
                    cmbTenantID.Top = flxAllocation.CellTop + flxAllocation.Top
                    cmbTenantID.Left = flxAllocation.CellLeft + flxAllocation.Left
                    cmbTenantID.Width = flxAllocation.ColWidth(flxAllocation.col)
                    cmbTenantID.text = flxAllocation.TextMatrix(flxAllocation.row, 2)
                    cmbTenantID.Visible = True
                    FocusControl cmbTenantID
                    SelTxtInCtrl cmbTenantID
              End If
               If flxAllocation.col = 3 Then
                    cmbFund.Top = flxAllocation.CellTop + flxAllocation.Top
                    cmbFund.Left = flxAllocation.CellLeft + flxAllocation.Left
                    cmbFund.Width = flxAllocation.ColWidth(flxAllocation.col)
                    cmbFund.text = flxAllocation.TextMatrix(flxAllocation.row, 3)
                    cmbFund.Visible = True
                    FocusControl cmbFund
                    SelTxtInCtrl cmbFund
              End If
              If flxAllocation.col = 4 Then
                    cmbAmountTypeGrid.Top = flxAllocation.CellTop + flxAllocation.Top
                    cmbAmountTypeGrid.Left = flxAllocation.CellLeft + flxAllocation.Left
                    cmbAmountTypeGrid.Width = flxAllocation.ColWidth(flxAllocation.col)
                    cmbAmountTypeGrid.text = RetAmountTypeText(flxAllocation.TextMatrix(flxAllocation.row, 4))
                    cmbAmountTypeGrid.Visible = True
                    FocusControl cmbAmountTypeGrid
                    SelTxtInCtrl cmbAmountTypeGrid
              End If
              If flxAllocation.col = 5 And flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
                    txtAmountGrid.Top = flxAllocation.CellTop + flxAllocation.Top
                    txtAmountGrid.Left = flxAllocation.CellLeft + flxAllocation.Left
                    txtAmountGrid.Width = flxAllocation.ColWidth(flxAllocation.col)
                    txtAmountGrid.Height = flxAllocation.RowHeight(flxAllocation.row) - 15
                    txtAmountGrid.text = flxAllocation.TextMatrix(flxAllocation.row, 5)
                    txtAmountGrid.Visible = True
                    FocusControl txtAmountGrid
                    SelTxtInCtrl txtAmountGrid
              End If
              If flxAllocation.col = 6 And flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
                txtPaymentDateGrid.Top = flxAllocation.CellTop + flxAllocation.Top
                txtPaymentDateGrid.Left = flxAllocation.CellLeft + flxAllocation.Left
                txtPaymentDateGrid.Width = flxAllocation.ColWidth(flxAllocation.col)
                txtPaymentDateGrid.Height = flxAllocation.RowHeight(flxAllocation.row) - 15
                txtPaymentDateGrid.text = flxAllocation.TextMatrix(flxAllocation.row, 6)
                txtPaymentDateGrid.Visible = True
                FocusControl txtPaymentDateGrid
                SelTxtInCtrl txtPaymentDateGrid
              End If
              If flxAllocation.col = 7 And flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
                txtReferenceGrid.Top = flxAllocation.CellTop + flxAllocation.Top
                txtReferenceGrid.Left = flxAllocation.CellLeft + flxAllocation.Left
                txtReferenceGrid.Width = flxAllocation.ColWidth(flxAllocation.col)
                txtReferenceGrid.Height = flxAllocation.RowHeight(flxAllocation.row) - 15
                txtReferenceGrid.text = flxAllocation.TextMatrix(flxAllocation.row, 7)
                txtReferenceGrid.Visible = True
                FocusControl txtReferenceGrid
                SelTxtInCtrl txtReferenceGrid
              End If
              If flxAllocation.col = 8 Then
                txtPostingDateGrid.Top = flxAllocation.CellTop + flxAllocation.Top
                txtPostingDateGrid.Left = flxAllocation.CellLeft + flxAllocation.Left
                txtPostingDateGrid.Width = flxAllocation.ColWidth(flxAllocation.col)
                txtPostingDateGrid.Height = flxAllocation.RowHeight(flxAllocation.row) - 15
                txtPostingDateGrid.text = flxAllocation.TextMatrix(flxAllocation.row, 8)
                txtPostingDateGrid.Visible = True
                FocusControl txtPostingDateGrid
                SelTxtInCtrl txtPostingDateGrid
              End If
      End If
End Sub
Public Sub FindComboString(vComboName As ComboBox, KeyCode As Integer)
    'Author  : Anol roy
    'Date    : 21 May 2015
    'Purpose : This function is used to match the written text with the Combobox Data
    'issue 571 validation
    If vComboName.ListCount = 0 Then Exit Sub
    Dim X As Integer, Y As Integer
    
    'For auto complete
    'Key 8=Back Space, 37=righr arow, 39= Left arow, 39=Delete etc...
    If Len(vComboName.text) = 0 Or KeyCode = 8 Or KeyCode = 36 Or KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46 Then Exit Sub
    Y = Len(vComboName.text)
    For X = 0 To vComboName.ListCount - 1
        If UCase(Left(vComboName.List(X), Y)) = UCase(vComboName.text) Then
            vComboName.text = vComboName.List(X)
            vComboName.SelStart = Y
            vComboName.SelLength = Len(vComboName.text) - Y
            Exit For
        End If
    Next X


End Sub

Private Sub flxAllocation_KeyPress(KeyAscii As Integer)
       
       If KeyAscii = 13 Then flxAllocation_DblClick
End Sub

Private Sub flxDmdLeaseList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then flxDmdLeaseList_Click
End Sub

Private Sub flxSPayment_Click()
    If flxSPayment.col < 5 Then
        If flxSPayment.TextMatrix(flxSPayment.row, 0) = "" And flxSPayment.TextMatrix(flxSPayment.row, 1) <> "" Then
           flxSPayment.TextMatrix(flxSPayment.row, 0) = "X"
        Else
           flxSPayment.TextMatrix(flxSPayment.row, 0) = ""
        End If
     End If
End Sub
Private Sub cmdDmdGridUnitLookup_Click()
   picDmdLeaseList.Visible = False
End Sub

Private Sub cmdDmdTenantLookup_Click()
On Error GoTo Err
   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchName.text = ""
   txtDmdTenantSearchUnitName.text = ""
   picDmdLeaseList.Top = txtTenantName.Top + txtTenantName.Height + 5
   picDmdLeaseList.Left = txtTenantName.Left + 5
   picDmdLeaseList.Visible = True
   picDmdLeaseList.ZOrder 0
   FocusControl txtDmdTenantSearchID
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub cmdSPayAll_Click()
   Dim i As Integer
   Dim j As Integer
   Dim szTran2Fix As String
   For i = 1 To flxSPayment.Rows - 1
      If Val(flxSPayment.TextMatrix(i, 10)) > 0 And flxSPayment.TextMatrix(i, 0) = "X" And flxSPayment.TextMatrix(i, 2) <> flxSPayment.TextMatrix(i, 26) Then
         flxSPayment.TextMatrix(i, 11) = flxSPayment.TextMatrix(i, 10)
      ElseIf Val(flxSPayment.TextMatrix(i, 10)) > 0 And flxSPayment.TextMatrix(i, 0) = "X" And flxSPayment.TextMatrix(i, 2) = flxSPayment.TextMatrix(i, 26) Then
           flxSPayment.TextMatrix(i, 0) = ""
            j = j + 1
            szTran2Fix = szTran2Fix & ". " & flxSPayment.TextMatrix(i, 1)
      End If
'      If frmBRPreForm.chkMultiple.Value = 1 Then
'         If flxSPayment.TextMatrix(i, 22) = "" Then
'            flxSPayment.TextMatrix(i, 22) = flxSPayment.TextMatrix(i, 6)
'         End If
'      End If
'       flxSPayment.TextMatrix(iCurRow, 2) = flxSPayment.TextMatrix(iCurRow, 26)  'if they are equal that means they have allocation problem
   Next i
   SumUpTotal
   If j > 0 Then
        MsgBox "A problem exists relating to a previous transaction entered against a selected lessee: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Transactions against this lessee have been cleared and cannot be processed. All other transactions will be processed as normal. Please contact PCM Consulting. ", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"
    End If
On Error GoTo Err
            If frmBRPreForm.chkMultiple.Value = 1 Then
               txtReveiptDate1.text = Date
               SelTxtInCtrl txtReveiptDate1
               txtReference1.text = ""
               picPaySelected.Top = 1755
               picPaySelected.Left = 4275
               picPaySelected.Visible = True
               picPaySelected.ZOrder 0
               SumUpTotal
               FocusControl txtReveiptDate1
              
               Exit Sub
Err:
                ShowMsgInTaskBar Err.description, "Y", "P"
            End If
End Sub

Private Sub cmdUploadReceipts_Click()
    If bSavedPayment = True Then
        If MsgBox("Processing a new upload file will clear your current batch receipt selections. Are you sure you wish to do this?", vbYesNo, "Clear batch receipt?") = vbYes Then
            cmdPaymentDiscard_Click
            bSavedPayment = False
        End If
    End If
    Load frmUploadReceipts
    frmUploadReceipts.Caption = "Upload Batch Receipts" + " - " + lblClient.Caption
    frmUploadReceipts.Left = 0
    frmUploadReceipts.Top = 1
    frmUploadReceipts.Show
End Sub




Private Sub flxSPayment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'added by anol 23 Feb 2016
        If flxSPayment.MouseCol = 2 Then
            flxSPayment.ToolTipText = flxSPayment.TextMatrix(flxSPayment.MouseRow, 4)
        End If
End Sub

Private Sub flxSPayment_RowColChange()
    Dim adoConn As New ADODB.Connection
    Dim rsLockDialog As New ADODB.Recordset
    Dim selcol As Integer
    Dim selRow As Integer
    Dim strSQL As String
    Dim colTransactionIDHere As String
    Dim i As Integer
    selcol = flxSPayment.col
    selRow = flxSPayment.row
    If Len(colTransactionIDOther) > 0 Then ' This procedure is only for unlock the record on each cell browsing written by anol 20190412
      'colTransactionIDOther varibale contains the transaction ID that is locked by other screen
        adoConn.Open getConnectionString
        strSQL = "Select DateTimeStamp ,UserSessionID,transactionID " & _
               "from tlbReceipt as Pt  where  (UserSessionID='' or isnull(UserSessionID='')) AND TransactionID in (" & colTransactionIDOther & ")"
        rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
        
        While Not rsLockDialog.EOF
                flxSPayment.col = 0
                For i = 1 To flxSPayment.Rows - 1
                    If flxSPayment.TextMatrix(i, 20) = rsLockDialog("transactionID").Value Then
                          flxSPayment.row = i
                          flxSPayment.CellBackColor = vbWhite
                          'now you need to lock it for this screen
                           colTransactionIDHere = colTransactionIDHere & flxSPayment.TextMatrix(i, 20) & ","
                           flxSPayment.TextMatrix(i, 27) = "" 'we are not loading sessionID in this column for current screen lock
                           flxSPayment.TextMatrix(i, 28) = ""
                           flxSPayment.TextMatrix(i, 29) = ""
                           flxSPayment.TextMatrix(i, 30) = ""
                           flxSPayment.TextMatrix(i, 31) = ""
                    End If
                 Next i
              rsLockDialog.MoveNext
        Wend
        flxSPayment.col = selcol
        flxSPayment.row = selRow
       
        If Len(colTransactionIDHere) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
            colTransactionIDHere = Left(colTransactionIDHere, Len(colTransactionIDHere) - 1)
        End If
        If Len(colTransactionIDHere) > 0 Then
            'again locking those records for current screen
            adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Batch Receipt',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                           SystemUser & "',MachineName='" & WS_Name & "'," & _
                           "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in  (" & colTransactionIDHere & ")"
        End If
        rsLockDialog.Close
        Set rsLockDialog = Nothing
        adoConn.Close
        Set adoConn = Nothing
   End If
End Sub

Private Sub Form_Activate()
    Call LockingDialog
End Sub

Private Sub Label20_Click(Index As Integer)
'MsgBox Label20(Index).Width
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    If IsNull(frmBRPreForm.txtClient.Tag) = True Then
         ShowMsgInTaskBar "Please select a client", "Y"
         'cboClient.SetFocus
         Exit Sub
   End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtReveiptDate1.text, frmBRPreForm.txtClient.Tag
End Sub

Private Sub txtAlloctotal_Change()
    txtGrandTotal.text = Format(Val(txtAlloctotal.text) + Val(txtGrossTotal.text), "0.00")
End Sub

Private Sub txtAmountGrid_GotFocus()
     flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub txtAmountGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      FocusControl flxAllocation
      txtAmountGrid.text = ""
      txtAmountGrid.Visible = False
   End If
   If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
        If flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then

            If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
                    MsgBox "Please select a lessee ID", vbInformation, "Sorry"
                    flxAllocation.col = 2
                    Exit Sub
                    If Trim(flxAllocation.TextMatrix(flxAllocation.row, 3)) = "" Then
                            MsgBox "Please select a fund code", vbInformation, "Sorry"
                            flxAllocation.col = 3
                            Exit Sub
                            If Trim(flxAllocation.TextMatrix(flxAllocation.row, 4)) = "" Then
                                MsgBox "Please select a payment type", vbInformation, "Sorry"
                                flxAllocation.col = 4
                                Exit Sub
                            End If
                    End If
            End If
            flxAllocation.TextMatrix(flxAllocation.row, 5) = Format(txtAmountGrid.text, "0.00")
            txtAmountGrid.Visible = False
              If IsNumeric(txtAmountGrid.text) = False Then
'                    MsgBox "Amount is not in correct format", vbInformation, "Warning"
'                    flxAllocation.CellForeColor = vbRed
'                    FocusControl flxAllocation
'                    Exit Sub
                Else
                    flxAllocation.CellForeColor = vbBlack
'                    If Val(txtAmountGrid.text) = 0 Then
'                        MsgBox "Amount cannot be zero", vbInformation, "Warning"
'                        flxAllocation.CellForeColor = vbRed
'                        FocusControl flxAllocation
'                        Exit Sub
'                    End If
                End If
            'flxAllocation.col = 6

            flxAllocation_DblClick
        End If
    End If
End Sub

Private Sub txtAmountGrid_LostFocus()
    flxAllocation.ScrollBars = flexScrollBarVertical
    'txtAmountGrid_KeyPress (13)
    'flxAllocation.col = 6
    If flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
            If Trim(flxAllocation.TextMatrix(flxAllocation.row, 2)) = "" Then
                    MsgBox "Please select a lessee ID", vbInformation, "Sorry"
                    flxAllocation.col = 2
                    Exit Sub
                    If Trim(flxAllocation.TextMatrix(flxAllocation.row, 3)) = "" Then
                            MsgBox "Please select a fund code", vbInformation, "Sorry"
                            flxAllocation.col = 3
                            Exit Sub
                            If Trim(flxAllocation.TextMatrix(flxAllocation.row, 4)) = "" Then
                                MsgBox "Please select a payment type", vbInformation, "Sorry"
                                flxAllocation.col = 4
                                Exit Sub
                            End If
                    End If
            End If
            flxAllocation.TextMatrix(flxAllocation.row, 5) = Format(txtAmountGrid.text, "0.00")
            txtAmountGrid.Visible = False
              If IsNumeric(txtAmountGrid.text) = False Then
'                    MsgBox "Amount is not in correct format", vbInformation, "Warning"
'                    flxAllocation.CellForeColor = vbRed
'                    FocusControl flxAllocation
'                    Exit Sub
                Else
                    flxAllocation.CellForeColor = vbBlack
'                    If Val(txtAmountGrid.text) = 0 Then
'                        MsgBox "Amount cannot be zero", vbInformation, "Warning"
'                        flxAllocation.CellForeColor = vbRed
'                        FocusControl flxAllocation
'                        Exit Sub
'                    End If
                End If
            'flxAllocation.col = 6

           ' flxAllocation_DblClick
        End If
End Sub

Private Sub txtDmdTenantSearchID_Change()
'   Dim i As Integer
'
'   If Len(txtDmdTenantSearchID.text) > 0 Then
'      txtDmdTenantSearchName.text = ""
'      txtDmdTenantSearchUnitName.text = ""
'   End If
'
'   For i = 1 To flxDmdLeaseList.Rows - 1
'      flxDmdLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 1), Len(txtDmdTenantSearchID.text))) <> UCase(txtDmdTenantSearchID.text) Then
'         flxDmdLeaseList.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 09 Aug 2014
   FilterTenantsList
   
End Sub

Private Sub txtDmdTenantSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDmdTenantSearchName
    End If
End Sub

Private Sub txtDmdTenantSearchName_Change()
'   Dim i As Integer
'
'   If Len(txtDmdTenantSearchName.text) > 0 Then
'      txtDmdTenantSearchID.text = ""
'      txtDmdTenantSearchUnitName.text = ""
'   End If
'
'   For i = 1 To flxDmdLeaseList.Rows - 1
'      flxDmdLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 2), Len(txtDmdTenantSearchName.text))) <> UCase(txtDmdTenantSearchName.text) Then
'         flxDmdLeaseList.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 09 Aug 2014
   FilterTenantsList
   
End Sub

Private Sub txtDmdTenantSearchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDmdTenantSearchUnitName
    End If
End Sub

Private Sub txtDmdTenantSearchUnitName_Change()
'   Dim i As Integer
'
'   If Len(txtDmdTenantSearchUnitName.text) > 0 Then
'      txtDmdTenantSearchID.text = ""
'      txtDmdTenantSearchName.text = ""
'   End If
'
'   For i = 1 To flxDmdLeaseList.Rows - 1
'      flxDmdLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 3), Len(txtDmdTenantSearchUnitName.text))) <> UCase(txtDmdTenantSearchUnitName.text) Then
'         flxDmdLeaseList.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 09 Aug 2014
   FilterTenantsList
   
End Sub

Private Sub flxDmdLeaseList_Click()
   Dim adoConn As New ADODB.Connection

   txtTenantName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) & " / " & flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)

   If cboTenant_.ListCount = 0 Then cboTenant_.AddItem ""
   cboTenant_.ListIndex = 0
   cboTenant_.Column(0) = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1)
   cboTenant_.Column(1) = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)

   RefreshGridSupp

   picDmdLeaseList.Visible = False
   FocusControl flxSPayment
End Sub

Private Function ValidationPostingDate() As Boolean
'issue 521
'system wrongly allows the posting date for all payments and receipt edits to be set before transaction date
' Date 26 Jan 2018 By anol
    Dim iRow As Integer
    If frmBRPreForm.chkMultiple.Value = False Then
        If DateDiff("d", Format(frmBRPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy"), Format(frmBRPreForm.txtDate.text, "dd/mm/yyyy")) > 0 Then
               MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
               Exit Function
        End If
    End If
    If frmBRPreForm.chkMultiple.Value = True Then
        For iRow = 1 To flxSPayment.Rows - 1
            If Trim(flxSPayment.TextMatrix(iRow, 25)) <> "" And flxSPayment.TextMatrix(iRow, 22) <> "" Then
                If DateDiff("d", Format(flxSPayment.TextMatrix(iRow, 25), "dd mmmm yyyy"), Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy")) > 0 Then
                    MsgBox "Posting date cannot be before the transaction date", vbInformation, flxSPayment.TextMatrix(iRow, 1)
                   Exit Function
                End If
            End If
        Next iRow
    End If
    ValidationPostingDate = True
End Function

Private Sub cmdSavePayment_Click()
   Call CheckField 'adding fundID and amount type to tblBtRptTran
   If frmBRPreForm.chkMultiple.Value = 1 Then _
      If Not CheckDataValidation Then Exit Sub

   On Error GoTo ErrHandler

   Dim iRow As Integer, szSQL As String, i As Integer
   Dim cTotalPI As Currency, cTotalPC As Currency
   Dim adoConn As New ADODB.Connection
   Dim adoBP As New ADODB.Recordset, adoBT As New ADODB.Recordset
   Dim paramBatchRptNo As String
   'issue 521
   If ValidationPostingDate = False Then
        Exit Sub
   End If
   If Val(txtGrandTotal.text) <> 0 Then
      If MsgBox("Do you wish to save your receipt selection and generate your receipt later?", vbQuestion + vbYesNo, "Batch Receipt") = vbNo Then Exit Sub
   Else
      MsgBox "There is no transaction to save.", vbInformation + vbOKOnly, "Batch Receipt"
      Exit Sub
   End If

'   Database has been connected.
   adoConn.Open getConnectionString

   paramBatchRptNo = SaveExpectedReceipt(adoConn, False)

   adoConn.Close
   Set adoConn = Nothing

   If MsgBox("Do you wish to print a list of your selected payments now?", vbQuestion + vbYesNo, "Batch Receipt") = vbNo Then Exit Sub
'  ********************************************************************************************************
'  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~     PRINT SUGGESTED RECEIPT        ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  ********************************************************************************************************

   'ShowReport App.Path & szReportPath & "\ExpectedReceipt.rpt"
   
   'Issue fixed 530 by anol date 20 apr 2015
   'Note 0001045 The batch receipts save report is printing all batch receipt transactions,
   'not just those that are currently saved. Please modify the report so it only displays the currently saved transactions.
   'I had added aa paramter paramBatchRptNo
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpectedReceipt.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Dim rep As New frmReport

      Report.ParameterFields(1).AddCurrentValue paramBatchRptNo
      Load rep
      rep.LoadReportViewer Report
      'End of modifications

   Exit Sub
ErrHandler:
   Debug.Print Err.Number & ": " & Err.description
End Sub

Private Function CheckDataValidation() As Boolean
   Dim iRow As Integer

   CheckDataValidation = True
   For iRow = 1 To flxSPayment.Rows - 1
      If (Val(flxSPayment.TextMatrix(iRow, 11)) = 0 And flxSPayment.TextMatrix(iRow, 22) <> "") Then
         MsgBox "Please input the amount for the transaction number " & flxSPayment.TextMatrix(iRow, 1), vbCritical + vbOKOnly, "Multiple Batch Receipt"
         CheckDataValidation = False
         Exit Function
      End If
      If (Val(flxSPayment.TextMatrix(iRow, 11)) <> 0 And flxSPayment.TextMatrix(iRow, 22) = "") Then
         MsgBox "Please input the payment date for the transaction number " & flxSPayment.TextMatrix(iRow, 1), vbCritical + vbOKOnly, "Multiple Batch Receipt"
         CheckDataValidation = False
         Exit Function
      End If
      
      'Resolved By BOSL. Issue: 0000550
      'Added by Asif. Date: 23 Mar 2015. Validate the reference and posting date.
'
'      If Val(flxSPayment.TextMatrix(iRow, 11)) <> 0 And flxSPayment.TextMatrix(iRow, 24) = "" Then
'        MsgBox "Please enter a reference for all receipts", vbExclamation, "No Reference"
'        CheckDataValidation = False
'        Exit Function
'      End If
    
      If Val(flxSPayment.TextMatrix(iRow, 11)) <> 0 And flxSPayment.TextMatrix(iRow, 25) = "" Then
        MsgBox "Please enter a posting date for all receipts", vbExclamation, "No Posting Date"
        CheckDataValidation = False
        Exit Function
      End If
      
   Next iRow
End Function

'  This method return the saved Batch Receipt header ID
Private Function SavedExpReceipt(adoConn As ADODB.Connection) As String
   Dim szSQL As String
   Dim adoBP As New ADODB.Recordset

   szSQL = "SELECT BR.BR " & _
           "FROM   tblBatchReceipt AS BR " & _
           "WHERE  BR.Generated = FALSE;"

'Debug.Print szSQL
   adoBP.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoBP.EOF Then
      SavedExpReceipt = adoBP.Fields.Item("BR").Value
   Else
      SavedExpReceipt = "NF"
   End If

   adoBP.Close
   Set adoBP = Nothing
End Function

'This method saves list of batch receipts list then it returns the header ID
Private Function SaveExpectedReceipt(adoConn As ADODB.Connection, bSettled As Boolean) As String
   'On Error GoTo ErrHandler

   Dim iRow As Integer, szSQL As String, i As Integer
   Dim adoBP As New ADODB.Recordset, adoBT As New ADODB.Recordset

   szSuppCnt = ""    'reset the supplier list

   szSQL = "DELETE * FROM tblBtRptTran " & _
           "WHERE BR IN (SELECT BR.BR FROM tblBatchReceipt AS BR " & _
                        "WHERE BR.Generated = FALSE);"
'Debug.Print szSQL
   adoConn.Execute szSQL
   adoConn.Execute "DELETE * FROM tblBatchReceipt WHERE Generated = FALSE;"

'  Save all Batch Receipt transactions into BatchReceipt table
   adoBP.Open "SELECT * FROM tblBatchReceipt;", adoConn, adOpenDynamic, adLockPessimistic

   With adoBP
      .AddNew
      SaveExpectedReceipt = UniqueID()
      .Fields.Item("BR").Value = SaveExpectedReceipt

      If frmBRPreForm.chkMultiple.Value Then
         .Fields.Item("BRDate").Value = Format(Now, "dd mmmm yyyy")
      Else
         .Fields.Item("BRDate").Value = Format(lblDate.Caption, "dd mmmm yyyy")
         .Fields.Item("PostingDate").Value = Format(frmBRPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
      End If

      .Fields.Item("RptOption").Value = IIf(optBR_Cheque.Value, "C", "B")
      .Fields.Item("ClientID").Value = frmBRPreForm.txtClient.Tag
      .Fields.Item("PropertyID").Value = frmBRPreForm.txtProperty.Tag
      .Fields.Item("TenantID").Value = cboTenant_.Column(0)
      .Fields.Item("BatchNo").Value = txtChqNo.text
      .Fields.Item("Generated").Value = bSettled
'      If frmBRPreForm.chkMultiple.Value Then
'         .Fields.Item("ChqNo").Value = "Mul Batch Receipt"
'      Else
      .Fields.Item("ChqNo").Value = txtReference_.text
'      End If

      .Update
      .Close
   End With
   Set adoBP = Nothing

   adoBT.Open "SELECT * FROM tblBtRptTran;", adoConn, adOpenDynamic, adLockPessimistic
'   tblBtRptTran table is used for saving the batch receipt
   With adoBT
      For iRow = 1 To flxSPayment.Rows - 1
         If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 And flxSPayment.RowHeight(flxSPayment.row) > 0 Then
            .AddNew
            .Fields.Item("BT").Value = UniqueID
            .Fields.Item("BR").Value = SaveExpectedReceipt
            .Fields.Item("TransactionID").Value = flxSPayment.TextMatrix(iRow, 20)
            .Fields.Item("TenantID").Value = flxSPayment.TextMatrix(iRow, 21)
            CountTenant (flxSPayment.TextMatrix(iRow, 21))
            .Fields.Item("TranType").Value = flxSPayment.TextMatrix(iRow, 15)
            .Fields.Item("UnitID").Value = flxSPayment.TextMatrix(iRow, 5)
            .Fields.Item("DueDate").Value = IIf(flxSPayment.TextMatrix(iRow, 6) = "", Null, flxSPayment.TextMatrix(iRow, 6))
            If frmBRPreForm.chkMultiple.Value Then
               .Fields.Item("Ref").Value = flxSPayment.TextMatrix(iRow, 24)
            Else
               .Fields.Item("Ref").Value = txtChqNo.text 'flxSPayment.TextMatrix(iRow, 7)
            End If
            .Fields.Item("Details").Value = flxSPayment.TextMatrix(iRow, 8)
            .Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 9)
            .Fields.Item("OSAmt").Value = flxSPayment.TextMatrix(iRow, 10)
            .Fields.Item("RptAmt").Value = IIf(Val(.Fields.Item("TranType").Value) = 2, flxSPayment.TextMatrix(iRow, 11) * (-1), flxSPayment.TextMatrix(iRow, 11))
            'trantype 2 means sales credit
            If frmBRPreForm.chkMultiple.Value = 1 Then
               .Fields.Item("RptDt").Value = Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy")
               .Fields.Item("PostingDate").Value = Format(flxSPayment.TextMatrix(iRow, 25), "dd mmmm yyyy")
            Else
               .Fields.Item("RptDt").Value = Format(frmBRPreForm.txtDate.text, "dd/mm/yyyy")
               .Fields.Item("PostingDate").Value = Format(frmBRPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
            End If
            .Update
         End If
      Next iRow
      '********************************Receipts on account Report**************************
           Dim strFundcode As String
           Dim lngtranscount  As Double
            Dim lRpt_ID As Long, lRptT_ID, lSlNumber As Long
            
            
            Dim rstSet As New ADODB.Recordset
            szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbReceipt"
            rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            lRpt_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
            rstSet.Close
       
                For i = 1 To flxAllocation.Rows - 1
                '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                     strFundcode = RetFundID(flxAllocation.TextMatrix(i, 3))
                    If flxAllocation.TextMatrix(i, 5) = "" Or flxAllocation.TextMatrix(i, 2) = "" Or _
                        flxAllocation.TextMatrix(i, 3) = "" Or flxAllocation.TextMatrix(i, 4) = "" Or Val(flxAllocation.TextMatrix(i, 5)) = 0 Or _
                            strFundcode = "" Or flxAllocation.TextMatrix(i, 9) = "D" Then GoTo Nextrec
                            lSlNumber = SlNumber("SA", "tlbReceipt", adoConn)
                            lRpt_ID = lRpt_ID + 1
                            .AddNew
                            .Fields.Item("BT").Value = UniqueID
                            .Fields.Item("BR").Value = SaveExpectedReceipt
                            .Fields.Item("TransactionID").Value = lRpt_ID
                            .Fields.Item("TenantID").Value = flxAllocation.TextMatrix(i, 2)
                            .Fields.Item("TranType").Value = 4 'CByte(sdoSA)                   'tlbTransactionType.TYPE_ID (4) = Sales Receipt on Account
                            .Fields.Item("UnitID").Value = GetUnitIDbyTenantID(flxAllocation.TextMatrix(i, 2), adoConn)
                            .Fields.Item("DueDate").Value = Format(flxAllocation.TextMatrix(i, 6), "dd mmmm yyyy")
                            If frmBRPreForm.chkMultiple.Value Then
                               .Fields.Item("Ref").Value = flxAllocation.TextMatrix(i, 7) ' "SA" & Format(Now, "yymmddhhmmss")
                            Else
                               .Fields.Item("Ref").Value = flxAllocation.TextMatrix(i, 7) '"SA" & Format(Now, "yymmddhhmmss")
                            End If
                            .Fields.Item("Details").Value = "Receipt on Account"
                            .Fields.Item("Amount").Value = 0 'flxAllocation.TextMatrix(i, 5)  'Amount
                            .Fields.Item("OSAmt").Value = 0 ' RoundingNumber(!amount, 2)
                            .Fields.Item("RptAmt").Value = flxAllocation.TextMatrix(i, 5)
                            If frmBRPreForm.chkMultiple.Value = 1 Then
                               .Fields.Item("RptDt").Value = flxAllocation.TextMatrix(i, 6)
                               .Fields.Item("PostingDate").Value = Format(flxAllocation.TextMatrix(i, 8), "dd mmmm yyyy")
                            Else
                               .Fields.Item("RptDt").Value = flxAllocation.TextMatrix(i, 6)
                               .Fields.Item("PostingDate").Value = Format(flxAllocation.TextMatrix(i, 8), "dd mmmm yyyy")
                            End If
                            .Fields.Item("FundId").Value = RetFundID(flxAllocation.TextMatrix(i, 3))
                            .Fields.Item("AmountType").Value = RetAmountType(flxAllocation.TextMatrix(i, 4))
                            
                            'saving Fund ID
                            'Payment Type
Nextrec:

            Next i
               
      '********************End *****************************************
   End With
   adoBT.Update
   adoBT.Close
   Set adoBT = Nothing

   Exit Function
ErrHandler:
   MsgBox Err.Number & ": " & Err.description, vbCritical + vbOKOnly, "Not Saved"
End Function

Private Sub CountTenant(szTenant As String)
   Dim szaTntCnt() As String, i As Integer

   szaTntCnt = Split(szSuppCnt, "#*#")

   For i = 0 To UBound(szaTntCnt)
      If szaTntCnt(i) <> szTenant Then
         If UBound(szaTntCnt) = 0 Then
            szSuppCnt = szTenant
         Else
            szSuppCnt = szSuppCnt & "#*#" & szTenant
         End If
      Else
         Exit For
      End If
   Next i
End Sub

Private Sub cmdSPClose_Click()
   Unload Me
End Sub

Public Sub cmdPaymentDiscard_Click()
'   Dim i As Integer, iFlxTRptCol As Integer
'
'   For i = 1 To flxSPayment.Rows - 1
'      If Val(flxSPayment.TextMatrix(i, 11)) > 0 Then
'         flxSPayment.TextMatrix(i, 11) = "0.00"
'      End If
'      If frmBRPreForm.chkMultiple.Value = 1 Then
'         If flxSPayment.TextMatrix(i, 22) <> "" Then
'            flxSPayment.TextMatrix(i, 22) = ""
'            'Resolved by BOSL
'            'Modified by Anol 03 Feb 2015
'            'issue 530 Posting date  is not clearing
'            flxSPayment.TextMatrix(i, 25) = ""
'         End If
'      End If
'   Next i
'   txtGrossTotal.text = "0.00"
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'   ClearSavedBR adoConn
'   adoConn.Close
'   Set adoConn = Nothing
   Dim i As Integer, iFlxTRptCol As Integer
        'Resolved by BOSL
         'Modified by Anol 16 Apr 2015
         'issue 530 Posting date  is not clearing
   For i = 1 To flxSPayment.Rows - 1
      flxSPayment.TextMatrix(i, 11) = "0.00"
      txtSPayment.text = ""
      txtSPayment.Visible = False
      txtRptDt.text = ""
      txtRptDt.Visible = False
      txtRefInput.text = ""
      txtRefInput.Visible = False
      txtPostingDate.text = ""
      txtPostingDate.Visible = False
      flxSPayment.TextMatrix(i, 0) = ""
      If frmBRPreForm.chkMultiple.Value = 1 Then
         'Resolved by BOSL
         'Modified by Anol 17 Feb 2015
         'issue 530 Posting date  is not clearing
         'If flxSPayment.TextMatrix(i, 22) <> "" Then
         flxSPayment.TextMatrix(i, 22) = ""
         flxSPayment.TextMatrix(i, 24) = ""
         flxSPayment.TextMatrix(i, 25) = ""
        ' End If
      End If
   Next i
   
   
   txtGrossTotal.text = "0.00"
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   ClearSavedBR adoConn
   adoConn.Close
   Set adoConn = Nothing
   bSavedPayment = False
   chkAll.Value = 0
End Sub

Private Sub flxSPayment_dblClick()
   On Error GoTo Err
   Dim i As Integer
   Dim selcol As Integer
  
   If flxSPayment.TextMatrix(flxSPayment.row, 3) = "" Then Exit Sub
 
   'added by anol for locking issue 749 will not be editable on double click
   selcol = flxSPayment.col
   flxSPayment.col = 0
   If flxSPayment.CellBackColor = vbRed Then
        MsgBox "The selected invoice is currently locked by '" & flxSPayment.TextMatrix(flxSPayment.row, 28) & _
             "' on '" & flxSPayment.TextMatrix(flxSPayment.row, 29) & "' in the '" & flxSPayment.TextMatrix(flxSPayment.row, 30) & "'" & vbCrLf & "" & _
             "screen for the Client '" & flxSPayment.TextMatrix(flxSPayment.row, 31) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
             Exit Sub
   End If
   flxSPayment.col = selcol
'
'issue 530 Smoothing operation on batch payment
'Modified by anol 19 apr 2015
If flxSPayment.col <= 11 Then
      iFlxSPayCol = 11
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound
   
      txtSPayment.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtSPayment.Top
      txtSPayment.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtSPayment.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtSPayment.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtSPayment.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      txtSPayment.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      FocusControl txtSPayment
   End If
If flxSPayment.col = 22 Then
      iFlxSPayCol = 22
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtRptDt.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtRptDt.Top
      txtRptDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtRptDt.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtRptDt.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtRptDt.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
       'line added by anol 30 Aug 2016
      SelTxtInCtrl txtRptDt
      txtRptDt.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      FocusControl txtRptDt
   End If
   If flxSPayment.col = 24 Then
      iFlxSPayCol = 24
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtRefInput.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtRefInput.Top
      txtRefInput.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtRefInput.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtRefInput.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtRefInput.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
       'line added by anol 30 Aug 2016
      SelTxtInCtrl txtRefInput
      txtRefInput.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      FocusControl txtRefInput
   End If
   If flxSPayment.col = 25 Then
      iFlxSPayCol = 25
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtPostingDate.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtRptDt.Top
      txtPostingDate.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtPostingDate.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtPostingDate.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtPostingDate.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      'line added by anol 30 Aug 2016
      SelTxtInCtrl txtPostingDate
      txtPostingDate.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      FocusControl txtPostingDate
   End If

'   bSavedPayment = False' rem by anol 20161031
   iCurRow = flxSPayment.row
   'Text1(0).Visible = True

   SelTxtInCtrl txtSPayment
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Function StarFound() As Integer
'   Dim iRow As Integer
'
'   For iRow = 1 To flxSPayment.Rows - 1
'      If flxSPayment.TextMatrix(iRow, 11) = "*" Then
'         StarFound = iRow
'         Exit Function
'      End If
'   Next iRow

   StarFound = flxSPayment.row
End Function

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<TenantID|<TenantName|<TenantPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 0          'Solid column
   conFlxGrid.ColWidth(1) = 1000       'Tenant ID
   conFlxGrid.ColWidth(2) = 3000       'Tenant Name
   conFlxGrid.ColWidth(3) = 1100       'Post Code
   conFlxGrid.Rows = 2

   conFlxGrid.RowHeight(0) = 0
End Sub

Private Sub flxSPayment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then flxSPayment_dblClick
End Sub

Private Sub Form_Load()
   bBRPreForm = True
   Me.Height = 10530
   Me.Top = 0
   Me.Left = 0
   UserSessionID = GetTimeStamp
   Me.BackColor = MODULEBACKCOLOR
   Shape1.BackColor = Me.BackColor
'   Frame5(5).Width = 2295
    bSavedPayment = False
    ConfigureFlxallocation
    LoadcmbTenantID
    LoadFund
    Call updateBankBalance
    Call WheelHook(Me.hWnd)
    
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
On Error GoTo AddFundID
   adoRst.Open "SELECT fundID from  tblBtRptTran;", adoConn, adOpenStatic, adLockReadOnly
   adoRst.Close
   GoTo Err
AddFundID:
   adoConn.Execute "ALTER TABLE tblBtRptTran ADD COLUMN fundID text(8);"
   adoConn.Execute "ALTER TABLE tblBtRptTran ADD COLUMN AmountType text(10);"
   adoConn.Close
Err:
    If adoConn.State = 1 Then
        adoConn.Close
    End If
  
   
End Sub
Public Sub updateBankBalance()
        Dim adoConn As New ADODB.Connection
        Dim adoRst As New ADODB.Recordset
        adoConn.Open getConnectionString
        Dim Balance As Double
        Dim szSQL As String
   ' find current Balance for the selected bank account and selected client ID by anol 2023-05-24
   szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) as AMTT from (" & _
            "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & Trim(frmBRPreForm.Label13(7).Caption) & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & Trim(frmBRPreForm.Label13(7).Caption) & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & Trim(frmBRPreForm.Label13(7).Caption) & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & frmBRPreForm.txtClient.Tag & "'   group by Type )"
                       
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
      txtBankBal1.text = IIf(IsNull(adoRst.Fields.Item("AMTT").Value), 0, adoRst.Fields.Item("AMTT").Value)
      txtBankBal1.text = Format(txtBankBal1.text, "0.00")
   End If
   adoRst.Close
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where isDeleted=false and BankCode='" & Trim(frmBRPreForm.Label13(7).Caption) & "' and ClientID='" & frmBRPreForm.txtClient.Tag & "'"
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        txtRetentions1.text = IIf(IsNull(adoRst.Fields.Item("DAmt").Value), 0, adoRst.Fields.Item("DAmt").Value)   'adoRst.Fields.Item("DAmt").Value
        txtRetentions1.text = Format(txtRetentions1.text, "0.00")
    End If
    adoRst.Close
    
    
'   szSQL = "SELECT * from tlbClientBanks where NominalCode and client_ID='" & txtClientList & "'"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   If Not adoRst.EOF Then
'   End If
   txtAvailableBankBal1.text = Val(txtBankBal1.text) - Val(txtRetentions1.text)
   txtAvailableBankBal1.text = Format(txtAvailableBankBal1.text, "0.00")
'   txtAvailableBankBal1
'   txtRetentions1
   adoConn.Close
End Sub
Private Sub LoadcmbTenantID()
    cmbTenantID.Clear
    Dim adoConn As New ADODB.Connection
    Dim rsTenantID As New ADODB.Recordset
    adoConn.Open getConnectionString
    LoadRptAmtType "RECEIPT AMOUNT TYPE", adoConn
    Dim szSQL As String
    szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
    rsTenantID.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsTenantID.EOF
        cmbTenantID.AddItem rsTenantID("SageAccountNumber").Value
    rsTenantID.MoveNext
    Wend
    adoConn.Close
End Sub
Public Function ValidtenantID(ID As String) As Boolean
    Dim adoConn As New ADODB.Connection
    Dim rsTenantID As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND Tenants.SageAccountNumber='" & ID & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
    rsTenantID.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsTenantID.EOF Then
       ValidtenantID = True
       Exit Function
    End If
    rsTenantID.Close
    adoConn.Close
End Function
Public Sub ConfigureFlxallocation()
    Dim szHeader As String

    frmBatchRpt.flxAllocation.Clear
   
    frmBatchRpt.flxAllocation.Cols = 9
    frmBatchRpt.flxAllocation.Rows = 2
    szHeader$ = "|<No|<LesseID|<FundCode|<Payment type|<Amount" & _
             "|<Receipt Date|<Reference|<Posting date|<Flag"
   
   frmBatchRpt.flxAllocation.FormatString = szHeader$
  Call ResizeallocaGrid
End Sub
Private Function ResizeallocaGrid()
    Dim intRow As Integer
    Dim intCol As Integer
   
    With frmBatchRpt.flxAllocation
        For intCol = 0 To .Cols - 1
           If intCol = 0 Then
               .ColWidth(intCol) = 350
           Else
              .ColWidth(intCol) = frmBatchRpt.lblAllocation(intCol).Left - frmBatchRpt.lblAllocation(intCol - 1).Left
           End If
           
        Next
    End With
End Function
' FromLoad_DataGrid method is called from the pre form.
Public Sub FromLoad_DataGrid()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   ConfigFlxSPayment

   adoConn.Open getConnectionString

'  Load all transactions in the grid
   LoadFlxSPayment adoConn

   LoadLastSavedData adoConn
   SumUpTotal

   PopulateDmdTenantLookup adoConn, szSQL

   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Sub LoadRptAmtType(szValue As String, adoConn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRst As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRst.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If
   
   cmbAmountTypeGrid.Clear
   While Not adoRst.EOF
      cmbAmountTypeGrid.AddItem adoRst!V
      adoRst.MoveNext
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub LoadLastSavedData(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset, adoBN As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, szDataPath As String

'  there might be only one record which 'Generated' value should be FALSE
   szSQL = "SELECT * FROM tblBatchReceipt WHERE Generated = FALSE;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      bSavedPayment = True
      txtReference_.text = adoRst.Fields.Item("ChqNo").Value
      adoRst.Close

      szSQL = "SELECT  B.BatchNo,T.* " & _
              "FROM tblBatchReceipt AS B, tblBtRptTran AS T " & _
              "WHERE B.BR = T.BR AND B.Generated = FALSE AND T.RptAmt > 0;"
'Debug.Print szSQLTransactionID, T.RptAmt, T.RptDt, T.Ref, T.PostingDate
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         txtChqNo.text = adoRst.Fields.Item("BatchNo").Value
         For iRow = 1 To flxSPayment.Rows - 1
            If flxSPayment.TextMatrix(iRow, 20) = adoRst.Fields.Item("TransactionID").Value Then
               flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("RptAmt").Value, "0.00")
               If frmBRPreForm.chkMultiple.Value = 1 Then
                  flxSPayment.TextMatrix(iRow, 22) = IIf(IsNull(adoRst.Fields.Item("RptDt").Value), "", _
                                                         Format(adoRst.Fields.Item("RptDt").Value, "dd/mm/yyyy"))
                  flxSPayment.TextMatrix(iRow, 24) = IIf(IsNull(adoRst.Fields.Item("Ref").Value), "", adoRst.Fields.Item("Ref").Value)
                  flxSPayment.TextMatrix(iRow, 25) = IIf(IsNull(adoRst.Fields.Item("PostingDate").Value), "", _
                                                         Format(adoRst.Fields.Item("PostingDate").Value, "dd/mm/yyyy"))
               End If
            End If
         Next iRow
         adoRst.MoveNext
      Wend
      If adoRst.RecordCount > 0 Then
          adoRst.MoveFirst
          ConfigureFlxallocation 'configure receipt on account
      End If
      Dim currow As Integer
      'loading saved receipt on account
      While Not adoRst.EOF
                If IsNull(adoRst("FundID").Value) Then GoTo Nextrec
                If adoRst("FundID").Value = "" Then GoTo Nextrec
                currow = currow + 1
                frmBatchRpt.flxAllocation.AddItem ""
                frmBatchRpt.flxAllocation.TextMatrix(currow, 0) = ""
                frmBatchRpt.flxAllocation.TextMatrix(currow, 1) = currow '1.No
                frmBatchRpt.flxAllocation.TextMatrix(currow, 2) = adoRst("TenantID").Value  '2.LesseID--combo
                frmBatchRpt.flxAllocation.TextMatrix(currow, 3) = RetFundCode(adoRst("FundID").Value) '3.FundCode--combo *****************************
                frmBatchRpt.flxAllocation.TextMatrix(currow, 5) = Format(adoRst.Fields.Item("RptAmt").Value, "0.00") '4.Amount
                frmBatchRpt.flxAllocation.TextMatrix(currow, 6) = IIf(IsNull(adoRst.Fields.Item("RptDt").Value), "", Format(adoRst.Fields.Item("RptDt").Value, "dd/mm/yyyy")) '5.Payment Date
                frmBatchRpt.flxAllocation.TextMatrix(currow, 7) = IIf(IsNull(adoRst.Fields.Item("Ref").Value), "", adoRst.Fields.Item("Ref").Value) '6.Reference
                frmBatchRpt.flxAllocation.TextMatrix(currow, 8) = IIf(IsNull(adoRst.Fields.Item("PostingDate").Value), "", Format(adoRst.Fields.Item("PostingDate").Value, "dd/mm/yyyy")) '7.Posting date
                frmBatchRpt.flxAllocation.TextMatrix(currow, 4) = RetAmountTypeDes(adoRst.Fields.Item("AmountType").Value) '4.Amount type*****
                frmBatchRpt.flxAllocation.TextMatrix(currow, 9) = "M" '0.B for automatic bank transaction
Nextrec:
                adoRst.MoveNext
      Wend
       Dim dblAmount As Double
             Dim i As Integer
             For i = 1 To flxAllocation.Rows - 1
                  If flxAllocation.TextMatrix(i, 9) <> "D" Then
                        dblAmount = dblAmount + Val(flxAllocation.TextMatrix(i, 5))
                 End If
              Next i
            
                txtAlloctotal.text = Format(dblAmount, "0.00")
   Else
      szSQL = "SELECT B.BatchNo " & _
              "FROM tblBatchReceipt AS B;"
      adoBN.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      txtChqNo.text = NextBatchNumber(adoBN)
      adoBN.Close
      Set adoBN = Nothing
      bSavedPayment = False
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Function NextBatchNumber(adoBN As ADODB.Recordset) As String
   Dim szaTemp() As String, iUnderSc As Integer

   iUnderSc = 0
   If Not adoBN.EOF Then
      While Not adoBN.EOF
         szaTemp = Split(adoBN.Fields.Item(0).Value, "_")
         If Format(Now, "ddmmyyyy") = szaTemp(0) Then
            If Val(szaTemp(1)) >= iUnderSc Then iUnderSc = Val(szaTemp(1))
         End If
         adoBN.MoveNext
      Wend
      iUnderSc = iUnderSc + 1
      NextBatchNumber = Format(Now, "ddmmyyyy") & "_" & iUnderSc
   Else
      NextBatchNumber = Format(Now, "ddmmyyyy") & "_1"
   End If
End Function
Private Sub LockingDialog()
    Dim strSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsLockDialog As New ADODB.Recordset
    Dim szHeader As String
    Dim iRow As Integer
    Dim StrWhere2 As String
    
    If frmLockingDialogisActive = False Then
        frmLockingDialogisActive = True
    Else
        Exit Sub
    End If
    adoConn.Open getConnectionString
    strSQL = "SELECT Rt.DateTimeStamp,Rt.Module,Rt.ClientID,MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber AS INV," & _
                "Rt.UserSessionID,Rt.WindowsUserName,Rt.MachineName,Rt.PrestigeUserName,Rt.ServerIPaddress," & _
                "Rt.TransactionID, Rt.SlNumber AS C_SL, Rt.DemandRef, Rt.AdjTag, " & _
                "Rt.SageAccountNumber, Rt.UnitID, Rt.DDate, Rt.Ref, Rt.Details, Rt.Amount, " & _
                "Rt.OSAmount, Rt.Type, TT.DESCRIPTION, T.Name, T.SageAccountNumber, " & _
                "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, U.PropertyID " & _
           "FROM (((tlbReceipt AS Rt INNER JOIN tlbTransactionTypes AS TT ON Rt.Type = TT.TYPE_ID) " & _
                "INNER JOIN Tenants AS T ON Rt.SageAccountNumber = T.SageAccountNumber) " & _
                "INNER JOIN Units AS U ON Rt.UnitID = U.UnitNumber) INNER JOIN " & _
                "Property AS P ON U.PropertyID = P.PropertyID " & _
           "WHERE Rt.OSAmount > 0 AND UserSessionID<>'" & UserSessionID & "' AND Len(DateTimeStamp)>0 AND " & _
                "Rt.ReceiptView = True AND " & _
                "(TT.TYPE_ID = 1) AND " & _
                "P.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
           "ORDER BY T.SageAccountNumber, Rt.TransactionID;"
    rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsLockDialog.EOF Then
        With frmLockingDialog
        .Show
        .flxLockedModule.Clear
        .flxLockedModule.Cols = 10
         szHeader$ = "|<DateTimeStamp|<Module|<Client|<Invoice|<UserSessionID|<WindowsUserName|<MachineName" & _
                  "|<PrestigeUserName|<ServerIPaddress"
           .flxLockedModule.FormatString = szHeader$
           
                  
        '.flxLockedModule.RowHeight(0) = 0
        .flxLockedModule.ColWidth(0) = 130 'selection grid
        .flxLockedModule.ColWidth(1) = 1600 'DateTimeStamp
        .flxLockedModule.ColWidth(2) = 1360 'Module
        .flxLockedModule.ColWidth(3) = 950 'Client
        .flxLockedModule.ColWidth(4) = 1000 'Invoice
        .flxLockedModule.ColWidth(5) = 0 '1800'UserSessionID
        .flxLockedModule.ColWidth(6) = 1200
        .flxLockedModule.ColWidth(7) = 1150
        .flxLockedModule.ColWidth(8) = 1000
        .flxLockedModule.ColWidth(9) = 1000
       
        iRow = 1
        .flxLockedModule.Rows = rsLockDialog.RecordCount + 1
        While Not rsLockDialog.EOF
               .flxLockedModule.TextMatrix(iRow, 1) = IIf(IsNull(rsLockDialog("DateTimeStamp").Value), "", rsLockDialog("DateTimeStamp").Value)
               .flxLockedModule.TextMatrix(iRow, 2) = IIf(IsNull(rsLockDialog("Module").Value), "", rsLockDialog("Module").Value)
               .flxLockedModule.TextMatrix(iRow, 3) = IIf(IsNull(rsLockDialog("ClientID").Value), "", rsLockDialog("ClientID").Value)
               .flxLockedModule.TextMatrix(iRow, 4) = IIf(IsNull(rsLockDialog("inv").Value), "", rsLockDialog("inv").Value)
               .flxLockedModule.TextMatrix(iRow, 5) = IIf(IsNull(rsLockDialog("UserSessionID").Value), "", rsLockDialog("UserSessionID").Value)
               .flxLockedModule.TextMatrix(iRow, 6) = IIf(IsNull(rsLockDialog("WindowsUserName").Value), "", rsLockDialog("WindowsUserName").Value)
               .flxLockedModule.TextMatrix(iRow, 7) = IIf(IsNull(rsLockDialog("MachineName").Value), "", rsLockDialog("MachineName").Value)
               .flxLockedModule.TextMatrix(iRow, 8) = IIf(IsNull(rsLockDialog("PrestigeUserName").Value), "", rsLockDialog("PrestigeUserName").Value)
               .flxLockedModule.TextMatrix(iRow, 9) = IIf(IsNull(rsLockDialog("ServerIPaddress").Value), "", rsLockDialog("ServerIPaddress").Value)
               iRow = iRow + 1
        rsLockDialog.MoveNext
        
        Wend
        End With
    Else
        'frmLockingDialog.Visible = False
    End If
    rsLockDialog.Close
    Set rsLockDialog = Nothing
    adoConn.Close
    Set adoConn = Nothing

End Sub
Private Sub LoadFlxSPayment(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset, rdoSplits As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, szDataPath As String, iCol As Integer
   Dim bSwitch As Boolean, iSupp As Integer
   Dim rsUserSessionID As String ' This variable shall store the value of timestamp from the recordset
   Dim colTransactionID As String 'this variable shall hold all the locked transaction number which i s locked by this screen
   
   'strTransactionID = "" 'this vaiable shall collect all the transactionID that shall be used for check some while saving the data
   szSQL = "SELECT Rt.TransactionID, Rt.SlNumber AS C_SL, Rt.DemandRef, Rt.AdjTag, " & _
             "Rt.UserSessionID,Rt.WindowsUserName,Rt.MachineName,Rt.PrestigeUserName,Rt.ServerIPaddress,RT.Module,RT.clientID, " & _
                "Rt.SageAccountNumber, Rt.UnitID, Rt.DDate, Rt.Ref, Rt.Details, Rt.Amount, " & _
                "Rt.OSAmount, Rt.Type, TT.DESCRIPTION, T.Name, T.SageAccountNumber, " & _
                "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, U.PropertyID " & _
           "FROM (((tlbReceipt AS Rt INNER JOIN tlbTransactionTypes AS TT ON Rt.Type = TT.TYPE_ID) " & _
                "INNER JOIN Tenants AS T ON Rt.SageAccountNumber = T.SageAccountNumber) " & _
                "INNER JOIN Units AS U ON Rt.UnitID = U.UnitNumber) INNER JOIN " & _
                "Property AS P ON U.PropertyID = P.PropertyID " & _
           "WHERE Rt.OSAmount > 0 AND " & _
                "Rt.ReceiptView = True AND " & _
                "(TT.TYPE_ID = 1) AND " & _
                "P.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
           "ORDER BY T.SageAccountNumber, Rt.TransactionID;"
'Debug.Print szSQL

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRst.EOF
      flxSPayment.TextMatrix(iRow, 20) = adoRst!TransactionID
'     Check the transaction type and changing the text color of the credit notes
      If adoRst.Fields.Item("Type").Value = 2 Or adoRst.Fields.Item("Type").Value = 4 Then
         flxSPayment.row = iRow
         For iCol = 2 To 11
            flxSPayment.col = iCol
            flxSPayment.CellForeColor = vbRed
         Next iCol
      End If

      flxSPayment.TextMatrix(iRow, 1) = adoRst!PF & IIf(IsNull(adoRst!C_SL), "", adoRst!C_SL)
      flxSPayment.TextMatrix(iRow, 2) = adoRst!SageAccountNumber 'adoRst!Name' Issue: 0000534 Modified By Asif.
      If InStr(adoRst!description, "Invoice") > 0 Then
         flxSPayment.TextMatrix(iRow, 3) = IIf(adoRst!AdjTag = "Y", "ADJI", adoRst!description)
      Else
         flxSPayment.TextMatrix(iRow, 3) = adoRst!description
      End If

      flxSPayment.TextMatrix(iRow, 4) = adoRst!Name ' adoRst!SageAccountNumber'fixed by anol for tooltip object
      flxSPayment.TextMatrix(iRow, 5) = adoRst!unitid
      flxSPayment.TextMatrix(iRow, 6) = IIf(Not IsNull(adoRst!dDate), Format(adoRst!dDate, "dd/mm/yyyy"), "")
      flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!ref), "", adoRst!ref)
      flxSPayment.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!Details), "", adoRst!Details)
      flxSPayment.TextMatrix(iRow, 9) = Format(adoRst!amount, "0.00")
      flxSPayment.TextMatrix(iRow, 10) = Format(adoRst!OSAmount, "0.00")
      flxSPayment.TextMatrix(iRow, 11) = "0.00"
      flxSPayment.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!DemandRef), "", adoRst!DemandRef)
      flxSPayment.TextMatrix(iRow, 15) = adoRst!Type
      flxSPayment.TextMatrix(iRow, 21) = adoRst!SageAccountNumber
      If flxSPayment.ColWidth(22) = 0 Then
         flxSPayment.TextMatrix(iRow, 22) = adoRst!propertyID
      Else
         flxSPayment.TextMatrix(iRow, 23) = adoRst!propertyID
      End If
      
'    flxSPayment.ColWidth(27) = 0     'userSessionID
'    flxSPayment.ColWidth(28) = 0     'WindowsuserName
'    flxSPayment.ColWidth(29) = 0     'Machine Name
'    flxSPayment.ColWidth(30) = 0     'Module
'    flxSPayment.ColWidth(31) = 0     'ClientID
     rsUserSessionID = IIf(IsNull(adoRst!UserSessionID), "", adoRst!UserSessionID)
     flxSPayment.TextMatrix(iRow, 28) = IIf(IsNull(adoRst!WindowsUserName), "", adoRst!WindowsUserName)
     flxSPayment.TextMatrix(iRow, 29) = IIf(IsNull(adoRst!MachineName), "", adoRst!MachineName)
     flxSPayment.TextMatrix(iRow, 30) = IIf(IsNull(adoRst!Module), "", adoRst!Module)
     flxSPayment.TextMatrix(iRow, 31) = IIf(IsNull(adoRst!ClientID), "", adoRst!ClientID)
     If Len(rsUserSessionID) > 0 And rsUserSessionID <> UserSessionID Then 'this means it is locked by other screen and now mark it red
             flxSPayment.col = 0
             flxSPayment.row = iRow
             flxSPayment.CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
             flxSPayment.TextMatrix(iRow, 27) = IIf(IsNull(adoRst!UserSessionID), "", adoRst!UserSessionID) 'Keeping the User SersesssionID to check the lock
             colTransactionIDOther = colTransactionIDOther & flxSPayment.TextMatrix(iRow, 20) & ","
      Else 'collect the transaction ID which needs to be locked
             colTransactionID = colTransactionID & flxSPayment.TextMatrix(iRow, 20) & ","
             Debug.Print flxSPayment.TextMatrix(iRow, 20)
      End If
      
      adoRst.MoveNext
      If Not adoRst.EOF Then
         flxSPayment.AddItem "" ' this procedure is taking 1 sec for 2k additem anol 20190424
         iRow = iRow + 1

         If adoRst!SageAccountNumber <> flxSPayment.TextMatrix(iRow - 1, 21) Then
            bSwitch = Not bSwitch
            iSupp = iSupp + 1
         End If

         If bSwitch Then UMarkRowFlxGrid flxSPayment, iRow
      End If
   Wend

   adoRst.Close
   Set adoRst = Nothing
   Call ChecksumValidationOnLoad(adoConn)
   
'lock this records when it is not red written by anol 20190425 Issue 756
    If Len(colTransactionIDOther) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOther = Left(colTransactionIDOther, Len(colTransactionIDOther) - 1)
    End If
                   
    If Len(colTransactionID) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionID = Left(colTransactionID, Len(colTransactionID) - 1)
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Batch Receipt',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                   "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionID & ")"
    End If

End Sub

Private Function CalculateTenantBalance() As Boolean
   Dim iRow As Integer
   Dim i As Integer, cTotalPay As Currency

   CalculateTenantBalance = False

   For i = 1 To flxSPayment.Rows - 2
      If flxSPayment.TextMatrix(flxSPayment.row, 21) = flxSPayment.TextMatrix(i, 21) Then
         If flxSPayment.TextMatrix(i, 15) = 2 Or flxSPayment.TextMatrix(i, 15) = 4 Then
            cTotalPay = cTotalPay - flxSPayment.TextMatrix(i, 11)
         Else
            cTotalPay = cTotalPay + flxSPayment.TextMatrix(i, 11)
         End If
      End If
   Next i

   If cTotalPay <= 0 Then
      MsgBox "Accumulated receipt value of this supplier is £" & cTotalPay & "" & Chr(13) & _
             "Receipt value must be positive.", vbCritical + vbOKOnly, "Batch Receipt"
      flxSPayment.TextMatrix(flxSPayment.row, 11) = "0.00"
   End If

   CalculateTenantBalance = True
End Function

Private Function ifanybankreceipt() As Boolean
    Dim i As Integer
    For i = 1 To flxAllocation.Rows - 1
        If flxAllocation.TextMatrix(i, 8) = "B" Then
            ifanybankreceipt = True
        Exit For
        End If
    Next i

End Function
Private Function CheckField()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Dim adoRst As New ADODB.Recordset
    On Error GoTo AddFundID
       adoRst.Open "SELECT fundID from  tblBtRptTran;", adoConn, adOpenStatic, adLockReadOnly
       adoRst.Close
       GoTo Err
AddFundID:
       adoConn.Execute "ALTER TABLE tblBtRptTran ADD COLUMN fundID text(8);"
       adoConn.Execute "ALTER TABLE tblBtRptTran ADD COLUMN AmountType text(10);"
       adoConn.Close
Err:
        If adoConn.State = 1 Then
            adoConn.Close
        End If
End Function
Private Sub cmdGPayment_Click()
'   When we use the save method/button then receipt is generated with Generated = FALSE
'   condition and when we book it generated with Generated = true condition
    If Val(txtGrandTotal.text) = 0 Then
      MsgBox "There is no transaction to save.", vbInformation + vbOKOnly, "Batch Receipt"
      Exit Sub
   End If
   Call CheckField 'adding fundID and amount type to tblBtRptTran
   'issue 521
   If ValidationPostingDate = False Then
        Exit Sub
   End If
   If frmBRPreForm.chkMultiple.Value = 1 Then _
      If Not CheckDataValidation Then Exit Sub
   If cboTenant_.text = "" Then Exit Sub

   Dim iRow As Integer, szSQL As String, i As Integer, szRecID As String
   Dim cTotalPI As Currency, cTotalPC As Currency, szOutPutFileLoc  As String
   Dim adoConn As New ADODB.Connection
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report, szTemp As String, szEB As String, szFileName As String, szFileEtn As String

   If MsgBox("Do you wish to generate your batch receipts now?", vbQuestion + vbYesNo, "Batch Receipt") = vbNo Then Exit Sub

   adoConn.Open getConnectionString
   adoConn.BeginTrans
   'The below line means there is no receipt at invoice but there is receipt on account
   'Val(txtGrossTotal.text) = 0 means first grid total is zero
    If Val(txtGrossTotal.text) = 0 Then GoTo SaveRecOnaccount
'  bSavedPayment : TRUE  --> there are some receipts found which are saved, not booked
'  LoadLastSavedData function is making this boolean true when there is some row in tblBtRptTran table
'  bSavedPayment : FALSE --> there are no receipt found which are undone
   If bSavedPayment Then
      'this condition is true when it is loaded with prevoiusly saved receipt
      szRecID = SavedExpReceipt(adoConn) 'returning tblBatchReceipt.BR WHERE  BR.Generated = FALSE 'but later no using of szRecID, so this is doing nothing
   Else
      'when there is no saved transaction
      'First save all receipts in the batch receipt table, which will help to print report
      szRecID = SaveExpectedReceipt(adoConn, True) 'saving master record tblBatchReceipt->Generated= true (with second parameter)
   End If

   ReceiptTotalTenantWise
  
   For i = 0 To UBound(szaTntID, 2) - 1
      BookNormalReceiptOfSI adoConn, i
   Next i

   adoConn.Execute "UPDATE tblBatchReceipt SET Generated = TRUE;"



SaveRecOnaccount:
'Added by Asif. Issue: 0000534. Date: 10 Apr 2015
'On Error GoTo FailedToLogUploadReceipt
      If (BankStatementFile <> "") And Val(txtAlloctotal.text) > 0 Then
         
'         Dim szSQL As String
         
         szSQL = "INSERT INTO tlbBatchReceiptUploadFile " & _
         "(" & _
            "FileName, UploadDate, UploadedBy, UploadedFrom " & _
         ") " & _
         "VALUES(" & _
            "'" & BankStatementFile & "', '" & Format(DateTime.Now, "dd/mm/yyyy") & "', '', '' " & _
         ")"
'         Debug.Print szSQL
         adoConn.Execute szSQL
      End If
      
'FailedToLogUploadReceipt:
'''
   If ChecksumValidationOnAllocation(adoConn) = False Then 'false means there is some Inconsistent data as receipt OSamount
         adoConn.RollbackTrans
'         adoconn.Close
'         Set adoconn = Nothing
'         Exit Sub
   Else
      adoConn.CommitTrans
   End If
   adoConn.Close
   Set adoConn = Nothing

   'If Val(txtAlloctotal.text) > 0 Then
            cmdSaveReceipttonAccount_Click
   'End If
   
   adoConn.Open getConnectionString
   '--------------------------------------------------------------------------------------------
   'anol have moved it here because receipt on account was not posting to NLPosting table 20161030
'  Export Transactions to Nominal Ledger (NLPosting table)
   Export_SRnSRR_2_NL adoConn
'--------------------------------------------------------------------------------------
'   adoConn.Close
'   Set adoConn = Nothing
   
    'modified by anol 20161026
   If IsFormLoaded(frmUploadReceipts.Name) = True And frmUploadReceipts.HasRecforBatchReceipt = True Then
            If Trim(frmUploadReceipts.txtInputFile.text) <> "" Then
'                 If frmUploadReceipts.cmbBank.text = "Barclays" Then
                      SaveCSVFSB Mid(frmUploadReceipts.txtInputFile.text, 1, Len(frmUploadReceipts.txtInputFile.text) - 4) & "_" & Format(Now, "yyyymmdd-hhMM") & ".csv"
'                 End If
'                 If frmUploadReceipts.cmbBank.text = "BIB" Then
'                      SaveCSVFSB Mid(frmUploadReceipts.txtInputFile.text, 1, Len(frmUploadReceipts.txtInputFile.text) - 4) & "_" & Format(Now, "yyyymmdd-hhMM") & ".csv"
'                 End If
'                 If frmUploadReceipts.cmbBank.text = "TSB" Then
'                      SaveCSVFSB Mid(frmUploadReceipts.txtInputFile.text, 1, Len(frmUploadReceipts.txtInputFile.text) - 4) & "_" & Format(Now, "yyyymmdd-hhMM") & ".csv"
'                 End If
                 If IsFormLoaded(frmUploadReceipts.Name) = True Then
                      Unload frmUploadReceipts
                 End If
            End If
   End If
'   ShowMsgInTaskBar "Batch receipt has been processed successfully.", "Y", "P"
'   Unload Me
   
   ConfigFlxSPayment
   LoadFlxSPayment adoConn
   Frame5(5).Refresh
   cmdGPayment.Refresh
   adoConn.Close
   Set adoConn = Nothing
   MsgBox "Batch receipt has been processed successfully.", vbInformation, "processed"
End Sub
Private Function ChecksumValidationOnAllocation(adoConn As ADODB.Connection) As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    'this function is written by anol 20181123 when found that (issue 673 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This functionshall compare allocation with receipt amount and outstanding amount
    On Error GoTo Err
    'when you do not enter any value in the first grid you will encounter an error to eescape that you are using on error
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    Dim strWhere As String
    Dim strTenantID As String
    For i = 0 To UBound(szaTntID, 2) - 1
        strTenantID = strTenantID & IIf(strTenantID = "", "'", ",'") & szaTntID(0, i) & "'"
    Next i
    If strTenantID <> "" Then
        strWhere = " AND R.sageaccountnumber IN (" & strTenantID & " )"
    End If
    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions where DeleteFlag=False group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID  " & strWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
                    

        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            rsChecksum.MoveNext
        Wend
        
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.DeleteFlag=False AND A.Totran IS NULL " & strWhere & " AND osamount>amount", adoConn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
                ChecksumValidationOnAllocation = True
        Else
                MsgBox "A problem occurred while creating this transaction: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Please contact PCM Consulting. This transaction has not been saved.", _
                     vbInformation + vbOKOnly, "Batch receipt not saved!"
        End If
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    Exit Function
Err:
End Function
Private Function ChecksumValidationOnLoad(adoConn As ADODB.Connection) As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    'this function is written by anol 20181123 when found that (issue 673 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This functionshall compare allocation with receipt amount and outstanding amount
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    Dim j As Integer
    Dim strWhere As String
    Dim strTenantID As String
    Dim temp
'    For i = 0 To UBound(szaTntID, 2) - 1
'        strTenantID = strTenantID & IIf(strTenantID = "", "'", ",'") & szaTntID(0, i) & "'"
'    Next i
'    If strTenantID <> "" Then
'        strWhere = " AND R.sageaccountnumber IN (" & strTenantID & " )"
'    End If
        strWhere = " AND R.ClientID='" & frmBRPreForm.txtClient.Tag & "'"
        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  where DeleteFlag=False group By ToTran ) as A " & _
                        "Where a.ToTran = r.TransactionID  " & strWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
                    

        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            strTenantID = strTenantID + IIf(strTenantID = "", "", "/") + rsChecksum("sageaccountnumber").Value
            i = i + 1
            rsChecksum.MoveNext
        Wend
        
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.DeleteFlag=False AND A.Totran IS NULL " & strWhere & " AND osamount>amount", adoConn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             strTenantID = strTenantID + IIf(strTenantID = "", "", "/") + adoRst("sageaccountnumber").Value
               i = i + 1
            rsChecksum.MoveNext
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
                ChecksumValidationOnLoad = True
        Else
                'ReDim ChecksumTenant(i) As String
                temp = Split(strTenantID, "/")
                For i = 0 To UBound(temp)
                    ' ChecksumTenant(i) = temp(i)
                     For j = 0 To flxSPayment.Rows - 1
                            If flxSPayment.TextMatrix(j, 2) = temp(i) Then
                                flxSPayment.TextMatrix(j, 26) = temp(i)
                            End If
                     Next j
                Next i
                MsgBox " A problem exists relating to a previous transaction entered against a lessee in this Batch: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Please contact PCM Consulting.", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"

        End If
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Private Function ChecksumValidationOnLostFocus(strTenantID As String) As Boolean  ' if returns true then true means data is fine and false means there is some inconsistent data
    'this function is written by anol 20181123 when found that (issue 673 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This functionshall compare allocation with receipt amount and outstanding amount
    Dim adoConn As New ADODB.Connection
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    Dim strWhere As String
    'Dim strTenantID As String
    adoConn.Open getConnectionString
'    For i = 0 To UBound(szaTntID, 2) - 1
'        strTenantID = strTenantID & IIf(strTenantID = "", "'", ",'") & szaTntID(0, i) & "'"
'    Next i
'    If strTenantID <> "" Then
'        strWhere = " AND R.sageaccountnumber IN (" & strTenantID & " )"
'    End If
    'strWhere = " AND R.ClientID='" & frmBRPreForm.txtClient.Tag & "'"
    strWhere = " AND R.sageaccountnumber='" & strTenantID & "'"
    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  where DeleteFlag=False group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID  " & strWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
                    

        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            rsChecksum.MoveNext
        Wend
        
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID  A.DeleteFlag=False AND A.Totran IS NULL " & strWhere & " AND osamount>amount", adoConn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
                   ChecksumValidationOnLostFocus = True
        Else
                   MsgBox "A problem exists relating to a previous transaction entered against the selected lessee: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Please contact PCM Consulting. ", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"


        End If
        adoConn.Close
        Set adoConn = Nothing
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Public Sub SaveCSV(ByVal strFileName As String, ByRef msFlex As MSHFlexGrid)
    Const SEPARATOR_CHAR As String = ","

    Dim intFreeFile As Integer
    Dim strLine As String
    Dim r As Integer
    Dim c As Integer

    intFreeFile = FreeFile
    
    Open strFileName For Output As #intFreeFile
    
    With msFlex
        ' Every row
        For r = 0 To .Rows - 1
            strLine = ""
            If r = 0 Then
                'First column names
                strLine = "Number,Date,Account,Amount,Subcategory,Memo"
            Else
                'Values
                'Date,Memo,Amount, Subcategory, Number,account
                '[F7],[F13],[F9],[F10],[F11],[F2]
                 If .TextMatrix(r, 10) = "No" Then
                     strLine = .TextMatrix(r, 5) 'Number
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 1) 'Date
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 6) 'Account
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 0) 'Amount =OS amount
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 4) 'Subcategory
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 2) 'Memo
                 End If
            End If
            If Trim(strLine) <> "" Then
                Print #intFreeFile, strLine
            End If
        Next r
    End With
    
    Close #intFreeFile
End Sub
Public Sub SaveCSVBIB(ByVal strFileName As String, ByRef msFlex As MSHFlexGrid)
    Const SEPARATOR_CHAR As String = ","

    Dim intFreeFile As Integer
    Dim strLine As String
    Dim r As Integer
    Dim c As Integer

    intFreeFile = FreeFile
    
    Open strFileName For Output As #intFreeFile
    
    With msFlex
        ' Every row
        For r = 0 To .Rows - 1
            strLine = ""
            If r = 0 Then
                'First column names
               ' strLine = "F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13"
                strLine = "GROUP,ACC ID,ACCOUNT NO,TYPE,BANK CODE,CURR,ENTRY DATE,AS AT,AMOUNT,TLA CODE,CHEQUE NO,STATUS,DESCRIPTION"

            Else
                'Values
                'Date,Memo,Amount, Subcategory, Number,account
                '[F7],[F13],[F9],[F10],[F11],[F2]
                 If .TextMatrix(r, 10) = "No" Then
                     strLine = ","
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 6) 'Account'[F2]
                     strLine = strLine & ",,,"
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 1) 'Date'[F7]
                     strLine = strLine & ","
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 0) 'Amount'[F9]=OS amount
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 4) 'Subcategory'[F10]
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 5) 'Number'[F11]
                     strLine = strLine & ","
                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 2) 'Memo'[F13]
                 End If
            End If
            If Trim(strLine) <> "" Then
                Print #intFreeFile, strLine
            End If
        Next r
    End With
    
    Close #intFreeFile
End Sub
Public Sub SaveCSVFSB(ByVal strFileName As String)
    Dim intFreeFile As Integer
    Dim strLine As String
    intFreeFile = FreeFile
    
    Open strFileName For Output As #intFreeFile
    
    Dim lCtr As Long
    Dim sTemp As String
    Dim sLine As String
    Dim amount As Double
    Dim sVar
    Dim oFSTR As Scripting.TextStream
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFSTR = fso.OpenTextFile(frmUploadReceipts.txtInputFile.text)
    
    lCtr = 1
    Do While Not oFSTR.AtEndOfStream
        sLine = oFSTR.ReadLine
        If lCtr = 1 Then
             sTemp = sTemp & sLine
        ElseIf lCtr > 1 And elegforwrite(lCtr, amount) Then
'             sVar = Split(sLine, ",")
'             If frmUploadReceipts.cmbBank.text = "TSB" Then
'                sVar(6) = amount
'             End If
'             sLine = Join(sVar, ",")
             sTemp = sTemp & sLine
        End If
        lCtr = lCtr + 1
        If Trim(sTemp) <> "" Then
            Print #intFreeFile, sTemp
        End If
        sTemp = ""
    Loop
   
    oFSTR.Close
    Close #intFreeFile
End Sub
Private Function elegforwrite(intLIneNO As Long, ByRef amount As Double) As Boolean
    Dim r As Integer
    With frmUploadReceipts.flxBankTransactions
            For r = 1 To .Rows - 1
                Debug.Print Val(.TextMatrix(r, 6))
                If Val(.TextMatrix(r, 6)) = intLIneNO And .TextMatrix(r, 10) = "No" Then 'column number 6 is holding the statemnt row number that has neeb not used . partially used are moved to account receipt.
                        elegforwrite = True
                        amount = .TextMatrix(r, 0)
                   Exit Function
                End If
            Next
    End With
End Function
'Public Sub SaveCSVFSB(ByVal strFileName As String, ByRef msFlex As MSHFlexGrid)
'    Const SEPARATOR_CHAR As String = ","
'
'    Dim intFreeFile As Integer
'    Dim strLine As String
'    Dim r As Integer
'    Dim c As Integer
'
'    intFreeFile = FreeFile
'
'    Open strFileName For Output As #intFreeFile
'
'    With msFlex
'        ' Every row
'        For r = 0 To .Rows - 1
'            strLine = ""
'            If r = 0 Then
'                'First column names
'               ' strLine = "F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13"
'                strLine = "Transaction Date,Transaction Description,Credit Amount,Transaction Type,Account Number"
'
'            Else
'                'Values
'                'Date,Memo,Amount, Subcategory, Number,account
'                '[F7],[F13],[F9],[F10],[F11],[F2]
'                'szHeader$ = "O/S Amt £|< Transaction Date|< Transaction Description|< Credit Amount|< Transaction Type|< Account Number|< ACCOUNT|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
'                 If .TextMatrix(r, 10) = "No" Then
'                     'strLine = ","
'
'                      'strLine = strLine & ","
'                     strLine = strLine & .TextMatrix(r, 1) 'Transaction Date|
'                      'strLine = strLine & ","
'                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 2) 'Transaction Description
'                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 0) 'Amount'[F9]=OS amount
'                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 4) 'Transaction Type
'                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 5) 'Account Number
'                     strLine = strLine & SEPARATOR_CHAR & .TextMatrix(r, 6) 'Account'[F2]
'                    ' strLine = strLine & ",,,"
'
'                     'strLine = strLine & ","
'
'
'
'                     strLine = strLine & ","
'
'                 End If
'            End If
'            If Trim(strLine) <> "" Then
'                Print #intFreeFile, strLine
'            End If
'        Next r
'    End With
'
'    Close #intFreeFile
'End Sub
Private Sub CreateBACS(adoConn As ADODB.Connection, szOutPutFileLoc As String, szEB As String, szRecID As String, szFileName As String, szFileEtn As String)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szLine As String, szaFormatLine() As String, szTemp As String
   Dim szFormatLine As String, szOutPutLine As String, szColHeading As String

   szTemp = szOutPutFileLoc & "\" & szFileName & Mid(szFileEtn, 2)
   Open szTemp For Output As #1

   FileFormat szFormatLine, szEB, szColHeading
   szaFormatLine = Split(szFormatLine, ", ")

   szSQL = "SELECT S.TenantName AS NAME, SUM(BT.PayAmt) AS AMT, S.BPR AS REF, S.SortCode AS SC, S.AcNo AS AC " & _
           "FROM (tblBatchReceipt AS BP INNER JOIN tblBtRptTran AS BT ON BP.BP = BT.BP) INNER JOIN " & _
               "Tenant AS S ON BT.TenantID = S.TenantID " & _
           "WHERE BP.BP = '" & szRecID & "' " & _
           "GROUP BY BT.TenantID, S.TenantName, S.BPR, S.SortCode, S.AcNo;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      If IsNull(adoRst.Fields.Item("SC").Value) Or IsNull(adoRst.Fields.Item("AC").Value) Then
         Close #1
         adoRst.Close
         Set adoRst = Nothing
         Exit Sub
      End If
   End If

   Print #1, szColHeading
   While Not adoRst.EOF
'  Write into the file.
      szOutPutLine = adoRst.Fields.Item(szaFormatLine(0)).Value & ", "
      szOutPutLine = szOutPutLine + adoRst.Fields.Item(szaFormatLine(1)).Value & ", "
      szOutPutLine = szOutPutLine + adoRst.Fields.Item(szaFormatLine(2)).Value & ", "
      szOutPutLine = szOutPutLine + CStr(adoRst.Fields.Item(szaFormatLine(3)).Value) & ", "
      szOutPutLine = szOutPutLine + adoRst.Fields.Item(szaFormatLine(4)).Value & ", 99"

      Print #1, szOutPutLine
      adoRst.MoveNext
   Wend

   Close #1
   adoRst.Close
   Set adoRst = Nothing

   MsgBox "The BACS file has been generated successfully.", vbInformation + vbOKOnly, "Batch Receipt"
End Sub

Private Sub FileFormat(ByRef szFormatLine As String, ByVal szEB As String, ByRef szColHeading As String)
   Dim szLine As String

   Open App.Path & "\BACS\bacs.txt" For Input As #2

   While Not EOF(2)
      Line Input #2, szLine

      If szLine = szEB Then
         Line Input #2, szLine
         Line Input #2, szLine
         Line Input #2, szColHeading

         szFormatLine = szLine
      Else
         Line Input #2, szLine
         Line Input #2, szLine
         Line Input #2, szColHeading
      End If
   Wend

   Close #2
End Sub

Private Function BACS_OPFLocation(adoConn As ADODB.Connection, ByRef szEB As String, ByRef szFileName As String, ByRef szFileEtn As String) As String
   On Error GoTo ERR_HANDLER

   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
'Below line modified by anol 13 Mar 2015
   szSQL = "SELECT FileLoc, EB, Indentifier, FileExten " & _
           "FROM tlbClientBanks " & _
           "WHERE MY_ID = " & frmBRPreForm.cmbBankAc.Column(2) & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If IsNull(adoRst.Fields.Item("EB").Value) Then
      MsgBox "Client's e-banking details are not updated.", vbCritical + vbOKOnly, "Batch Receipt"
      adoRst.Clone
      Set adoRst = Nothing
      Exit Function
   End If
   szEB = adoRst.Fields.Item("EB").Value
   szFileName = adoRst.Fields.Item("Indentifier").Value
   szFileEtn = adoRst.Fields.Item("FileExten").Value
   BACS_OPFLocation = adoRst.Fields.Item("FileLoc").Value

   adoRst.Close
   Set adoRst = Nothing
   Exit Function

ERR_HANDLER:

   Set adoRst = Nothing
End Function
'
'Private Sub ClearSuggestedPayment(adoConn As ADODB.Connection)
'   adoConn.Execute "UPDATE tblBatchReceipt SET Generated = TRUE;"
'End Sub

Private Sub AutoAlloc(adoConn As ADODB.Connection, szTenant As String)
   Dim i As Integer, j As String, cTotalPay As Currency

   SeperateDrCr szTenant

   cTotalPay = 0
   For i = 0 To UBound(iaPI_RowNo) - 1
      cTotalPay = cTotalPay + CCur(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))
   Next i

   j = 0
   For i = 0 To UBound(iaPI_RowNo) - 1
      If Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) = Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Simply allocate Cr -> PI
         Book1ReceiptOfSI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))

         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))

         flxSPayment.TextMatrix(iaPC_RowNo(j), 11) = Format(Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)), "0.00")
         j = j + 1
      ElseIf Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) > Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Simply allocate Cr -> PI
'        PI = PI - Cr
         Book1ReceiptOfSI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))

'        Update the credit Out-Standing amount by receipt amount
         flxSPayment.TextMatrix(iaPC_RowNo(j), 10) = CCur(flxSPayment.TextMatrix(iaPC_RowNo(j), 10)) - CCur(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))
         
         flxSPayment.TextMatrix(iaPI_RowNo(i), 11) = Format(Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)), "0.00")
         i = i - 1
         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))
         flxSPayment.TextMatrix(iaPC_RowNo(j), 11) = "0.00"
         j = j + 1
      ElseIf Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) < Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Cr_ = PI
'        Simply allocate Cr_ -> PI
'        Cr = Cr - Cr_
         Book1ReceiptOfSI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))
         flxSPayment.TextMatrix(iaPC_RowNo(j), 11) = Format(Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)), "0.00")

         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))

      End If
      If j > UBound(iaPC_RowNo) - 1 Then Exit For
   Next i
End Sub

Private Sub Book1ReceiptOfSI(adoConn As ADODB.Connection, iFlxRowPI As Integer, iFlxRowPC As Integer, cAmt As Currency)
   Dim iRow  As Integer, szSQL  As String
   Dim lRT_ID As Long
   Dim rstRst As New ADODB.Recordset

   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM RptTransactions;"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
   rstRst.Close

'  Update the invoice out-standing amount by allocated amount
   szSQL = "UPDATE tlbReceipt " & _
           "SET OSAmount = " & CCur(flxSPayment.TextMatrix(iFlxRowPI, 10)) - CCur(flxSPayment.TextMatrix(iFlxRowPI, 11)) & ", " & _
               "ReceiptView = IIF(OSAmount > 0, TRUE, FALSE) " & _
           "WHERE TransactionID = " & CLng(flxSPayment.TextMatrix(iFlxRowPI, 20)) & ";"
   adoConn.Execute szSQL

   szSQL = "UPDATE tlbReceipt " & _
           "SET OSAmount = " & CCur(flxSPayment.TextMatrix(iFlxRowPC, 10)) - cAmt & ", " & _
               "ReceiptView = IIF(" & CCur(flxSPayment.TextMatrix(iFlxRowPC, 10)) - cAmt & " > 0, TRUE, FALSE) " & _
           "WHERE TransactionID = " & CLng(flxSPayment.TextMatrix(iFlxRowPC, 20)) & ";"
   adoConn.Execute szSQL

'  Update the Invoice out standing amount by receipt amount
   szSQL = "SELECT * FROM RptTransactions;"
   rstRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

'         ~CREDIT NOTE~
   With rstRst
      .AddNew
      !TranType = "AL"
      !TransactionID = lRT_ID
      !Alloc_Unalloc = 1
      !FromTran = CLng(flxSPayment.TextMatrix(iFlxRowPI, 20))
      !ToTran = CLng(flxSPayment.TextMatrix(iFlxRowPC, 20))
      !AllocDate = Format(lblDate.Caption, "dd mmmm yyyy")
      !receiptAmount = cAmt
      !Discount = 0
      !UpdateSage = False
      'Modified by 13 mar 2015 anol
      !BankCode = frmBRPreForm.cmbBankAc.Column(0) 'Nominal code
      !nominalCode = rstRst!BankCode
      .Update
   End With

   rstRst.Close
   Set rstRst = Nothing
End Sub

Private Sub ReceiptTotalTenantWise()
   Dim cTotalRpt As Currency, j As Integer, i As Integer, bFlag As Boolean

   If frmBRPreForm.chkMultiple.Value = 1 Then
      ReDim szaTntID(4, 1) As String

      For j = 1 To flxSPayment.Rows - 1            'Pick the first receipt information
         If CCur(flxSPayment.TextMatrix(j, 11)) > 0 Then
            szaTntID(0, 0) = flxSPayment.TextMatrix(j, 21)
            szaTntID(2, 0) = flxSPayment.TextMatrix(j, 22)
            szaTntID(3, 0) = flxSPayment.TextMatrix(j, 24)
            szaTntID(4, 0) = flxSPayment.TextMatrix(j, 25)
            Exit For
         End If
      Next j

      For j = j To flxSPayment.Rows - 1            'Main Loop through out the grid
         bFlag = False
         If CCur(flxSPayment.TextMatrix(j, 11)) > 0 Then
            For i = 0 To UBound(szaTntID, 2) - 1

'        Resolved by BOSL. Modified by Asif.
'        Issue 0000550. Date: 23 Mar 2015
'        The last condition was wrong as it tried to compare the
'        Posting Date with the Reference, hence was not able to save the amout and therefore
'        throwing an exception while saving the receipts.
                
'               If szaTntID(0, i) = flxSPayment.TextMatrix(j, 21) And _
'                     szaTntID(2, i) = flxSPayment.TextMatrix(j, 22) And _
'                     szaTntID(3, i) = flxSPayment.TextMatrix(j, 24) And _
'                     szaTntID(3, i) = flxSPayment.TextMatrix(j, 25) Then
               If szaTntID(0, i) = flxSPayment.TextMatrix(j, 21) And _
                     szaTntID(2, i) = flxSPayment.TextMatrix(j, 22) And _
                     szaTntID(3, i) = flxSPayment.TextMatrix(j, 24) And _
                     szaTntID(4, i) = flxSPayment.TextMatrix(j, 25) Then
                  
                  szaTntID(1, i) = Val(szaTntID(1, i)) + CCur(flxSPayment.TextMatrix(j, 11))
                  bFlag = True
               End If
            Next i

            If Not bFlag Then
               ReDim Preserve szaTntID(4, UBound(szaTntID, 2) + 1)

               szaTntID(0, i) = flxSPayment.TextMatrix(j, 21)
               szaTntID(2, i) = flxSPayment.TextMatrix(j, 22)
               szaTntID(3, i) = flxSPayment.TextMatrix(j, 24)
               szaTntID(4, i) = flxSPayment.TextMatrix(j, 25)
               szaTntID(1, i) = Val(szaTntID(1, i)) + CCur(flxSPayment.TextMatrix(j, 11))
            End If
         End If
      Next j
      Exit Sub
   End If

   ReDim szaTntID(4, 1) As String

   For j = 1 To flxSPayment.Rows - 1            'Pick the first receipt information
      If CCur(flxSPayment.TextMatrix(j, 11)) > 0 Then
         szaTntID(0, 0) = flxSPayment.TextMatrix(j, 21)
         Exit For
      End If
   Next j
   For j = j To flxSPayment.Rows - 1            'Main Loop through out the grid
      bFlag = False
      If CCur(flxSPayment.TextMatrix(j, 11)) > 0 Then
         For i = 0 To UBound(szaTntID, 2) - 1
            If szaTntID(0, i) = flxSPayment.TextMatrix(j, 21) Then
               szaTntID(1, i) = Val(szaTntID(1, i)) + CCur(flxSPayment.TextMatrix(j, 11))
               bFlag = True
            End If
         Next i

         If Not bFlag Then
            ReDim Preserve szaTntID(4, UBound(szaTntID, 2) + 1)

            szaTntID(0, i) = flxSPayment.TextMatrix(j, 21)
            szaTntID(1, i) = Val(szaTntID(1, i)) + CCur(flxSPayment.TextMatrix(j, 11))
         End If
      End If
   Next j
End Sub

Private Sub BookNormalReceiptOfSI(adoConn As ADODB.Connection, iTenant As Integer)
   Dim cRptAmt       As Currency
   Dim lSlNumber     As Long
   Dim lSp_ID        As Long
   Dim lSPTran_ID    As Long
   Dim iRow          As Integer
   Dim i             As Integer
   Dim j             As Integer
   Dim szRef         As String
   Dim szSQL         As String
   Dim szaNewSplit() As String
   Dim aSupplier()   As String
   Dim aSup(2)       As String
   Dim bFlag         As Boolean
   Dim rstSet        As New ADODB.Recordset
   Dim rstSplit      As New ADODB.Recordset
   Dim rsInvoice   As New ADODB.Recordset
   Dim cSumSplits As Double
   Dim rsAddAllocSplit As New ADODB.Recordset
   Dim rsRptTransactionsSplit As New ADODB.Recordset
   Dim lSPTran_ID_split As Long
   On Error GoTo ErrHandler
   
   ReDim aTenant(flxSPayment.Rows - 1, 2) As String
'  Generate the next Transaction Id from the tlbReceipt

   szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbReceipt;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSp_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close

   szSQL = "SELECT MAX(TransactionID) AS TID FROM RptTransactions;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
   
   szSQL = "SELECT MAX(TransactionID) AS TID FROM RptTransactionsSplit;"
   rsRptTransactionsSplit.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID_split = CLng(IIf(IsNull(rsRptTransactionsSplit!TID), 1, rsRptTransactionsSplit!TID))
   rsRptTransactionsSplit.Close
    
   frmMMain.Leasee1_LesseList_isUptoDate = False
   frmMMain.Leasee4_LesseList_isUptoDate = False
   frmMMain.frmDemand3_LesseList_isUptoDate = False
 
'****************************************** RECEIPT HEADER ***********************************
'   ===============       ADD NEW RECEIPT      ========================
   szSQL = "SELECT * FROM tlbReceipt;"
   rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   lSlNumber = SlNumber("SR", "tlbReceipt", adoConn)
      
   rstSet.AddNew
   lSp_ID = lSp_ID + 1
   rstSet!TransactionID = lSp_ID
   rstSet!CreatedBy = User
   rstSet!CreatedDate = Now
   rstSet!szTransactionID = rstSet!TransactionID
   rstSet!Type = 3 'CByte(sdoPP)
   rstSet!SageAccountNumber = szaTntID(0, iTenant) 'flxSPayment.TextMatrix(iRow, 21)

   'Resolved by BOSL. Modified by Asif. Date: 05-12-2014
   'Issue No: 0000512
      Dim rstUnit As New ADODB.Recordset
      rstUnit.Open "SELECT UnitNumber FROM LeaseDetails WHERE SageAccountNumber = '" & rstSet!SageAccountNumber & "' and status=true ", adoConn, adOpenDynamic, adLockOptimistic
      
      If Not rstUnit.EOF Then
         rstSet!unitid = rstUnit.Fields.Item("UnitNumber").Value
         rstUnit.Close
         Set rstUnit = Nothing
      End If
   'End
   
   If frmBRPreForm.chkMultiple.Value = 0 Then
      rstSet!dDate = Format(frmBRPreForm.txtDate.text, "dd mmmm yyyy")
      rstSet!RDate = Format(lblDate.Caption, "dd mmmm yyyy")
      rstSet!postingDate = Format(frmBRPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
   Else
      rstSet!RDate = Format(szaTntID(2, iTenant), "dd mmmm yyyy")
      rstSet!postingDate = Format(szaTntID(4, iTenant), "dd mmmm yyyy")
   End If

   If frmBRPreForm.chkMultiple.Value Then
      rstSet!ref = szaTntID(3, iTenant)
   Else
      rstSet!ref = txtReference_.text
   End If

   rstSet!Details = "BATCH RECEIPT"
   rstSet!amount = CCur(szaTntID(1, iTenant))
   rstSet!IsSageUpdate = False
   rstSet!UpdateSage = False
   rstSet!ReceiptView = False
   rstSet!ExtRef = txtChqNo.text
   If optBR_Bank.Value Then
      rstSet!RptAmtType = "Bank"
   Else
      rstSet!RptAmtType = "CHQ"
   End If
   'Modified by anol 13 Mar 2015
   rstSet!BankCode = frmBRPreForm.cmbBankAc.Column(0) 'Nominal COde
   rstSet!nominalCode = rstSet!BankCode
   'issue 973 clientID was not writing 2021-07-27
   rstSet!ClientID = frmBRPreForm.txtClient.Tag
   rstSet!SlNumber = lSlNumber
   rstSet!LastModifiedBy = User
   rstSet!LastModifiedDate = Now
   rstSet.Update
   rstSet.Close

'*******************************************  RECEIPT SPLITS  ******************************************
   For iRow = 1 To flxSPayment.Rows - 1
      If frmBRPreForm.chkMultiple.Value = 0 Then
         bFlag = True
      Else
         bFlag = False
         szSQL = flxSPayment.TextMatrix(iRow, 22)
         szRef = flxSPayment.TextMatrix(iRow, 24)
      End If
      '" flxSPayment.RowHeight(iRow) > 0 And " Has been remmed issue 399
      If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 And _
            (flxSPayment.TextMatrix(iRow, 15) = 1 Or flxSPayment.TextMatrix(iRow, 15) = 23) And _
            flxSPayment.TextMatrix(iRow, 21) = szaTntID(0, iTenant) And _
            IIf(bFlag, True, (szSQL = szaTntID(2, iTenant) And _
                              szRef = szaTntID(3, iTenant))) Then
        'grid col 15 is type
        'grid 21 is sageaccount number
        '     Saving the split(s) of the header. Each split will keep track of the SI's split ID in @AllocTranID
        '  First Update SI's split outstanding amount
                 szSQL = "SELECT * FROM tlbReceiptSplit " & _
                         "WHERE RptHeader = " & CLng(flxSPayment.TextMatrix(iRow, 20)) & " AND OSAmount > 0 " & _
                         "ORDER BY SplitID;"
        'Debug.Print szSQL
                 rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
        
                 cRptAmt = Round(CCur(flxSPayment.TextMatrix(iRow, 11)), 2)
                 ReDim szaNewSplit(RecordCount(rstSet), 4) As String      'Amt, SplitID, TransID,propertyID
                 i = 0
                 While cRptAmt > 0
                    With rstSet
                       If cRptAmt >= Round(CCur(.Fields.Item("OSAmount").Value), 2) Then
                          cRptAmt = cRptAmt - Round(CCur(.Fields.Item("OSAmount").Value), 2)
                          szaNewSplit(i, 0) = Round(CCur(.Fields.Item("OSAmount").Value), 2)
                          .Fields.Item("OSAmount").Value = 0
                       Else
                          .Fields.Item("OSAmount").Value = Round(CCur(.Fields.Item("OSAmount").Value), 2) - cRptAmt
                          szaNewSplit(i, 0) = cRptAmt
                          cRptAmt = 0
                       End If
                       szaNewSplit(i, 1) = .Fields.Item("SplitID").Value
                       szaNewSplit(i, 2) = .Fields.Item("TransactionID").Value
                       szaNewSplit(i, 3) = .Fields.Item("FundID").Value
                       If flxSPayment.ColWidth(22) = 0 Then
                           szaNewSplit(i, 4) = flxSPayment.TextMatrix(iRow, 22)
                        Else
                           szaNewSplit(i, 4) = flxSPayment.TextMatrix(iRow, 23)
                        End If
                       .Update
                       .MoveNext
                       i = i + 1
                    End With
                 Wend '
                 rstSet.Close

'  Now create new splits and show the allocation mapping
                 szSQL = "SELECT * FROM tlbReceiptSplit"
                 rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
        
                 szSQL = "SELECT * FROM RptTransactionsSplit"
                 rsAddAllocSplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                 
                 i = i - 1
                 While i >= 0
                     lSPTran_ID_split = lSPTran_ID_split + 1
                    With rstSet
                       .AddNew
                       .Fields.Item("TransactionID").Value = UniqueID()
                       .Fields.Item("RptHeader").Value = lSp_ID
                       .Fields.Item("FundID").Value = szaNewSplit(i, 3)
                       .Fields.Item("Amount").Value = szaNewSplit(i, 0)
                       .Fields.Item("SplitID").Value = szaNewSplit(i, 1)
                       .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
                       .Fields.Item("Description").Value = flxSPayment.TextMatrix(iRow, 7)
                       .Fields.Item("AllocTranID").Value = szaNewSplit(i, 2)
                       .Fields.Item("PropertyID").Value = szaNewSplit(i, 4)
                       .Fields.Item("RptTransactionsIDSplit").Value = lSPTran_ID_split
                       .Update
                    End With
                    'add record in split allocation table
                             rsAddAllocSplit.AddNew
                             rsAddAllocSplit!TransactionID = lSPTran_ID_split
                             rsAddAllocSplit!FromTran = lSp_ID
                             rsAddAllocSplit!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
                             rsAddAllocSplit!AllocDate = Format(Date, "DD MMMM YYYY")
                             rsAddAllocSplit!receiptAmount = szaNewSplit(i, 0)
                             rsAddAllocSplit!BankCode = frmBRPreForm.cmbBankAc.Column(0) 'Nominal Code
                             rsAddAllocSplit!nominalCode = rsAddAllocSplit!BankCode
                             rsAddAllocSplit!fundID = szaNewSplit(i, 3)
'                             If Left(flxSPayment.TextMatrix(iRow, 1), 3) = "SRR" Then
'                                     rsAddAllocSplit!VATAMOUNT = 0
'                                     rsAddAllocSplit!NetAmount = rsAddAllocSplit!receiptAmount
'                             Else
                            rsAddAllocSplit!VATAMOUNT = CalculateVatAmountFormdmsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 13)), szaNewSplit(i, 1), rstSet.Fields.Item("Amount").Value)
                            rsAddAllocSplit!NetAmount = CalculateNetAmountFormdmsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 13)), szaNewSplit(i, 1), rstSet.Fields.Item("Amount").Value)
'                             End If
                             rsAddAllocSplit!VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
                             rsAddAllocSplit!SplitIDofSI = szaNewSplit(i, 1)
                             rsAddAllocSplit!deleteFlag = False
                             rsAddAllocSplit.Update
                    'end of adding
                    i = i - 1
                 Wend
                 rstSet.Close
                 rsAddAllocSplit.Close

'**************************************************************************************************************

                 UpdateBankAcBal_Plus adoConn, CCur(flxSPayment.TextMatrix(iRow, 11)), frmBRPreForm.cmbBankAc.Column(0), frmBRPreForm.txtClient.Tag
        
        '        =============     UPDATE OS BALANCE OF SI     =================
                 szSQL = "SELECT * " & _
                         "FROM tlbReceipt " & _
                         "WHERE TransactionID = " & flxSPayment.TextMatrix(iRow, 20) & ";"
                 rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
        
                 rstSet!OSAmount = CCur(flxSPayment.TextMatrix(iRow, 10)) - _
                                    CCur(IIf(flxSPayment.TextMatrix(iRow, 11) = "", 0, flxSPayment.TextMatrix(iRow, 11))) ''O/S Amount-receipt amount
                 rstSet!ReceiptView = IIf(rstSet!OSAmount > 0, True, False)
                 If rstSet!ReceiptView = False Then
                   rstSet!DateTimeStamp = ""
                   rstSet!Module = ""
                   rstSet!UserSessionID = ""
                   rstSet!WindowsUserName = ""
                   rstSet!MachineName = ""
                   rstSet!PrestigeUserName = ""
                   rstSet!ServerIPaddress = ""
                End If
                 rstSet.Update
                 rstSet.Close
        
       'here
       '        ============      RECORD THE ALLOCATION      ==================
                 szSQL = "SELECT * " & _
                         "FROM RptTransactions;"
                 rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                 rstSet.AddNew
                 rstSet!TranType = "AL"
                 lSPTran_ID = lSPTran_ID + 1
                 rstSet!TransactionID = lSPTran_ID
                 rstSet!Alloc_Unalloc = 1
                 rstSet!FromTran = lSp_ID
                 rstSet!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
                 rstSet!AllocDate = Format(Date, "DD MMMM YYYY")
                 rstSet!receiptAmount = CCur(flxSPayment.TextMatrix(iRow, 11))
                 'Modified by anol 15 Mar 2015
                 rstSet!BankCode = frmBRPreForm.cmbBankAc.Column(0) 'Nominal Code
                 rstSet!nominalCode = rstSet!BankCode
                 rstSet!SlNumber = lSlNumber
                 rstSet!deleteFlag = False
                 rstSet.Update
                 rstSet.Close
      End If
   Next iRow
   Set rstSet = Nothing
   Exit Sub

ErrHandler:
'   Debug.Print ERR.Number & " " & ERR.description
   MsgBox "System could not run the batch payment. Please contact with PCM."
End Sub
Private Function CalculateVatAmountFormdmsplit(adoConn As ADODB.Connection, szDemandID As String, ByVal szSpitID As Integer, dblSIVATAmount As Double) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select AMOUNT+VATAMOUNT as amt,VATAMOUNT from DemandSplitRecords where DemandID=" & szDemandID & " and SplitID=" & szSpitID & " ", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
            CalculateVatAmountFormdmsplit = dblSIVATAmount / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
             CalculateVatAmountFormdmsplit = CalculateVatAmountFormdmsplit * IIf(IsNull(rsDemandSplitRecords("VATAMOUNT").Value), 0, rsDemandSplitRecords("VATAMOUNT").Value)
    End If
    rsDemandSplitRecords.Close
    
End Function
Private Function CalculateNetAmountFormdmsplit(adoConn As ADODB.Connection, szDemandID As String, ByVal szSpitID As Integer, dblSINetAmount) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select AMOUNT+VATAMOUNT as amt,AMOUNT  from DemandSplitRecords where DemandID=" & szDemandID & "  and SplitID=" & szSpitID & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
             CalculateNetAmountFormdmsplit = dblSINetAmount / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
             CalculateNetAmountFormdmsplit = CalculateNetAmountFormdmsplit * IIf(IsNull(rsDemandSplitRecords("AMOUNT").Value), 0, rsDemandSplitRecords("AMOUNT").Value)
    End If
    rsDemandSplitRecords.Close

End Function

'Private Function CalculateVatAmountFormdmsplit(adoconn As ADODB.Connection, szDemandID As String, szSpitID As Integer, dblReceiptAmount As Double) As Double
'    Dim rsDemandSplitRecords As New ADODB.Recordset
'    rsDemandSplitRecords.Open "Select S.AMOUNT+VATAMOUNT as amt,VATAMOUNT from DemandSplitRecords S,tlbReceipt R,tlbReceiptSplit RS where " & _
'                    "R.DemandRef=S.DemandID  AND R.TransactionID=RS.RptHeader AND R.TransactionID  =" & szDemandID & "  AND RS.SplitID  =" & szSpitID & "  AND Rs.SPLITID=S.SplitID ", adoconn, adOpenStatic, adLockReadOnly
'    If Not rsDemandSplitRecords.EOF Then
'             CalculateVatAmountFormdmsplit = dblReceiptAmount / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
'             CalculateVatAmountFormdmsplit = dblReceiptAmount * IIf(IsNull(rsDemandSplitRecords("VATAMOUNT").Value), 0, rsDemandSplitRecords("VATAMOUNT").Value)
'    End If
'    rsDemandSplitRecords.Close
'
'End Function
'Private Function CalculateNetAmountFormdmsplit(adoconn As ADODB.Connection, szDemandID As String, szSpitID As Integer, dblReceiptAmount) As Double
'    Dim rsDemandSplitRecords As New ADODB.Recordset
'    rsDemandSplitRecords.Open "Select S.AMOUNT+VATAMOUNT as amt,S.AMOUNT from DemandSplitRecords S,tlbReceipt R,tlbReceiptSplit RS where " & _
'                    "R.DemandRef=S.DemandID  AND R.TransactionID=RS.RptHeader AND R.TransactionID  =" & szDemandID & "  AND RS.SplitID  =" & szSpitID & "  AND Rs.SPLITID=S.SplitID", adoconn, adOpenStatic, adLockReadOnly
'
'    If Not rsDemandSplitRecords.EOF Then
'             CalculateNetAmountFormdmsplit = dblReceiptAmount / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
'             CalculateNetAmountFormdmsplit = CalculateNetAmountFormdmsplit * IIf(IsNull(rsDemandSplitRecords("AMOUNT").Value), 0, rsDemandSplitRecords("AMOUNT").Value)
'    End If
'    rsDemandSplitRecords.Close
'
'End Function

Private Sub ConfigFlxSPayment()
   Dim szHeader As String

   flxSPayment.Clear
   'flxSPayment.Cols = 27
   flxSPayment.Cols = 32
   If frmBRPreForm.chkMultiple.Value = 1 Then
      Label1(11).Visible = True
      Label1(12).Visible = True
      Label20(18).Width = 14415
      Label19(3).Visible = False
      Label19(5).Visible = False

      txtReference_.Visible = False

      
      szHeader$ = "|<No.|<Tenant|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
                  "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
                  "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo" & _
                  "|<RptDate|PropID|<Reference|<Posting Date"
   Else
      Me.Width = 13245
      Label1(11).Visible = False
      Label1(12).Visible = False
      Label20(18).Width = 11875
      Label19(3).Visible = True
      Label19(5).Visible = True
      
      'Resolved By BOSL. Issue: 0000550
      'Added by Asif. Date: 23 Mar 2015. 1.5 The reference field should be hidden from the batch receipt without multiple screen.
'      txtReference_.Visible = True
      
      Label1(8).Left = Label1(8).Left + 400
      Label1(9).Left = Label1(9).Left + 400
      Label1(10).Left = Label1(10).Left + 400

      'flxSPayment.Cols = 23
      szHeader$ = "|<No.|<Tenant|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
                  "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
                  "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo.|PropID"
   End If

   Me.Width = Label20(18).Width + 320

   flxSPayment.Width = Label20(18).Width
   Shape1.Width = Label20(18).Width
   flxSPayment.Rows = 2
   flxSPayment.RowHeight(0) = 0

   flxSPayment.FormatString = szHeader$

   flxSPayment.ColAlignment(0) = vbCenter
   flxSPayment.ColWidth(0) = Label1(1).Left - flxSPayment.Left    'Sign
   flxSPayment.ColWidth(1) = Label1(2).Left - Label1(1).Left      'No
   flxSPayment.ColWidth(2) = Label1(3).Left - Label1(2).Left      'Tenant
   flxSPayment.ColWidth(3) = Label1(4).Left - Label1(3).Left      'Type
   flxSPayment.ColWidth(4) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
   flxSPayment.ColWidth(5) = Label1(5).Left - Label1(4).Left      'Unit ID
   flxSPayment.ColWidth(6) = Label1(6).Left - Label1(5).Left      'Date
   flxSPayment.ColWidth(7) = Label1(7).Left - Label1(6).Left      'Ref
   flxSPayment.ColWidth(8) = Label1(8).Left - Label1(7).Left      'Details
   flxSPayment.ColWidth(9) = Label1(9).Left - Label1(8).Left      'Amount
   flxSPayment.ColWidth(10) = Label1(10).Left - Label1(9).Left    'O/S Amount
   If frmBRPreForm.chkMultiple.Value = 1 Then
      flxSPayment.ColWidth(11) = Label1(11).Left - Label1(10).Left 'Receipt Amt
   Else
      flxSPayment.ColWidth(11) = flxSPayment.Width + flxSPayment.Left - Label1(10).Left - 300 'Receipt
   End If

   flxSPayment.ColWidth(12) = 0     'Discount
   flxSPayment.ColWidth(13) = 0     'DemandID
   flxSPayment.ColWidth(14) = 0     'SAGE O/S £
   flxSPayment.ColWidth(15) = 0     'Transaction Type - linked with column 1 Type
   flxSPayment.ColWidth(16) = 0     'R/A; R -> receipt, A -> allocation
   flxSPayment.ColWidth(17) = 0     'allocation ref
   flxSPayment.ColWidth(18) = 0     'allocation amount
   flxSPayment.ColWidth(19) = 0     'Sage Department
   flxSPayment.ColWidth(20) = 0     'TransactionID
   flxSPayment.ColWidth(21) = 0     'Tenant ID
   If frmBRPreForm.chkMultiple.Value = 1 Then
      flxSPayment.ColWidth(22) = Label1(12).Left - Label1(11).Left 'Receipt Date
      flxSPayment.ColWidth(23) = 0
      flxSPayment.ColWidth(24) = Label1(13).Left - Label1(12).Left 'Receipt Refence
      flxSPayment.ColAlignment(24) = vbLeftJustify
      flxSPayment.ColWidth(25) = flxSPayment.Width + flxSPayment.Left - Label1(13).Left - 300 'Posting Date
   Else
      flxSPayment.ColWidth(22) = 0
      flxSPayment.ColWidth(23) = 0
      flxSPayment.ColWidth(24) = 0
      flxSPayment.ColWidth(25) = 0
   End If
   flxSPayment.ColWidth(26) = 0     'Mark checksumof allocation
   'adding five more column to the grid for locking issue
    flxSPayment.ColWidth(27) = 0     'userSessionID
    flxSPayment.ColWidth(28) = 0     'WindowsuserName
    flxSPayment.ColWidth(29) = 0     'Machine Name
    flxSPayment.ColWidth(30) = 0     'Module
    flxSPayment.ColWidth(31) = 0     'ClientID
        
   txtGrossTotal.Left = Label1(10).Left
   txtAlloctotal.Left = txtGrossTotal.Left
   txtGrandTotal.Left = txtGrossTotal.Left
   txtGrossTotal.Width = flxSPayment.ColWidth(11)
   txtAlloctotal.Width = flxSPayment.ColWidth(11)
   txtGrandTotal.Width = flxSPayment.ColWidth(11)
   cmdSPClose.Left = flxSPayment.Left + flxSPayment.Width - cmdSPClose.Width
   txtGrossTotal.text = "0.00"
   lbl11.Left = txtGrossTotal.Left - lbl11.Width - 100
   lblbEE.Left = txtGrossTotal.Left - lblbEE.Width - 100
   lblbb.Left = txtGrossTotal.Left - lblbb.Width - 100
'   Frame5(5).Left = txtGrossTotal.Left + txtGrossTotal.Width + 100
'   Frame5(5).Top = txtAlloctotal.Top
   grpUploadReceipts(0).Left = txtGrossTotal.Left + txtGrossTotal.Width + 100
   grpUploadReceipts(0).Top = txtAlloctotal.Top
   'If frmBRPreForm.chkMultiple.Value = 0 Then
        Frame5(5).Left = 4410
        Frame5(5).Top = Frame2.Top
   'End If
End Sub

Public Sub FilterGrid()
   Dim iRow As Integer

   If lblProperty.ToolTipText <> "ALL" Then
      For iRow = 1 To flxSPayment.Rows - 1
         If flxSPayment.ColWidth(22) = 0 Then
            If flxSPayment.TextMatrix(iRow, 22) <> lblProperty.ToolTipText Then
               flxSPayment.RowHeight(iRow) = 0
            End If
         Else
            If flxSPayment.TextMatrix(iRow, 23) <> lblProperty.ToolTipText Then
               flxSPayment.RowHeight(iRow) = 0
            End If
         End If
      Next iRow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
   
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
   adoConn.Close
   Set adoConn = Nothing
   UserSessionID = ""
   frmLockingDialogisActive = False
   
   Call WheelUnHook(Me.hWnd)
   bBRPreForm = False
   Unload frmBRPreForm
   Unload frmUploadReceipts
   Exit Sub
Err:
End Sub

Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo Err
   If flxSPayment.row < flxSPayment.Rows And frmBRPreForm.chkMultiple.Value = 1 And Not txtRptDt.Visible And flxSPayment.col = 11 Then
      txtSPayment.Visible = False

      flxSPayment.col = 22
      iFlxSPayCol = 22
      txtRptDt.Top = flxSPayment.CellTop + flxSPayment.Top
      txtRptDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      txtRptDt.Width = flxSPayment.ColWidth(22)
      txtRptDt.Height = flxSPayment.RowHeight(iCurRow) - 15
      txtRptDt.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      FocusControl txtRptDt
      SumUpTotal
   End If
'  Move to the next row to enter amount
   If flxSPayment.row < flxSPayment.Rows And Not txtRptDt.Visible And (flxSPayment.col = 24 Or flxSPayment.col = 11) Then
      If MoveDownPosition Then
         flxSPayment.col = 11
         FocusControl flxSPayment
      Else
         'cmdSavePayment.SetFocus
         If Val(cmdAll.Tag) = 0 Then
             FocusControl cmdGPayment
        Else
            FocusControl cmdAll
        End If
      End If
   End If
'Debug.Print Index

Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub txtDmdTenantSearchUnitName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl flxDmdLeaseList
    End If
End Sub

Private Sub txtGrossTotal_Change()
    txtGrandTotal.text = Format(Val(txtAlloctotal.text) + Val(txtGrossTotal.text), "0.00")
End Sub

Private Sub txtPaymentDateGrid_Change()
    TextBoxChangeDate txtPaymentDateGrid
End Sub

Private Sub txtPaymentDateGrid_GotFocus()
    flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub txtPaymentDateGrid_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      FocusControl flxAllocation
      txtPaymentDateGrid.text = ""
      txtPaymentDateGrid.Visible = False
   End If
    If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
        If flxAllocation.col = 6 And flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
             Dim dblAmount As Double
              Dim i As Integer
              If Trim(txtPaymentDateGrid.text) = "" Then
                MsgBox "Please enter a Receipt Date", vbInformation, "Receipt Date"
                FocusControl txtPaymentDateGrid
                Exit Sub
            End If
             For i = 1 To flxAllocation.Rows - 1
                  If flxAllocation.TextMatrix(i, 9) <> "D" Then
                        dblAmount = dblAmount + Val(flxAllocation.TextMatrix(i, 5))
                 End If
              Next i
               txtAlloctotal.text = Format(dblAmount, "0.00")
               If IsDate(txtPaymentDateGrid.text) = False Then
                    MsgBox "Receipt Date is not in the correct format", vbInformation, "Warning"
                    'flxAllocation.CellForeColor = vbRed
                    txtPaymentDateGrid.text = ""
                    FocusControl txtPaymentDateGrid
                    Exit Sub
                    
                Else
                    flxAllocation.CellForeColor = vbBlack
                End If
                flxAllocation.TextMatrix(flxAllocation.row, 6) = Format(txtPaymentDateGrid.text, "dd/mm/yyyy")
                flxAllocation.TextMatrix(flxAllocation.row, 8) = Format(txtPaymentDateGrid.text, "dd/mm/yyyy")
                
                  
            flxAllocation.col = 7
            'flxAllocation_DblClick
            txtPaymentDateGrid.Visible = False
            'flxAllocation.SetFocus
            flxAllocation_DblClick
        End If
    End If
End Sub

Private Sub txtPaymentDateGrid_LostFocus()
    flxAllocation.ScrollBars = flexScrollBarVertical
    If IsDate(txtPaymentDateGrid.text) = False Then
        If txtPaymentDateGrid.text <> "" Then
            MsgBox "Receipt Date is not in correct format", vbInformation, "Warning"
        End If
        Exit Sub
    End If
    txtPaymentDateGrid_KeyPress (13)
End Sub

Private Sub txtPostingDate_Change()
     TextBoxChangeDate txtPostingDate
End Sub
Private Sub txtPostingDate_GotFocus()
     txtRptDt.Visible = False
     txtRefInput.Visible = False
     txtSPayment.Visible = False
    If iFlxSPayCol = 25 Then
      If flxSPayment.TextMatrix(iCurRow, 25) = "" Then
            txtPostingDate.text = Format(Now, "dd/mm/yyyy")
      Else
            txtPostingDate.text = Format(flxSPayment.TextMatrix(iCurRow, 25), "dd/mm/yyyy")
      End If
   End If
End Sub

Private Sub txtPostingDate_KeyPress(KeyAscii As Integer)
On Error GoTo Err
    If KeyAscii = 27 Then
      FocusControl flxSPayment
      txtPostingDate.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

     If KeyAscii = 13 Then
            '13 Apr 2015 issue 530 Modified by anol
            flxSPayment.TextMatrix(iCurRow, 25) = Format(txtPostingDate.text, "dd/mm/yyyy")
             If IsDate(txtPostingDate.text) = False Then
                   ShowMsgInTaskBar "Date format is not correct!"
                    Exit Sub 'added by anol 20160911
              End If
            If MoveDownPosition Then
                 flxSPayment.col = 11
                 FocusControl flxSPayment
            Else
                If Val(cmdAll.Tag) = 0 Then
                    FocusControl cmdGPayment
                Else
                    FocusControl cmdAll
                End If
            End If
        txtPostingDate.Visible = False
        flxSPayment.ScrollBars = flexScrollBarVertical
   End If

   TextBoxKeyPrsDate txtPostingDate, KeyAscii
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub txtPostingDate_LostFocus()
    If txtPostingDate.text <> "" Then TextBoxFormatDate txtPostingDate
    flxSPayment.TextMatrix(iCurRow, 25) = txtPostingDate.text
    txtPostingDate.Visible = False
    flxSPayment.ScrollBars = flexScrollBarVertical
End Sub

Private Sub txtPostingDateGrid_Change()
    TextBoxChangeDate txtPostingDateGrid
    
End Sub

Private Sub txtPostingDateGrid_GotFocus()
    flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub txtPostingDateGrid_KeyPress(KeyAscii As Integer)
    'TextBoxKeyPrsDate txtPostingDateGrid, KeyAscii
   If KeyAscii = 27 Then
      FocusControl flxAllocation
      txtPostingDateGrid.text = ""
      txtPostingDateGrid.Visible = False
   End If
    If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
            If flxAllocation.row = flxAllocation.Rows - 1 Then
                    flxAllocation.AddItem ""
             End If
             Dim dblAmount As Double
             Dim i As Integer
             For i = 1 To flxAllocation.Rows - 1
                  If flxAllocation.TextMatrix(i, 9) <> "D" Then
                        dblAmount = dblAmount + Val(flxAllocation.TextMatrix(i, 5))
                 End If
              Next i
            
                txtAlloctotal.text = Format(dblAmount, "0.00")
            If flxAllocation.col = 8 And flxAllocation.TextMatrix(flxAllocation.row + 1, 0) <> "" Then
            If Trim(txtPostingDateGrid.text) = "" Then
                MsgBox "Please enter posting date", vbInformation, "posting date"
                FocusControl txtPostingDateGrid
                Exit Sub
            End If
                flxAllocation.TextMatrix(flxAllocation.row, 8) = Format(txtPostingDateGrid.text, "dd/mm/yyyy")
                If IsDate(txtPostingDateGrid.text) = False Then
                    MsgBox "Posting  Date is not in correct format", vbInformation, "Warning"
                    flxAllocation.CellForeColor = vbRed
                Else
                    flxAllocation.CellForeColor = vbBlack
                End If
                

                flxAllocation.row = flxAllocation.row + 1
                flxAllocation.col = 2
                'flxAllocation_DblClick
                txtPostingDateGrid.Visible = False
                FocusControl flxAllocation
                Exit Sub
            End If
            If flxAllocation.col = 8 And flxAllocation.TextMatrix(flxAllocation.row + 1, 0) = "" Then
                flxAllocation.TextMatrix(flxAllocation.row, 8) = Format(txtPostingDateGrid.text, "dd/mm/yyyy")
                If IsDate(txtPostingDateGrid.text) = False Then
                    MsgBox "Posting Date is not in the correct format", vbInformation, "Warning"
                    'flxAllocation.CellForeColor = vbRed
                    txtPostingDateGrid.text = ""
                    FocusControl txtPostingDateGrid
                    Exit Sub
                Else
                    flxAllocation.CellForeColor = vbBlack
                End If
                flxAllocation.col = 2
                flxAllocation.row = flxAllocation.row + 1
                flxAllocation.TextMatrix(flxAllocation.row, 9) = "M"
                txtPostingDateGrid.Visible = False
                FocusControl flxAllocation
            End If
        
        
'    ElseIf KeyAscii = 13 And flxAllocation.row >= flxAllocation.Rows - 2 Then
'        If flxAllocation.col = 8 Then
'            If flxAllocation.TextMatrix(flxAllocation.row + 1, 0) = "" Then
'                flxAllocation.AddItem ""
'                BoolManualMode = True
'                flxAllocation.col = 1
'                txtPostingDateGrid.Visible = False
'                flxAllocation.SetFocus
'            End If
'        End If
    End If
End Sub

Private Sub txtPostingDateGrid_LostFocus()
    flxAllocation.ScrollBars = flexScrollBarVertical
End Sub

Private Sub txtReference1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdOK
    End If
End Sub

Private Sub txtReference1_LostFocus()
            If Len(txtReference1.text) = 0 Then
                    If MsgBox("You have not entered a reference. Do you wish to leave this blank?", vbYesNo, "Warning") = vbNo Then
                        Exit Sub
                    Else
                        FocusControl cmdOK
                    End If
            End If
End Sub

Private Sub txtReferenceGrid_GotFocus()
    flxAllocation.ScrollBars = flexScrollBarNone
End Sub

Private Sub txtReferenceGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      FocusControl flxAllocation
      txtReferenceGrid.text = ""
      txtReferenceGrid.Visible = False
   End If
    If KeyAscii = 13 And flxAllocation.row <= flxAllocation.Rows - 1 Then
        If flxAllocation.col = 7 And flxAllocation.TextMatrix(flxAllocation.row, 9) = "M" Then
'             If Trim(txtReferenceGrid.text) = "" Then
'                MsgBox "Please enter a Reference", vbInformation, "Reference"
'                txtReferenceGrid.SetFocus
'                Exit Sub
'            End If
            If Len(txtReferenceGrid.text) = 0 Then
                    If MsgBox("You have not entered a reference. Do you wish to leave this blank?", vbYesNo, "Warning") = vbNo Then
                        FocusControl txtReferenceGrid
                        Exit Sub
                    End If
            End If
            flxAllocation.TextMatrix(flxAllocation.row, 7) = txtReferenceGrid.text
            flxAllocation.col = 8
            txtReferenceGrid.Visible = False
            'flxAllocation.SetFocus
            flxAllocation_DblClick
        End If
       
    End If
End Sub

Private Sub txtReferenceGrid_LostFocus()
    flxAllocation.ScrollBars = flexScrollBarVertical
    txtReferenceGrid_KeyPress (13)
End Sub

Private Sub txtRefInput_GotFocus()
   iCurRow = StarFound

   txtRefInput.text = flxSPayment.TextMatrix(iCurRow, 24)
   'Added by anol 13 apr 2015
    txtRptDt.Visible = False
    txtPostingDate.Visible = False
    txtSPayment.Visible = False
    
End Sub

Private Sub txtRefInput_LostFocus()
   txtRefInput.Visible = False
   flxSPayment.TextMatrix(iCurRow, 24) = txtRefInput.text
   txtRefInput.Visible = False
   flxSPayment.ScrollBars = flexScrollBarVertical
End Sub

Private Sub txtReveiptDate1_Change()
    TextBoxChangeDate txtReveiptDate1
    lblPostingDate.ToolTipText = txtReveiptDate1.text
End Sub

Private Sub txtReveiptDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtReference1
    End If
End Sub

Private Sub txtReveiptDate1_LostFocus()
    If IsDate(txtReveiptDate1.text) = False Then
        MsgBox "Receipt date is not in valid format", vbInformation, "Warning"
    Else
        txtReveiptDate1.text = Format(txtReveiptDate1.text, "dd/mm/yyyy")
    End If
End Sub

Private Sub txtRptDt_Change()
   TextBoxChangeDate txtRptDt
End Sub

Private Sub txtRptDt_GotFocus()
   iCurRow = StarFound

   If iFlxSPayCol = 22 Then
      If flxSPayment.TextMatrix(iCurRow, 22) = "" Then
         txtRptDt.text = Format(Now, "dd/mm/yyyy")
      Else
         txtRptDt.text = Format(flxSPayment.TextMatrix(iCurRow, 22), "dd/mm/yyyy")
      End If
   End If
   txtSPayment.Visible = False
   txtRefInput.Visible = False
   txtPostingDate.Visible = False
   SelTxtInCtrl txtRptDt
End Sub

Private Sub txtRefInput_KeyPress(KeyAscii As Integer)
On Error GoTo Err
   If KeyAscii = 27 Then
      FocusControl flxSPayment
      txtRefInput = ""
      txtRefInput.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

'   If KeyAscii = 13 Then
''      If txtRefInput.text <> "" Then
'      flxSPayment.TextMatrix(iCurRow, 24) = txtRefInput.text
''      Else
''         flxSPayment.TextMatrix(iCurRow, 24) = ""
''      End If
'
'      txtRefInput.Visible = False
'      flxSPayment.ScrollBars = flexScrollBarVertical
'      'issue 458 anol 02 Feb 2015 Invalid procedure call
'      If Text1(0).Visible = True Then
'            Text1(0).SetFocus
'      End If
'   End If
        If KeyAscii = 13 Then
            'added by anol 13 Apr 2015
            If Len(txtRefInput.text) = 0 Then
                    If MsgBox("You have not entered a reference. Do you wish to leave this blank?", vbYesNo, "Warning") = vbNo Then
                    FocusControl txtRefInput
                    Exit Sub
                    End If
            End If
            flxSPayment.TextMatrix(iCurRow, 24) = txtRefInput.text
            txtRefInput.Visible = False
            flxSPayment.ScrollBars = flexScrollBarVertical
            flxSPayment.col = 25
            flxSPayment_dblClick
       End If
       Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub txtRptDt_KeyPress(KeyAscii As Integer)
On Error GoTo Err
   If KeyAscii = 27 Then
      FocusControl flxSPayment
      txtRptDt.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

'   If KeyAscii = 13 Then
'      If txtRptDt.text <> "" Then
'         If TextBoxFormatDate(txtRptDt) Then
'            If iFlxSPayCol = 22 Then
'               flxSPayment.TextMatrix(iCurRow, 22) = Format(txtRptDt.text, "dd/mm/yyyy")
'               'Resolved by BOSL
'               'added by anol 03 Feb 2015
'               'issue 468 Note 915 Posting date not changeing
'                flxSPayment.TextMatrix(iCurRow, 25) = flxSPayment.TextMatrix(iCurRow, 22)
''               If flxSPayment.TextMatrix(iCurRow, 25) = "" Then _
''                     flxSPayment.TextMatrix(iCurRow, 25) = flxSPayment.TextMatrix(iCurRow, 22)
'               flxSPayment.col = 24
'               flxSPayment_dblClick
'            ElseIf iFlxSPayCol = 24 Then
'               flxSPayment.TextMatrix(iCurRow, 22) = Format(txtRptDt.text, "dd/mm/yyyy")
'            Else
'               flxSPayment.TextMatrix(iCurRow, 25) = Format(txtRptDt.text, "dd/mm/yyyy")
'            End If
'         End If
'      Else
'
'          flxSPayment.TextMatrix(iCurRow, 22) = ""
'
'      End If
'
'      txtRptDt.Visible = False
'      flxSPayment.ScrollBars = flexScrollBarVertical
'   End If
         If KeyAscii = 13 Then
              If txtRptDt.text <> "" Then
                 If TextBoxFormatDate(txtRptDt) Then
                    If iFlxSPayCol = 22 Then
                       flxSPayment.TextMatrix(iCurRow, 22) = Format(txtRptDt.text, "dd/mm/yyyy")
                       'Resolved by BOSL
                       'added by anol 03 Feb 2015
                       'issue 468 Note 915 Posting date not changing
                        flxSPayment.TextMatrix(iCurRow, 25) = flxSPayment.TextMatrix(iCurRow, 22)
        '               If flxSPayment.TextMatrix(iCurRow, 25) = "" Then _
        '                     flxSPayment.TextMatrix(iCurRow, 25) = flxSPayment.TextMatrix(iCurRow, 22)
                       flxSPayment.col = 24
                       flxSPayment_dblClick
                       '17 Feb 2015 issue 530 Modified by anol
                    ElseIf iFlxSPayCol = 24 Then
                       flxSPayment.TextMatrix(iCurRow, 22) = Format(txtRptDt.text, "dd/mm/yyyy")
                    Else
                        '13 Apr 2015 issue 530 Modified by anol
                       flxSPayment.TextMatrix(iCurRow, 25) = Format(txtRptDt.text, "dd/mm/yyyy")
                       If MoveDownPosition Then
                            flxSPayment.col = 11
                            FocusControl flxSPayment
                       End If
                    End If
                 End If
              Else
                       '13 Apr 2015 issue 530 Modified by anol
                       flxSPayment.TextMatrix(iCurRow, 25) = Format(txtRptDt.text, "dd/mm/yyyy")
                       If MoveDownPosition Then
                            flxSPayment.col = 11
                            FocusControl flxSPayment
                       End If
              End If
        
              txtRptDt.Visible = False
             
              flxSPayment.ScrollBars = flexScrollBarVertical
           End If
   TextBoxKeyPrsDate txtRptDt, KeyAscii
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub txtRptDt_LostFocus()
    If txtRptDt.text <> "" Then TextBoxFormatDate txtRptDt
    flxSPayment.TextMatrix(iCurRow, 22) = txtRptDt.text
    
    If flxSPayment.TextMatrix(iCurRow, 22) <> "" Then
            flxSPayment.TextMatrix(iCurRow, 25) = flxSPayment.TextMatrix(iCurRow, 22)
        End If
    txtRptDt.Visible = False
    flxSPayment.ScrollBars = flexScrollBarVertical
'   If flxSPayment.col = 22 Then txtRptDt_KeyPress 13
'
'   If flxSPayment.col = 25 Then txtRptDt_KeyPress 13
'    '17 Feb 2015 issue 530
'   If flxSPayment.col = 24 Then txtRptDt_KeyPress 13
End Sub

Private Sub txtSPayment_GotFocus()
   If Val(txtSPayment.text) = 0 Then
      txtSPayment.text = Format(flxSPayment.TextMatrix(flxSPayment.row, 10), "0.00")
      SelTxtInCtrl txtSPayment
   End If
   'anol 13 Apr 2015
   txtRptDt.Visible = False
   txtPostingDate.Visible = False
   txtRefInput.Visible = False
End Sub

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer
On Error GoTo Err
   If flxSPayment.row < flxSPayment.Rows - 1 Then
      If flxSPayment.RowHeight(flxSPayment.row + 1) = 0 Then
      '22 Jun 2015
         'txtSPayment.text = "0.00"
         txtSPayment.Visible = False
         Text1(0).Visible = False
         MoveDownPosition = False
         Exit Function
      End If
      flxSPayment.row = flxSPayment.row + 1
   Else
      'txtSPayment.text = "0.00"
      txtSPayment.Visible = False
      Text1(0).Visible = False
      MoveDownPosition = False
      'cmdGPayment.SetFocus
      If Val(cmdAll.Tag) = 0 Then
             FocusControl cmdGPayment
        Else
             FocusControl cmdAll
        End If
      Exit Function
   End If

   iLeft = flxSPayment.CellLeft + flxSPayment.Left
   iTop = flxSPayment.CellTop + flxSPayment.Top
   MoveDownPosition = True
   
Exit Function
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Function

Private Sub txtSPayment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
   'If KeyCode = 13 Then Text1(0).SetFocus
   'Re written by anol 13 apr 2015
   'issue 530 Smoothing operation on batch payment
    If KeyCode = 13 Then
            If flxSPayment.row < flxSPayment.Rows And frmBRPreForm.chkMultiple.Value = 1 And Not txtRptDt.Visible And flxSPayment.col = 11 Then
                txtSPayment.Visible = False
                flxSPayment.col = 22
                iFlxSPayCol = 22
                txtRptDt.Top = flxSPayment.CellTop + flxSPayment.Top
                txtRptDt.Left = flxSPayment.CellLeft + flxSPayment.Left
                txtRptDt.Width = flxSPayment.ColWidth(22)
                txtRptDt.Height = flxSPayment.RowHeight(iCurRow) - 15
                txtRptDt.Visible = True
                flxSPayment.ScrollBars = flexScrollBarNone
                FocusControl txtRptDt
                SumUpTotal
           End If
       
        '  Move to the next row to enter amount
           If flxSPayment.row < flxSPayment.Rows And Not txtRptDt.Visible And flxSPayment.col = 11 Then  'flxSPayment.col = 24 Or
              If MoveDownPosition Then
                 flxSPayment.col = 11
                 FocusControl flxSPayment
              Else
                 'cmdSavePayment.SetFocus
                 '22 Jun 2015
                 If Val(cmdAll.Tag) = 0 Then
                    FocusControl cmdGPayment
                 Else
                    FocusControl cmdAll
                 End If
              End If
           End If
   
   End If
   
Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub txtSPayment_KeyPress(KeyAscii As Integer)
On Error GoTo Err
   If KeyAscii = 13 Then
        KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      FocusControl flxSPayment
      txtSPayment.text = "0.00"
      txtSPayment.Visible = False
      Text1(0).Visible = False
   End If
   DigitTextKeyPress txtSPayment, KeyAscii
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
End Sub

Private Sub SumUpTotal()
   Dim i As Integer, cGT As Currency

   For i = 1 To flxSPayment.Rows - 1
      If Left(flxSPayment.TextMatrix(i, 1), 2) = "SI" Then
         cGT = cGT + Val(flxSPayment.TextMatrix(i, 11))
      Else
         cGT = cGT - Val(flxSPayment.TextMatrix(i, 11))
      End If
   Next i

   txtGrossTotal.text = Format(cGT, "0.00")
End Sub

Private Sub SeperateDrCr(szSupp As String)
   Dim i As Integer, cTotalPay As Currency
   Dim iC As Integer, iI As Integer

   ReDim iaPI_RowNo(0)
   ReDim iaPC_RowNo(0)

   For i = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(i, 21) = szSupp And flxSPayment.TextMatrix(i, 11) > 0 Then
         If flxSPayment.TextMatrix(i, 15) = 2 Or flxSPayment.TextMatrix(i, 15) = 4 Then
            iaPC_RowNo(iC) = i
            iC = iC + 1
            ReDim Preserve iaPC_RowNo(UBound(iaPC_RowNo) + 1)
         Else
            iaPI_RowNo(iI) = i
            iI = iI + 1
            ReDim Preserve iaPI_RowNo(UBound(iaPI_RowNo) + 1)
         End If
      End If
   Next i
End Sub

Private Function NoCrTrn(szSuppID As String) As Boolean
   Dim i As Integer

   NoCrTrn = True
   flxSPayment.col = 10

   For i = 1 To flxSPayment.Rows - 2
      If flxSPayment.TextMatrix(i, 21) = szSuppID And Val(flxSPayment.TextMatrix(i, 11)) > 0 Then
         flxSPayment.row = i
         If flxSPayment.CellForeColor = vbRed Then
            NoCrTrn = False
            Exit For
         End If
      End If
   Next i
End Function

Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim Data() As String

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

   adoRst.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadProperties(adoConn As ADODB.Connection, cboP As Control, szClientID As String)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, j As Integer
   Dim i As Integer, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer


   On Error GoTo ErrorHandler

'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & szClientID & "' " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadTenant(ByVal adoConn As ADODB.Connection, cboS As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, iTotalRow As Integer, j As Integer
   Dim i As Integer, iTotalCol As Integer, Data() As String

   On Error GoTo ErrorHandler

   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag = "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If frmBRPreForm.txtClient.Tag <> "ALL" And frmBRPreForm.txtProperty.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & frmBRPreForm.txtClient.Tag & "' AND " & _
               "Units.PropertyID = '" & frmBRPreForm.txtProperty.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox vbTab & "Either there are no supplier records entered in the system or " & vbCrLf & _
             vbTab & "supplier payement type does not match with you selection." & vbCrLf & vbCrLf & _
             "Please enter supplier in the supplier module or set the receipt type for each supplier.", vbInformation + vbOKOnly, "Batch Receipt"

      GoTo NoRes
   End If

   iTotalRow = adoRst.RecordCount
   iTotalCol = adoRst.Fields.Count

   ReDim Data(iTotalCol - 1, iTotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Tenants"
   For i = 1 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboS.Column() = Data()
   cboS.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRst = Nothing
End Sub

Private Sub txtSPayment_LostFocus()
On Error GoTo Err
   If Val(flxSPayment.TextMatrix(iCurRow, 10)) < Val(txtSPayment.text) Then
      MsgBox "Payment amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
      txtSPayment.text = "0.00"
      If txtSPayment.Visible = True Then
            FocusControl txtSPayment
      End If
      Exit Sub
   End If

   If txtSPayment.text = "" Then txtSPayment.text = "0.00"
   If flxSPayment.TextMatrix(iCurRow, 2) = flxSPayment.TextMatrix(iCurRow, 26) Then 'if they are equal that means they have allocation problem
        flxSPayment.TextMatrix(iCurRow, 11) = "0.00"
        MsgBox "A problem exists relating to a previous transaction entered against the selected lessee: " & _
                     Chr(13) & _
                     "Please contact PCM Consulting. ", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"

   Else
        flxSPayment.TextMatrix(iCurRow, 11) = Format(txtSPayment.text, "0.00")
   End If
   SumUpTotal
   txtSPayment.Visible = False
   flxSPayment.ScrollBars = flexScrollBarVertical
   Exit Sub
Err:
    ShowMsgInTaskBar Err.description, "Y", "P"
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
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
           'bHandled = False

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

