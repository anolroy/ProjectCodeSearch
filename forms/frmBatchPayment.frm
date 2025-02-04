VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBatchPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Payments"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16575
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   16575
   ShowInTaskbar   =   0   'False
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
      Left            =   13140
      TabIndex        =   70
      Top             =   0
      Width           =   2400
      Begin VB.TextBox txtBankBal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   195
         Width           =   1125
      End
      Begin VB.TextBox txtRetentions1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   525
         Width           =   1125
      End
      Begin VB.TextBox txtAvailableBankBal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Balance  £"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retentions  £"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avail.Bank Bal£"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   74
         Top             =   855
         Width           =   1050
      End
   End
   Begin VB.Frame grpUploadReceipts 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Upload Payments"
      Height          =   750
      Index           =   0
      Left            =   8820
      TabIndex        =   67
      Top             =   6660
      Width           =   1815
      Begin VB.CommandButton cmdUploadReceipts 
         BackColor       =   &H00F0F0F0&
         Caption         =   "&Upload Payments"
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   225
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Selection"
      Height          =   735
      Left            =   135
      TabIndex        =   66
      Top             =   6660
      Width           =   6645
      Begin VB.CommandButton cmdPaySelected 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Pay Selected"
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
         Left            =   5265
         TabIndex        =   10
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton cmdSavePayment 
         BackColor       =   &H00F0F0F0&
         Caption         =   "&Save"
         Height          =   375
         Left            =   30
         TabIndex        =   6
         ToolTipText     =   "Generate Payment later"
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelAll 
         BackColor       =   &H00F0F0F0&
         Caption         =   "&Pay All"
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
         Left            =   4005
         TabIndex        =   9
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton cmdClearSel 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1260
         TabIndex        =   7
         Top             =   225
         Width           =   1395
      End
      Begin VB.CommandButton cmdPaymentDiscard 
         BackColor       =   &H00F0F0F0&
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
         Left            =   2655
         TabIndex        =   8
         Top             =   225
         Width           =   1350
      End
   End
   Begin VB.TextBox txtRef 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   14535
      TabIndex        =   5
      Top             =   1890
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtPostingDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   13275
      TabIndex        =   4
      Top             =   1890
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
      Left            =   4920
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   49
      Top             =   7800
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   53
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   465
         End
         Begin MSForms.ComboBox ComboBox2 
            Height          =   315
            Left            =   3675
            TabIndex        =   55
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
         Begin MSForms.ComboBox ComboBox1 
            Height          =   315
            Left            =   480
            TabIndex        =   54
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
         Left            =   1560
         TabIndex        =   52
         Top             =   300
         Width           =   2415
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
         Left            =   120
         TabIndex        =   51
         Top             =   300
         Width           =   1335
      End
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
         TabIndex        =   50
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   58
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4392
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
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
         Caption         =   "Account"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   61
         Top             =   75
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   60
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   59
         Top             =   75
         Width           =   540
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
      Left            =   12720
      TabIndex        =   21
      Text            =   "this control is using to use tab in the grid"
      Top             =   8040
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
      Left            =   12720
      TabIndex        =   18
      Text            =   "this control is using to use tab in the grid"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox txtPayDt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12120
      TabIndex        =   3
      Top             =   1890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment:"
      ForeColor       =   &H00000000&
      Height          =   750
      Index           =   5
      Left            =   6795
      TabIndex        =   38
      Top             =   6660
      Width           =   1995
      Begin VB.CommandButton cmdSPayAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pay &All"
         Height          =   375
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdSPFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pay in &Full"
         Height          =   375
         Left            =   2085
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdGPayment 
         BackColor       =   &H00F0F0F0&
         Caption         =   "&Generate Payment Now"
         Height          =   375
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdSPClose 
      BackColor       =   &H00F0F0F0&
      Caption         =   "C&lose"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6855
      Width           =   1380
   End
   Begin VB.TextBox txtSPayment 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10800
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Payment Method"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   75
      Width           =   2100
      Begin VB.OptionButton optBP_MULT 
         Caption         =   "Multiple"
         Height          =   300
         Left            =   960
         TabIndex        =   46
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton optBP_Cheque 
         Caption         =   "Cheque"
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optBP_BACS 
         Caption         =   "BACS"
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   1545
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8493
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
   Begin VB.Label lblBC 
      BackStyle       =   0  'Transparent
      Caption         =   "lblBC"
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
      Left            =   6390
      TabIndex        =   69
      Top             =   135
      Width           =   735
   End
   Begin VB.Label lblRef 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   195
      Left            =   14760
      TabIndex        =   65
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   195
      Index           =   12
      Left            =   13560
      TabIndex        =   64
      Top             =   1260
      Width           =   1005
   End
   Begin MSForms.ComboBox cboFund 
      Height          =   285
      Left            =   3120
      TabIndex        =   63
      Top             =   800
      Visible         =   0   'False
      Width           =   2520
      VariousPropertyBits=   1820346395
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4445;503"
      TextColumn      =   2
      ColumnCount     =   6
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "705;70555"
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund:"
      Height          =   195
      Index           =   5
      Left            =   2400
      TabIndex        =   62
      Top             =   800
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSForms.CommandButton cmdDmdSuppLookup 
      Height          =   240
      Left            =   10050
      TabIndex        =   47
      Top             =   510
      Width           =   255
      Caption         =   """"
      Size            =   "450;423"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtSupplierName 
      Height          =   285
      Left            =   6480
      TabIndex        =   48
      Top             =   480
      Width           =   3840
      VariousPropertyBits=   746604575
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "6773;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      Height          =   195
      Index           =   1
      Left            =   5820
      TabIndex        =   37
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date"
      Height          =   195
      Index           =   11
      Left            =   12420
      TabIndex        =   45
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment:"
      Height          =   195
      Index           =   2
      Left            =   10440
      TabIndex        =   44
      Top             =   6420
      Width           =   1020
   End
   Begin MSForms.TextBox txtGrossTotal 
      Height          =   300
      Left            =   11760
      TabIndex        =   43
      Top             =   6420
      Width           =   1380
      VariousPropertyBits=   679495711
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2434;529"
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
      Left            =   11220
      TabIndex        =   42
      Top             =   120
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
      Left            =   7200
      TabIndex        =   41
      Top             =   120
      Width           =   3390
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
      Left            =   3120
      TabIndex        =   40
      Top             =   480
      Width           =   2505
   End
   Begin MSForms.ComboBox cboSupplier_ 
      Height          =   285
      Left            =   7320
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2963;503"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
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
      Left            =   2400
      TabIndex        =   36
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   12
      Left            =   2400
      TabIndex        =   35
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      Height          =   195
      Index           =   2
      Left            =   1050
      TabIndex        =   34
      Top             =   1260
      Width           =   600
   End
   Begin MSForms.TextBox txtChqNo 
      Height          =   285
      Left            =   11220
      TabIndex        =   17
      Top             =   480
      Width           =   1860
      VariousPropertyBits=   679495707
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3281;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref/Chq:"
      Height          =   195
      Index           =   4
      Left            =   10440
      TabIndex        =   33
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   195
      Index           =   3
      Left            =   10455
      TabIndex        =   32
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment £"
      Height          =   195
      Index           =   10
      Left            =   11490
      TabIndex        =   30
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Amt. £"
      Height          =   195
      Index           =   9
      Left            =   10500
      TabIndex        =   29
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      Height          =   195
      Index           =   8
      Left            =   9420
      TabIndex        =   28
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   195
      Index           =   7
      Left            =   7380
      TabIndex        =   27
      Top             =   1260
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   195
      Index           =   6
      Left            =   6120
      TabIndex        =   26
      Top             =   1260
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Index           =   5
      Left            =   5040
      TabIndex        =   25
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   24
      Top             =   1260
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   3
      Left            =   2490
      TabIndex        =   23
      Top             =   1260
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      Height          =   195
      Index           =   0
      Left            =   5820
      TabIndex        =   19
      Top             =   120
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
      TabIndex        =   31
      Top             =   1260
      Width           =   15585
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
      Left            =   3120
      TabIndex        =   39
      Top             =   120
      Width           =   2505
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Left            =   2280
      Top             =   75
      Width           =   10845
   End
End
Attribute VB_Name = "frmBatchPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bBPPreForm    As Boolean        'This is used to monitor form status (loaded/unloaded), in the frmMMain
Private bMultiple    As Boolean

Dim szaSuppID()      As String
Dim snFileHandeling  As Single
Dim iaPI_RowNo()     As Integer
Dim iaPC_RowNo()     As Integer
Dim iLeft            As Integer
Dim iTop             As Integer
Dim iCurRow          As Integer
Dim bSavedPayment    As Boolean
Dim szSuppCnt        As String
Dim cOpeningBal      As Currency
Dim szSubject        As String
Dim szBody           As String
Dim iFlxSPayCol      As Integer

Dim szaSupplierBalance()  As String
Public vPropertyName As String
Public vClientName As String
Public vPropertyID As String
Public vClientID As String
Dim UserSessionID As String
Dim frmLockingDialogisActive As Boolean

Dim colTransactionIDOther As String 'this variable shall hold all the locked transaction number which i s locked by other screen
Private Sub RefreshGridSupp()
   Dim iRow As Integer

   For iRow = 1 To flxSPayment.Rows - 1
      flxSPayment.RowHeight(iRow) = 240
   Next iRow

   If cboSupplier_.Column(0) = "ALL" Then Exit Sub

   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 2) <> cboSupplier_.Column(1) Then
         flxSPayment.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub

Private Sub ConfigureDmdFlxLeaseList()
   Dim szHeader As String

   flxDmdLeaseList.Clear
   flxDmdLeaseList.Cols = 4
   flxDmdLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Tenant ID|<Tenant Name|<Unit Name"
   flxDmdLeaseList.FormatString = szHeader$
   flxDmdLeaseList.ColWidth(0) = Label20(9).Left - flxDmdLeaseList.Left   '240        Solid column
   flxDmdLeaseList.ColWidth(1) = Label20(8).Left - Label20(9).Left - 20  '1400       'Tenant ID
   flxDmdLeaseList.ColWidth(2) = Label20(7).Left - Label20(8).Left - 20         'Tenant Name
   flxDmdLeaseList.ColWidth(3) = flxDmdLeaseList.Left + flxDmdLeaseList.Width - Label20(7).Left - 300 'Unit Name
   flxDmdLeaseList.Rows = 2
End Sub

Private Sub cmdClearSel_Click()
     'added by anol 21 Apr 2015
    'issue 547
      txtSPayment.text = ""
      txtSPayment.Visible = False
      txtPayDt.text = ""
      txtPayDt.Visible = False
      txtRef.text = ""
      txtRef.Visible = False
      txtPostingDate.text = ""
      txtPostingDate.Visible = False
      Dim i As Integer
      For i = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(i, 0) = "X" Then
            flxSPayment.TextMatrix(i, 0) = ""
            flxSPayment.TextMatrix(i, 11) = "0.00"
            If bMultiple = True Then
                flxSPayment.TextMatrix(i, 22) = ""
                flxSPayment.TextMatrix(i, 24) = ""
                flxSPayment.TextMatrix(i, 23) = ""
            End If
        End If
      Next i
      flxSPayment.col = 11
      flxSPayment.SetFocus
      SumUpTotal
End Sub

Private Sub cmdDmdGridUnitLookup_Click()
   picDmdLeaseList.Visible = False
End Sub

Private Sub cmdDmdSuppLookup_Click()
   Me.MousePointer = vbHourglass

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   ConfigureDmdFlxLeaseList
   If frmBPPreForm.optBP_BACS.Value Then
      szSQL = "SELECT SupplierID, SupplierName, AcBalance " & _
              "FROM Supplier " & _
              "WHERE PaymentType = 'BACS' " & _
              "ORDER BY SupplierName;"
   Else
      szSQL = "SELECT SupplierID, SupplierName, AcBalance  " & _
              "FROM Supplier " & _
              "WHERE PaymentType = 'CHQ' " & _
              "ORDER BY SupplierName;"
   End If
'Debug.Print szSQL
   PopulateDmdTenantLookup adoConn, szSQL
   UpdateBalance

   adoConn.Close
   Set adoConn = Nothing

   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchName.text = ""

   picDmdLeaseList.Top = txtSupplierName.Top + txtSupplierName.Height + 5
   picDmdLeaseList.Left = txtSupplierName.Left + 5
   picDmdLeaseList.Visible = True
   picDmdLeaseList.ZOrder 0
   flxDmdLeaseList.SetFocus

   Me.MousePointer = vbArrow
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxDmdLeaseList.Rows - 1
      For j = 0 To UBound(szaSupplierBalance, 2) - 1
         If flxDmdLeaseList.TextMatrix(i, 1) = szaSupplierBalance(0, j) Then
            flxDmdLeaseList.TextMatrix(i, 3) = Format(szaSupplierBalance(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalance, 2) Then flxDmdLeaseList.TextMatrix(i, 3) = "0.00"
   Next i
End Sub

Public Function PopulateDmdTenantLookup(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox vbTab & "Either there are no supplier records entered in the system or " & vbCrLf & _
             vbTab & "there are no suppliers with payment type that matches you selection." & vbCrLf & vbCrLf & _
             "Please enter a supplier in the supplier module or set a payment type on the" & vbCrLf & vbCrLf & _
             vbTab & "supplier record for the supplier you wish to pay.", vbInformation + vbOKOnly, "Batch Payment"

      GoTo NoRes
   End If

   Dim iRow As Integer
   iRow = 1
   flxDmdLeaseList.TextMatrix(iRow, 1) = "ALL"
   flxDmdLeaseList.TextMatrix(iRow, 2) = "All Suppliers"
   If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   iRow = 2

   While Not adoRst.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRst!SupplierID
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRst!SupplierName
'      flxDmdLeaseList.TextMatrix(iRow, 3) = adoRst!AcBalance

      iRow = iRow + 1
      adoRst.MoveNext

      If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   Wend

NoRes:
   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub cmdGPayment_GotFocus()
     'added by anol 21 Apr 2015
    'issue 547 Making operation smooth of batch payment form
    txtPayDt.text = ""
    txtPostingDate.text = ""
    txtRef.text = ""
    txtSPayment.text = ""
    txtPayDt.Visible = False
    txtPostingDate.Visible = False
    txtRef.Visible = False
    txtSPayment.Visible = False
End Sub

Private Sub cmdPaymentDiscard_Click()
'   Dim i As Integer, iFlxTRptCol As Integer
'
'   For i = 1 To flxSPayment.Rows - 1
'      If Val(flxSPayment.TextMatrix(i, 11)) > 0 Then
'         flxSPayment.TextMatrix(i, 11) = "0.00"
'      End If
'      If optBP_MULT.Value Then
'         If flxSPayment.TextMatrix(i, 22) <> "" Then
'            flxSPayment.TextMatrix(i, 22) = ""
'         End If
'         'added by anol 23 March 'Reset button not clearing posting date
'         flxSPayment.TextMatrix(i, 23) = ""
'      End If
'   Next i
'   'Below line has been added by anol 09 Apr 2015
'   'issue 547 batch payment reset button is not working
'   txtPayDt.text = ""
'   txtGrossTotal.text = "0.00"
      Dim i As Integer, iFlxTRptCol As Integer
      txtSPayment.text = ""
      txtSPayment.Visible = False
      txtPayDt.text = ""
      txtPayDt.Visible = False
      txtRef.text = ""
      txtRef.Visible = False
      txtPostingDate.text = ""
      txtPostingDate.Visible = False
      For i = 1 To flxSPayment.Rows - 1
           flxSPayment.TextMatrix(i, 0) = ""
           If Val(flxSPayment.TextMatrix(i, 11)) > 0 Then
              flxSPayment.TextMatrix(i, 11) = "0.00"
           End If
           If optBP_MULT.Value Then
                 flxSPayment.TextMatrix(i, 0) = ""
                 flxSPayment.TextMatrix(i, 11) = "0.00"
                 flxSPayment.TextMatrix(i, 22) = ""
                 flxSPayment.TextMatrix(i, 24) = ""
                 flxSPayment.TextMatrix(i, 23) = ""
           End If
      Next i
      txtGrossTotal.text = "0.00"
End Sub

Private Function CheckDataValidation(adoConn As ADODB.Connection) As Boolean
   Dim iRow As Integer

   CheckDataValidation = True
   For iRow = 1 To flxSPayment.Rows - 1
      If (Val(flxSPayment.TextMatrix(iRow, 11)) = 0 And flxSPayment.TextMatrix(iRow, 22) <> "") Then
         MsgBox "Please input the amount for the transaction number " & flxSPayment.TextMatrix(iRow, 1), vbCritical + vbOKOnly, "Multiple Batch Payment"
         CheckDataValidation = False
         Exit Function
      End If
      If (Val(flxSPayment.TextMatrix(iRow, 11)) <> 0 And flxSPayment.TextMatrix(iRow, 22) = "") Then
         MsgBox "Please input the payment date for the transaction number " & flxSPayment.TextMatrix(iRow, 1), vbCritical + vbOKOnly, "Multiple Batch Payment"
         CheckDataValidation = False
         Exit Function
      End If
      'added by anol 21 Apr 2015
      'issue 547
      If (Val(flxSPayment.TextMatrix(iRow, 11)) <> 0 And flxSPayment.TextMatrix(iRow, 23) = "") Then
         MsgBox "Please input the posting date for the transaction number " & flxSPayment.TextMatrix(iRow, 1), vbCritical + vbOKOnly, "Multiple Batch Payment"
         CheckDataValidation = False
         Exit Function
      End If
       'added by anol 201609011
      If IsDate(flxSPayment.TextMatrix(iRow, 23)) = True Then
              If IsPeriodStatus(flxSPayment.TextMatrix(iRow, 23), vClientID, adoConn) = 0 Then
                  ShowMsgInTaskBar "The posting date cannot fall within a closed financial period at row " & iRow, "Y", "N"
                  CheckDataValidation = False
                  Exit Function
              ElseIf IsPeriodStatus(flxSPayment.TextMatrix(iRow, 23), vClientID, adoConn) = 9 Then
                  ShowMsgInTaskBar "The posting date does not fall in any existing financial period at row " & iRow, "Y", "N"
                  CheckDataValidation = False
                  Exit Function
              End If
      End If
   Next iRow
End Function

Private Sub cmdPaySelected_Click()
   Dim i As Integer
   flxSPayment.col = 0
   For i = 1 To flxSPayment.Rows - 1
      flxSPayment.row = i
      If Val(flxSPayment.TextMatrix(i, 10)) > 0 And flxSPayment.TextMatrix(i, 0) = "X" And flxSPayment.CellBackColor <> vbRed Then
         flxSPayment.TextMatrix(i, 11) = flxSPayment.TextMatrix(i, 10)
      Else
         flxSPayment.TextMatrix(i, 11) = "0.00"
      End If
      If optBP_MULT.Value And flxSPayment.TextMatrix(i, 0) = "X" And flxSPayment.CellBackColor <> vbRed Then
         If flxSPayment.TextMatrix(i, 22) = "" Then
            flxSPayment.TextMatrix(i, 22) = flxSPayment.TextMatrix(i, 6)
            'Fixed by anol 25 Mar 2015
            'It was not filling posting date when user selects select all
            flxSPayment.TextMatrix(i, 23) = flxSPayment.TextMatrix(i, 6)
         End If
      Else
         flxSPayment.TextMatrix(i, 22) = ""
         flxSPayment.TextMatrix(i, 23) = ""
      End If
   Next i

   SumUpTotal
End Sub

Private Sub cmdSavePayment_Click()
  

   On Error GoTo ErrHandler

   Dim iRow As Integer, szSQL As String, i As Integer
   Dim cTotalPI As Currency, cTotalPC As Currency
   Dim adoConn As New ADODB.Connection
   Dim adoBP As New ADODB.Recordset, adoBT As New ADODB.Recordset
   'issue 521
   If ValidationPostingDate = False Then
        Exit Sub
   End If
   adoConn.Open getConnectionString
   If bMultiple Then
      If Not CheckDataValidation(adoConn) Then
            adoConn.Close
            Exit Sub
      End If
   End If
      
      
   If Val(txtGrossTotal.text) <> 0 Then
      If MsgBox("Do you wish to save your payment selection and generate your payment later?", vbQuestion + vbYesNo, "Batch Payment") = vbNo Then
            adoConn.Close
            Exit Sub
       End If
   Else
      adoConn.Close
      MsgBox "No transaction to save.", vbInformation + vbOKOnly, "Batch Payment"
      Exit Sub
   End If

'   Database has been connected.
 

   Call SaveSuggestedPayment(adoConn, False)
   MsgBox "This Batch has been saved.", vbInformation + vbOKOnly, "Batch Payments"
   adoConn.Close
   Set adoConn = Nothing

   If MsgBox("Do you wish to print a list of your selected payments now?", vbQuestion + vbYesNo, "Batch Payment") = vbNo Then Exit Sub
'  ********************************************************************************************************
'  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~     PRINT SUGGESTED PAYMENT        ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  ********************************************************************************************************

   ShowReport App.Path & szReportPath & "\SuggestedPayment.rpt"

   Exit Sub
ErrHandler:
   Debug.Print Err.Number & ": " & Err.description
End Sub

Private Function SavedSuggestedPayment(adoConn As ADODB.Connection) As String
   Dim szSQL As String
   Dim adoBP As New ADODB.Recordset

   szSQL = "SELECT BP.BP " & _
           "FROM tblBatchPayment AS BP " & _
           "WHERE BP.Generated = FALSE;"

'Debug.Print szSQL
   adoBP.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoBP.EOF Then
      SavedSuggestedPayment = adoBP.Fields.Item("BP").Value
   Else
      SavedSuggestedPayment = "NF"
   End If

   adoBP.Close
   Set adoBP = Nothing
End Function

'  bSettled @
Private Function ValidationPostingDate() As Boolean
'issue 521
'system wrongly allows the posting date for all payments and receipt edits to be set before transaction date
' Date 26 Jan 2018 By anol
    Dim iRow As Integer
    If frmBPPreForm.optBP_MULT.Value = False Then
        If DateDiff("d", Format(frmBPPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy"), Format(frmBPPreForm.txtDate.text, "dd/mm/yyyy")) > 0 Then
               MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
               Exit Function
        End If
    End If
    If frmBPPreForm.optBP_MULT.Value = True Then
        For iRow = 1 To flxSPayment.Rows - 1
            If Trim(flxSPayment.TextMatrix(iRow, 23)) <> "" And flxSPayment.TextMatrix(iRow, 22) <> "" Then
                If DateDiff("d", Format(flxSPayment.TextMatrix(iRow, 23), "dd mmmm yyyy"), Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy")) > 0 Then
                    MsgBox "Posting date cannot be before the transaction date", vbInformation, flxSPayment.TextMatrix(iRow, 1)
                   Exit Function
                End If
            End If
        Next iRow
    End If
    ValidationPostingDate = True
End Function
Private Function SaveSuggestedPayment(adoConn As ADODB.Connection, bSettled As Boolean) As String
   On Error GoTo ErrHandler

   Dim iRow As Integer, szSQL As String, i As Integer
   Dim adoBP As New ADODB.Recordset, adoBT As New ADODB.Recordset

   szSuppCnt = ""    'reset the supplier list

   szSQL = "DELETE * FROM tblBatchTransaction " & _
           "WHERE BP IN (SELECT BP.BP FROM tblBatchPayment AS BP " & _
                        "WHERE BP.Generated = FALSE);"
'Debug.Print szSQL
   adoConn.Execute szSQL
   adoConn.Execute "DELETE * FROM tblBatchPayment WHERE Generated = FALSE;"

   adoBP.Open "SELECT * FROM tblBatchPayment;", adoConn, adOpenDynamic, adLockPessimistic

   With adoBP
      .AddNew
      SaveSuggestedPayment = UniqueID()
      .Fields.Item("BP").Value = SaveSuggestedPayment

      If frmBPPreForm.optBP_MULT.Value Then
         .Fields.Item("BPDate").Value = Format(Now, "dd mmmm yyyy")
         
      Else
         .Fields.Item("BPDate").Value = Format(lblDate.Caption, "dd mmmm yyyy")
         '.Fields.Item("PostingDate").Value = Format(frmBPPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
      End If
       'issue 547
       'anol  06 apr 2015
       'Below line shall update posting date for multiple
      .Fields.Item("PostingDate").Value = Format(frmBPPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
      .Fields.Item("PayOption").Value = IIf(optBP_Cheque.Value, "C", "B")
      .Fields.Item("ClientID").Value = frmBPPreForm.txtClient.Tag
      .Fields.Item("PropertyID").Value = frmBPPreForm.txtProperty.Tag
      .Fields.Item("Bank_ID").Value = frmBPPreForm.cmbBankAc.Column(2)
      .Fields.Item("SupplierID").Value = cboSupplier_.Column(0)
      .Fields.Item("BatchNo").Value = txtChqNo.text
      .Fields.Item("Generated").Value = bSettled
      .Fields.Item("ChqNo").Value = frmBPPreForm.txtCheqNo.text
      .Update
      .Close
   End With
   Set adoBP = Nothing

   adoBT.Open "SELECT * FROM tblBatchTransaction;", adoConn, adOpenDynamic, adLockPessimistic

   With adoBT
      For iRow = 1 To flxSPayment.Rows - 1
         If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
            .AddNew
            .Fields.Item("BT").Value = UniqueID()
            .Fields.Item("BP").Value = SaveSuggestedPayment
            .Fields.Item("TransactionID").Value = flxSPayment.TextMatrix(iRow, 20)
            .Fields.Item("SupplierID").Value = flxSPayment.TextMatrix(iRow, 21)
            CountSupplier (flxSPayment.TextMatrix(iRow, 21))
            .Fields.Item("TranType").Value = 8
            .Fields.Item("PropertyID").Value = flxSPayment.TextMatrix(iRow, 5)
            .Fields.Item("DueDate").Value = IIf(flxSPayment.TextMatrix(iRow, 6) = "", Null, flxSPayment.TextMatrix(iRow, 6))
            'Below line modified by anol 22 Apr 2015
            'Issue 547 Reference is not saving
            If Not bMultiple Then
                .Fields.Item("Ref").Value = flxSPayment.TextMatrix(iRow, 7)
            Else
                .Fields.Item("Ref").Value = flxSPayment.TextMatrix(iRow, 24)
                'adoConn.Execute "Update tlbPayment set ref='" & flxSPayment.TextMatrix(iRow, 24) & "' where TransactionID=" & flxSPayment.TextMatrix(iRow, 20) & ""
            End If
            'End of modification
            .Fields.Item("Details").Value = flxSPayment.TextMatrix(iRow, 8)
            .Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 9)
            .Fields.Item("OSAmt").Value = flxSPayment.TextMatrix(iRow, 10)
            .Fields.Item("PayAmt").Value = flxSPayment.TextMatrix(iRow, 11)
            If bMultiple Then
               .Fields.Item("PayDt").Value = Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy")
               .Fields.Item("PostingDate").Value = Format(flxSPayment.TextMatrix(iRow, 23), "dd mmmm yyyy")
            Else
               .Fields.Item("PayDt").Value = Format(frmBPPreForm.txtDate.text, "dd/mm/yyyy")
               .Fields.Item("PostingDate").Value = Format(frmBPPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
            End If
            .Update
         End If
      Next iRow
   End With
   adoBT.Close
   Set adoBT = Nothing

   Exit Function
ErrHandler:
   Debug.Print Err.Number & ": " & Err.description
End Function

Private Sub CountSupplier(szSupplier As String)
   Dim szaSuppCnt() As String, i As Integer

   szaSuppCnt = Split(szSuppCnt, "#*#")

   For i = 0 To UBound(szaSuppCnt)
      If szaSuppCnt(i) <> szSupplier Then
         If UBound(szaSuppCnt) = 0 Then
            szSuppCnt = szSupplier
         Else
            szSuppCnt = szSuppCnt & "#*#" & szSupplier
         End If
      Else
         Exit For
      End If
   Next i
End Sub

Private Sub cmdSelAll_Click() 'actually this does the work of pay all
   Dim i As Integer
   flxSPayment.col = 0
   For i = 1 To flxSPayment.Rows - 1
      flxSPayment.row = i
      If Val(flxSPayment.TextMatrix(i, 10)) > 0 And flxSPayment.CellBackColor <> vbRed Then 'red items are locked . So I am avoiding those to fill
         flxSPayment.TextMatrix(i, 11) = flxSPayment.TextMatrix(i, 10) 'fill the amounts
      End If
      If optBP_MULT.Value And flxSPayment.CellBackColor <> vbRed Then
         If flxSPayment.TextMatrix(i, 22) = "" Then
            flxSPayment.TextMatrix(i, 22) = flxSPayment.TextMatrix(i, 6)
            'Fixed by anol 25 Mar 2015
            'It was not filling posting date when user selects select all
            flxSPayment.TextMatrix(i, 23) = flxSPayment.TextMatrix(i, 6)
         End If
      End If
   Next i

   SumUpTotal
End Sub

Private Sub cmdSPClose_Click()
   Unload Me
End Sub

Private Sub flxDmdLeaseList_Click()
   Dim adoConn As New ADODB.Connection

   txtSupplierName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) & " / " & flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)

   If cboSupplier_.ListCount = 0 Then cboSupplier_.AddItem ""
   cboSupplier_.ListIndex = 0
   cboSupplier_.Column(0) = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1)
   cboSupplier_.Column(1) = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)

   RefreshGridSupp

   picDmdLeaseList.Visible = False
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

Private Sub flxSPayment_dblClick()
   Dim i As Integer
   Dim selcol As Integer
   If flxSPayment.TextMatrix(flxSPayment.row, 3) = "" Then Exit Sub
   'added by anol for locking issue 749 will not be editable on double click
   selcol = flxSPayment.col
   flxSPayment.col = 0
   If flxSPayment.CellBackColor = vbRed Then
        'MsgBox "Selected invoice is locked by another user. Please wait untill other user release this record.", vbInformation, "Warning"
'        MsgBox "The selected invoice is currently locked by the user " & flxSPayment.TextMatrix(flxSPayment.row, 27) & " on " & flxSPayment.TextMatrix(flxSPayment.row, 28) & " and cannot be edited." & _
'                "Please wait until it is released.", vbInformation, "Warning"
                
        MsgBox "The selected invoice is currently locked by '" & flxSPayment.TextMatrix(flxSPayment.row, 27) & _
                "' on '" & flxSPayment.TextMatrix(flxSPayment.row, 28) & "' in the '" & flxSPayment.TextMatrix(flxSPayment.row, 29) & "'" & vbCrLf & "" & _
                        "screen for the Client '" & flxSPayment.TextMatrix(flxSPayment.row, 30) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
        Exit Sub
   End If
   flxSPayment.col = selcol
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
      txtSPayment.SetFocus
   End If
   If flxSPayment.col = 22 Then
      iFlxSPayCol = 22
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtPayDt.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtPayDt.Top
      txtPayDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtPayDt.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtPayDt.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtPayDt.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      txtPayDt.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      txtPayDt.SetFocus
   End If
   If flxSPayment.col = 23 Then
      iFlxSPayCol = 23
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtPostingDate.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtPostingDate.Top
      txtPostingDate.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtPostingDate.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtPostingDate.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtPostingDate.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      txtPostingDate.Visible = True
      'line added by anol 30 Aug 2016
      SelTxtInCtrl txtPostingDate
      flxSPayment.ScrollBars = flexScrollBarNone
      txtPostingDate.SetFocus
   End If
If flxSPayment.col = 24 Then
      iFlxSPayCol = 24
      flxSPayment.col = iFlxSPayCol
      flxSPayment.row = StarFound

      txtRef.Top = flxSPayment.CellTop + flxSPayment.Top
      iTop = txtRef.Top
      txtRef.Left = flxSPayment.CellLeft + flxSPayment.Left
      iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtRef.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtRef.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtRef.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      txtRef.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      txtRef.SetFocus
   End If

   bSavedPayment = False
   iCurRow = flxSPayment.row

   SelTxtInCtrl txtSPayment
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
   szHeader$ = "|<SupplierID|<SupplierName|<SupplierPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 0          'Solid column
   conFlxGrid.ColWidth(1) = 1000       'Supplier ID
   conFlxGrid.ColWidth(2) = 3000       'Supplier Name
   conFlxGrid.ColWidth(3) = 1100       'Post Code
   conFlxGrid.Rows = 2

   conFlxGrid.RowHeight(0) = 0
End Sub

Private Sub flxSPayment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      flxSPayment_dblClick
   End If
End Sub

Private Sub flxSPayment_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 33 Then
'      MsgBox ""
'   End If
End Sub

Public Sub Testing_Method()
   cmdSelAll_Click
   cboFund.ListIndex = 1
   cmdGPayment_Click
End Sub

Private Sub flxSPayment_RowColChange() 'this procedure is written by anol 20190412 for instant unlocking all rows
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
               "from tlbPayment as Pt  where  (UserSessionID='' or isnull(UserSessionID='')) AND TransactionID in (" & colTransactionIDOther & ")"
        rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
        
        While Not rsLockDialog.EOF
                flxSPayment.col = 0
                For i = 1 To flxSPayment.Rows - 1
                    If flxSPayment.TextMatrix(i, 20) = rsLockDialog("transactionID").Value Then
                          flxSPayment.row = i
                          flxSPayment.CellBackColor = vbWhite
                          'now you need to lock it for this screen
                           colTransactionIDHere = colTransactionIDHere & flxSPayment.TextMatrix(i, 20) & ","
                           flxSPayment.TextMatrix(i, 26) = "" 'we are not loading sessionID in this column for current screen lock
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
            adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Batch Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
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
    If frmBPPreForm.cmdACType.Value = "Supplier" Then
        StrWhere2 = " AND S.Type = 'SUPPLIER'"
    End If
    If frmBPPreForm.cmdACType.Value = "Client" Then
         StrWhere2 = " AND S.Type = 'CLIENT'"
    End If
    If frmBPPreForm.cmdACType.Value = "Managing Agent" Then
         StrWhere2 = " AND S.Type = 'AGENT'"
    End If
    If frmBPPreForm.cmdACType.Value = "Landlord" Then
         StrWhere2 = " AND S.Type = 'LLORD'"
    End If
    StrWhere2 = StrWhere2 & " AND Pt.ClientID='" & frmBPPreForm.txtClient.Tag & "' "
    strSQL = "Select DateTimeStamp ,Module ,Pt.ClientID,MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber AS INV,UserSessionID,WindowsUserName,MachineName,PrestigeUserName,ServerIPaddress " & _
            "from (tlbPayment as Pt INNER JOIN  tlbTransactionTypes AS TT ON Pt.Type = TT.TYPE_ID) INNER JOIN Supplier AS S ON  " & _
            "Pt.SageAccountNumber = S.SupplierID where UserSessionID<>'" & UserSessionID & "' AND DateTimeStamp<>'' AND " & _
            "Pt.OSAmount>0 AND S.PaymentType = '" & IIf(frmBPPreForm.optBP_BACS.Value, "BACS", "CHQ") & "' " & StrWhere2 & _
             "group by DateTimeStamp ,Module ,Pt.ClientID,SlNumber,UserSessionID,WindowsUserName,MachineName,PrestigeUserName," & _
             "ServerIPaddress,MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber order by MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber"
             'rsLockDialog.Close
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
'    Call ErrorlogNoPaymentsplitFound(adoconn)
    adoConn.Close
    Set adoConn = Nothing
    
End Sub
'Private Sub ErrorlogNoPaymentsplitFound(adoconn As ADODB.Connection)
'    'we dont' Nedd this function we are writing split on progrma startup when this condition arise
'    'SELECT tlbPayment.TransactionID, tlbPaymentSplit.PayHeader FROM tlbPayment Left JOIN tlbPaymentsplit ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID ;
'    Dim rsNopaymentsplit As New ADODB.Recordset
'    Dim rsNopayment As New ADODB.Recordset
'    rsNopaymentsplit.Open "select count(*) as cnt from (SELECT tlbPayment.TransactionID " & _
'                           "FROM tlbPaymentSplit right JOIN tlbPayment ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID where  tlbPayment.TransactionID is null);", adoconn, adOpenStatic, adLockReadOnly
'
'    If rsNopaymentsplit("cnt").Value > 0 Then
'        rsNopayment.Open "SELECT tlbPayment.TransactionID " & _
'                           "FROM tlbPaymentSplit right JOIN tlbPayment ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID where  tlbPayment.TransactionID is null", adoconn, adOpenStatic, adLockReadOnly
'        adoconn.Execute "Insert into SpareTable5 values(ClientID,Code,CC) values('Batch Payment'," & Now & " ,'There is no split for some transactions' & SQL2String(rsNopayment,0))"
'        rsNopayment.Close
'        Set rsNopayment = Nothing
'    End If
'    rsNopaymentsplit.Close
'    Set rsNopaymentsplit = Nothing
'End Sub
Private Sub Form_Load()
   bBPPreForm = True
   UserSessionID = GetTimeStamp
   bMultiple = frmBPPreForm.optBP_MULT.Value
   'cascading the form
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   Me.Height = 7935
   Me.Width = 13245
'   Me.Top = 0
'   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   optBP_Cheque.BackColor = MODULEBACKCOLOR
   optBP_BACS.BackColor = MODULEBACKCOLOR
   optBP_MULT.BackColor = MODULEBACKCOLOR
   Frame5(5).Width = 2295
   bSavedPayment = False

   ConfigFlxSPayment

   adoConn.Open getConnectionString

'  Check the O/S of the payment split with the payment header
   CheckPaySpOS adoConn
'  Load all supplier's name in the dropdown menu
   LoadSupplier adoConn, cboSupplier_
'  Load all transactions in the grid
   LoadFlxSPayment adoConn
'  Load saved data if any
   LoadLastSavedData adoConn
'   LoadDept adoConn
   SupplierAccountBalance adoConn

   adoConn.Close
   Set adoConn = Nothing

   SumUpTotal
   cOpeningBal = CCur(txtGrossTotal.text)
   Call updateBankBalance
   Call WheelHook(Me.hWnd)
   cmdGPayment.Refresh
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
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & Trim(frmBPPreForm.Label13(7).Caption) & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & frmBPPreForm.txtClient.Tag & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & Trim(frmBPPreForm.Label13(7).Caption) & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & frmBPPreForm.txtClient.Tag & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & Trim(frmBPPreForm.Label13(7).Caption) & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & frmBPPreForm.txtClient.Tag & "'   group by Type )"
                       
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
      txtBankBal1.text = IIf(IsNull(adoRst.Fields.Item("AMTT").Value), 0, adoRst.Fields.Item("AMTT").Value)
      txtBankBal1.text = Format(txtBankBal1.text, "0.00")
   End If
   adoRst.Close
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where isDeleted=false and BankCode='" & Trim(frmBPPreForm.Label13(7).Caption) & "' and ClientID='" & frmBPPreForm.txtClient.Tag & "'"
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
Private Sub LoadLastSavedData(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset, adoBN As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, szDataPath As String

'  there might be only one record which 'Generated' value should be FALSE
   szSQL = "SELECT * FROM tblBatchPayment WHERE Generated = FALSE;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      bSavedPayment = True
      adoRst.Close
'field T.PostingDate and T.ref has been added by anol 22 Apr 2015 issue 547 ref is not saving
      szSQL = "SELECT B.*, T.TransactionID, T.PayAmt, T.PayDt,T.PostingDate,T.ref " & _
              "FROM tblBatchPayment AS B, tblBatchTransaction AS T " & _
              "WHERE B.BP = T.BP AND B.Generated = FALSE AND T.PayAmt > 0;"
'Debug.Print szSQL
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         txtChqNo.text = adoRst.Fields.Item("BatchNo").Value
         For iRow = 1 To flxSPayment.Rows - 1
            If flxSPayment.TextMatrix(iRow, 20) = adoRst.Fields.Item("TransactionID").Value Then
               flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("PayAmt").Value, "0.00")
                    If bMultiple Then
                            flxSPayment.TextMatrix(iRow, 22) = IIf(IsNull(adoRst.Fields.Item("PayDt").Value), "", _
                                                                   Format(adoRst.Fields.Item("PayDt").Value, "dd/mm/yyyy"))
                            'Below 2 line has been added by anol 22 Apr 2015 issue 547 ref is not saving
                            flxSPayment.TextMatrix(iRow, 23) = IIf(IsNull(adoRst.Fields.Item("PostingDate").Value), "", _
                                                                   Format(adoRst.Fields.Item("PostingDate").Value, "dd/mm/yyyy"))
                            flxSPayment.TextMatrix(iRow, 24) = IIf(IsNull(adoRst.Fields.Item("ref").Value), "", _
                                                                   adoRst.Fields.Item("ref").Value)
                    End If
            End If
         Next iRow
         adoRst.MoveNext
      Wend
   Else
      szSQL = "SELECT B.BatchNo " & _
              "FROM tblBatchPayment AS B;"
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

   On Error Resume Next

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

Private Sub LoadFlxSPayment(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset, rdoSplits As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, szDataPath As String
   Dim bSwitch As Boolean, iSupp As Integer, iKount As Integer
   Dim rsUserSessionID As String ' This variable shall store the value of timestamp from the recordset
   Dim colTransactionID As String 'this variable shall hold all the locked transaction number which i s locked by this screen
    'Resolved by BOSL
    'Issue number 0000447
    'Batch Payments - Incorrect Filltering
    'Modified my Anol 05-08-2014
    Dim strWhere As String
    Dim StrWhere2 As String
     If vClientName <> "All Clients" Then
        strWhere = " And Client.ClientName= '" & vClientName & "' "
    End If
    If vPropertyName <> "All Properties" Then
        strWhere = strWhere & " And Property.PropertyName= '" & vPropertyName & "' "
    End If
    If frmBPPreForm.cmdACType.Value = "Supplier" Then
        StrWhere2 = " AND S.Type = 'SUPPLIER' "
    End If
    If frmBPPreForm.cmdACType.Value = "Client" Then
         StrWhere2 = " AND S.Type = 'CLIENT' "
    End If
    If frmBPPreForm.cmdACType.Value = "Managing Agent" Then
         StrWhere2 = " AND S.Type = 'AGENT' "
    End If
    If frmBPPreForm.cmdACType.Value = "Landlord" Then
         StrWhere2 = " AND S.Type = 'LLORD' "
    End If
  'Check If any record is locked by other screen/user, Mark that as red so that user cannot process
  
   colTransactionIDOther = ""
  
 
  
    
 szSQL = "SELECT Pt.TransactionID,PI.isRentPayable, PI.SlNumber AS D_SL, Pt.SlNumber AS C_SL, Pt.PI, Pt.AdjTag, " & _
                  "Pt.SageAccountNumber, Pt.UnitID, Pt.DDate, Pt.Ref,Pt.ExtRef, Pt.Details, Pt.Amount, " & _
                  "Pt.OSAmount, Pt.Type, TT.DESCRIPTION, S.SupplierName, S.SupplierID, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF,Pt.ClientID,UserSessionID,WindowsUserName,MachineName,Module " & _
               "FROM  ((((tlbPayment AS Pt INNER JOIN tlbTransactionTypes AS TT ON Pt.Type = TT.TYPE_ID)" & _
               "LEFT JOIN tblPurInv AS PI ON Pt.PI = PI.MY_ID)" & _
                "INNER JOIN Supplier AS S ON Pt.SageAccountNumber = S.SupplierID) LEFT JOIN Client  " & _
                "ON Client.ClientID = PI.CL_ID) LEFT JOIN Property ON PI.PropertyID =Property.PropertyID " & _
               "WHERE " & _
                  "Pt.PaymentView=True AND Pt.OSAmount>0 AND TT.TYPE_ID=6 AND " & _
                  "S.PaymentType = '" & IIf(frmBPPreForm.optBP_BACS.Value, "BACS", "CHQ") & "' " & _
                   strWhere & StrWhere2 & "ORDER BY S.SupplierID, Pt.TransactionID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'  Save the first Supp ID. Look at below
   If Not adoRst.EOF Then
      ReDim Preserve szaSuppID(1) As String
      szaSuppID(0) = adoRst!SupplierID
   End If

   iRow = 1
   While Not adoRst.EOF
      flxSPayment.TextMatrix(iRow, 20) = adoRst!TransactionID
      If adoRst.Fields.Item("Type").Value = 6 Or adoRst.Fields.Item("Type").Value = 24 Then
         flxSPayment.TextMatrix(iRow, 1) = adoRst!PF & adoRst!D_SL
      End If

      If adoRst.Fields.Item("Type").Value = 7 Or adoRst.Fields.Item("Type").Value = 9 Then
         flxSPayment.row = iRow
         For iKount = 2 To 11
            flxSPayment.col = iKount
            flxSPayment.CellForeColor = vbRed
         Next iKount
         flxSPayment.TextMatrix(iRow, 1) = adoRst!PF & IIf(IsNull(adoRst!C_SL), "", adoRst!C_SL)
      End If
      flxSPayment.TextMatrix(iRow, 2) = adoRst!SupplierName
      If InStr(adoRst!description, "Invoice") > 0 Then
         flxSPayment.TextMatrix(iRow, 3) = IIf(adoRst!AdjTag = "Y", "ADJI", adoRst!description)
      Else
         flxSPayment.TextMatrix(iRow, 3) = adoRst!description
      End If

      flxSPayment.TextMatrix(iRow, 4) = adoRst!SageAccountNumber
      flxSPayment.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!unitid), "", adoRst!unitid)
      flxSPayment.TextMatrix(iRow, 6) = IIf(Not IsNull(adoRst!dDate), Format(adoRst!dDate, "dd/mm/yyyy"), "")
      flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!ref), "", adoRst!ref)
      flxSPayment.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!Details), "", adoRst!Details)
      flxSPayment.TextMatrix(iRow, 9) = Format(adoRst!amount, "0.00")
      flxSPayment.TextMatrix(iRow, 10) = Format(adoRst!OSAmount, "0.00")
      flxSPayment.TextMatrix(iRow, 11) = "0.00"
      flxSPayment.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!Pi), "", adoRst!Pi)
      flxSPayment.TextMatrix(iRow, 15) = adoRst!Type
      flxSPayment.TextMatrix(iRow, 21) = adoRst!SupplierID
      'Below line was added by anol 22 Apr 2015
      'issue 547 Reference is not saving
      If bMultiple Then
         ' flxSPayment.TextMatrix(iRow, 24) = IIf(IsNull(adoRst!extref), "", adoRst!extref)
      End If
      rsUserSessionID = IIf(IsNull(adoRst!UserSessionID), "", adoRst!UserSessionID)
      flxSPayment.TextMatrix(iRow, 27) = IIf(IsNull(adoRst!WindowsUserName), "", adoRst!WindowsUserName)
      flxSPayment.TextMatrix(iRow, 28) = IIf(IsNull(adoRst!MachineName), "", adoRst!MachineName)
      flxSPayment.TextMatrix(iRow, 29) = IIf(IsNull(adoRst!Module), "", adoRst!Module)
      flxSPayment.TextMatrix(iRow, 30) = IIf(IsNull(adoRst!ClientID), "", adoRst!ClientID)
      flxSPayment.TextMatrix(iRow, 31) = IIf(IsNull(adoRst!isRentPayable), "", adoRst!isRentPayable)
      If Len(rsUserSessionID) > 0 And rsUserSessionID <> UserSessionID Then 'this means it is locked by other screen and now mark it red
             flxSPayment.col = 0
             flxSPayment.row = iRow
             flxSPayment.CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
             flxSPayment.TextMatrix(iRow, 26) = IIf(IsNull(adoRst!UserSessionID), "", adoRst!UserSessionID) 'Keeping the USersesssionID to check the lock
             colTransactionIDOther = colTransactionIDOther & flxSPayment.TextMatrix(iRow, 20) & ","
      Else 'collect the transaction ID which needs to be locked
             colTransactionID = colTransactionID & flxSPayment.TextMatrix(iRow, 20) & ","
      End If
      adoRst.MoveNext
      If Not adoRst.EOF Then
         flxSPayment.AddItem ""
         iRow = iRow + 1

'  if the next supp ID does not match with the first supp ID then it will take a copy of the supp.
         If adoRst!SupplierID <> flxSPayment.TextMatrix(iRow - 1, 21) Then
            bSwitch = Not bSwitch
            iSupp = iSupp + 1
            szaSuppID(iSupp) = adoRst!SupplierID
            ReDim Preserve szaSuppID(UBound(szaSuppID) + 1)
         End If

         If bSwitch Then UMarkRowFlxGrid flxSPayment, iRow 'And flxSPayment.CellBackColor <> RGB(255, 255, 153)
      End If
   Wend
   Call ChecksumValidationOnLoad(adoConn)
   adoRst.Close
   Set adoRst = Nothing
    'lock this records when it is not red written by anol 20190406 Issue 749
    If Len(colTransactionIDOther) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOther = Left(colTransactionIDOther, Len(colTransactionIDOther) - 1)
    End If
                   
    If Len(colTransactionID) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionID = Left(colTransactionID, Len(colTransactionID) - 1)
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Batch Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                   "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionID & ")"
    End If
     
                   
    'rsUserSessionID <> UserSessionID
End Sub
Private Sub FixType1Problem(adoConn As ADODB.Connection, TransactionID As String)
     Dim rsChecksum As New ADODB.Recordset
     Dim rsPIAmount As New ADODB.Recordset
     Dim PIAmount As Double
     Dim PaymentSum As Double
     'Note: This procedure is updating OSamount of PI which has the problem: 'not updating osAmount of PI'
     'If no payment is found , program will not come to this far, because on calling I have a inner  join with paytrans
    'Note: Pay transaction table fromtran contains payment information and totran contains PI Invoice information.
    'there can be many payment(fromtrans) agains invoice(ToTrans)
    'Do you have entry in the tlbpayment table for payment for this PI transaction ID?
    'Now there may have many payment agains an invoice so you have to sum the paid amount and match it with invoice amount minus outstanding amount
    
     rsChecksum.Open "Select (Sum(amount)-sum(osamount)) as TAmount From tlbPayment P Inner Join PayTransactions T " & _
                    "on P.TransactionID=T.FromTran where T.ToTran=" & TransactionID & "", adoConn, adOpenStatic, adLockReadOnly
                    'Here TAmount shall contains actual paid against the invoice which came came in parameter by invoice id
     If Not rsChecksum.EOF Then
            PaymentSum = IIf(IsNull(rsChecksum("TAmount").Value), 0, rsChecksum("TAmount").Value)
     End If
     rsChecksum.Close
     rsPIAmount.Open "Select amount from tlbPayment where TransactionId=" & TransactionID & "", adoConn, adOpenStatic, adLockReadOnly
     If Not rsPIAmount.EOF Then
        PIAmount = IIf(IsNull(rsPIAmount("amount").Value), 0, rsPIAmount("amount").Value)
     End If
     rsPIAmount.Close
     adoConn.Execute "Update tlbPayment P set osamount = " & (PIAmount - PaymentSum) & " where P.TransactionID=" & TransactionID & " AND osamount<=amount"
     adoConn.Execute "Update tlbPaymentSplit P set osamount = " & (PIAmount - PaymentSum) & " where P.PayHeader=" & TransactionID & " AND osamount<=amount"
     adoConn.Execute "Insert into SpareTable5(ClientID,Code,CC) values('B.Payment','" & Date & "' ,'Updating PI osamount with " & PaymentSum & " TransactionID: " & TransactionID & "  ' )"
     
End Sub

Private Function ChecksumValidationOnLoad(adoConn As ADODB.Connection) As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    'this function is written by anol 20181205 when found that (issue 695 )Updated OS amount  extra incorrectly
    'This function shall prevent saving the data if when outstading amount on payment is not updated.
    'This functionshall compare allocation with receipt payment and outstanding amount
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    Dim j As Integer
    Dim strWhere As String
    Dim strTenantID As String
    Dim hasFixed As Boolean
    Dim temp
        'Note: Pay transaction table fromtran contains payment information and totran contains PI Invoice information.
        strWhere = " AND R.ClientID='" & frmBPPreForm.txtClient.Tag & "'"
'        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
'                        "Where a.ToTran = r.TransactionID  " & StrWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
         
         'In a case, I have found the os-amount of PI is not updated.But there is an entry in the payment and allocation table.
         'For that case I am manually fixing the transaction which I am calling type1 problem
         
        rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
                  " ToTran from PayTransactions  where DeleteFlag=False group By ToTran ) as A where A.ToTran=R.transactionID " & strWhere & " AND round((amount-amt),2)<>round(osamount,2)", adoConn, adOpenStatic, adLockReadOnly
        While Not rsChecksum.EOF
            Call FixType1Problem(adoConn, CStr(rsChecksum("transactionID").Value))
            rsChecksum.MoveNext
            hasFixed = True
        Wend
        rsChecksum.Close
        If hasFixed = True Then
            Call LoadFlxSPayment(adoConn)
        End If
        rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
                  " ToTran from PayTransactions where DeleteFlag=False group By ToTran ) as A where A.ToTran=R.transactionID " & strWhere & " AND round((amount-amt),2)<>round(osamount,2)", adoConn, adOpenStatic, adLockReadOnly
        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            strTenantID = strTenantID + IIf(strTenantID = "", "", "/") + rsChecksum("sageaccountnumber").Value
            i = i + 1
            rsChecksum.MoveNext
        Wend
        
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
'         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
'                "A.Totran=R.transactionID where A.Totran IS NULL " & StrWhere & " AND osamount>amount", adoConn, adOpenKeyset, adLockReadOnly
'        While Not adoRst.EOF
'             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
'             strTenantID = strTenantID + IIf(strTenantID = "", "", "/") + adoRst("sageaccountnumber").Value
'               i = i + 1
'            rsChecksum.MoveNext
'             adoRst.MoveNext
'        Wend
'        adoRst.Close
'        Set adoRst = Nothing
        If szTran2Fix = "" Then
                ChecksumValidationOnLoad = True
        Else
                'ReDim ChecksumTenant(i) As String
                temp = Split(strTenantID, "/")
                For i = 0 To UBound(temp)
                    ' ChecksumTenant(i) = temp(i)
                     For j = 0 To flxSPayment.Rows - 1
                            If flxSPayment.TextMatrix(j, 2) = temp(i) Then
                                flxSPayment.TextMatrix(j, 25) = temp(i)
                            End If
                     Next j
                Next i
                MsgBox " A problem exists relating to a previous transaction entered against a Supplier in this Batch: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Please contact PCM Consulting.", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"

        End If

                    
End Function
Private Function CalculateSupplierBalance() As Boolean
   Dim iRow As Integer
   Dim i As Integer, cTotalPay As Currency

   CalculateSupplierBalance = False

   For i = 1 To flxSPayment.Rows - 2
      If flxSPayment.TextMatrix(iCurRow, 21) = flxSPayment.TextMatrix(i, 21) Then
         If flxSPayment.TextMatrix(i, 15) = 7 Or flxSPayment.TextMatrix(i, 15) = 9 Then
            cTotalPay = cTotalPay - flxSPayment.TextMatrix(i, 11)
         Else
            cTotalPay = cTotalPay + flxSPayment.TextMatrix(i, 11)
         End If
      End If
   Next i

   If cTotalPay <= 0 Then
      MsgBox "Accumulated payment value of this supplier is £" & cTotalPay & "" & Chr(13) & _
             "Payment value must be positive.", vbCritical + vbOKOnly, "Batch Payment"
      flxSPayment.TextMatrix(iCurRow, 11) = "0.00"
   End If

   CalculateSupplierBalance = True
End Function
Private Sub UpdateOverDraftStatus(adoConn As ADODB.Connection, szBank As String, szClientID As String)
      Dim rsClientBank As New ADODB.Recordset
      rsClientBank.Open "Select AllowOverDraft,OverDraftLimit from tlbClientBanks where NominalCode='" & szBank & "' and Client_ID='" & szClientID & "'", adoConn, adOpenStatic, adLockReadOnly
      If Not rsClientBank.EOF Then
            frmBPPreForm.isOverDratAllowed.Caption = rsClientBank("AllowOverDraft").Value
            frmBPPreForm.lblOverDraftAmount.Caption = IIf(IsNull(rsClientBank("OverDraftLimit").Value), "0", rsClientBank("OverDraftLimit").Value)
       End If
       rsClientBank.Close
       Set rsClientBank = Nothing
End Sub
Private Function DeleteALLFiles(strFolder As String)
    Dim oFs As New FileSystemObject
    Dim oFolder As Folder
    Dim oFile As file
    
    If oFs.FolderExists(strFolder) Then
        Set oFolder = oFs.GetFolder(strFolder)
    
        'caution!
        On Error Resume Next
    
        For Each oFile In oFolder.Files
            oFile.Delete True 'setting force to true will delete a read-only file
        Next
    
        DeleteALLFiles = oFolder.Files.Count = 0
    End If
    

End Function
Private Sub cmdGPayment_Click()
'~#~#~#~#~#~#~#      WARNINGS   ~#~#~#~#~#~#~#
'     DO NOT OPEN THE RELAVENT DATABASE WHILE YOU ARE TESTING THE BATCH PAYMENT
'----------------------------------------------
   Dim szOutPutFileLoc  As String
   Dim strProcessFileLocaton As String
   Dim adoConn    As New ADODB.Connection
   Dim reportApp  As New CRAXDRT.Application
   Dim Report     As CRAXDRT.Report
   Dim fso        As New Scripting.FileSystemObject
   Dim fso1        As New Scripting.FileSystemObject
   Dim isBACSFileMove As Boolean
   Dim CRXFormulaFields As CRAXDRT.FormulaFieldDefinitions
   Dim CRXFormulaField As CRAXDRT.FormulaFieldDefinition

   Dim szaSuppliers() As String, szSuppliers As String
   Dim szaSuppEmail() As String, szSuppEmail As String
   Dim colAtt As New Collection
   'issue 521
   If ValidationPostingDate = False Then
        Exit Sub
   End If
   If adoConn.State = 0 Then
        adoConn.Open getConnectionString
   End If
   
   If bMultiple Then
      If Not CheckDataValidation(adoConn) Then
            adoConn.Close
            Exit Sub
      End If
   End If
   If cboSupplier_.text = "" Then
        adoConn.Close
        Exit Sub
   End If

   Dim iRow       As Integer
   Dim szSQL      As String
   Dim i          As Integer
   Dim szRecID    As String
   Dim szEB       As String
   Dim cTotalPI   As Currency
   Dim cTotalPC   As Currency
   Dim szTemp     As String
   Dim szFileName As String
   Dim szFileEtn  As String
   Dim bEmailSucc As Boolean
   Dim iEmailSent As Integer
   Dim iEmailFail As Integer
   Dim adoRst As New ADODB.Recordset
   'Dim iRow As Integer
   Dim rsBankCheck As New ADODB.Recordset
  

   If Val(txtGrossTotal.text) <= 0 Then
      adoConn.Close
      MsgBox "No payment has been made.", vbInformation + vbOKOnly, "Batch Payment"
      Exit Sub
   End If
   'Resolved By BOSL
   '0000525: Bank overdrawn message showing on transaction entry incorrectly
   'Overdraft Checking By anol 26 Apr 2015
   Dim dblBankBanlance As Double

   
   'frmBPPreForm.cmbBankAc.Column (0)

   dblBankBanlance = BankAccBalance(adoConn, frmBPPreForm.cmbBankAc.Column(0), frmBPPreForm.cmbBankAc.Column(9)) 'szBank As String, szClientID As String
    
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where  isDeleted=false and  BankCode='" & frmBPPreForm.cmbBankAc.Column(0) & "' " & _
            " and ClientID='" & frmBPPreForm.cmbBankAc.Column(9) & "' "
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        dblBankBanlance = dblBankBanlance - IIf(IsNull(adoRst.Fields.Item("DAmt").Value), 0, adoRst.Fields.Item("DAmt").Value)
    End If
    adoRst.Close
    
   Call UpdateOverDraftStatus(adoConn, frmBPPreForm.cmbBankAc.Column(0), frmBPPreForm.cmbBankAc.Column(9)) ''szBank As String, szClientID As String
   'adoConn.Close
   If dblBankBanlance - Val(txtGrossTotal.text) < 0 Then 'Account balance-Current Amount
          'Call UpdateOverDraftStatus(adoConn, frmBPPreForm.cmbBankAc.Column(0), frmBPPreForm.cmbBankAc.Column(9)) ''szBank As String, szClientID As String
          If frmBPPreForm.isOverDratAllowed.Caption = "True" Then
             If Val(frmBPPreForm.lblOverDraftAmount.Caption) > 0 Then
                If (dblBankBanlance - Val(txtGrossTotal.text)) * (-1) > Val(frmBPPreForm.lblOverDraftAmount.Caption) Then
                   If MsgBox("This Bank Account is over its overdraft limit. Do you wish to continue?", vbQuestion + vbYesNo, "Bank Overdrawn") = vbNo Then Exit Sub
                End If
             End If
          Else
             MsgBox "This Bank Account cannot go overdrawn", vbInformation + vbOKOnly, "Bank Overdraft"
             Exit Sub
          End If
   End If
  'End of modification
  
  'Here is the logic if you are selecting wrong bank account in which differes from Rentpayable bank account.
   'We should check from the client statement , the Bank code selected for this Rent Payable an d compare with selected bank account for paying current PI
   'written by anol 20230728
  For iRow = 1 To flxSPayment.Rows - 1
         If flxSPayment.TextMatrix(iRow, 31) = "" Then 'isrentpayable flag is  empty then do not check anything
         Else
                Debug.Print Val(flxSPayment.TextMatrix(iRow, 11))
                Debug.Print flxSPayment.TextMatrix(iRow, 31)
                    If flxSPayment.TextMatrix(iRow, 31) = True And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
                          szSQL = "SELECT R.BankCode FROM tblpurinv AS T INNER JOIN rentsummarystatement AS R ON " & _
                          "(T.inv_no = 'SS'& R.StatementID) WHERE MY_ID='" & flxSPayment.TextMatrix(iRow, 13) & "'"
                          rsBankCheck.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                          If RecordCount(rsBankCheck) > 0 Then
                              If rsBankCheck("BankCode").Value <> Trim(lblBC.Caption) Then
                                  rsBankCheck.Close
                                  adoConn.Close
                                  MsgBox "Please select a bank account that matches the bank account on the client statement used to generate this rent payable invoice", vbInformation, "Warning"
                                  Exit Sub
                              End If
                          End If
                     End If
          End If

 Next
  adoConn.Close
   If MsgBox("Do you wish to generate your batch payments now?", vbQuestion + vbYesNo, "Batch Payment") = vbNo Then Exit Sub
   
''rem by anol 20190410 issue 749 Locking mechanism implementation
''   If IsLoadedAndVisible("frmReport") Then
''      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
''      Exit Sub
''   End If
''   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
''   If FileExists(szTemp) Then
''      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
''      Exit Sub
''   End If

   adoConn.Open getConnectionString
   
   If optBP_BACS.Value Then      'BACS option is selected, BACS file is being created here
      szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn, strProcessFileLocaton)
'if there is file exists (pass parameter string folder location) then it return vbNUllstring-anol
' if folder does not exists then it returns vbnullstring-anol
'So I dont find any meaning of this  conditions
      'Dir$(szOutPutFileLoc)<>vbNullString has been replaced with following line
      If FolderExists(szOutPutFileLoc) = False Then
         MsgBox "The BACS file path does not exists.", vbCritical + vbOKOnly, "BACS File not found"
         ShowMsgInTaskBar "BACS run was not successful.", "Y", "N"
         GoTo CloseConnection
      End If
'      If FolderExists(strProcessFileLocaton) = False Then
'         MsgBox "The BACS '" & strProcessFileLocaton & "' path does not exists.", vbCritical + vbOKOnly, "BACS File not found"
'         ShowMsgInTaskBar "BACS run was not successful.", "Y", "N"
'         GoTo CloseConnection
'      End If
      
      szTemp = szOutPutFileLoc & "\" & szFileName & Mid(szFileEtn, 2)

      If Right(szOutPutFileLoc, 1) = "\" Then szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)

      szTemp = szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
      If Len(szTemp) = 2 Then GoTo CloseConnection
      snFileHandeling = 6
    'szTemp is the file name with full path ,Dir$(szTemp) shall return only file name - anol
      If Dir$(szTemp) <> "" And optBP_BACS.Value Then
         If MsgBox("Do you wish to create a new BACS output file?" & vbCrLf & _
                   "If you choose NO, the system will append these payments to the existing BACS file.", _
                   vbQuestion + vbYesNo + vbDefaultButton2, "BACS file") = vbNo Then
                   'Here  append file begin
                    If MsgBox("You have chosen to append these payments to the existing BACS file." & vbCrLf & _
                               "Please confirm again.", vbQuestion + vbYesNo, "BACS file") = vbYes Then
                       snFileHandeling = vbNo
'                       If MsgBox("Do you want to move this final BACS file to a specified '" & strProcessFileLocaton & "' for processing?", vbYesNo, "Please confirm.") = vbYes Then
'                            isBACSFileMove = True
'                       End If
            Else
                
                GoTo CloseConnection
            End If
         Else
            If MsgBox("You have chosen to create a new BACS file. Please confirm.", vbQuestion + vbYesNo, "BACS file") = vbNo Then
               GoTo CloseConnection
            Else
               Call AutoArchiveBACSFile
               If Dir$(szTemp) <> "" Then           'Sys checks 4 BACS file again. If file doesn't exist, sys will create BACS file
                  MsgBox "You must transmit your existing BACS file before you can create a new BACS file.", vbInformation + vbOKOnly, "BACS file"
                  GoTo CloseConnection
               End If
            End If
         End If
      Else
'            If MsgBox("Do you want to move this final BACS file to a specified '" & strProcessFileLocaton & "' for processing?", vbYesNo, "Please confirm.") = vbYes Then
'                   isBACSFileMove = True
'            End If
        
      End If
   End If                        'BACS option is selected - End IF

'  First save all payments in the batch payment table, which will help to print report
   If Not bSavedPayment Then
      szRecID = SaveSuggestedPayment(adoConn, True)   'Save in the tblBatchTransaction first.
   Else
      szRecID = SavedSuggestedPayment(adoConn)
   End If

'  IF  -->>  User's selection is CHEQUE
'  System will run the following procedure. System will generate a report with payment details.
'  there will be additional work if user selects BACS.

'  System will update the PI's os amount and book a payment.
'  If there is any allocation then system will run auto allocaion procedure.

'  FIRST STEP: Check is there any credit transaction needed to allocate against PI?
'      If 'no' then book normal PI.
'      If 'yes' then automatic allocation needed to run.
'  -------------------------------------------------------------------------------------------------------------------
   If NoCrTrn(szaSuppID(i)) Then 'Processing payment for each supplier
        adoConn.Close
        adoConn.Open getConnectionString
        adoConn.BeginTrans
        If (BookNormalPaymentOfPI(adoConn, szaSuppID(i))) = True Then ' Here the main Batch Payment function begins
            adoConn.CommitTrans
            frmMMain.frmSupplier_SupplierListBCL_isUptoDate = False
            frmMMain.frmSupplier_SupplierList_isUptoDate = False
        Else
            adoConn.RollbackTrans
            MsgBox "The batch payment for Supplier " & szaSuppID(i) & " could not be generated. Please check your payment information is correctly entered and try again.", vbInformation, "Unable to generate batch payment"
            Exit Sub
        End If
   Else
     'Disabled /rem by anol 2019 07 20
     ' AutoAlloc adoConn, szaSuppID(i)
   End If
   
'Resolved By BOSL. Modified by Asif.
'Issue:0000509. Date: 07 Dec 2014.
'This fix is brought from the Main form so that the payment is updated with correct fundid as soon
'the transaction is saved and before posting to NL.

'********************************************************************************************
'   FIXING DATA: There are some PI found which FundID don't match with split FundID
'********************************************************************************************
'FIXING_FUNDID_tlbPayment_tlbPaymentSplit:
   szSQL = "SELECT P.* " & _
           "FROM tlbPaymentSplit AS S, tlbPayment AS P " & _
           "WHERE P.TransactionID = S.PayHeader AND " & _
                 "P.FundID <> S.FundID AND S.Splitid = 1;"
                 
   Dim Rst1 As New ADODB.Recordset
   Rst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
'   Debug.Print szSQL
   If Not Rst1.EOF Then
    Debug.Print time
      adoConn.Execute "UPDATE tlbPayment AS P, tlbPaymentSplit AS S " & _
                    "SET    P.FundID = S.FundID " & _
                    "WHERE  P.TransactionID = S.PayHeader AND " & _
                           "S.Splitid = 1;"
    
    Debug.Print time
                           
   End If
   Rst1.Close
' END OF MODIFICATION


'--------------------------------------------------------------------------------------------
'  Export Transactions to Nominal Ledger (NLPosting table)
   Export_PPnPPR_2_NL adoConn
'--------------------------------------------------------------------------------------
'issue 523
'added by anol 20 Jan 2015
'modified by anol 13 Mar 2015
   
   UpdateBankAcBal_Minus adoConn, Val(txtGrossTotal.text), frmBPPreForm.cmbBankAc.Column(0), frmBPPreForm.txtClient.Tag

   adoConn.Execute "UPDATE tblBatchPayment SET Generated = TRUE;"

'  When user selects Cheque, they can select either 'Remittance Only' or 'Cheque with Remittance'
'  However if they select BACS, 'Remittance Only' is always true
''And optBP_MULT.Value = False option has been added by anol on 22 Apr 2015 as it does not make sense when it is multiple
   Call ChangeReportODBC
   If frmBPPreForm.optSCPO(0).Value And optBP_MULT.Value = False Then                              '****    REMITTANCE ONLY   ********
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RemittanceAdviceBACS.rpt")

      Set CRXFormulaFields = Report.FormulaFields

      Report.EnableParameterPrompting = False
      If Report.HasSavedData Then Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue szRecID

      Load frmReport
      frmReport.CRViewer91.DisplayBorder = False
      frmReport.CRViewer91.DisplayTabs = False
      frmReport.LoadReportViewer Report

'      CreateNonExistsFolder App.Path & "\Temp"
      iRow = TotalSuppliers(adoConn, szRecID, szSuppliers, szSuppEmail)
      szaSuppliers = Split(szSuppliers, ", ")
      szaSuppEmail = Split(szSuppEmail, ", ")

      If szFromEmail <> "" And szSMTPserver <> "" And szSuppEmail <> "" Then
         BACS_EmailText adoConn, szSubject, szBody, lblClient.Caption

         iEmailFail = 0
         iEmailSent = 0

         For i = 1 To iRow
            If szaSuppEmail(i - 1) <> "" Then
               szSQL = szaSuppliers(i - 1) & "_" & UniqueID() & ".pdf"
               Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
               Report.ExportOptions.DestinationType = crEDTDiskFile
               Report.ExportOptions.FormatType = crEFTPortableDocFormat
               Report.ExportOptions.PDFExportAllPages = False
               Report.ExportOptions.PDFFirstPageNumber = i
               Report.ExportOptions.PDFLastPageNumber = i

               Report.Export False

               If colAtt.Count > 0 Then colAtt.Remove (1)
               colAtt.Add DB_PATH & "\AllStuff\Temp\" & szSQL
               bEmailSucc = SendEmail(szFromEmail, Trim(szaSuppEmail(i - 1)), _
                             szSubject, _
                             szBody, , , _
                             colAtt, szaSuppliers(i - 1), "PI", lblClient.Caption)
               iEmailSent = iEmailSent + IIf(bEmailSucc, 1, 0)
               iEmailFail = iEmailFail + IIf(bEmailSucc, 0, 1)
            End If
         Next i
         If iEmailFail + iEmailSent > 0 Then
            MsgBox "Emails have been successfully sent:" & iEmailSent & " and failed:" & iEmailFail & ".", vbInformation + vbOKOnly, "Email Sent Notification"
         End If
      End If
   End If
'And optBP_MULT.Value = False option has been added by anol on 22 Apr 2015 'as it does not make sense when it is multiple
   If frmBPPreForm.optSCPO(1).Value And optBP_MULT.Value = False Then                              '****    CHEQUE REMITTANCE   *******
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ChequeRemittance.rpt")

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue szRecID

      Load frmReport
      frmReport.LoadReportViewer Report
   End If

'  When user selects BACS
   If optBP_BACS.Value Then
'        If isBACSFileMove = True Then
'                'Program will first copy BACS file to BACS archive location in AllStuff using the path AllStuff\BACS_Archive
'                If FolderExists(DB_PATH & "\AllStuff\BACS_Archive") = False Then
'                    CreateNonExistsFolder DB_PATH & "\AllStuff\BACS_Archive"
'                End If
'                szOutPutFileLoc = strProcessFileLocaton
'        End If
        'Call AutoArchiveBACSFile
        If CreateBACS(adoConn, szOutPutFileLoc, szEB, szRecID, szFileName, szFileEtn) = True Then
            If szEB = "2" Then
                 If MsgBox("Do you Plan to Append further Batch Payments to this BACS file?", vbYesNo, "Are you sure?") = vbNo Then
                    'If BACS File has been created then ask a question.
                          If Trim(strProcessFileLocaton) <> "" And szEB = 2 Then
                             If MsgBox("Do you wish to move this Final BACS file to the location specified '" & strProcessFileLocaton & "' for processing?", vbYesNo, "Please confirm.") = vbYes Then
                            ' If isBACSFileMove = True Then
                                    If szOutPutFileLoc = strProcessFileLocaton Then
                                          MsgBox "This file has not been moved because the 'unprocessed file' location is the same as the 'Process file' location", vbCritical + vbOKOnly, "Warning"
                                          Exit Sub
                                    End If
                             '7.  Then program will MOVE current BACS file to Process file location specified for the current bank account being processed with confirmation
                                 'of "process file location using the following message. "This BACS file has been copied to <<specified process file location path>> for processing"
                                 If FolderExists(strProcessFileLocaton) = False Then
                                         CreateNonExistsFolder strProcessFileLocaton
                                 End If
                                 ' fso1.CopyFile; szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3), strProcessFileLocaton
                                 ' fso1.DeleteFile; szOutPutFileLoc
                                 fso1.GetFolder(szOutPutFileLoc).Copy strProcessFileLocaton, True
                                ' fso1.DeleteFile szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
                                 'Call DeleteALLFiles(szOutPutFileLoc)
'                                 fso1.GetFolder(szOutPutFileLoc).Delete
                                 Set fso1 = Nothing

                                 'Program will first copy BACS file to BACS archive location in AllStuff using the path AllStuff\BACS_Archive
                                 If FolderExists(DB_PATH & "\AllStuff\BACS_Archive") = False Then
                                     CreateNonExistsFolder DB_PATH & "\AllStuff\BACS_Archive"
                                 End If
                                 'Call AutoArchiveBACSFile
                                 'Here is the code for old archive method(Working Code)
                                 'Here the program is moving only one file which has been metioned in the settings. I need to read all file from the folder and move them anol 04-25-2020
                                 i = 1
                                 While i > 0
                                   'here checking at archive location for a file before copy
                                   'Now salia has defined a new location for being archived
                                    If Dir$(DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(i) & "." & Mid(szFileEtn, 3)) = "" Then
                                       'if file info equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                                       'if actual file not exists in disk it shall return emptystring by DIR function
                                       fso.CopyFile szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3), DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(i) & "." & Mid(szFileEtn, 3)
                                       fso.DeleteFile szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
                                       Set fso = Nothing
                                       'szOutPutFileLoc value is U:\BACSFILES\Savoy Stewart\PTX\UNPROCESSED BACS FILES\436120.csv
                                       i = 0
                                    Else
                                       'if file info not equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                                       'if some file exists with the name it shall return the file name and it shall come here
                                       i = i + 1
                                    End If
                                 Wend
                                 Call DeleteALLFiles(szOutPutFileLoc)
                                 MsgBox "This BACS file has been moved to the location '" & strProcessFileLocaton & "' for processing", vbInformation + vbOKOnly, "File moved"
                             Else
                                 'Do nothing
                             End If
                        Else ' if ans is yes for the first question(Do you Plan to Append further Batch Payments to this BACS file?)
                            'do nothing
                        End If ' end if for the first question
               Else ' If process file location is empty  or has not been inputtted
                    'Call AutoArchiveBACSFile
                    'Do nothing
               End If ' end if for plan to append
            End If
        End If 'end if for create bacs file true
   End If 'end if for back opt true
   'Do not Clear the lock flag here as the screen is still showing some other values
  
                   
   ConfigFlxSPayment
   LoadFlxSPayment adoConn
   Frame5(5).Refresh
   cmdGPayment.Refresh

CloseConnection:
   adoConn.Close
   Set adoConn = Nothing
    MsgBox "Batch Payment has been processed successfully.", vbInformation, "processed"
End Sub
Private Sub AutoArchiveBACSFile()
        Dim i As Integer
        Dim iFileLoop As Integer
        Dim szEB As String
        Dim szOutPutFileLoc As String
        Dim iFilecount As Long
        Dim szFileName As String
        Dim adoConn As New ADODB.Connection
        Dim szFileEtn As String
        Dim FS As New FileSystemObject
        Dim FSfolder As Folder
        Dim file As file
        Dim szOutPutFilePath As String
        Dim strProcessFilelocation As String
        Dim fso As New Scripting.FileSystemObject
        adoConn.Open getConnectionString
        szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn, strProcessFilelocation)
        adoConn.Close
        Set adoConn = Nothing

        If Right(szOutPutFileLoc, 1) = "\" Then szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)
    
        iFilecount = -1
        If szOutPutFileLoc = "" Then Exit Sub
        Set FSfolder = FS.GetFolder(szOutPutFileLoc)
        For Each file In FSfolder.Files
                If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                       iFilecount = iFilecount + 1
                End If
        Next file
        If iFilecount = -1 Then
            
            Exit Sub
        Else
            
        End If
        Dim filearray() As String
        ReDim filearray(iFilecount) As String
        iFilecount = 0
        For Each file In FSfolder.Files
                If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                       filearray(iFilecount) = file
                       iFilecount = iFilecount + 1
                End If
        Next file
        Call CreateNonExistsFolder(DB_PATH & "\AllStuff\BACS_Archive")
        While i < iFilecount
                 iFileLoop = 1
                 While iFileLoop > 0
                     'here checking at archive location for a file before copy
                     'Now salia has defined a new location for being archived
                      szFileName = Dir$(filearray(i))
                      szOutPutFilePath = filearray(i)
                      szFileName = Left(szFileName, Len(szFileName) - 4)
                      
                      If Dir$(DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(iFileLoop) & "." & Mid(szFileEtn, 3)) = "" Then
                         'if file info equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                         'if actual file not exists in disk it shall return emptystring by DIR function
                         
                         If (Dir$(szOutPutFilePath) <> "") Then 'if source not found check it else u shall get an error
                              'fso.copy source, destination
                               fso.CopyFile szOutPutFilePath, DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(iFileLoop) & "." & Mid(szFileEtn, 3)
                               fso.DeleteFile szOutPutFilePath
                              'szOutPutFilePath value is U:\BACSFILES\Savoy Stewart\PTX\UNPROCESSED BACS FILES\436120.csv
                               iFileLoop = 0
                         End If
                      Else
                         'if file info not equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                         'if some file exists with the name it shall return the file name and it shall come here
                         iFileLoop = iFileLoop + 1
                      End If
                 Wend
                 i = i + 1
        Wend
End Sub
Private Function TotalSuppliers(adoConn As ADODB.Connection, _
                                szRecID As String, _
                                ByRef szSuppleirs As String, _
                                ByRef szSuppEmail As String) As Integer
   Dim szSQL As String, i As Integer
   Dim adoRst As New ADODB.Recordset

'   szSQL = "SELECT DISTINCT S.SupplierID, S.SupplierOfficeEmail " & _
           "From Supplier AS S, tblBatchTransaction AS BT " & _
           "Where S.SupplierOfficeEmail LIKE '%" & "@" & "%' AND " & _
               "S.SupplierOfficeEmail LIKE '%" & "." & "%' AND " & _
               "S.SupplierID = BT.SupplierID AND " & _
               "BT.BP = '" & szRecID & "' " & _
           "ORDER BY S.SupplierID;"
   szSQL = "SELECT DISTINCT S.SupplierID, S.SupplierOfficeEmail " & _
           "From Supplier AS S, tblBatchTransaction AS BT " & _
           "Where S.SupplierID = BT.SupplierID AND " & _
               "BT.BP = '" & szRecID & "' " & _
           "ORDER BY S.SupplierID;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalSuppliers = Val(adoRst.RecordCount)

   For i = 0 To Val(adoRst.RecordCount) - 1
      szSuppleirs = szSuppleirs & adoRst.Fields.Item(0).Value & ", "
      szSuppEmail = szSuppEmail & IIf(IsNull(adoRst.Fields.Item(1).Value), "", adoRst.Fields.Item(1).Value) & ", "
      adoRst.MoveNext
   Next i

   If i > 0 Then
      szSuppleirs = Left(szSuppleirs, Len(szSuppleirs) - 2)
      szSuppEmail = Left(szSuppEmail, Len(szSuppEmail) - 2)
   End If

   Set adoRst = Nothing
End Function
Private Function ReturnIDBACSPaymentRun(adoConn As ADODB.Connection) As Long
    Dim rsBACSPaymentRun As New ADODB.Recordset
    rsBACSPaymentRun.Open "Select RunNo from BACSPaymentRun order by RunNo desc", adoConn, adOpenKeyset, adLockReadOnly
    If rsBACSPaymentRun.EOF Then
        ReturnIDBACSPaymentRun = 1001
        rsBACSPaymentRun.Close
        Exit Function
    End If
    If Not rsBACSPaymentRun.EOF Then
        ReturnIDBACSPaymentRun = rsBACSPaymentRun("RunNo").Value + 1
        rsBACSPaymentRun.Close
    End If
    
End Function
Private Sub WriteBACSPaymentRun(adoConn As ADODB.Connection, szFileName As String, szEB As String)
    Dim rsBACSPaymentRun As New ADODB.Recordset
    Dim RunNo As Long
    RunNo = ReturnIDBACSPaymentRun(adoConn)
    rsBACSPaymentRun.Open "Select * from BACSPaymentRun where 1=2", adoConn, adOpenDynamic, adLockOptimistic
    
    Dim szLine As String
    Dim i As Integer
    Open szFileName For Input As #3
    i = 1
    While Not EOF(3)
        Line Input #3, szLine
        Debug.Print szLine
        With rsBACSPaymentRun
            .AddNew
            !RunNo = RunNo
            !RunDate = Date
            !LineNo = i
            !EB = szEB
            !description = szLine
            .Update
        End With
        i = i + 1
    Wend
    Close #3
    rsBACSPaymentRun.Close
    Set rsBACSPaymentRun = Nothing
End Sub

Private Function CreateBACS(adoConn As ADODB.Connection, szOutPutFileLoc As String, szEB As String, szRecID As String, szFileName As String, szFileEtn As String) As Boolean
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szLine As String, szaFormatLine() As String, szTemp As String
   Dim szFormatLine As String, szOutPutLine As String, szColHeading As String

   On Error GoTo ErrorHanlder
   If szEB = "1" Then
        szTemp = szOutPutFileLoc & "\" & szFileName & Mid(szFileEtn, 2)
   Else
        szTemp = szOutPutFileLoc & "\" & szFileName & Mid(szFileEtn, 2)
   End If

   If snFileHandeling = 6 Then Open szTemp For Output As #1
   If snFileHandeling = 7 Then Open szTemp For Append As #1

   FileFormat szFormatLine, szEB, szColHeading
   szaFormatLine = Split(szFormatLine, ", ")

   szSQL = "SELECT S.SupplierName AS NAME, SUM(BT.PayAmt) AS AMT, S.BPR AS REF, S.SortCode AS SC, S.AcNo AS AC, S.AcName " & _
           "FROM (tblBatchPayment AS BP INNER JOIN tblBatchTransaction AS BT ON BP.BP = BT.BP) INNER JOIN " & _
               "Supplier AS S ON BT.SupplierID = S.SupplierID " & _
           "WHERE BP.BP = '" & szRecID & "' " & _
           "GROUP BY BT.SupplierID, S.SupplierName, S.BPR, S.SortCode, S.AcNo, S.AcName;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      If IsNull(adoRst.Fields.Item("SC").Value) Or IsNull(adoRst.Fields.Item("AC").Value) Then
         Close #1
         adoRst.Close
         Set adoRst = Nothing
         Exit Function
      End If
   End If
   Dim strTemp As String
   If szEB = "1" Then                                    'Barclays.Net
     ' If snFileHandeling = 6 Then Print #1, szColHeading
      strTemp = """"
      While Not adoRst.EOF
'     Write into the file.
         ' col 0 is Sort Code and col2 is Account Number like a CSV file
         'new changes on 2021-12-05
          ' szOutPutLine = ""
         szOutPutLine = strTemp & Left(adoRst.Fields.Item(szaFormatLine(0)).Value, 6) & strTemp & ","
         szOutPutLine = szOutPutLine + strTemp & Left(adoRst.Fields.Item(szaFormatLine(1)).Value, 18) & strTemp & ","
         szOutPutLine = szOutPutLine + strTemp & Left(adoRst.Fields.Item(szaFormatLine(2)).Value, 8) & strTemp & ","
         szOutPutLine = szOutPutLine + strTemp & CStr(Format(adoRst.Fields.Item(szaFormatLine(3)).Value, "0.00")) & strTemp & ","
         'szOutPutLine = szOutPutLine + """IIf(IsNull(adoRst.Fields.Item(szaFormatLine(4)).Value), "", adoRst.Fields.Item(szaFormatLine(4)).Value) & "", ""99"""
         szOutPutLine = szOutPutLine + strTemp & Left(IIf(IsNull(adoRst.Fields.Item(szaFormatLine(4)).Value), "", adoRst.Fields.Item(szaFormatLine(4)).Value), 18) & strTemp & ","
         szOutPutLine = szOutPutLine + strTemp & "99" & strTemp
         'szOutPutLine = szOutPutLine + "" & 99 & ""

         Print #1, szOutPutLine
         adoRst.MoveNext
      Wend
   End If


   If szEB = "2" Then                                       'PTX BACS
      While Not adoRst.EOF
'     Write into the file.
         szOutPutLine = adoRst.Fields.Item("SC").Value                     'Destination sort code
         szOutPutLine = szOutPutLine + adoRst.Fields.Item("AC").Value      'Destination account number
         szOutPutLine = szOutPutLine + "0"                                 'Destination account type
         szOutPutLine = szOutPutLine + "99"                                'Transaction code
         szOutPutLine = szOutPutLine + frmBPPreForm.cmbBankAc.Column(4)    'Payee's sort code
         szOutPutLine = szOutPutLine + frmBPPreForm.cmbBankAc.Column(3)    'Payee's account number
         szOutPutLine = szOutPutLine + "    "                              'Free format
         szOutPutLine = szOutPutLine + Format(adoRst.Fields.Item("AMT").Value * 100, "00000000000")       'Amount, 11 char long
        'issue 547
        'Below line has been modified by anol 06 Apr 2015
         'szSQL = UCase(MakeFixedLenString(frmBPPreForm.cmbBankAc.Column(1), " ", 18))
         szSQL = UCase(MakeFixedLenString(frmBPPreForm.cmbBankAc.Column(6), " ", 18))
         szOutPutLine = szOutPutLine + szSQL                                'Payee's account name

'         If frmBPPreForm.cmbBankAc.Column(5) = "" Then                                                    'Reference
         szOutPutLine = szOutPutLine + UCase(MakeFixedLenString(IIf(IsNull(adoRst.Fields.Item("REF").Value), "", adoRst.Fields.Item("REF").Value), " ", 18))
'         Else
'            szOutPutLine = szOutPutLine + UCase(MakeFixedLenString(IIf(IsNull(frmBPPreForm.cmbBankAc.Column(5)), "", frmBPPreForm.cmbBankAc.Column(5)), " ", 18))
'         End If
         szOutPutLine = szOutPutLine + UCase(MakeFixedLenString(adoRst.Fields.Item("AcName").Value, " ", 18)) 'Destination account name

         Print #1, szOutPutLine
         adoRst.MoveNext
      Wend
   End If
   If szEB = "3" Then                                    'Natwest Bankline
      If snFileHandeling = 6 Then Print #1, szColHeading

      While Not adoRst.EOF
'     Write into the file.
         ' col 0 is Sort Code and col2 is Account Number like a CSV file
         szOutPutLine = adoRst.Fields.Item(szaFormatLine(0)).Value & ","
         szOutPutLine = szOutPutLine + adoRst.Fields.Item(szaFormatLine(1)).Value & ","
         szOutPutLine = szOutPutLine + adoRst.Fields.Item(szaFormatLine(2)).Value & ","
         szOutPutLine = szOutPutLine + CStr(adoRst.Fields.Item(szaFormatLine(3)).Value) & ","
         szOutPutLine = szOutPutLine + IIf(IsNull(adoRst.Fields.Item(szaFormatLine(4)).Value), _
                                           "", adoRst.Fields.Item(szaFormatLine(4)).Value) & ",99"

         Print #1, szOutPutLine
         adoRst.MoveNext
      Wend
   End If
   Close #1
   adoRst.Close
   Set adoRst = Nothing
   Call WriteBACSPaymentRun(adoConn, szTemp, szEB)
   MsgBox "The BACS file has been successfully generated.", vbInformation + vbOKOnly, "Batch Payment"
   CreateBACS = True
   Exit Function

ErrorHanlder:
   MsgBox Err.description & " : " & Err.Number, vbCritical + vbOKOnly, "BACS run. Path Not found while export: " & szOutPutFileLoc
End Function

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

Private Function BACS_OPFLocation(adoConn As ADODB.Connection, ByRef szEB As String, ByRef szFileName As String, ByRef szFileEtn As String, ByRef strProcessFileLocaton As String) As String
   On Error GoTo ERR_HANDLER

   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim strDate As String
   Dim intFileNo As Integer

   szSQL = "SELECT FileLoc, EB, Indentifier,BANK_AC_NUM,C.ClientName, FileExten,ProcessFileLoc,LASTBACSFDATE,LASTBACSFNO " & _
           "FROM tlbClientBanks B,Client C " & _
           "WHERE C.ClientID=B.Client_ID AND MY_ID = " & frmBPPreForm.cmbBankAc.Column(2) & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If IsNull(adoRst.Fields.Item("EB").Value) Then
      MsgBox "Client's e-banking details are not updated.", vbCritical + vbOKOnly, "Batch Payment"
      adoRst.Clone
      Set adoRst = Nothing
      Exit Function
   End If
   szEB = adoRst.Fields.Item("EB").Value
   If szEB = "1" Then
        If Format(Date, "ddMMyyyy") = Format(adoRst.Fields.Item("LASTBACSFDATE").Value, "ddMMyyyy") Then
                intFileNo = IIf(IsNull(adoRst.Fields.Item("LASTBACSFNO").Value), "0", adoRst.Fields.Item("LASTBACSFNO").Value)
        Else
                intFileNo = "0"
        End If
        intFileNo = intFileNo + 1
        szFileName = adoRst.Fields.Item("Indentifier").Value & "_" & adoRst.Fields.Item("ClientName").Value & "_" & adoRst.Fields.Item("BANK_AC_NUM").Value & "_" & Format(Date, "ddMMyyyy") & "_" & intFileNo
        szFileEtn = "*.csv"
        adoConn.Execute "Update tlbClientBanks set LASTBACSFNO=" & intFileNo & ",LASTBACSFDATE=#" & Format(Date, "dd MMM yyyy") & "# where  MY_ID = " & frmBPPreForm.cmbBankAc.Column(2) & ";"
   Else
        szFileName = adoRst.Fields.Item("Indentifier").Value
        szFileEtn = adoRst.Fields.Item("FileExten").Value
   End If
   BACS_OPFLocation = adoRst.Fields.Item("FileLoc").Value
   strProcessFileLocaton = adoRst.Fields.Item("ProcessFileLoc").Value

   adoRst.Close
   Set adoRst = Nothing
   Exit Function

ERR_HANDLER:
   
   Set adoRst = Nothing
End Function
'
'Private Sub ClearSuggestedPayment(adoConn As ADODB.Connection)
'   adoConn.Execute "UPDATE tblBatchPayment SET Generated = TRUE;"
'End Sub

Private Sub AutoAlloc(adoConn As ADODB.Connection, szSupplier As String)
   On Error GoTo Err
   Dim i As Integer, j As String, cTotalPay As Currency

   SeperateDrCr szSupplier

   cTotalPay = 0
   For i = 0 To UBound(iaPI_RowNo) - 1
      cTotalPay = cTotalPay + CCur(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))
   Next i

   j = 0
   For i = 0 To UBound(iaPI_RowNo) - 1
      If Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) = Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Simply allocate Cr -> PI
         Book1PaymentOfPI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))

         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))

         flxSPayment.TextMatrix(iaPC_RowNo(j), 11) = Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))
         j = j + 1
      ElseIf Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) > Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Simply allocate Cr -> PI
'        PI = PI - Cr
         Book1PaymentOfPI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))
         flxSPayment.TextMatrix(iaPI_RowNo(i), 11) = Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))
         i = i - 1
         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11))
         j = j + 1
      ElseIf Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11)) < Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) Then
'        Cr_ = PI
'        Simply allocate Cr_ -> PI
'        Cr = Cr - Cr_
         Book1PaymentOfPI adoConn, iaPI_RowNo(i), iaPC_RowNo(j), Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))
         flxSPayment.TextMatrix(iaPC_RowNo(j), 11) = Val(flxSPayment.TextMatrix(iaPC_RowNo(j), 11)) - _
                                                     Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))

         cTotalPay = cTotalPay - Val(flxSPayment.TextMatrix(iaPI_RowNo(i), 11))

      End If
      If j > UBound(iaPC_RowNo) - 1 Then Exit For
   Next i

   If cTotalPay > 0 Then BookNormalPaymentOfPI adoConn, szSupplier
   Exit Sub
Err:
   MsgBox Err.description, vbInformation, "Error from Auto Alloc,Please contact PCM consulting LTD."
End Sub

Private Sub Book1PaymentOfPI(adoConn As ADODB.Connection, iFlxRowPI As Integer, iFlxRowPC As Integer, cAmt As Currency)
   Dim iRow  As Integer, szSQL  As String
   Dim lRT_ID As Long
   Dim rstRst As New ADODB.Recordset

   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM PayTransactions;"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
   rstRst.Close

'  Update the invoice out-standing amount by allocated amount
   szSQL = "UPDATE tlbPayment " & _
           "SET OSAmount = " & CCur(flxSPayment.TextMatrix(iFlxRowPI, 10)) - CCur(flxSPayment.TextMatrix(iFlxRowPI, 11)) & ", " & _
               "PaymentView = IIF(OSAmount > 0, TRUE, FALSE) " & _
           "WHERE TransactionID = " & CLng(flxSPayment.TextMatrix(iFlxRowPI, 20)) & ";"
   adoConn.Execute szSQL

'  Update the credit Out-Standing amount by receipt amount
   flxSPayment.TextMatrix(iFlxRowPC, 10) = CCur(flxSPayment.TextMatrix(iFlxRowPC, 10)) - cAmt

   szSQL = "UPDATE tlbPayment " & _
           "SET OSAmount = " & CCur(flxSPayment.TextMatrix(iFlxRowPC, 10)) & ", " & _
               "PaymentView = IIF(OSAmount > 0, TRUE, FALSE) " & _
           "WHERE TransactionID = " & CLng(flxSPayment.TextMatrix(iFlxRowPC, 20)) & ";"
   adoConn.Execute szSQL
   
'  Update the Invoice out standing amount by receipt amount
   szSQL = "SELECT * FROM PayTransactions;"
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
      !PaymentAmount = cAmt
      !Discount = 0
      !UpdateSage = False
      'Add bank receipt form: add bank code below the bank account
'Modified by anol 13 Mar 2015
'BankCodeByBankID
      '!BankCode = BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
     ' !BankCode = frmBPPreForm.cmbBankAc.Column(2)
      'Modified by anol 06 apr 2015
      !BankCode = frmBPPreForm.cmbBankAc.Column(0)
      !nominalCode = rstRst!BankCode
      .Update
   End With

   rstRst.Close
   Set rstRst = Nothing
End Sub

Private Sub LoadDept(adoConn As ADODB.Connection)
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName FROM Fund;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
      ReDim Data(1, adoRst.RecordCount) As String

      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
         Data(1, rRow) = adoRst.Fields.Item("FundName").Value
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
      cboFund.Clear
      cboFund.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Exit Sub

   ' Error Handling Code
Error_Handler:

   ' Destroy Objects
   Set adoRst = Nothing
End Sub
Private Function returnRef(SageAccNo As String, issueDate As String, postingDate As String) As String
    'This function i written by anol 20 Apr 2015
    'This function shall return the reference number from grid
    'issue 0000547: Batch payments not working correctly
    '2/ Add reference column to batch payments
    returnRef = ""
    Dim iRow As Integer
    
        For iRow = 1 To flxSPayment.Rows - 1
            If bMultiple Then
                If flxSPayment.TextMatrix(iRow, 4) = SageAccNo And flxSPayment.TextMatrix(iRow, 22) = issueDate And flxSPayment.TextMatrix(iRow, 23) = postingDate Then
                returnRef = flxSPayment.TextMatrix(iRow, 24)
                Exit Function
                End If
            Else
                If flxSPayment.TextMatrix(iRow, 4) = SageAccNo And flxSPayment.TextMatrix(iRow, 11) = issueDate Then
                returnRef = flxSPayment.TextMatrix(iRow, 7)
                Exit Function
                End If
            End If
        Next iRow
   
        
   
End Function
Private Function CalculateVatAmountFormPIsplit(adoConn As ADODB.Connection, szPITranID As String, szSpitID As Integer, dblPayment As Double) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select VAT+NET_AMOUNT as amt,VAT from tblPurInvSRec S,tlbPayment P where P.PI =S.ParentID AND P.TransactionID=" & szPITranID & " and TRAN_ID='" & szSpitID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
            CalculateVatAmountFormPIsplit = dblPayment / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
             CalculateVatAmountFormPIsplit = CalculateVatAmountFormPIsplit * IIf(IsNull(rsDemandSplitRecords("VAT").Value), 0, rsDemandSplitRecords("VAT").Value)
    End If
    rsDemandSplitRecords.Close
    
End Function
Private Function CalculateNetAmountFormPIsplit(adoConn As ADODB.Connection, szDemandID As String, szSpitID As Integer, dblPayment As Double) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select VAT+NET_AMOUNT as amt,NET_AMOUNT,VAT  from tblPurInvSRec S,tlbPayment P where P.PI =S.ParentID AND P.TransactionID=" & szDemandID & " and TRAN_ID='" & szSpitID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
             If IIf(IsNull(rsDemandSplitRecords("VAT").Value), 0, rsDemandSplitRecords("VAT").Value) = 0 Then
                 CalculateNetAmountFormPIsplit = dblPayment
             Else
                    CalculateNetAmountFormPIsplit = dblPayment / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
                    CalculateNetAmountFormPIsplit = CalculateNetAmountFormPIsplit * IIf(IsNull(rsDemandSplitRecords("NET_AMOUNT").Value), 0, rsDemandSplitRecords("NET_AMOUNT").Value)
             End If
    End If
    rsDemandSplitRecords.Close

End Function
Private Function BookNormalPaymentOfPI(adoConn As ADODB.Connection, szSupplier As String) As Boolean
   Dim rstSet     As New ADODB.Recordset
   Dim rstAlloc   As New ADODB.Recordset
   Dim rstPS      As New ADODB.Recordset
   Dim rsInvoice As New ADODB.Recordset
   Dim rsSSR As New ADODB.Recordset
   Dim cSumSplits As Double
   Dim iRow          As Integer
   Dim iBottom       As Integer
   Dim cTotalPay     As Currency
   Dim szBankCode    As String
   Dim lSlNumber     As Long, iCT As Integer     'iCT -> number of child transactions
   Dim lTranID       As Long
   Dim lSp_ID        As Long, lSPTran_ID As Long, szSQL As String
   Dim i As Integer, j As Integer, aSupplier() As String, aSup() As String
   Dim t As Integer, b As Integer
   Dim intSplitId As Integer
   Dim lSPTran_ID_split As Long
   Dim rsPaytransaction As New ADODB.Recordset
   Dim rstPayTransactionsSplit As New ADODB.Recordset

'  Generate the next Transaction Id from the tlbPayment
On Error GoTo ErrHandler
   szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbPayment;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSp_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close

   szSQL = "SELECT MAX(TransactionID) AS TID FROM PayTransactions;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
    szSQL = "SELECT MAX(TransactionID) AS TID FROM PayTransactionsSplit;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID_split = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close

   If bMultiple Then
'      ReDim aSupplier(flxSPayment.Rows - 1, 2) As String
'      ReDim aSup(2) As String
      ReDim aSupplier(flxSPayment.Rows - 1, 3) As String
      ReDim aSup(3) As String

      j = 1
      i = 0
'      Copied all value entered rows in the array
      For iRow = 1 To flxSPayment.Rows - 1
         If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
            aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4)
            aSupplier(i, 1) = Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy")
            aSupplier(i, 2) = flxSPayment.TextMatrix(iRow, 11)
            aSupplier(i, 3) = flxSPayment.TextMatrix(iRow, 23)
            i = i + 1
         End If
      Next iRow
      iRow = i - 1

'      Sort the array by suppliers' id                                                          .
      For i = 1 To iRow
         For j = 0 To iRow - i
            If aSupplier(j, 0) > aSupplier(j + 1, 0) Then
               aSup(0) = aSupplier(j, 0)
               aSup(1) = aSupplier(j, 1)
               aSup(2) = aSupplier(j, 2)
               aSup(3) = aSupplier(j, 3)

               aSupplier(j, 0) = aSupplier(j + 1, 0)
               aSupplier(j, 1) = aSupplier(j + 1, 1)
               aSupplier(j, 2) = aSupplier(j + 1, 2)
               aSupplier(j, 3) = aSupplier(j + 1, 3)

               aSupplier(j + 1, 0) = aSup(0)
               aSupplier(j + 1, 1) = aSup(1)
               aSupplier(j + 1, 2) = aSup(2)
               aSupplier(j + 1, 3) = aSup(3)
            End If
         Next j
      Next i

'  group each supplier's payments lines
      For i = 0 To iRow - 1
'      mark top and bottom index of the a supplier
         For j = i + 1 To iRow
            If aSupplier(i, 0) <> aSupplier(j, 0) And aSupplier(j, 0) <> "" Then
               j = j - 1
               Exit For
            End If
         Next j
'         if the supplier has same date then sum payment to a single amount
         If i < j Then
            For t = i To j - 1
               For b = t + 1 To j
'               Debug.Print aSupplier(b, 0)
'               Debug.Print aSupplier(b, 1)
'               Debug.Print aSupplier(b, 2)
                  If aSupplier(t, 1) = aSupplier(b, 1) And aSupplier(t, 3) = aSupplier(b, 3) And Val(aSupplier(b, 2)) > 0 Then     'if the date is same for the same supplier
                     aSupplier(t, 2) = CStr(Val(aSupplier(t, 2)) + Val(aSupplier(b, 2)))
                     Debug.Print aSupplier(t, 2)
                     aSupplier(b, 2) = "0.00"
                  End If
               Next b
            Next t
            i = j
         End If
      Next i

'      squeeze the array where there is a 0.00 value
      t = 0       't counts how many lines are 0 value
      For i = 0 To iRow
         If Val(aSupplier(i, 2)) = 0 Then
            t = t + 1
            If i = iRow Then
               iRow = iRow - 1
               Exit For
            End If
            b = i
            For j = b + 1 To iRow
               aSupplier(b, 0) = aSupplier(j, 0)
               aSupplier(b, 1) = aSupplier(j, 1)
               aSupplier(b, 2) = aSupplier(j, 2)
               aSupplier(b, 3) = aSupplier(j, 3)
               b = b + 1
            Next j
            i = i - 1
            iRow = iRow - 1
         End If
      Next i
      
      Debug.Print iRow
   Else
      ReDim aSupplier(flxSPayment.Rows - 1, 2) As String
      ReDim aSup(2) As String
      
      j = 1
      i = 0
'      Copied all value entered rows in the array
      For iRow = 1 To flxSPayment.Rows - 1
         If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
            aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4)
            aSupplier(i, 2) = flxSPayment.TextMatrix(iRow, 11)
            i = i + 1
         End If
      Next iRow
      iRow = i - 1

'      Sort the array by suppliers' id                                                          .
      For i = 0 To iRow - 1
         For j = i To iRow
            If aSupplier(i, 0) > aSupplier(j, 0) Then
               aSup(0) = aSupplier(i, 0)
               aSup(2) = aSupplier(i, 2)
               
               aSupplier(i, 0) = aSupplier(j, 0)
               aSupplier(i, 2) = aSupplier(j, 2)
               
               aSupplier(j, 0) = aSup(0)
               aSupplier(j, 2) = aSup(2)
            End If
         Next j
      Next i

      i = 0
      j = i + 1
      While i <= iRow And j <= iRow
         j = i + 1
         Do While j <= iRow         '      mark top (i) and bottom (j) index of a supplier
            If aSupplier(i, 0) <> aSupplier(j, 0) Then
               i = i + 1
               j = i + 1
            Else
               While j <= iRow And aSupplier(i, 0) = aSupplier(j, 0)
                  j = j + 1
               Wend
               j = j - 1
               Exit Do
            End If
         Loop

'         sum all payments of same supplier to make a single payment
         If i < j And aSupplier(i, 0) = aSupplier(j, 0) Then
            For t = i To j - 1
               For b = t + 1 To j
                  aSupplier(t, 2) = CStr(Val(aSupplier(t, 2)) + Val(aSupplier(b, 2)))
                  aSupplier(b, 2) = "0.00"
               Next b
            Next t
            i = j + 1
         Else
            i = i + 1
         End If
      Wend

'      squeeze the array where there is a 0.00 value
      i = 0
      While i <= iRow
         If Val(aSupplier(i, 2)) = 0 Then
            b = i
            For j = b + 1 To iRow
               aSupplier(b, 0) = aSupplier(j, 0)
               aSupplier(b, 2) = aSupplier(j, 2)
               aSupplier(j, 2) = 0
               b = b + 1
            Next j
            i = i - 1
            iRow = iRow - 1
         End If
         i = i + 1
      Wend
      iBottom = iRow
   End If

'        ===============       ADD NEW PAYMENTS      ========================
'Add bank receipt form: add bank code below the bank account
'added by anol 13 Mar 2015
   'szBankCode = BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
   szBankCode = frmBPPreForm.cmbBankAc.Column(0)
'BankCodeByBankID
   szSQL = "SELECT * FROM tlbPayment;"
   rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   lSlNumber = SlNumber("PP", "tlbPayment", adoConn)
   'Resolved by BOSL
   'issue 547
   'Multiple payment was not working
   'Fixed by anol 06 Aug 2015
    iBottom = iRow
    'End of modification
      
    szSQL = "SELECT MAX(TransactionID) AS TID FROM PayTransactionsSplit;"
   rstPayTransactionsSplit.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID_split = CLng(IIf(IsNull(rstPayTransactionsSplit!TID), 1, rstPayTransactionsSplit!TID))
   rstPayTransactionsSplit.Close
   
   For i = 0 To iBottom
      intSplitId = 1
      rstSet.AddNew
      lSp_ID = lSp_ID + 1
      rstSet!TransactionID = lSp_ID
      rstSet!CreatedBy = User
      rstSet!CreatedDate = Now
      lTranID = lSp_ID
      rstSet!szTransactionID = CStr(rstSet!TransactionID)
      rstSet!Type = 8
      rstSet!SageAccountNumber = aSupplier(i, 0)
      If bMultiple Then
         rstSet!PDate = Format(aSupplier(i, 1), "dd mmmm yyyy")
         rstSet!postingDate = Format(aSupplier(i, 3), "dd mmmm yyyy")
      Else
         rstSet!PDate = Format(lblDate.Caption, "dd mmmm yyyy")
         rstSet!postingDate = Format(frmBPPreForm.lblPostingDate.ToolTipText, "dd mmmm yyyy")
      End If
      rstSet!Details = "BATCH PAYMENT"
      rstSet!amount = CCur(aSupplier(i, 2))
      'below line has been added by anol 20181031 issue 669
      rstSet!OSAmount = 0
      rstSet!IsSageUpdate = False
      rstSet!UpdateSage = False
      rstSet!PaymentView = False
      rstSet!LastModifiedBy = User
      rstSet!LastModifiedDate = Now
      'issue 547
      'Next cheque was not incrementing
      'Modified by anol 06 Apr 2015
      If bMultiple Then
             rstSet!ExtRef = returnRef(aSupplier(i, 0), aSupplier(i, 1), aSupplier(i, 3))
             If Len(rstSet!ExtRef) = 0 Then
                    If Val(txtChqNo.text) = 0 Then
                        rstSet!ExtRef = txtChqNo.text
                    Else
                        rstSet!ExtRef = Val(txtChqNo.text) + i
                    End If
            End If
      Else
            If Val(txtChqNo.text) = 0 Then
                rstSet!ExtRef = txtChqNo.text 'returnRef(aSupplier(i, 0), aSupplier(i, 2), "")
            Else
                rstSet!ExtRef = Val(txtChqNo.text) + i
            End If
      End If
      'added by anol 21 Apr 2015
      'If bMultiple Then
            'rstSet!ref = returnRef(aSupplier(i, 0), aSupplier(i, 1), aSupplier(i, 3))
     ' Else
            'rstSet!ref = returnRef(aSupplier(i, 0), aSupplier(i, 2), "")
      'End If
      'End of modification
      rstSet!PayAmtType = IIf(optBP_BACS.Value, "BACS", "CHQ")
      rstSet!BankCode = szBankCode
      rstSet!nominalCode = rstSet!BankCode
      rstSet!SlNumber = lSlNumber
      rstSet!ClientID = frmBPPreForm.txtClient.Tag
      rstSet.Update

'  ################################################# Create PP splits  ###############################################
      szSQL = "SELECT * FROM tlbPaymentSplit where 1=2;"
      rstPS.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic 'this record set shall be used to create new record in tlbPaymentSplit table
      lSPTran_ID_split = lSPTran_ID_split + 1
      For iRow = 1 To flxSPayment.Rows - 1
        If bMultiple Then
            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And aSupplier(i, 1) = flxSPayment.TextMatrix(iRow, 22) And IsDate(flxSPayment.TextMatrix(iRow, 22)) And _
               Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
            
            cTotalPay = Val(flxSPayment.TextMatrix(iRow, 11))

            szSQL = "SELECT * FROM tlbPaymentSplit AS S " & _
                    "WHERE  S.OSAmount > 0 AND " & _
                          " S.PayHeader = " & flxSPayment.TextMatrix(iRow, 20) & " " & _
                    "ORDER BY S.SplitID;" 'This is a PI transactionID

            rstAlloc.Open szSQL, adoConn, adOpenStatic, adLockReadOnly 'creating PP splits
            'This procedure is creating same number of PP splits as in th PI
            While cTotalPay > 0
               'Note by anol what happens is there is no PI splits? this should not happened by the SOP.if happenes then
               'I should log that in a table and look for that problem later
               With rstPS 'This (" rstPS ")recordset shall be used to create new record in tlbPaymentSplit table
                  .AddNew
                  .Fields.Item("TransactionID").Value = UniqueID()
                  .Fields.Item("PayHeader").Value = lTranID
                  .Fields.Item("FundID").Value = rstAlloc.Fields.Item("FundID").Value
                  If cTotalPay >= CCur(rstAlloc.Fields.Item("OSAmount").Value) Then
                     .Fields.Item("Amount").Value = CCur(rstAlloc.Fields.Item("OSAmount").Value)
                     cTotalPay = cTotalPay - CCur(rstAlloc.Fields.Item("OSAmount").Value)
                  Else
                     .Fields.Item("Amount").Value = cTotalPay
                     cTotalPay = 0
                  End If
                  '.Fields.Item("SplitID").Value = 1
                  ' Modified by anol 2019-07-27
                  .Fields.Item("SplitID").Value = intSplitId
                   intSplitId = intSplitId + 1
                  .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
                  .Fields.Item("Description").Value = rstAlloc.Fields.Item("Description").Value
                  .Fields.Item("AllocTranID").Value = rstAlloc.Fields.Item("TransactionID").Value 'this PI split transaction ID putting into PP paypent split
                  .Fields.Item("PayTransactionIDSplit").Value = lSPTran_ID_split
                  .Fields.Item("PropertyID").Value = flxSPayment.TextMatrix(iRow, 5) 'property ID
                  .Update
               End With
               'added by anol 2021-10-23 writing on PayTransactionsSplit table
               With rsPaytransaction
                    szSQL = "SELECT * FROM PayTransactionsSplit;"
                    rsPaytransaction.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                    .AddNew
                    !TranType = "AL"
                    !TransactionID = lSPTran_ID_split
                    !Alloc_Unalloc = 1
                    !FromTran = lTranID                   'Payment transaction ID
                    !ToTran = CLng(flxSPayment.TextMatrix(iRow, 20)) 'PI transaction ID
                    !AllocDate = Format(aSupplier(i, 1), "DD MMMM YYYY")
                    !PaymentAmount = rstPS.Fields.Item("Amount").Value
                    !BankCode = szBankCode
                    !nominalCode = !BankCode
                   ' !SlNumber = lSlNumber
                    !fundID = rstAlloc.Fields.Item("FundID").Value
                    !VATAMOUNT = CalculateVatAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 20)), rstPS.Fields.Item("SplitID").Value, !PaymentAmount)
                    !NetAmount = CalculateNetAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 20)), rstPS.Fields.Item("SplitID").Value, !PaymentAmount)
                    !VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
                    !SplitIDofPI = rstPS.Fields.Item("SplitID").Value
                    !deleteFlag = False
                    .Update
                End With
                rsPaytransaction.Close
                lSPTran_ID_split = lSPTran_ID_split + 1
                rstAlloc.MoveNext
            Wend
            rstAlloc.Close
         End If
       Else 'else when not  bmultiple
            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And _
               Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
                lSPTran_ID_split = lSPTran_ID_split + 1
                cTotalPay = Val(flxSPayment.TextMatrix(iRow, 11))
    
                szSQL = "SELECT * FROM tlbPaymentSplit AS S " & _
                        "WHERE  S.OSAmount > 0 AND " & _
                              " S.PayHeader = " & flxSPayment.TextMatrix(iRow, 20) & " " & _
                        "ORDER BY S.SplitID;"
   
                rstAlloc.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                   'This procedure is creating same number of PP splits as in th PI
                While cTotalPay > 0
                   With rstPS
                      .AddNew
                      .Fields.Item("TransactionID").Value = UniqueID()
                      .Fields.Item("PayHeader").Value = lTranID
                      .Fields.Item("FundID").Value = rstAlloc.Fields.Item("FundID").Value
                      If cTotalPay >= CCur(rstAlloc.Fields.Item("OSAmount").Value) Then
                         .Fields.Item("Amount").Value = CCur(rstAlloc.Fields.Item("OSAmount").Value)
                         cTotalPay = cTotalPay - CCur(rstAlloc.Fields.Item("OSAmount").Value)
                      Else
                         .Fields.Item("Amount").Value = cTotalPay
                         cTotalPay = 0
                      End If
                      '.Fields.Item("SplitID").Value = 1
                        ' Modified by anol 2019-07-27
                      .Fields.Item("SplitID").Value = intSplitId
                       intSplitId = intSplitId + 1
                      .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
                      .Fields.Item("Description").Value = rstAlloc.Fields.Item("Description").Value
                      .Fields.Item("AllocTranID").Value = rstAlloc.Fields.Item("TransactionID").Value
                      .Fields.Item("PayTransactionIDSplit").Value = lSPTran_ID_split
                      .Fields.Item("PropertyID").Value = flxSPayment.TextMatrix(iRow, 5) 'property ID
                      .Update
                   End With
                   'added by anol 2021-10-23 writing on PayTransactionsSplit table
                    With rsPaytransaction
                         szSQL = "SELECT * FROM PayTransactionsSplit;"
                         rsPaytransaction.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                         .AddNew
                         !TranType = "AL"
                         !TransactionID = lSPTran_ID_split
                         !Alloc_Unalloc = 1
                         !FromTran = lTranID                   'Payment transaction ID
                         !ToTran = CLng(flxSPayment.TextMatrix(iRow, 20)) 'PI transaction ID
                         !AllocDate = Format(Date, "DD MMMM YYYY")
                         !PaymentAmount = rstPS.Fields.Item("Amount").Value
                         !BankCode = szBankCode
                         !nominalCode = !BankCode
                        ' !SlNumber = lSlNumber
                         !fundID = rstAlloc.Fields.Item("FundID").Value
                         !VATAMOUNT = CalculateVatAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 20)), rstPS.Fields.Item("SplitID").Value, !PaymentAmount)
                         !NetAmount = CalculateNetAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 20)), rstPS.Fields.Item("SplitID").Value, !PaymentAmount)
                         !VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
                         !SplitIDofPI = rstPS.Fields.Item("SplitID").Value
                         !deleteFlag = False
                         .Update
                     End With
                     rsPaytransaction.Close
                     lSPTran_ID_split = lSPTran_ID_split + 1
                   rstAlloc.MoveNext
                Wend
                rstAlloc.Close
            End If
       End If 'end if for bmultiple
      Next iRow

      rstPS.Close
      
'     ====================  UPDATE PI SPLITS in tlbPaymentSplit    ===============================
      For iRow = 1 To flxSPayment.Rows - 1
      'Below line was fixed by anol 25 Mar 2015 tlbPaymentSplit was not updating properly
      'Rollbacked 06 apr 2015
      'And aSupplier(i, 1) = flxSPayment.TextMatrix(iRow, 22) And IsDate(flxSPayment.TextMatrix(iRow, 22))
            If bMultiple Then
                  If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And aSupplier(i, 1) = flxSPayment.TextMatrix(iRow, 22) And IsDate(flxSPayment.TextMatrix(iRow, 22)) And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
                  
                     'If aSupplier(iRow - 1, 0) = flxSPayment.TextMatrix(iRow, 4) And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
                        cTotalPay = Val(flxSPayment.TextMatrix(iRow, 11))
           
                        szSQL = "SELECT * FROM tlbPaymentSplit AS S WHERE S.PayHeader = " & flxSPayment.TextMatrix(iRow, 20) & ";"
                        rstAlloc.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
                       
                      If Not rstAlloc.EOF Then
                            Do While cTotalPay > 0
                                If cTotalPay >= CCur(rstAlloc.Fields.Item("OSAmount").Value) Then
                                   cTotalPay = cTotalPay - CCur(rstAlloc.Fields.Item("OSAmount").Value)
                                   rstAlloc.Fields.Item("OSAmount").Value = 0
                                Else
                                   rstAlloc.Fields.Item("OSAmount").Value = rstAlloc.Fields.Item("OSAmount").Value - cTotalPay
                                   cTotalPay = 0
                                End If
                                rstAlloc.Update
                                rstAlloc.MoveNext
                                If rstAlloc.EOF Then Exit Do
                             Loop
                        End If
                        rstAlloc.Close
                     End If
            Else 'else when not bmultiple
                  If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
                        cTotalPay = Val(flxSPayment.TextMatrix(iRow, 11))
           
                        szSQL = "SELECT * FROM tlbPaymentSplit AS S " & _
                                "WHERE S.PayHeader = " & flxSPayment.TextMatrix(iRow, 20) & ";"
                        rstAlloc.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
                        'Below line added by anol 26 Mar 2015
                      If Not rstAlloc.EOF Then
                            Do While cTotalPay > 0
                                If cTotalPay >= CCur(rstAlloc.Fields.Item("OSAmount").Value) Then
                                   cTotalPay = cTotalPay - CCur(rstAlloc.Fields.Item("OSAmount").Value)
                                   rstAlloc.Fields.Item("OSAmount").Value = 0
                                Else
                                   rstAlloc.Fields.Item("OSAmount").Value = rstAlloc.Fields.Item("OSAmount").Value - cTotalPay
                                   cTotalPay = 0
                                End If
                                rstAlloc.Update
                                rstAlloc.MoveNext
                                If rstAlloc.EOF Then Exit Do
                             Loop
                        End If
                        rstAlloc.Close
                     End If
            End If
      Next iRow
'     ============      RECORD THE ALLOCATION      ==================
'#
      szSQL = "SELECT * " & _
              "FROM   PayTransactions;"
      rstAlloc.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      For iRow = 1 To flxSPayment.Rows - 1
         If bMultiple Then
            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And _
                  Val(flxSPayment.TextMatrix(iRow, 11)) > 0 And _
                  aSupplier(i, 1) = Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy") Then
               rstAlloc.AddNew
               rstAlloc!TranType = "AL"
               lSPTran_ID = lSPTran_ID + 1
               rstAlloc!TransactionID = lSPTran_ID
               rstAlloc!Alloc_Unalloc = 1
               rstAlloc!FromTran = lSp_ID
               rstAlloc!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
               rstAlloc!AllocDate = Format(aSupplier(i, 1), "DD MMMM YYYY")
               rstAlloc!PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 11))
               rstAlloc!BankCode = szBankCode 'BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
               rstAlloc!nominalCode = rstAlloc!BankCode
               rstAlloc!SlNumber = lSlNumber

               rstAlloc.Update
            End If
         Else
            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
               rstAlloc.AddNew
               rstAlloc!TranType = "AL"
               lSPTran_ID = lSPTran_ID + 1
               rstAlloc!TransactionID = lSPTran_ID
               rstAlloc!Alloc_Unalloc = 1
               rstAlloc!FromTran = lSp_ID
               rstAlloc!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
               rstAlloc!AllocDate = Format(Date, "DD MMMM YYYY")
               rstAlloc!PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 11))
               rstAlloc!BankCode = szBankCode            'BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
               rstAlloc!nominalCode = rstAlloc!BankCode
               rstAlloc!SlNumber = lSlNumber

               rstAlloc.Update
            End If
         End If
      Next iRow
      rstAlloc.Close

      lSlNumber = lSlNumber + 1
   
'''     ============      RECORD THE ALLOCATION   SPLIT   ==================
'''#
''      szSQL = "SELECT * " & _
''              "FROM   PayTransactionsSplit;"
''      rstAlloc.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
''
''      For iRow = 1 To flxSPayment.Rows - 1
''         If bMultiple Then
''            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And _
''                  Val(flxSPayment.TextMatrix(iRow, 11)) > 0 And _
''                  aSupplier(i, 1) = Format(flxSPayment.TextMatrix(iRow, 22), "dd/mm/yyyy") Then
''                  rsInvoice.Open "Select PayHeader,Amount,SplitID,fundID,P.unitID  FROM tlbPaymentSplit,tlbPayment P where P.transactionID=payHeader AND payHeader=" & CLng(flxSPayment.TextMatrix(iRow, 20)) & " ", adoconn, adOpenDynamic, adLockOptimistic
''                  If Not rsInvoice.EOF Then
''                        cSumSplits = CCur(flxSPayment.TextMatrix(iRow, 11))
''                  End If
''                  While Not rsInvoice.EOF And cSumSplits > 0
''                            rstAlloc.AddNew
''                            rstAlloc!TranType = "AL"
''                            lSPTran_ID_split = lSPTran_ID_split + 1
''                            rstAlloc!transactionID = lSPTran_ID_split
''                            rstAlloc!Alloc_Unalloc = 1
''                            rstAlloc!FromTran = lSp_ID
''                            rstAlloc!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
''                            rstAlloc!AllocDate = Format(aSupplier(i, 1), "DD MMMM YYYY")
''                            'rstAlloc!PaymentAmount = rsInvoice!amount
''
''                            If rsInvoice!amount >= cSumSplits Then
''                                rstAlloc!PaymentAmount = cSumSplits
''                                cSumSplits = 0
''                            Else
''                                 rstAlloc!PaymentAmount = rsInvoice!amount
''                                 cSumSplits = cSumSplits - rsInvoice!amount
''                            End If
''
''
''                            'cSumSplits = cSumSplits - rsInvoice!amount
''                            rstAlloc!BankCode = szBankCode 'BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
''                            rstAlloc!nominalCode = rstAlloc!BankCode
''                           ' rstAlloc!SlNumber = lSlNumber
''                            rstAlloc!fundID = rsInvoice!fundID
''                            rstAlloc!SplitIDofPI = rsInvoice!splitID
''                            rstAlloc!propertyID = rsInvoice!unitid
''                            If Left(flxSPayment.TextMatrix(iRow, 1), 3) = "PPR" Then  'allocating PPR
''                                    rsSSR.Open "Select * from PayTransactions where SplitIDofPI=" & rsInvoice!splitID & " and  FromTran=" & _
''                                    CLng(flxSPayment.TextMatrix(iRow, 20)) & " and DeleteFlag=True", adoconn, adOpenDynamic, adLockOptimistic
''                                    If Not rsSSR.EOF Then
''                                         rstAlloc!VATAMOUNT = rsSSR!VATAMOUNT
''                                    Else
''                                         rstAlloc!VATAMOUNT = 0
''                                    End If
''                                    rsSSR.Close
''                                    rsSSR.Open "Select * from PayTransactions where SplitIDofPI=" & rsInvoice!splitID & " and  FromTran=" & _
''                                    CLng(flxSPayment.TextMatrix(iRow, 20)) & " and DeleteFlag=True", adoconn, adOpenDynamic, adLockOptimistic
''                                    If Not rsSSR.EOF Then
''                                        rstAlloc!NetAmount = rsSSR!PaymentAmount
''                                    Else
''                                        rstAlloc!NetAmount = rsInvoice!amount
''                                    End If
''                                    rsSSR.Close
''                            Else    'allocating PI
''                                'Exit Function
''                                    rstAlloc!VATAMOUNT = CalculateVatAmountFormPIsplit(adoconn, CLng(flxSPayment.TextMatrix(iRow, 20)), rsInvoice!splitID, rstAlloc!PaymentAmount)
''                                    rstAlloc!NetAmount = CalculateNetAmountFormPIsplit(adoconn, CLng(flxSPayment.TextMatrix(iRow, 20)), rsInvoice!splitID, rstAlloc!PaymentAmount)
''                            End If
''                           rstAlloc!VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
''                          rstAlloc!SplitIDofPI = rsInvoice!splitID
''                          rstAlloc!deleteFlag = False
''                           rstAlloc.Update
''                        rsInvoice.MoveNext
''                    Wend
''               rsInvoice.Close
''            End If
''         Else
''            If aSupplier(i, 0) = flxSPayment.TextMatrix(iRow, 4) And Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
''                rsInvoice.Open "Select PayHeader,Amount,SplitID,fundID  FROM tlbPaymentSplit where payHeader=" & CLng(flxSPayment.TextMatrix(iRow, 20)) & " ", adoconn, adOpenDynamic, adLockOptimistic
''                  If Not rsInvoice.EOF Then
''                        cSumSplits = CCur(flxSPayment.TextMatrix(iRow, 11))
''                  End If
''                  While Not rsInvoice.EOF And cSumSplits > 0
''                            rstAlloc.AddNew
''                            rstAlloc!TranType = "AL"
''                            lSPTran_ID_split = lSPTran_ID_split + 1
''                            rstAlloc!transactionID = lSPTran_ID_split
''                            rstAlloc!Alloc_Unalloc = 1
''                            rstAlloc!FromTran = lSp_ID
''                            rstAlloc!ToTran = CLng(flxSPayment.TextMatrix(iRow, 20))
''                            rstAlloc!AllocDate = Format(Date, "DD MMMM YYYY")
''                            'rstAlloc!PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 11))
''                            rstAlloc!PaymentAmount = rsInvoice!amount
''                            cSumSplits = cSumSplits - rsInvoice!amount
''                            rstAlloc!BankCode = szBankCode            'BankCodeByBankID(adoConn, frmBPPreForm.cmbBankAc.Column(0))
''                            rstAlloc!nominalCode = rstAlloc!BankCode
''                           ' rstAlloc!SlNumber = lSlNumber
''                            rstAlloc!SplitIDofPI = rsInvoice!splitID
''                            rstAlloc!fundID = rsInvoice!fundID
''                            rstAlloc.Update
''                            If Left(flxSPayment.TextMatrix(iRow, 1), 3) = "PPR" Then  'allocating PPR
''                                    rsSSR.Open "Select * from PayTransactions where SplitIDofPI=" & rsInvoice!splitID & " and  FromTran=" & _
''                                    CLng(flxSPayment.TextMatrix(iRow, 20)) & " and DeleteFlag=True", adoconn, adOpenDynamic, adLockOptimistic
''                                    If Not rsSSR.EOF Then
''                                         rstAlloc!VATAMOUNT = rsSSR!VATAMOUNT
''                                    Else
''                                         rstAlloc!VATAMOUNT = 0
''                                    End If
''                                    rsSSR.Close
''                                    rsSSR.Open "Select * from PayTransactions where SplitIDofPI=" & rsInvoice!splitID & " and  FromTran=" & _
''                                    CLng(flxSPayment.TextMatrix(iRow, 20)) & " and DeleteFlag=True", adoconn, adOpenDynamic, adLockOptimistic
''                                    If Not rsSSR.EOF Then
''                                        rstAlloc!NetAmount = rsSSR!PaymentAmount
''                                    Else
''                                        rstAlloc!NetAmount = rsInvoice!amount
''                                    End If
''                                    rsSSR.Close
''                            Else    'allocating PI
''                                    rstAlloc!VATAMOUNT = CalculateVatAmountFormPIsplit(adoconn, CLng(flxSPayment.TextMatrix(iRow, 20)), rsInvoice!splitID, rsInvoice!amount)
''                                    rstAlloc!NetAmount = CalculateNetAmountFormPIsplit(adoconn, CLng(flxSPayment.TextMatrix(iRow, 20)), rsInvoice!splitID, rsInvoice!amount)
''                            End If
''                           rstAlloc!VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
''                          rstAlloc!SplitIDofPI = rsInvoice!splitID
''                          rstAlloc!deleteFlag = False
''                        rsInvoice.MoveNext
''                    Wend
''                    rsInvoice.Close
''            End If
''         End If
''      Next iRow
''      rstAlloc.Close

''      lSlNumber = lSlNumber + 1
   Next i

    '    =============     UPDATE OS BALANCE OF PI  (tlbPayment)   =================
   For iRow = 1 To flxSPayment.Rows - 1
      If Val(flxSPayment.TextMatrix(iRow, 11)) > 0 Then
         rstSet.Find "TransactionID = " & flxSPayment.TextMatrix(iRow, 20), , , 1
'Debug.Print rstSet.RecordCount
         rstSet!OSAmount = CCur(flxSPayment.TextMatrix(iRow, 10)) - _
                           CCur(IIf(flxSPayment.TextMatrix(iRow, 11) = "", 0, _
                           flxSPayment.TextMatrix(iRow, 11)))
         rstSet!PaymentView = IIf(rstSet!OSAmount > 0, True, False)
         If rstSet!PaymentView = False Then
            rstSet!DateTimeStamp = ""
            rstSet!Module = ""
            rstSet!UserSessionID = ""
            rstSet!WindowsUserName = ""
            rstSet!MachineName = ""
            rstSet!PrestigeUserName = ""
            rstSet!ServerIPaddress = ""
         End If
         rstSet.Update
      End If
   Next iRow
   rstSet.Close
  ' adoConn.Execute "Update tlbpaymentsplit set OSAMOUNT=200"
   Set rstSet = Nothing
   If SiPi_Check(adoConn, "PI") = False Then Exit Function
   If ChecksumValidationOnPayAllocation(adoConn) = False Then Exit Function
   BookNormalPaymentOfPI = True
   Exit Function
ErrHandler:
   MsgBox Err.description, vbInformation, "An error occured in BookNormalPayment of PI, Please contact PCM consulting LTD."
End Function
Private Function SiPi_Check(adoConn As ADODB.Connection, szSiPi As String) As Boolean
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
        SiPi_Check = False
         If szSiPi = "PI" Then
            MsgBox "Payment and Payment Split are diffent, it is going to rollback this transaction"
         End If
         If szSiPi = "SI" Then
            MsgBox "Receipt and Receipt Split are diffent, it is going to rollback this transaction"
         End If
   Else
        SiPi_Check = True
        'MsgBox "HI"
   End If
      
End Function

Private Sub ConfigFlxSPayment()
   Dim szHeader As String

   flxSPayment.Clear
   flxSPayment.Cols = 32
   If bMultiple Then
    'Modified by Anol 21 Apr 2015
    'issue 547
      Me.Width = 16465
      Label1(11).Visible = True
      Label20(18).Width = 16465
    'Modified by Anol 21 Apr 2015
    'issue 547
    'flxSPayment.Cols = 24
      'flxSPayment.Cols = 25
      szHeader$ = "|<No.|<Supplier|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
                  "|<Ref|<Details|>Amount £|>O/S Amt. £|>Receipt £|>Discount" & _
                  "|<DemandID|>SAGE O/S £|<RptNo|<PayDate|<PostingDate"
   Else
      Me.Width = 13245
      Label1(11).Visible = False
      Label20(18).Width = 12975

      'flxSPayment.Cols = 22
      szHeader$ = "|<No.|<Supplier|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
                  "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
                  "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo."
   End If
   Dim n As Integer
   flxSPayment.Width = Label20(18).Width
   flxSPayment.Rows = 2
   flxSPayment.RowHeight(0) = 0

   flxSPayment.FormatString = szHeader$

   flxSPayment.ColAlignment(0) = vbCenter
   flxSPayment.ColWidth(0) = Label1(1).Left - flxSPayment.Left    'Sign
   flxSPayment.ColWidth(1) = Label1(2).Left - Label1(1).Left      'No
   flxSPayment.ColWidth(2) = Label1(3).Left - Label1(2).Left      'Supplier
   flxSPayment.ColWidth(3) = Label1(4).Left - Label1(3).Left      'Type
   flxSPayment.ColWidth(4) = 0      'Supplier A/c - no need to show it in the grid, its already in the header part
   flxSPayment.ColWidth(5) = Label1(5).Left - Label1(4).Left      'Unit ID
   flxSPayment.ColWidth(6) = Label1(6).Left - Label1(5).Left      'Date
   flxSPayment.ColWidth(7) = Label1(7).Left - Label1(6).Left
   flxSPayment.ColWidth(8) = Label1(8).Left - Label1(7).Left      'Details
   flxSPayment.ColWidth(9) = Label1(9).Left - Label1(8).Left      'Amount
   flxSPayment.ColWidth(10) = Label1(10).Left - Label1(9).Left    'O/S Amount
   If bMultiple Then
      flxSPayment.ColWidth(11) = Label1(11).Left - Label1(10).Left 'Payment Amt
   Else
      flxSPayment.ColWidth(11) = flxSPayment.Width + flxSPayment.Left - Label1(10).Left - 300 'Receipt Amt
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
   flxSPayment.ColWidth(21) = 0     'Supplier ID
   If bMultiple Then
        flxSPayment.ColWidth(22) = Label1(12).Left - Label1(11).Left 'Date
        flxSPayment.ColWidth(23) = 960 'flxSPayment.Width + flxSPayment.Left - Label1(12).Left - 300 'Posting date
        flxSPayment.ColWidth(24) = 1360  'flxSPayment.Width + flxSPayment.Left - Label1(12).Left - 300 'Referrence
   Else
        flxSPayment.ColWidth(22) = 0 'not in use when not multiplae
        flxSPayment.ColWidth(23) = 0 'not in use when not multiplae
        flxSPayment.ColWidth(24) = 0 'not in use when not multiplae
   End If
   flxSPayment.ColWidth(25) = 0 'Shall keep the supplier ID if checksum problematic transactions are found
   flxSPayment.ColWidth(26) = 0 'Shall keep UserSessionID if this row is locked by other user /different module in system
   flxSPayment.ColWidth(27) = 0 'Shall keep ComputerUser if this row is locked by other user /different module in system
   flxSPayment.ColWidth(28) = 0 'Shall keep MachineName if this row is locked by other user /different module in system
   flxSPayment.ColWidth(29) = 0 'Shall keep Module if this row is locked by other user /different module in system
   flxSPayment.ColWidth(30) = 0 'Shall keep ClientID if this row is locked by other user /different module in system
   flxSPayment.ColWidth(31) = 0 ' I shall Keep isRentpayable here
   txtGrossTotal.Width = flxSPayment.ColWidth(11)
   cmdSPClose.Width = txtGrossTotal.Width
   txtGrossTotal.text = "0.00"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   If cOpeningBal <> CCur(Val(txtGrossTotal.text)) And CCur(Val(txtGrossTotal.text)) = 0 Then
      
      Dim szSQL   As String

   '   Database has been connected.
      
      szSQL = "DELETE * FROM tblBatchTransaction " & _
              "WHERE BP IN (SELECT BP.BP FROM tblBatchPayment AS BP " & _
                           "WHERE BP.Generated = FALSE);"
   'Debug.Print szSQL
      adoConn.Execute szSQL
      adoConn.Execute "DELETE * FROM tblBatchPayment WHERE Generated = FALSE;"

     
   End If
   bBPPreForm = False
   adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
   adoConn.Close
   Set adoConn = Nothing
   UserSessionID = ""
   frmLockingDialogisActive = False
   UnLoadForm Me
   Unload frmBPPreForm
    
'   Call WheelUnHook(Me.hwnd)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
'  Move to the next column for enter the date
   If flxSPayment.row < flxSPayment.Rows And bMultiple And Not txtPayDt.Visible And flxSPayment.col = 11 Then
      txtSPayment.Visible = False

      flxSPayment.col = 22
      iFlxSPayCol = 22
      txtPayDt.Top = flxSPayment.CellTop + flxSPayment.Top
      txtPayDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      txtPayDt.Width = flxSPayment.ColWidth(22)
      txtPayDt.Height = flxSPayment.RowHeight(iCurRow) - 15
      txtPayDt.Visible = True
      flxSPayment.ScrollBars = flexScrollBarNone
      txtPayDt.SetFocus
      SumUpTotal
   End If
'  Move to the next row to enter amount
   If flxSPayment.row < flxSPayment.Rows And Not txtPayDt.Visible And flxSPayment.col = 22 Then
      If MoveDownPosition Then
         flxSPayment.col = 11

         flxSPayment.SetFocus
      Else
         cmdSavePayment.SetFocus
      End If
   End If
End Sub

Private Sub txtGrossTotal_GotFocus()
   cmdSavePayment.SetFocus
End Sub

Private Sub txtPayDt_Change()
   TextBoxChangeDate txtPayDt
End Sub

Private Sub txtPayDt_GotFocus()
   iCurRow = StarFound

   If flxSPayment.col = 22 Then
      If flxSPayment.TextMatrix(iCurRow, 22) = "" Then
         txtPayDt.text = Format(Now, "dd/mm/yyyy")
      Else
         txtPayDt.text = Format(flxSPayment.TextMatrix(iCurRow, 22), "dd/mm/yyyy")
      End If
   End If
'   If flxSPayment.col = 23 Then
'      If flxSPayment.TextMatrix(iCurRow, 23) = "" Then
'         txtPayDt.text = Format(Now, "dd/mm/yyyy")
'      Else
'         txtPayDt.text = Format(flxSPayment.TextMatrix(iCurRow, 23), "dd/mm/yyyy")
'      End If
'   End If

    txtPostingDate.Visible = False
    txtRef.Visible = False
    txtSPayment.Visible = False
    
   SelTxtInCtrl txtPayDt
End Sub

Private Sub txtPayDt_LostFocus()
'    flxSPayment.TextMatrix(iCurRow, 22) = txtPayDt.text
'fixed date formating by ANOL 30 Aug 2016
        If txtPayDt.text <> "" Then TextBoxFormatDate txtPayDt
        flxSPayment.TextMatrix(iCurRow, 22) = txtPayDt.text
        If flxSPayment.TextMatrix(iCurRow, 22) <> "" Then
            flxSPayment.TextMatrix(iCurRow, 23) = flxSPayment.TextMatrix(iCurRow, 22)
        End If
End Sub

Private Sub txtPostingDate_Change()
       TextBoxChangeDate txtPostingDate
End Sub

Private Sub txtPostingDate_GotFocus()
     txtPayDt.Visible = False
     txtRef.Visible = False
     txtSPayment.Visible = False
    
    If iFlxSPayCol = 24 Then
      If flxSPayment.TextMatrix(iCurRow, 24) = "" Then
            txtPostingDate.text = Format(Now, "dd/mm/yyyy")
      Else
            txtPostingDate.text = Format(flxSPayment.TextMatrix(iCurRow, 24), "dd/mm/yyyy")
      End If
'      SelTxtInCtrl txtPostingDate
   End If
End Sub

Private Sub txtPostingDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      flxSPayment.SetFocus
      txtPostingDate.text = ""
      txtPostingDate.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

   If KeyAscii = 13 Then
      flxSPayment.TextMatrix(iCurRow, 23) = txtPostingDate.text
      If IsDate(txtPostingDate.text) = False Then
            ShowMsgInTaskBar "Date format is not correct!"
            Exit Sub 'added by anol 20160911
      End If
      txtPostingDate.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
      
      flxSPayment.col = 24
      flxSPayment_dblClick
   End If
   TextBoxKeyPrsDate txtPostingDate, KeyAscii
End Sub

Private Sub txtPostingDate_LostFocus()
'fixed date formating by ANOL 30 Aug 2016
     If txtPostingDate.text <> "" Then TextBoxFormatDate txtPostingDate
     flxSPayment.TextMatrix(iCurRow, 23) = txtPostingDate.text
     
End Sub

Private Sub txtRef_GotFocus()
    txtPayDt.Visible = False
    txtPostingDate.Visible = False
    txtSPayment.Visible = False
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      flxSPayment.SetFocus
      txtPostingDate.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

   If KeyAscii = 13 Then
        If Len(txtRef.text) = 0 Then
                If MsgBox("You have not entered a reference. Do you wish to leave this blank?", vbYesNo, "Warning") = vbNo Then
                    txtRef.SetFocus
                    Exit Sub
                End If
        End If
        flxSPayment.TextMatrix(iCurRow, 24) = txtRef.text
        If MoveDownPosition Then
             flxSPayment.col = 11
             flxSPayment.SetFocus
        End If
        'iCurRow = iCurRow - 1
        txtRef.Visible = False
        flxSPayment.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub txtRef_LostFocus()
    flxSPayment.TextMatrix(iCurRow, 24) = txtRef.text
    txtRef.Visible = False
End Sub

Private Sub txtSPayment_GotFocus()
   iCurRow = StarFound

   If Val(txtSPayment.text) = 0 Then
      txtSPayment.text = Format(flxSPayment.TextMatrix(iCurRow, 10), "0.00")
      SelTxtInCtrl txtSPayment
   End If
   txtPayDt.Visible = False
   txtPostingDate.Visible = False
   txtRef.Visible = False
End Sub

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer

   If flxSPayment.row < flxSPayment.Rows - 1 Then
      If flxSPayment.RowHeight(flxSPayment.row + 1) = 0 Then
         txtSPayment.Visible = False
         flxSPayment.ScrollBars = flexScrollBarVertical
         MoveDownPosition = False
         Exit Function
      End If
      flxSPayment.row = flxSPayment.row + 1
      'iCurRow = flxSPayment.row
   Else
      txtSPayment.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
      MoveDownPosition = False
      Exit Function
   End If

   iTop = flxSPayment.CellTop + flxSPayment.Top
   MoveDownPosition = True
End Function

Private Sub txtSPayment_KeyDown(KeyCode As Integer, Shift As Integer)
    'Smoothing neviagtion
    'Wriiten by anol 21 Apr 2015
    'issue 547
     'If KeyCode = 13 Then Text1(0).SetFocus
   '  Move to the next column for enter the date
   If KeyCode = 13 Then
       If flxSPayment.row < flxSPayment.Rows And bMultiple And Not txtPayDt.Visible And flxSPayment.col = 11 Then
          txtSPayment.Visible = False
    
          flxSPayment.col = 22
          iFlxSPayCol = 22
          txtPayDt.Top = flxSPayment.CellTop + flxSPayment.Top
          txtPayDt.Left = flxSPayment.CellLeft + flxSPayment.Left
          txtPayDt.Width = flxSPayment.ColWidth(22)
          txtPayDt.Height = flxSPayment.RowHeight(iCurRow) - 15
          txtPayDt.Visible = True
          flxSPayment.ScrollBars = flexScrollBarNone
          txtPayDt.SetFocus
          SumUpTotal
       End If
    '  Move to the next row to enter amount
       If flxSPayment.row < flxSPayment.Rows And Not txtPayDt.Visible And flxSPayment.col = 11 Then
          If MoveDownPosition Then
             flxSPayment.col = 11
             flxSPayment.SetFocus
          Else
             cmdSavePayment.SetFocus
          End If
          'iCurRow = iCurRow - 1
       End If
   End If
End Sub

Private Sub txtPayDt_KeyPress(KeyAscii As Integer)
      If KeyAscii = 27 Then
      flxSPayment.SetFocus
      txtPayDt.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

   If KeyAscii = 13 Then
      If txtPayDt.text <> "" Then
         If TextBoxFormatDate(txtPayDt) Then
            If iFlxSPayCol = 22 Then 'if on paydate
               flxSPayment.TextMatrix(iCurRow, 22) = Format(txtPayDt.text, "dd/mm/yyyy")
               flxSPayment.TextMatrix(iCurRow, 23) = flxSPayment.TextMatrix(iCurRow, 22)
               flxSPayment.col = 23 'goto posting date and fit the text box
               flxSPayment_dblClick
             End If
'            ElseIf iFlxSPayCol = 24 Then
'               flxSPayment.TextMatrix(iCurRow, 22) = Format(txtRptDt.text, "dd/mm/yyyy")
'            Else
'                '13 Apr 2015 issue 530 Modified by anol
'               flxSPayment.TextMatrix(iCurRow, 25) = Format(txtRptDt.text, "dd/mm/yyyy")
'               If MoveDownPosition Then
'                    flxSPayment.col = 11
'                    flxSPayment.SetFocus
'               End If
'            End If
         End If
      Else
         flxSPayment.TextMatrix(iCurRow, 22) = ""
      End If

      txtPayDt.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If

   TextBoxKeyPrsDate txtPayDt, KeyAscii
End Sub

Private Sub txtSPayment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 45 Then
         KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      flxSPayment.SetFocus
      txtSPayment.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If
   
   DigitTextKeyPress txtSPayment, KeyAscii
End Sub

Private Sub txtSPayment_LostFocus()
   On Error Resume Next
   If Val(flxSPayment.TextMatrix(iCurRow, 10)) < Val(txtSPayment.text) Then
      MsgBox "Payment amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
      txtSPayment.text = "0.00"
      txtSPayment.SetFocus
      Exit Sub
   End If

   If txtSPayment.text = "" Then txtSPayment.text = "0.00"
   If flxSPayment.TextMatrix(iCurRow, 4) = flxSPayment.TextMatrix(iCurRow, 25) Then 'if they are equal that means they have allocation problem
        flxSPayment.TextMatrix(iCurRow, 11) = "0.00"
        MsgBox "A problem exists relating to a previous transaction entered against the selected Supplier: " & _
                     Chr(13) & _
                     "Please contact PCM Consulting. ", _
                     vbInformation + vbOKOnly, "Warning! Problem Transaction Found!"

   Else
        flxSPayment.TextMatrix(iCurRow, 11) = Format(txtSPayment.text, "0.00")
   End If
   
   SumUpTotal
   txtSPayment.Visible = False
   flxSPayment.ScrollBars = flexScrollBarVertical
End Sub
Private Function ChecksumValidationOnPayAllocation(adoConn As ADODB.Connection) As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    'this function is written by anol 20181205 when found that (issue 673 )Updated OS amount  extra incorrectly
    'This function shall prevent saving the data if when outstading amount on payment is not updated.
    'This functionshall compare allocation with payment amount and outstanding amount
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    Dim strWhere As String
    Dim strTenantID As String
    For i = 1 To flxSPayment.Rows - 1
        strTenantID = strTenantID & IIf(strTenantID = "", "'", ",'") & flxSPayment.TextMatrix(i, 4) & "'"
    Next i
    If strTenantID <> "" Then
        strWhere = " AND R.sageaccountnumber IN (" & strTenantID & " )"
    End If
'        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
'                    "Where a.ToTran = r.TransactionID  " & StrWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
                    
      rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
                  " ToTran from PayTransactions where DeleteFlag=False group By ToTran ) as A where A.ToTran=R.transactionID " & strWhere & " AND round((amount-amt),2)<>round(osamount,2)", adoConn, adOpenStatic, adLockReadOnly

        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            rsChecksum.MoveNext
        Wend
        
        rsChecksum.Close
        Set rsChecksum = Nothing
     
        If szTran2Fix = "" Then
                ChecksumValidationOnPayAllocation = True
        Else
                MsgBox "A problem occurred while writing allocation for this payment transaction: " & _
                     Chr(13) & szTran2Fix & "." & _
                     "Please contact PCM Consulting. This transaction has not been saved.", _
                     vbInformation + vbOKOnly, "Batch Payment not saved!"
        End If
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Private Sub SumUpTotal()
   Dim i As Integer, cGT As Currency

   For i = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(i, 3) = "Purchase Credit" Then
         cGT = cGT - Val(flxSPayment.TextMatrix(i, 11))
      Else
         cGT = cGT + Val(flxSPayment.TextMatrix(i, 11))
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
         If flxSPayment.TextMatrix(i, 15) = 7 Or flxSPayment.TextMatrix(i, 15) = 9 Then
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

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()
'   cboC.ListIndex = 0
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

Private Sub LoadSupplier(ByVal adoConn As ADODB.Connection, cboS As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, iTotalRow As Integer, j As Integer
   Dim i As Integer, iTotalCol As Integer, Data() As String

   On Error GoTo ErrorHandler

   If frmBPPreForm.optBP_BACS.Value Then
      szSQL = "SELECT SupplierID, SupplierName " & _
              "FROM Supplier " & _
              "WHERE PaymentType = 'BACS' " & _
              "ORDER BY SupplierName;"
   Else
      szSQL = "SELECT SupplierID, SupplierName  " & _
              "FROM Supplier " & _
              "WHERE PaymentType = 'CHQ' " & _
              "ORDER BY SupplierName;"
   End If

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox vbTab & "Either there are no supplier records entered in the system or " & vbCrLf & _
             vbTab & "there are no suppliers with payment type that matches you selection." & vbCrLf & vbCrLf & _
             "Please enter a supplier in the supplier module or set a payment type on the" & vbCrLf & vbCrLf & _
             vbTab & "supplier record for the supplier you wish to pay.", vbInformation + vbOKOnly, "Batch Payment"

      GoTo NoRes
   End If

   iTotalRow = adoRst.RecordCount
   iTotalCol = adoRst.Fields.Count

   ReDim Data(iTotalCol - 1, iTotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Suppliers"
   For i = 1 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboS.Column() = Data()
   cboS.ListIndex = 0

   txtSupplierName.text = "ALL / All Suppliers"

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRst = Nothing
End Sub

Private Sub txtDmdTenantSearchID_Change()
   Dim i As Integer

   If Len(txtDmdTenantSearchID.text) > 0 Then txtDmdTenantSearchName.text = ""

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 1), Len(txtDmdTenantSearchID.text))) <> UCase(txtDmdTenantSearchID.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtDmdTenantSearchName_Change()
   Dim i As Integer

   If Len(txtDmdTenantSearchName.text) > 0 Then txtDmdTenantSearchID.text = ""

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 2), Len(txtDmdTenantSearchName.text))) <> UCase(txtDmdTenantSearchName.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

'  Build up lessee's Account History
Private Sub SupplierAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "From tlbPayment " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaSupplierBalance(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type = 6 OR Type = 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalance(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBalance(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalance(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaSupplierBalance(1, i) = szaSupplierBalance(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaSupplierBalance(0, iIndex) = adoPayCr.Fields.Item("Cr").Value
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
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
