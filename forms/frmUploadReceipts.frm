VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUploadReceipts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Batch Receipts"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   17400
   Begin VB.Frame Frame6 
      Caption         =   "Matched Bank Statement Transactions"
      Height          =   1815
      Left            =   90
      TabIndex        =   68
      Top             =   8730
      Width           =   17160
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxStatement2 
         Height          =   1530
         Left            =   45
         TabIndex        =   69
         Top             =   225
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   2699
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorFixed  =   16761024
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   9
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
   End
   Begin VB.ComboBox cmbTenant 
      Height          =   315
      Left            =   2205
      TabIndex        =   14
      Top             =   6165
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ComboBox cmbAmountTypeGrid 
      Height          =   315
      Left            =   5355
      TabIndex        =   16
      Top             =   6165
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ComboBox cmbFundGrid 
      Height          =   315
      Left            =   3915
      TabIndex        =   15
      Top             =   6165
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   90
      TabIndex        =   56
      Top             =   10620
      Width           =   17115
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   14505
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   15705
         TabIndex        =   23
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdUnmatch 
         Caption         =   "&Unmatch"
         Height          =   375
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdMatch 
         Caption         =   "&Match"
         Height          =   375
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdateReference 
         Caption         =   "&Save Reference"
         Height          =   375
         Left            =   13005
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         Width           =   1395
      End
      Begin VB.CommandButton cmdUnmatchAll 
         Caption         =   "Unmatch &All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearReference 
         Caption         =   "Clear Reference"
         Height          =   375
         Left            =   11475
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1395
      End
   End
   Begin VB.PictureBox picDmdLeaseList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   10125
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   42
      Top             =   2205
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
         TabIndex        =   47
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
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   0
            Width           =   465
         End
         Begin MSForms.ComboBox ComboBox2 
            Height          =   315
            Left            =   3675
            TabIndex        =   49
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
            TabIndex        =   48
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
         TabIndex        =   46
         Top             =   300
         Width           =   1935
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   52
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant ID"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   55
         Top             =   70
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Name"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   54
         Top             =   70
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   53
         Top             =   75
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select the Input File Criteria"
      Height          =   1875
      Left            =   90
      TabIndex        =   27
      Top             =   90
      Width           =   17250
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   480
         Left            =   13050
         TabIndex        =   7
         Top             =   990
         Width           =   1035
      End
      Begin VB.ComboBox cmbFund 
         Height          =   315
         Left            =   12195
         TabIndex        =   5
         Top             =   630
         Width           =   1905
      End
      Begin VB.OptionButton optDueDate 
         Appearance      =   0  'Flat
         Caption         =   "By Due Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   675
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optInvoiceDate 
         Appearance      =   0  'Flat
         Caption         =   "By Invoice Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   29
         Top             =   675
         Width           =   1695
      End
      Begin VB.CheckBox chkSearchPattern 
         Appearance      =   0  'Flat
         Caption         =   "Search for the transactions where the search criteria matches any part of the reference in the bank statement transactions."
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4545
         TabIndex        =   28
         ToolTipText     =   $"frmUploadReceipts.frx":0000
         Top             =   1485
         Value           =   1  'Checked
         Width           =   9240
      End
      Begin VB.CommandButton cmdBrowseFile 
         Caption         =   "..."
         Height          =   390
         Left            =   12555
         TabIndex        =   6
         Top             =   1065
         Width           =   375
      End
      Begin MSForms.CommandButton cmdAll 
         Height          =   330
         Left            =   3960
         TabIndex        =   61
         Top             =   1440
         Width           =   480
         Caption         =   "All"
         Size            =   "847;582"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Tenant"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   41
         Top             =   1530
         Width           =   1290
      End
      Begin MSForms.TextBox txtTenantName 
         Height          =   285
         Left            =   1800
         TabIndex        =   40
         Top             =   1485
         Width           =   1725
         VariousPropertyBits=   746604575
         BackColor       =   16777215
         BorderStyle     =   1
         Size            =   "3043;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdDmdTenantLookup 
         Height          =   285
         Left            =   3510
         TabIndex        =   39
         Top             =   1485
         Width           =   390
         Caption         =   """"
         Size            =   "688;503"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Payment Type"
         Height          =   195
         Index           =   11
         Left            =   10485
         TabIndex        =   38
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Default Fund"
         Height          =   195
         Index           =   5
         Left            =   10485
         TabIndex        =   37
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select &Bank:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   35
         Top             =   315
         Width           =   840
      End
      Begin MSForms.ComboBox cmbBank 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   2700
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4762;503"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;4233"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Match Record By:"
         Height          =   195
         Index           =   0
         Left            =   6000
         TabIndex        =   34
         Top             =   270
         Width           =   1185
      End
      Begin MSForms.ComboBox cmbMatchBy 
         Height          =   285
         Left            =   7440
         TabIndex        =   1
         Top             =   270
         Width           =   2700
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4762;503"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;4233"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Transaction Aging:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   33
         Top             =   675
         Width           =   1305
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Receipt Allocation:"
         Height          =   195
         Index           =   3
         Left            =   6000
         TabIndex        =   32
         Top             =   675
         Width           =   1320
      End
      Begin MSForms.ComboBox cmbReceiptAllocation 
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Top             =   630
         Width           =   2700
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4762;503"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;4233"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input &File (.csv, .xls):"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   31
         Top             =   1125
         Width           =   1395
      End
      Begin MSForms.TextBox txtInputFile 
         Height          =   300
         Left            =   1800
         TabIndex        =   30
         Top             =   1065
         Width           =   10680
         VariousPropertyBits=   679495707
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "18838;529"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbRptAmtType 
         Height          =   315
         Left            =   12195
         TabIndex        =   4
         Top             =   270
         Width           =   1905
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3360;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1762"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sales Invoice Transactions"
      Height          =   2955
      Left            =   90
      TabIndex        =   25
      Tag             =   "(Double click to enter the default reference for the tenant and press ENTER. Applicable when matching transactions by Reference)"
      Top             =   1980
      Width           =   17205
      Begin VB.CheckBox chkUnmachInvoice 
         Appearance      =   0  'Flat
         Caption         =   "Show Unmatched Records Only"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   13860
         TabIndex        =   67
         ToolTipText     =   $"frmUploadReceipts.frx":00AA
         Top             =   0
         Width           =   2850
      End
      Begin VB.CheckBox chkMatchedRec 
         Appearance      =   0  'Flat
         Caption         =   "Show Only Matched records"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11160
         TabIndex        =   65
         ToolTipText     =   $"frmUploadReceipts.frx":0154
         Top             =   0
         Width           =   2400
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Clear selection"
         Height          =   300
         Left            =   15705
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chkShowAllCol 
         Appearance      =   0  'Flat
         Caption         =   "All Columns"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9765
         TabIndex        =   36
         Top             =   0
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.TextBox txtRefInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10680
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
         Height          =   2595
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   17010
         _ExtentX        =   30004
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   9
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
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankTransactions 
      Height          =   2340
      Left            =   180
      TabIndex        =   13
      Top             =   5580
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   4128
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   16761024
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   9
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
      Caption         =   "Unmatched Bank Statement Transactions"
      Height          =   3015
      Left            =   90
      TabIndex        =   24
      Top             =   4995
      Width           =   17205
      Begin VB.CheckBox chkUnassignedtransactions 
         Appearance      =   0  'Flat
         Caption         =   "Show unassigned transactions Only"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   13950
         TabIndex        =   66
         ToolTipText     =   $"frmUploadReceipts.frx":01FE
         Top             =   225
         Width           =   3120
      End
      Begin VB.CheckBox chkStatement 
         Appearance      =   0  'Flat
         Caption         =   "Show Only Assigned transactions"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11160
         TabIndex        =   12
         ToolTipText     =   $"frmUploadReceipts.frx":02A8
         Top             =   225
         Width           =   2895
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sales Receipts:"
         Height          =   195
         Index           =   2
         Left            =   6165
         TabIndex        =   63
         Top             =   225
         Width           =   1515
      End
      Begin MSForms.TextBox txtGrossTotal 
         Height          =   300
         Left            =   7935
         TabIndex        =   11
         Top             =   225
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
      Begin MSForms.TextBox txtSearchMemo 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   260
         Width           =   1725
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         Size            =   "3043;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Statement Ref"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   62
         Top             =   270
         Width           =   1575
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   330
         Left            =   4365
         TabIndex        =   10
         Top             =   225
         Width           =   480
         Caption         =   "All"
         Size            =   "847;582"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.TextBox txtGrandTotal 
      Height          =   300
      Left            =   7845
      TabIndex        =   60
      Top             =   8370
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total :"
      Height          =   195
      Index           =   1
      Left            =   5580
      TabIndex        =   59
      Top             =   8415
      Width           =   930
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipts on account:"
      Height          =   195
      Index           =   0
      Left            =   5580
      TabIndex        =   58
      Top             =   8055
      Width           =   1935
   End
   Begin MSForms.TextBox txtReceiptsOnAccount 
      Height          =   300
      Left            =   7845
      TabIndex        =   57
      Top             =   8010
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
End
Attribute VB_Name = "frmUploadReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified By Asif. Date: 03 Apr 2015
Option Explicit

Dim ccBanks As New Collection
'Dim ccUnitColumns As New Collection
'Dim ccLesseeColumns As New Collection
Dim ccDateColumns As New Collection
Dim ccReferenceColumns As New Collection
Dim ccAmountColumns As New Collection
Dim ccExt1 As New Collection
Dim ccExt2 As New Collection
Public bankAmountColIndex As Integer
Dim bankReferenceColIndex As Integer
Dim bankReferenceColIndexDev As Integer ' Column position in the grid has a deviation with statement file because I am deleting the column in the array
Dim postingDateColIndex As Integer
Dim iFlxSPayCol As Integer
Dim Filename As String
Dim LastSelGridnumber As Integer ' This variable shall contain last selected grid (ranges 1,2,3)
Dim iCurRow As Integer
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public HasRecforBatchReceipt As Boolean
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Function FillBatchReceipts(adoconn As ADODB.Connection, filepath As String) As Boolean
    ' Note Anol 20161128
    ' While reading the statement file I am not loading all the information of the statement file in the grid
    ' Rather I am keeping the row index in StatementCol at column 12 which shall be used for writtingnew unused staement line
    ' FOR Bank TSB, I can show maximum 5 columns as per structure of this procedure
    ' So I am Deleting the un neccessary  row from the array which loads statement sColValue()
    Dim ifullAllocated As Integer
'    Dim con As New ADODB.Connection
'    Dim rsStatementRecords As New ADODB.Recordset
    Dim rsBatchReceipts As New ADODB.Recordset
    Dim rsMasterBatchReceipt As New ADODB.Recordset
    Dim X
    
    
    Dim directory As String
    Dim szHeader As String
    Dim szFilter As String
    
    Dim oFSTR As Scripting.TextStream
    Dim sColValue() As String
    Dim sLine As String
    Dim lCtr As Long
     
    flxBankTransactions.Clear
    flxStatement2.Clear
'    On Error GoTo ErrorHandler
    
    If Dir(filepath, vbDirectory) = vbNullString Then
        MsgBox "The file does not exist. Please select a valid file.", vbInformation, "Invalid Directory"
        FillBatchReceipts = False
        Exit Function
    End If
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set oFSTR = fso.OpenTextFile(filepath) 'Reading the CSV file into a text stream.
    directory = fso.GetParentFolderName(filepath)
    Filename = fso.GetFileName(filepath)
    
    If CheckIfUploaded(adoconn) Then
       MsgBox "WARNING: The selected bank statement file is already processed."
       FillBatchReceipts = False
       Exit Function
    End If
    If cmbBank.text = "BIB" Then
       
        'Deleting first line on csv file.
        X = DeleteLine(txtInputFile.text, 1)
    End If
    If cmbBank.text = "Santander" Then
'         x = DeleteLineSantandar(txtInputFile.text)
        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
        Exit Function
    End If
'    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'              "Data Source=" & "" & directory & ";" & _
'                "Extended Properties=""Text;HDR=Yes;"""
    flxBankTransactions.Cols = 11
'    If cmbBank.text = "Barclays" Then
''        rsStatementRecords.Open "Select Date,Memo,Amount, Subcategory, Number,account from [" & Filename & "]", con, adOpenStatic, adLockReadOnly
'        flxBankTransactions.Cols = 5 + 4
'    ElseIf cmbBank.text = "BIB" Then
'     '[ENTRY DATE] , description, amount, [TLA CODE], [CHEQUE NO], [ACC ID]
'     'below line comment out on 20161128 By anol
''        rsStatementRecords.Open "Select [ENTRY DATE] as EDate , description, amount, [TLA CODE], [CHEQUE NO], [ACC ID] from [" & Filename & "] ", con, adOpenStatic, adLockReadOnly '
'    '    rsStatementRecords.Open "Select [F7],[F13],[F9],[F10],[F11],[F2] from [" & Filename & "] where [F10] is not null Or [F10]<>'DESCRIPTION'", con, adOpenStatic, adLockReadOnly '
'    '    rsStatementRecords.Close
'       ' rsStatementRecords.Open "Select * from [" & Filename & "]", con, adOpenStatic, adLockReadOnly
'        flxBankTransactions.Cols = 6 + 4
'    ElseIf cmbBank.text = "TSB" Then
'    'Format$([Sold Count], '#') as  [Sold Count]
''        rsStatementRecords.Open "Select [Transaction Date] as EDate ,CSTR([Transaction Description]) as description, [Credit Amount] as amount, [Transaction Type]," & _
'            "[Account Number] from [" & Filename & "]", con, adOpenStatic, adLockReadOnly
'          flxBankTransactions.Cols = 5 + 4
'   Else
    If cmbBank.text = "Barclays" Or cmbBank.text = "TSB" Or cmbBank.text = "Lloyds" Or cmbBank.text = "BIB" Or cmbBank.text = "Santander" Or cmbBank.text = "HSBC" Then
    Else
        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
        Exit Function
    End If
    
    
    '*****************************rsMasterBatchReceipts****************************
    Dim szSQL As String
       szSQL = "SELECT Rt.TransactionID, Rt.SlNumber AS C_SL, Rt.DemandRef, Rt.AdjTag, " & _
                    "Rt.SageAccountNumber, Rt.UnitID, Rt.DDate, '' as Reference, Rt.Details, Rt.Amount, " & _
                    "Rt.OSAmount, '0.00' as Receipt, '' as TranDate, Rt.Type, TT.DESCRIPTION, T.Name, " & _
                    "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, U.PropertyID, 'No' as Match, " & _
                    "T.Ref " & _
               "FROM (((tlbReceipt AS Rt INNER JOIN tlbTransactionTypes AS TT ON Rt.Type = TT.TYPE_ID) " & _
                    "INNER JOIN Tenants AS T ON Rt.SageAccountNumber = T.SageAccountNumber) " & _
                    "INNER JOIN Units AS U ON Rt.UnitID = U.UnitNumber) INNER JOIN " & _
                    "Property AS P ON U.PropertyID = P.PropertyID " & _
               "WHERE Rt.OSAmount > 0 AND " & _
                    "Rt.ReceiptView = True AND " & _
                    "(TT.TYPE_ID = 1) AND " & _
                    "P.ClientID = '" & frmBRPreForm.txtClient.Tag & "' "
    szFilter = ""
    
    'Filter the transactions based on either Due Date or Invoice Date
    If optDueDate.Value = True Then
        szFilter = " ORDER BY T.SageAccountNumber,DDate "
    Else
        szFilter = " ORDER BY T.SageAccountNumber,DDate "
    End If
    
    'Order the transactions based on either ascending or descending
    If cmbReceiptAllocation.Value = "1" Then
        szFilter = szFilter & " ASC "
    ElseIf cmbReceiptAllocation.Value = "2" Then
        szFilter = szFilter & " DESC "
    Else
        szFilter = ""
    End If
    
    szSQL = szSQL & szFilter & ";"
       rsMasterBatchReceipt.Open szSQL, adoconn, adOpenStatic, adLockBatchOptimistic
    Dim i As Integer
    flxBankTransactions.Rows = 2
    If cmbBank.text = "Barclays" Then
           'Date,Memo,Amount
           'Number  Date    Account Amount  Subcategory Memo
           'Date    Account Amount  Subcategory Memo
           postingDateColIndex = 1 ''location in the grid is 1 but the statement file that is 0
           bankReferenceColIndex = 5
           bankAmountColIndex = 3
           szHeader$ = "O/S Amt £|<Date|<Account|<Amount|<Subcategory|<Memo|<StatementRowNo|<Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "BIB" Then
            postingDateColIndex = 2
            bankReferenceColIndex = 5
            bankAmountColIndex = 3
'            Exit Function
            'GROUP   ACC ID  ACCOUNT NO  TYPE    BANK CODE   CURR    ENTRY DATE  AS AT   AMOUNT  TLA CODE    CHEQUE NO   STATUS  DESCRIPTION
            szHeader$ = "O/S Amt £|<ACCOUNT NO|< ENTRY DATE|< AMOUNT|< TLA CODE|< DESCRIPTION|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "TSB" Then
            'in the file
            '1Transaction Date    2Transaction Type    3Sort Code   4Account Number  5Transaction Description 6Debit Amount    7Credit Amount   8Balance
            '1Transaction Date    2Transaction Type       4Account Number  5Transaction Description     7Credit Amount
            ' rsStatementRecords.Open "Select [Transaction Date] as EDate ,CSTR([Transaction Description]) as description, [Credit Amount] as amount, [Transaction Type]," & _
            "[Account Number] from [" & Filename & "]", con, adOpenStatic, adLockReadOnly
            'date index in the actual statement file
            postingDateColIndex = 1 'location in the grid is 1 but the statement file that is 0
            bankReferenceColIndex = 4
            bankAmountColIndex = 5 ' THIS 5 IS THE LOCATION OF GRID AMOUNT, IN THE FILE IT WILL BE 4
            'F7,F13,F9,F10,F11
    'Date    Memo    Amount  Subcategory Number
    'ENTRY DATE DESCRIPTION AMOUNT  TLA CODE    CHEQUE NO
'            szHeader$ = "O/S Amt £|< Transaction Date|< Transaction Description|< Credit Amount|< Transaction Type|< Account Number|< ACCOUNT|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
            szHeader$ = "O/S Amt £|<Transaction Date|< Transaction Type|< Account Number|<  Transaction Description|< Credit Amount|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
'            szHeader$ = "O/S Amt £|< Transaction Date|<Transaction Type|<Sort Code|<Account Number|<Transaction Description|<Debit Amount|<Credit Amount|<Balance|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "Lloyds" Then
            'In File Header "Transaction Date    Transaction Type    Sort Code   Account Number  Transaction Description Debit Amount    Credit Amount   Balance" same as TSB
             szHeader$ = "O/S Amt £|<Transaction Date|<Transaction Type|<Account Number|<Transaction Description|<Credit Amount|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 1 'location in the grid is 1 but the statement file that is 0
             bankReferenceColIndex = 4 '356
             bankAmountColIndex = 5 ' THIS 5 IS THE LOCATION OF GRID AMOUNT
    ElseIf cmbBank.text = "Santander" Then
             'In File Header "Date    Narrative   Transaction Type    Debit   Credit  Current Balance
             szHeader$ = "O/S Amt £|< Date|<Narrative |<Transaction Type|<Credit Amount|<Current Balance|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 1 'location in the grid is 1 but the statement file that is 0
             bankReferenceColIndex = 2 '356
             bankAmountColIndex = 4 ' THIS 5 IS THE LOCATION OF GRID AMOUNT
    ElseIf cmbBank.text = "HSBC" Then
             'In File Header "Date    Type    Description Paid out    Paid in Balance
             szHeader$ = "O/S Amt £|< Date|<Type |<Description|<Paid in|<Balance|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 0 'I have taken this index from the receipt file
             bankReferenceColIndex = 2 'I have taken this index from the receipt file
             bankAmountColIndex = 4 ' I have taken this index from the receipt file
             bankReferenceColIndexDev = 1
    End If
    flxBankTransactions.FormatString = szHeader$
    flxBankTransactions.ColWidth(6) = 0 '7 StatementRowNo
    flxBankTransactions.ColWidth(7) = 1700 '7 Tenant ID
    flxBankTransactions.ColWidth(8) = 1800 '8 Fund Code
    flxBankTransactions.ColWidth(9) = 1500 '9 Amount Type
    flxBankTransactions.ColWidth(10) = 1500 '10 Assigned
    flxBankTransactions.ColWidth(11) = 0 '11 TransID
    flxBankTransactions.ColAlignment(bankAmountColIndex) = vbRightJustify
'    Exit Function
    'added by anol 20161021
    flxStatement2.Cols = flxBankTransactions.Cols
    flxStatement2.Rows = 2
    flxStatement2.FormatString = szHeader$
    flxStatement2.ColWidth(6) = 0 '7 Account
    flxStatement2.ColWidth(7) = 1700 '7 Tenant ID
    flxStatement2.ColWidth(8) = 1800 '8 Fund Code
    flxStatement2.ColWidth(9) = 1500 '9 Amount Type
    flxStatement2.ColWidth(10) = 1500 '10 Assigned
    flxStatement2.ColWidth(11) = 0 '11 TransID
    flxStatement2.ColAlignment(bankAmountColIndex) = vbRightJustify
    'End of addition
    
    Dim referenceValue, record, unit, lessee, filter As String
    Dim iBankTranRow As Integer
    iBankTranRow = 1
    ifullAllocated = 1
    filter = ""
    
    Dim receiptBalance, receiptAmount As Double
    receiptBalance = 0
    receiptAmount = 0
    lCtr = 1
'    Do While Not rsStatementRecords.EOF() 'loop with each line of bank statement
        Do While Not oFSTR.AtEndOfStream
                sLine = oFSTR.ReadLine
               ' Debug.Print sLine
                'GoTo NextReceipt
           If lCtr = 1 Then
                'Compare the header for diffrent banks
                If cmbBank.text = "TSB" Then
                    If UCase(Trim(sLine)) = UCase("Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance") Then
                        'so Header format is fine here
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                ElseIf cmbBank.text = "Barclays" Then
                    If UCase(Trim(sLine)) = UCase("Number,Date,Account,Amount,Subcategory,Memo") Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                 ElseIf cmbBank.text = "Lloyds" Then
                    If UCase(Trim(sLine)) = UCase("Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance") Or _
                    UCase(Trim(sLine)) = UCase("Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance,") Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                 ElseIf cmbBank.text = "BIB" Then
                    If UCase(Trim(sLine)) = UCase("GROUP,ACC ID,ACCOUNT NO,TYPE,BANK CODE,CURR,ENTRY DATE,AS AT,AMOUNT,TLA CODE,CHEQUE NO,STATUS,DESCRIPTION") Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                 ElseIf cmbBank.text = "Santander" Then
                    If UCase(Trim(sLine)) = UCase("Date,Narrative,Transaction Type,Debit,Credit,Current Balance") Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                 ElseIf cmbBank.text = "HSBC" Then
                    If UCase(Trim(sLine)) = UCase("Date,Type,Description,Paid out,Paid in,Balance") Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                    End If
                 Else
                       MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Function
                 End If
           Else
            '    reference = Split(rsStatementRecords.Fields(referenceCol).Value, " "Lloyds
                'null reference shall raise an error . this has benn fixed by anol 20161125
'                Debug.Print rsStatementRecords.AbsolutePosition
'                Debug.Print rsStatementRecords.Fields(referenceCol).Value
'                Debug.Print rsStatementRecords.Fields.count
'                If IsNull(rsStatementRecords.Fields(referenceCol).Value) Then
'                    GoTo NextReceipt
'                End If
                 sColValue = Split(sLine, ",")
'                referenceValue = rsStatementRecords.Fields(referenceCol).Value
                If cmbBank.text = "TSB" Then
                    'making the array as same format as grid
                    DeleteArrayItem sColValue, 2
                    DeleteArrayItem sColValue, 4
                    DeleteArrayItem sColValue, 5
                ElseIf cmbBank.text = "Barclays" Then
                    DeleteArrayItem sColValue, 0 'suppresing null column
                ElseIf cmbBank.text = "Lloyds" Then
''                    DeleteArrayItem sColValue, 3
''                    DeleteArrayItem sColValue, 5
''                    'DeleteArrayItem sColValue, 6
                    'making the array as same format as grid
                    DeleteArrayItem sColValue, 2
                    DeleteArrayItem sColValue, 4
                    DeleteArrayItem sColValue, 5
                ElseIf cmbBank.text = "BIB" Then
                    DeleteArrayItem sColValue, 0
                    DeleteArrayItem sColValue, 0
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 2
                    DeleteArrayItem sColValue, 4
                    DeleteArrayItem sColValue, 4
                ElseIf cmbBank.text = "Santander" Then
                    DeleteArrayItem sColValue, 3
                ElseIf cmbBank.text = "HSBC" Then
                    DeleteArrayItem sColValue, 3 'deleting paid out
                End If
                 referenceValue = sColValue(bankReferenceColIndex - 1)
                 'issue 320
                ' Upload Batch receipts error
                '(Support – WPM)
                'An error occurs when the user tries to upload a batch receipts file if there was a appostrophe in the value
                 referenceValue = Replace(referenceValue, "'", "''")
'                 Debug.Print referenceValue
                'Exit Function
                If IsNull(sColValue(bankReferenceColIndex - 1)) Or sColValue(bankReferenceColIndex - 1) = "" Then
                    GoTo NextReceipt
                End If
                If IsNull(sColValue(bankAmountColIndex - 1)) Or sColValue(bankAmountColIndex - 1) = "" Then
                    GoTo NextReceipt
                End If
'                bankAmountColIndex = 4
                receiptAmount = Val(sColValue(bankAmountColIndex - 1))
'                Exit Function
                If receiptAmount <= 0 Then
                    GoTo NextReceipt
                End If
                
                receiptBalance = receiptAmount
       '**************************rsBatchReceipts***************************
               szSQL = "SELECT DD.* FROM " & _
               "(SELECT Rt.TransactionID, Rt.SlNumber AS C_SL, Rt.DemandRef, Rt.AdjTag, " & _
                            "Rt.SageAccountNumber, Rt.UnitID, Rt.DDate, '' as Reference, Rt.Details, Rt.Amount, " & _
                            "Rt.OSAmount, '0.00' as Receipt, '' as TranDate, Rt.Type, TT.DESCRIPTION, T.Name, " & _
                            "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, U.PropertyID, 'No' as Match, " & _
                            "T.Ref,'' AS  FlagTran " & _
                       "FROM (((tlbReceipt AS Rt INNER JOIN tlbTransactionTypes AS TT ON Rt.Type = TT.TYPE_ID) " & _
                            "INNER JOIN Tenants AS T ON Rt.SageAccountNumber = T.SageAccountNumber) " & _
                            "INNER JOIN Units AS U ON Rt.UnitID = U.UnitNumber) INNER JOIN " & _
                            "Property AS P ON U.PropertyID = P.PropertyID " & _
                       "WHERE Rt.OSAmount > 0 AND " & _
                            "Rt.ReceiptView = True AND " & _
                            "(TT.TYPE_ID = 1) AND " & _
                            "P.ClientID = '" & frmBRPreForm.txtClient.Tag & "') AS DD "
               
               szFilter = ""
'            TT.TYPE_ID = 1 means sales invoice, PF means short form of transaction type
               If chkSearchPattern.Value = 0 Then
                    If cmbMatchBy.Value = "1" Then
                        szFilter = " WHERE SageAccountNumber = '" & CStr(referenceValue) & "' OR Ref = '" & CStr(referenceValue) & "'"
                    ElseIf cmbMatchBy.Value = "2" Then
                        szFilter = " WHERE UnitID = '" & CStr(referenceValue) & "' OR Ref = '" & CStr(referenceValue) & "'"
                    ElseIf cmbMatchBy.Value = "3" Then
                        szFilter = " WHERE Ref = '" & CStr(referenceValue) & "' OR Ref = '" & CStr(referenceValue) & "'"
                    Else
                        szFilter = " WHERE (SageAccountNumber = '" & CStr(referenceValue) & "' OR UnitID = '" & CStr(referenceValue) & "' OR Ref = '" & CStr(referenceValue) & "') "
                    End If
                Else
                    If cmbMatchBy.Value = "1" Then
                        'szFilter = " WHERE Instr(1,'" + referenceValue + "', SageAccountNumber) >= 1 "
'                        szFilter = " WHERE Instr(1,'" + referenceValue + "', SageAccountNumber) >= 1 Or T.Ref='" + referenceValue + "' "
                    'below line was showing error when not converting reference into string fixed by anol 20161125
                        szFilter = " WHERE Instr(1,'" + CStr(referenceValue) + "', SageAccountNumber) >= 1 Or T.Ref='" + CStr(referenceValue) + "' "
                       
                    ElseIf cmbMatchBy.Value = "2" Then
                        szFilter = " WHERE Instr(1,'" + CStr(referenceValue) + "', UnitID) >= 1  "
                    ElseIf cmbMatchBy.Value = "3" Then
                        szFilter = " WHERE Len(Ref) > 0 AND Instr(1,'" + CStr(referenceValue) + "', Ref) >= 1 Or T.Ref='" + CStr(referenceValue) + "'"
                    Else
                        szFilter = " WHERE Instr(1,'" + CStr(referenceValue) + "', SageAccountNumber) >= 1 OR Instr(1,'" + CStr(referenceValue) + "', UnitID) >= 1 OR Instr(1,'" + CStr(referenceValue) + "', Ref) >= 1 Or T.Ref='" + CStr(referenceValue) + "' "
                    End If
                End If
            
                  'Filter the transactions based on either Due Date or Invoice Date
                If optDueDate.Value = True Then
                    szFilter = szFilter & " ORDER BY DDate "
                Else
                    szFilter = szFilter & " ORDER BY DDate "
                End If
            
                'Order the transactions based on either ascending or descending
                If cmbReceiptAllocation.Value = "1" Then
                    szFilter = szFilter & " ASC "
                ElseIf cmbReceiptAllocation.Value = "2" Then
                    szFilter = szFilter & " DESC "
                Else
                    'Get invoices in the order in which they are displayed based on the grid.
            '          szFilter = ""
                End If
                
                szSQL = szSQL & szFilter & ";"
                rsBatchReceipts.Open szSQL, adoconn, adOpenStatic, adLockBatchOptimistic

                
                Dim matchedBatchReceiptFound As Integer
                matchedBatchReceiptFound = 0
                
                While Not rsBatchReceipts.EOF
                    'rsBatchReceipts is the record set that is filtered by Reference or other criteria fro each line of the statement
'                    rsBatchReceipts not EOF means it has some reletion or cause for allocation
                    If receiptBalance > 0 Then ' The amount got from the statement after allocation
                        rsMasterBatchReceipt.filter = " TransactionID = '" & rsBatchReceipts!TransactionID & "'"
                        While Not rsMasterBatchReceipt.EOF
                            If rsMasterBatchReceipt!match = "Yes" Then
                                matchedBatchReceiptFound = matchedBatchReceiptFound + 1
                            Else
                                If rsBatchReceipts!OSAmount < receiptBalance Then
                                   rsMasterBatchReceipt!receipt = rsBatchReceipts!OSAmount
                                Else
                                   rsMasterBatchReceipt!receipt = receiptBalance
                                End If
                                receiptBalance = receiptBalance - rsBatchReceipts!OSAmount
                                rsMasterBatchReceipt!Reference = sColValue(bankReferenceColIndex - 1) 'rsStatementRecords.Fields(referenceCol).Value changed by anol 20161128
                                rsMasterBatchReceipt!match = "Yes"
                                rsBatchReceipts!FlagTran = rsMasterBatchReceipt!PF & IIf(IsNull(rsMasterBatchReceipt!C_SL), "", rsMasterBatchReceipt!C_SL)
                                rsMasterBatchReceipt!TranDate = Format(CDate(sColValue(postingDateColIndex - 1)), "dd/mm/yyyy") 'Format(CDate(rsStatementRecords.Fields(dateCol).Value), "dd/mm/yyyy")changed by anol 20161128
                            End If
                            rsMasterBatchReceipt.MoveNext
                        Wend
                    End If
                    rsBatchReceipts.MoveNext
                Wend
                
                If rsBatchReceipts.RecordCount = matchedBatchReceiptFound Then
                   
                    flxBankTransactions.RowHeight(iBankTranRow) = 280
'                    For i = 0 To rsStatementRecords.Fields.count - 1
                    For i = 0 To 4 'used generally -1, now 4+2=5
                        ' Filling single line of second grid one by one all cells
'                        Debug.Print sColValue(i)
'                        Exit Function
                       flxBankTransactions.TextMatrix(iBankTranRow, i + 1) = IIf(IsNull(sColValue(i)), "", sColValue(i)) 'rsStatementRecords.Fields(i).Value
                    Next i
                    
                    flxBankTransactions.TextMatrix(iBankTranRow, 0) = Format(sColValue(bankAmountColIndex - 1), "0.00") 'rsStatementRecords.Fields(amountCol).Value 'IsNull ommited
                    flxBankTransactions.TextMatrix(iBankTranRow, bankAmountColIndex) = Format(flxBankTransactions.TextMatrix(iBankTranRow, bankAmountColIndex), "0.00")
                    flxBankTransactions.TextMatrix(iBankTranRow, 6) = lCtr
                    If rsBatchReceipts.RecordCount > 0 Then
                      rsBatchReceipts.MoveFirst
                          If Len(IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)) > 0 Then
                              flxBankTransactions.TextMatrix(iBankTranRow, 7) = IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)
                              flxBankTransactions.TextMatrix(iBankTranRow, 10) = "Yes"
                              While Not rsBatchReceipts.EOF
                                    If flxBankTransactions.TextMatrix(iBankTranRow, 11) <> "" Then
                                        flxBankTransactions.TextMatrix(iBankTranRow, 11) = flxBankTransactions.TextMatrix(iBankTranRow, 11) + "," + rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    Else
                                        flxBankTransactions.TextMatrix(iBankTranRow, 11) = rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    End If
                                    rsBatchReceipts.MoveNext
                              Wend
                          End If
                      End If
                     'flxBankTransactions.ColWidth(7) = 1700 '7 Tenant ID
                     flxBankTransactions.TextMatrix(iBankTranRow, 8) = cmbFund.text
                     flxBankTransactions.TextMatrix(iBankTranRow, 9) = cmbRptAmtType.text
                     If flxBankTransactions.TextMatrix(iBankTranRow, 10) = "" Then
                          flxBankTransactions.TextMatrix(iBankTranRow, 10) = "No"
                     End If
                    flxBankTransactions.AddItem ""
                    iBankTranRow = iBankTranRow + 1
                End If
                
                If rsBatchReceipts.RecordCount > 0 And rsBatchReceipts.RecordCount > matchedBatchReceiptFound And receiptBalance > 0 Then
                     ' Filling single line of second grid one by one all cells
                    
                     flxBankTransactions.RowHeight(iBankTranRow) = 280
'                    For i = 0 To rsStatementRecords.Fields.count - 1
                     For i = 0 To 4
                       flxBankTransactions.TextMatrix(iBankTranRow, i + 1) = sColValue(i) 'IIf(IsNull(rsStatementRecords.Fields(i).Value), "", rsStatementRecords.Fields(i).Value)
                    Next i
                    flxBankTransactions.TextMatrix(iBankTranRow, 0) = Format(receiptBalance, "0.00") 'Here is the difference
                    flxBankTransactions.TextMatrix(iBankTranRow, bankAmountColIndex) = Format(flxBankTransactions.TextMatrix(iBankTranRow, bankAmountColIndex), "0.00")
                    'Putting the statement Col reference
                    flxBankTransactions.TextMatrix(iBankTranRow, 6) = lCtr
                    
                        If rsBatchReceipts.RecordCount > 0 Then
                            rsBatchReceipts.MoveFirst
                            If Len(IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)) > 0 Then
                                flxBankTransactions.TextMatrix(iBankTranRow, 7) = IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)
                                flxBankTransactions.TextMatrix(iBankTranRow, 10) = "Yes"
                               While Not rsBatchReceipts.EOF
                                    If flxBankTransactions.TextMatrix(iBankTranRow, 11) <> "" Then
                                        flxBankTransactions.TextMatrix(iBankTranRow, 11) = flxBankTransactions.TextMatrix(iBankTranRow, 11) + "," + rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    Else
                                        flxBankTransactions.TextMatrix(iBankTranRow, 11) = rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    End If
                                    rsBatchReceipts.MoveNext
                              Wend
                            End If
                        End If
                        'Putting default Fund
                         flxBankTransactions.TextMatrix(iBankTranRow, 8) = cmbFund.text
                         flxBankTransactions.TextMatrix(iBankTranRow, 9) = cmbRptAmtType.text
                         If flxBankTransactions.TextMatrix(iBankTranRow, 10) = "" Then
                              flxBankTransactions.TextMatrix(iBankTranRow, 10) = "No"
                         End If
                         flxBankTransactions.AddItem ""
                         iBankTranRow = iBankTranRow + 1
               End If
               If rsBatchReceipts.RecordCount > 0 And rsBatchReceipts.RecordCount > matchedBatchReceiptFound And receiptBalance <= 0 Then 'receiptBalance <  0 has been changed to  receiptBalance <= 0 by anol 20161128
                       
                    flxStatement2.TextMatrix(ifullAllocated, 0) = 0
                    flxStatement2.RowHeight(ifullAllocated) = 280
'                    For i = 0 To rsStatementRecords.Fields.count -     1
                    For i = 0 To 4
                       flxStatement2.TextMatrix(ifullAllocated, i + 1) = sColValue(i) 'IIf(IsNull(rsStatementRecords.Fields(i).Value), "", rsStatementRecords.Fields(i).Value)
                    Next i
                     flxStatement2.TextMatrix(ifullAllocated, 0) = Format(flxStatement2.TextMatrix(ifullAllocated, 0), "0.00")
                      flxStatement2.TextMatrix(ifullAllocated, bankAmountColIndex) = Format(flxStatement2.TextMatrix(ifullAllocated, bankAmountColIndex), "0.00")
                       flxStatement2.TextMatrix(ifullAllocated, 6) = lCtr
                            rsBatchReceipts.MoveFirst
                            If Len(IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)) > 0 Then
                                flxStatement2.TextMatrix(ifullAllocated, 7) = IIf(IsNull(rsBatchReceipts.Fields("SageaccountNumber").Value), "", rsBatchReceipts.Fields("SageaccountNumber").Value)
'                                Debug.Print flxStatement2.TextMatrix(ifullAllocated, 7)
                                flxStatement2.TextMatrix(ifullAllocated, 10) = "Yes"
                               While Not rsBatchReceipts.EOF
                                    If flxStatement2.TextMatrix(ifullAllocated, 11) <> "" Then
                                        flxStatement2.TextMatrix(ifullAllocated, 11) = flxStatement2.TextMatrix(ifullAllocated, 11) + "," + rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    Else
                                        flxStatement2.TextMatrix(ifullAllocated, 11) = rsBatchReceipts!FlagTran 'rsBatchReceipts!PF & IIf(IsNull(rsBatchReceipts!C_SL), "", rsBatchReceipts!C_SL)
                                    End If
                                    rsBatchReceipts.MoveNext
                              Wend
                            End If
                      
                        'Putting default Fund
                         flxStatement2.TextMatrix(ifullAllocated, 8) = cmbFund.text
                         flxStatement2.TextMatrix(ifullAllocated, 9) = cmbRptAmtType.text
                         flxStatement2.AddItem ""
                         ifullAllocated = ifullAllocated + 1
               End If

                    
               flxStatement2.row = 0
               rsBatchReceipts.Close
NextReceipt:

        End If
         lCtr = lCtr + 1
    Loop
    oFSTR.Close

   
    rsMasterBatchReceipt.filter = ""
    'MsgBox rsBatchReceipts.RecordCount
    
    Dim iRow, iCol As Integer
    Dim bSwitch As Boolean, iSupp As Integer
    
    'flxSPayment.Clear
    'flxSPayment.Cols = rsBatchReceipts.Fields.count
    flxSPayment.Rows = rsMasterBatchReceipt.RecordCount + 1
'    flxSPayment.FixedCols = 1
    
    'For i = 0 To rsBatchReceipts.Fields.count - 1
    '    flxSPayment.TextMatrix(0, i) = rsBatchReceipts.Fields(i).Name
    'Next i
    
    
    'Dim iRow, iCol As Integer
    'Dim bSwitch As Boolean, iSupp As Integer
    
    iRow = 1
    
    'szHeader$ = "|<ID|<Tenant|<Type|<Unit ID|<Due Date" & _
    '         "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
    '         "|>Receipt £|<DemandID|<Reference|<Posting Date"
    
       While Not rsMasterBatchReceipt.EOF
    '      flxSPayment.TextMatrix(iRow, 1) = rsBatchReceipts!TransactionID
    ''     Check the transaction type and changing the text color of the credit notes
    '      If rsBatchReceipts.Fields.Item("Type").Value = 2 Or rsBatchReceipts.Fields.Item("Type").Value = 4 Then
    '         flxSPayment.row = iRow
    '         For iCol = 2 To 11
    '            flxSPayment.col = iCol
    '            flxSPayment.CellForeColor = vbRed
    '         Next iCol
    '      End If
    
          flxSPayment.TextMatrix(iRow, 1) = rsMasterBatchReceipt!PF & IIf(IsNull(rsMasterBatchReceipt!C_SL), "", rsMasterBatchReceipt!C_SL)
          flxSPayment.TextMatrix(iRow, 2) = rsMasterBatchReceipt!SageAccountNumber
    '      If InStr(rsBatchReceipts!description, "Invoice") > 0 Then
    '         flxSPayment.TextMatrix(iRow, 3) = IIf(rsBatchReceipts!AdjTag = "Y", "ADJI", rsBatchReceipts!description)
    '      Else
    '         flxSPayment.TextMatrix(iRow, 3) = rsBatchReceipts!description
    '      End If
            
          flxSPayment.TextMatrix(iRow, 3) = rsMasterBatchReceipt!unitid
          flxSPayment.TextMatrix(iRow, 4) = IIf(Not IsNull(rsMasterBatchReceipt!dDate), Format(rsMasterBatchReceipt!dDate, "dd/mm/yyyy"), "")
    '      flxSPayment.TextMatrix(iRow, 5) = IIf(IsNull(rsBatchReceipts!Ref), "", rsBatchReceipts!Ref)
          flxSPayment.TextMatrix(iRow, 5) = IIf(IsNull(rsMasterBatchReceipt!Details), "", rsMasterBatchReceipt!Details)
          flxSPayment.TextMatrix(iRow, 6) = Format(rsMasterBatchReceipt!amount, "0.00")
          flxSPayment.TextMatrix(iRow, 7) = Format(rsMasterBatchReceipt!OSAmount, "0.00")
          flxSPayment.TextMatrix(iRow, 8) = Format(rsMasterBatchReceipt!receipt, "0.00") '"0.00"
    '      flxSPayment.TextMatrix(iRow, 9) = IIf(IsNull(rsBatchReceipts!DemandRef), "", rsBatchReceipts!DemandRef)
          flxSPayment.TextMatrix(iRow, 9) = rsMasterBatchReceipt!TranDate
          flxSPayment.TextMatrix(iRow, 10) = Mid$(rsMasterBatchReceipt!Reference, 1, 18)
          flxSPayment.TextMatrix(iRow, 11) = rsMasterBatchReceipt!match
          'flxSPayment.TextMatrix(iRow, 12) = IIf(IsNull(rsMasterBatchReceipt!ref), "", rsMasterBatchReceipt!ref)
          flxSPayment.TextMatrix(iRow, 12) = IIf(IsNull(rsMasterBatchReceipt!Reference), "", rsMasterBatchReceipt!Reference)
          
          rsMasterBatchReceipt.MoveNext
          iRow = iRow + 1
          
    '      If rsBatchReceipts!SageAccountNumber <> frmBatchRpt.flxSPayment.TextMatrix(iRow - 1, 21) Then
    '          bSwitch = Not bSwitch
    '          iSupp = iSupp + 1
    '      End If
       Wend
    
    ResizeGrid
    'flxSPayment.RowSel = 1
    FillBatchReceipts = False
    
    
'    rsStatementRecords.Close
    rsMasterBatchReceipt.Close
    If rsBatchReceipts.State = 1 Then
        rsBatchReceipts.Close
    End If
'    con.Close
'    MsgBox ifullAllocated
    Exit Function
    
ErrorHandler:
    
       If ERR.Number = 3265 Then
            MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
       ElseIf ERR.Number = -2147467259 Then
            MsgBox ERR.description & ":  :" & ERR.Number, , "N"
       Else
            MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
            flxBankTransactions.Clear
            txtGrandTotal.text = "0.00"
       End If
       
'       If rsStatementRecords.State = 1 Then
'          rsStatementRecords.Close
'          Set rsStatementRecords = Nothing
'       End If
       
       If rsMasterBatchReceipt.State = 1 Then
          rsMasterBatchReceipt.Close
          Set rsMasterBatchReceipt = Nothing
       End If
       
'       If con.State = 1 Then
'          con.Close
'       End If
End Function

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
End Sub
Public Sub DeleteArrayItem(ItemArray As Variant, ByVal ItemElement As Long)
    Dim i  As Long
    
    If Not IsArray(ItemArray) Then
        ERR.Raise 13, , "Type Mismatch"
        Exit Sub
    End If
 
    If ItemElement < LBound(ItemArray) Or ItemElement > UBound(ItemArray) Then
        ERR.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If
 
    For i = ItemElement To UBound(ItemArray) - 1
        ItemArray(i) = ItemArray(i + 1)
    Next
    On Error GoTo ErrorHandler:
 
    ReDim Preserve ItemArray(LBound(ItemArray) To UBound(ItemArray) - 1)
 
    Exit Sub
ErrorHandler:
  '~~> An error will occur if array is fixed
    ERR.Raise ERR.Number, , _
    "Array not resizable."
End Sub
'Private Sub LoadMappinginCollections()
''The program does not require mapping file any more. Because I am certain that
'What the header should be for individual bank. I am checking that header definition
'Whileon button apply corresponding  with the bank name


''Dim cc As New Collection
''
''cc.Add Item:="Barclays", key:="1"
''cc.Add Item:="HSBC", key:="2"
''cc.Add Item:="Lloyds", key:="3"
''
''MsgBox cc("2")
'
'On Error GoTo ErrorHandler
'
'Dim adoconn As New ADODB.Connection
'Dim rsCSV As New ADODB.Recordset
'
'
'adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'          "Data Source=" & "" & App.Path & "\BACS;" & _
'            "Extended Properties=""Text;HDR=Yes;"""
'
'rsCSV.Open "Select * from [BatchReceiptFileMapping.csv]", adoconn, adOpenStatic, adLockReadOnly
'Dim record As String
'Dim count As Integer
'count = 1
'
'If Not ccBanks.count > 0 Then
'    Do While Not rsCSV.EOF()
'        ccBanks.Add Item:=rsCSV.Fields(0).Value, key:=CStr(count)
'        ccDateColumns.Add Item:=rsCSV.Fields("Date").Value, key:=rsCSV.Fields(0).Value
''        ccExt1.Add Item:=rsCSV.Fields("Ext1").Value, key:=rsCSV.Fields(0).Value
''        ccExt2.Add Item:=rsCSV.Fields("Ext2").Value, key:=rsCSV.Fields(0).Value
'        ccReferenceColumns.Add Item:=rsCSV.Fields("Reference").Value, key:=rsCSV.Fields(0).Value
'        ccAmountColumns.Add Item:=rsCSV.Fields("Amount").Value, key:=rsCSV.Fields(0).Value
'
'        rsCSV.MoveNext
'        count = count + 1
'    Loop
'End If
'
'rsCSV.Close
'adoconn.Close
'Exit Sub
'
'ErrorHandler:
'
'    MsgBox "The Batch Receipt Mapping File appears to be missing or not in valid format. " & vbNewLine & ERR.description & vbNewLine & "" & _
'    "Please make sure the mapping file in valid format is copied in the BACS folder in the program directory"
'    If rsCSV.State = 1 Then
'        rsCSV.Close
'    End If
'    Set adoconn = Nothing
'End Sub

Private Sub ConfigureFlxSPayment()
    Dim szHeader As String

    flxSPayment.Clear
   
    flxSPayment.Cols = 14
'    szHeader$ = "|<No.|<Tenant|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
'             "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
'             "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo" & _
'             "|<RptDate|PropID|<Reference|<Posting Date"
    
    szHeader$ = "|<ID|<Tenant|<Unit ID|<Due Date" & _
             "|<Details|>Amount £|>O/S Amt. £" & _
             "|>Receipt £|<Posting Date|<Statement Ref|<Match|<Saved Ref|<Unallocated"
   
   flxSPayment.FormatString = szHeader$
   iFlxSPayCol = 12
   flxSPayment.ColWidth(13) = 0
End Sub

Private Function LoadDataDropDowns() As Boolean
   
   On Error GoTo ErrorHandler

   'Bank
   Dim BankData() As String
   ReDim BankData(2, 5) As String
   
   Dim i As Integer
'   For i = 1 To ccBanks.count
'        BankData(0, i - 1) = CStr(i)
'        BankData(1, i - 1) = CStr(ccBanks(i))
'   Next i
    BankData(0, 0) = "1"
   BankData(1, 0) = "Barclays"
   
   BankData(0, 1) = "2"
   BankData(1, 1) = "Lloyds"
   
   BankData(0, 2) = "3"
   BankData(1, 2) = "Santander"
    
   BankData(0, 3) = "4"
   BankData(1, 3) = "BIB"
   
   BankData(0, 4) = "5"
   BankData(1, 4) = "TSB"
   
   BankData(0, 5) = "6"
   BankData(1, 5) = "HSBC"
   
'        cmbBank.Clear
'    cmbBank.AddItem "Barclays"
'    cmbBank.AddItem "Lloyds"
'    cmbBank.AddItem "Santander"
'    cmbBank.AddItem "BIB"
'    cmbBank.AddItem "TSB"
   cmbBank.Column() = BankData()
   cmbBank.ListIndex = 0
   
   ' Account Matching
   Dim AccountMatching() As String

   ReDim AccountMatching(4, 4) As String
   
   AccountMatching(0, 0) = "1"
   AccountMatching(1, 0) = "Lessee A/C"
   
   AccountMatching(0, 1) = "2"
   AccountMatching(1, 1) = "Unit ID"
   
   AccountMatching(0, 2) = "3"
   AccountMatching(1, 2) = "Reference"
    
   AccountMatching(0, 3) = "4"
   AccountMatching(1, 3) = "ALL"
   
   cmbMatchBy.Column() = AccountMatching()
   cmbMatchBy.ListIndex = 0
   '
   
   ' Account Matching
   Dim ReceiptAllocation() As String

   ReDim ReceiptAllocation(3, 3) As String
   
   ReceiptAllocation(0, 0) = "1"
   ReceiptAllocation(1, 0) = "Oldest Invoice First"
   
   ReceiptAllocation(0, 1) = "2"
   ReceiptAllocation(1, 1) = "Recent Invoice First"
   
   ReceiptAllocation(0, 2) = "3"
   ReceiptAllocation(1, 2) = "Settle All"
   
   cmbReceiptAllocation.Column() = ReceiptAllocation()
   cmbReceiptAllocation.ListIndex = 0
   '

   LoadDataDropDowns = True
   Exit Function


ErrorHandler:
   'ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
   MsgBox "The Batch Receipt Mapping File appears to be not in valid format. " & vbNewLine & vbNewLine & "" & _
    "Please make sure the mapping file is in valid format"
   
End Function

Private Sub Check1_Click()
    
End Sub

Private Sub chkMatchedRec_Click()
   Dim iRow As Integer
   If chkMatchedRec.Value = 0 Then
        For iRow = 1 To flxSPayment.Rows - 1
              flxSPayment.RowHeight(iRow) = 240
        Next iRow
   End If
    
  If chkMatchedRec.Value = 1 Then
        flxSPayment.row = 0
        chkUnmachInvoice.Value = 0
        For iRow = 1 To flxSPayment.Rows - 1
            If flxSPayment.TextMatrix(iRow, 11) = "Yes" Then
                flxSPayment.RowHeight(iRow) = 240
            Else
                flxSPayment.RowHeight(iRow) = 0
            End If
        Next iRow
   End If
End Sub

Private Sub chkShowAllCol_Click()
If chkShowAllCol.Value = 1 Then
   flxSPayment.ColWidth(1) = 550
   flxSPayment.ColWidth(5) = 1500
Else
   flxSPayment.ColWidth(1) = 0
   flxSPayment.ColWidth(5) = 0
End If
End Sub



Private Sub chkStatement_Click()
   Dim iRow As Integer
   cmbTenant.Visible = False
    cmbFundGrid.Visible = False
    cmbAmountTypeGrid.Visible = False
   If chkStatement.Value = 0 Then
        For iRow = 1 To flxBankTransactions.Rows - 1
              flxBankTransactions.RowHeight(iRow) = 240
        Next iRow
   End If
    
  If chkStatement.Value = 1 Then
        flxBankTransactions.row = 0
        chkUnassignedtransactions.Value = 0
        For iRow = 1 To flxBankTransactions.Rows - 1
            If flxBankTransactions.TextMatrix(iRow, 10) = "Yes" Then
                flxBankTransactions.RowHeight(iRow) = 240
            Else
                flxBankTransactions.RowHeight(iRow) = 0
            End If
        Next iRow
   End If
End Sub

Private Sub chkUnassignedtransactions_Click()
    Dim iRow As Integer
    cmbTenant.Visible = False
    cmbFundGrid.Visible = False
    cmbAmountTypeGrid.Visible = False
    If chkUnassignedtransactions.Value = 0 Then
         For iRow = 1 To flxBankTransactions.Rows - 1
               flxBankTransactions.RowHeight(iRow) = 240
         Next iRow
    End If
    
    If chkUnassignedtransactions.Value = 1 Then
          chkStatement.Value = 0
          flxBankTransactions.row = 0
          For iRow = 1 To flxBankTransactions.Rows - 1
              If flxBankTransactions.TextMatrix(iRow, 10) = "No" Then
                  flxBankTransactions.RowHeight(iRow) = 240
              Else
                  flxBankTransactions.RowHeight(iRow) = 0
              End If
          Next iRow
     End If
End Sub

Private Sub chkUnmachInvoice_Click()
   Dim iRow As Integer
   If chkUnmachInvoice.Value = 0 Then
        For iRow = 1 To flxSPayment.Rows - 1
              flxSPayment.RowHeight(iRow) = 240
        Next iRow
   End If
    
  If chkUnmachInvoice.Value = 1 Then
        flxSPayment.row = 0
        chkMatchedRec.Value = 0
        For iRow = 1 To flxSPayment.Rows - 1
            If flxSPayment.TextMatrix(iRow, 11) = "No" Then
                flxSPayment.RowHeight(iRow) = 240
            Else
                flxSPayment.RowHeight(iRow) = 0
            End If
        Next iRow
   End If
End Sub

Private Sub cmbAmountTypeGrid_GotFocus()
    flxBankTransactions.ScrollBars = flexScrollBarNone
End Sub

Private Sub cmbAmountTypeGrid_KeyPress(KeyAscii As Integer)
    Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    Dim i As Integer
   If KeyAscii = 27 Then
        flxBankTransactions.SetFocus
        cmbAmountTypeGrid.text = ""
        cmbAmountTypeGrid.Visible = False
   End If
    If KeyAscii = 13 And flxBankTransactions.row < flxBankTransactions.Rows - 1 Then
        
        'If flxBankTransactions.col = 9 Then
             flxBankTransactions.col = 7
            If cmbAmountTypeGrid.text <> "" And RetAmountTypeID(cmbAmountTypeGrid.text) = "" Then
                    MsgBox "Please select valid amount type to proceed", vbInformation, "Sorry"
                    'flxBankTransactions.CellForeColor = vbRed
                    cmbAmountTypeGrid.Visible = True
                    cmbAmountTypeGrid.SetFocus
                    flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                    
                dblAmountReceiptTotal = 0
                
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                     'flxBankTransactions.SetFocus
                     Exit Sub
             Else
                flxBankTransactions.CellForeColor = vbBlack
             End If
            flxBankTransactions.TextMatrix(flxBankTransactions.row, 9) = cmbAmountTypeGrid.text
            If Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 7)) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 9)) = "" Then
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
            Else
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "Yes"
                
                dblAmountReceiptTotal = 0
                
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
            End If
            'flxBankTransactions_DblClick
            cmbAmountTypeGrid.Visible = False
            'Fixed by anol 30 Jun 2015
            Do While flxBankTransactions.row < flxBankTransactions.Rows - 1
                flxBankTransactions.row = flxBankTransactions.row + 1
                If flxBankTransactions.RowHeight(flxBankTransactions.row) > 0 Then
                    Exit Do
                End If
            Loop
            'End of modification
           
             If flxBankTransactions.row >= flxBankTransactions.Rows - 1 Then
                'flxBankTransactions.TextMatrix(flxBankTransactions.row, 9) = cmbAmountTypeGrid.text
                cmbAmountTypeGrid.Visible = False
                cmdOK.SetFocus
             Else
                flxBankTransactions.SetFocus
             End If
        'End If
     ElseIf KeyAscii = 13 And flxBankTransactions.row >= flxBankTransactions.Rows - 2 Then
       ' flxBankTransactions.TextMatrix(flxBankTransactions.row, 9) = cmbAmountTypeGrid.text
        cmbAmountTypeGrid.Visible = False
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmbAmountTypeGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    Call FindComboString(cmbAmountTypeGrid, KeyCode)
End Sub

Private Sub cmbAmountTypeGrid_LostFocus()
'#
    Dim bolvalid As Boolean
    flxBankTransactions.TextMatrix(iCurRow, 9) = cmbAmountTypeGrid.text
    flxBankTransactions.col = 9
    
    If cmbAmountTypeGrid.text <> "" And validAmountTypeID(cmbAmountTypeGrid.text) = False Then
                
                MsgBox "Please select valid amount type to proceed", vbInformation, "Sorry"
                flxBankTransactions.CellForeColor = vbRed
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                
    Else
            flxBankTransactions.CellForeColor = vbBlack
             bolvalid = True
    End If
    '#
        
        
        
    
     If flxBankTransactions.TextMatrix(iCurRow, 7) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 9)) = "" Or bolvalid = False Then
            flxBankTransactions.TextMatrix(iCurRow, 10) = "No"
        Else
             flxBankTransactions.TextMatrix(iCurRow, 10) = "Yes"
        End If
    flxBankTransactions.ScrollBars = flexScrollBarBoth
        Dim dblAmountReceiptTotal As Double
        Dim dblAmountReceipt As Double
        Dim i As Integer
        dblAmountReceiptTotal = 0
            
           For i = 1 To flxBankTransactions.Rows - 1
                If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                   dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                   If dblAmountReceipt > 0 Then
                       dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                   End If
                 End If
           Next i
            txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
End Sub

Private Sub cmbBank_LostFocus()
    If cmbBank.text <> "" And Not IsNull(cmbBank.Value) Then
        SaveSetting "PropertyManagement", "ChoosedOption", "ULR", cmbBank.Value
     Else
        SaveSetting "PropertyManagement", "ChoosedOption", "ULR", ""
    End If
End Sub

Private Sub cmbFund_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBrowseFile.SetFocus
    End If
End Sub

Private Sub cmbFund_KeyUp(KeyCode As Integer, Shift As Integer)
     Call FindComboString(cmbFund, KeyCode)
End Sub

Private Sub cmbFund_LostFocus()
    If Trim(cmbFund.text) <> "" Then
        If RetFundID(Trim(cmbFund.text)) = "" Then
            MsgBox "Please select valid fund code to proceed", vbInformation, "Sorry"
            cmbFund.SetFocus
        End If
    End If
End Sub

Private Function RetAmountTypeID(Code As String) As String
    Dim adoconn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode AND SecondaryCode.Value='" & Code & "' " & _
             "ORDER BY SecondaryCode.Value ;", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        RetAmountTypeID = rsAmountType.Fields("V").Value
       
   End If
   rsAmountType.Close
   adoconn.Close
    
End Function
Private Function validAmountTypeID(Code As String) As Boolean
    Dim adoconn As New ADODB.Connection
    Dim rsAmountType As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsAmountType.Open "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode AND SecondaryCode.Value='" & Code & "' " & _
             "ORDER BY SecondaryCode.Value ;", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsAmountType.EOF Then
        validAmountTypeID = True
   End If
   rsAmountType.Close
   adoconn.Close
    
End Function
Public Function ValidtenantID(ID As String) As Boolean
    Dim adoconn As New ADODB.Connection
    Dim rsTenantID As New ADODB.Recordset
    adoconn.Open getConnectionString
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
    rsTenantID.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsTenantID.EOF Then
       ValidtenantID = True
       
    End If
    rsTenantID.Close
    adoconn.Close
End Function
Private Function RetFundID(Code As String) As String
    Dim adoconn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND where FundCode='" & Code & "';", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsFund.EOF Then
        RetFundID = rsFund.Fields("FundID").Value
        Exit Function
   End If
   rsFund.Close
   adoconn.Close
    
End Function

Private Function validFundCode(Code As String) As Boolean
    Dim adoconn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND where FundCode='" & Code & "';", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsFund.EOF Then
        validFundCode = True
   Else
        validFundCode = False
   End If
   rsFund.Close
   adoconn.Close
    
End Function
Private Sub cmbFundGrid_GotFocus()
    flxBankTransactions.ScrollBars = flexScrollBarNone
End Sub

Private Sub cmbFundGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      flxBankTransactions.SetFocus
      cmbFundGrid.text = ""
      cmbFundGrid.Visible = False
   End If
    If KeyAscii = 13 And flxBankTransactions.row < flxBankTransactions.Rows - 1 Then
        
'        If flxBankTransactions.col = 8 Then
            flxBankTransactions.TextMatrix(flxBankTransactions.row, 8) = cmbFundGrid.text
            'flxBankTransactions_DblClick
            If Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 7)) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 9)) = "" Then
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
            Else
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "Yes"
                Dim dblAmountReceipt, dblAmountReceiptTotal As Double
                dblAmountReceiptTotal = 0
                Dim i As Integer
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
            End If
            
            cmbFundGrid.Visible = False
            If cmbFundGrid.text <> "" And RetFundID(cmbFundGrid.text) = "" Then
                    MsgBox "Please select valid fund code to proceed", vbInformation, "Sorry"
                    'flxBankTransactions.CellForeColor = vbRed
                    cmbFundGrid.Visible = True
                    cmbFundGrid.SetFocus
                    flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                   
                    dblAmountReceiptTotal = 0
                    
                    For i = 1 To flxBankTransactions.Rows - 1
                         If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                            dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                            If dblAmountReceipt > 0 Then
                                dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                            End If
                          End If
                    Next i
                    txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                     'flxBankTransactions.SetFocus
'                     Exit Sub
             Else
                flxBankTransactions.CellForeColor = vbBlack
             End If
            'flxBankTransactions.row = flxBankTransactions.row + 1
            flxBankTransactions.col = 9
            'flxBankTransactions.SetFocus
            flxBankTransactions_DblClick
'        End If
'     ElseIf KeyAscii = 13 And flxSPayment.row >= flxSPayment.Rows - 2 Then
'        flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = cmbFundGrid.text
'        cmbFundGrid.Visible = False
'        cmdOK.SetFocus
    End If
End Sub

Private Sub cmbFundGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    Call FindComboString(cmbFundGrid, KeyCode)
End Sub

Private Sub cmbFundGrid_LostFocus()
'#
    Dim bolvalid As Boolean
    flxBankTransactions.TextMatrix(iCurRow, 8) = cmbFundGrid.text
        flxBankTransactions.col = 8
        
        If cmbFundGrid.text <> "" And validFundCode(cmbFundGrid.text) = False Then
                    
                    MsgBox "Please select valid fund Code to proceed", vbInformation, "Sorry"
                    flxBankTransactions.CellForeColor = vbRed
                    flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                    
        Else
                flxBankTransactions.CellForeColor = vbBlack
                bolvalid = True
        End If
        '#
        
        
        
    If flxBankTransactions.TextMatrix(iCurRow, 7) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 9)) = "" Or bolvalid = False Then
            flxBankTransactions.TextMatrix(iCurRow, 10) = "No"
        Else
             flxBankTransactions.TextMatrix(iCurRow, 10) = "Yes"
        End If
    flxBankTransactions.ScrollBars = flexScrollBarBoth
     Dim dblAmountReceiptTotal As Double
    Dim dblAmountReceipt As Double
    Dim i As Integer
 dblAmountReceiptTotal = 0
           
           For i = 1 To flxBankTransactions.Rows - 1
                If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                   dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                   If dblAmountReceipt > 0 Then
                       dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                   End If
                 End If
           Next i
            txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
End Sub

Private Sub cmbRptAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmbFund.SetFocus
    End If
End Sub

Private Sub cmbRptAmtType_LostFocus()
    If Trim(cmbRptAmtType.text) <> "" Then
        If RetAmountTypeID(cmbRptAmtType.text) = "" Then
            MsgBox "Please select valid payment type to proceed", vbInformation, "Sorry"
            cmbRptAmtType.SetFocus
        End If
    End If
End Sub

Private Sub cmbTenant_GotFocus()
    flxBankTransactions.ScrollBars = flexScrollBarNone
End Sub

Private Sub cmbTenant_KeyPress(KeyAscii As Integer)
    Dim i As Integer
     Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    If KeyAscii = 27 Then
      flxBankTransactions.SetFocus
      cmbTenant.text = ""
      cmbTenant.Visible = False
   End If
    If KeyAscii = 13 And flxBankTransactions.row < flxBankTransactions.Rows - 1 Then
            
        If flxBankTransactions.col = 7 Then
            flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = cmbTenant.text
            If flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(flxBankTransactions.row, 9)) = "" Then
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                dblAmountReceiptTotal = 0
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                
            Else
                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "Yes"
                dblAmountReceiptTotal = 0
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
            End If
            'flxBankTransactions_DblClick
            cmbTenant.Visible = False
             If cmbTenant.text <> "" And ValidtenantID(cmbTenant.text) = False Then
                    ShowMsgInTaskBar "Please select valid tenant ID to proceed", vbInformation, "Sorry"
                    flxBankTransactions.CellForeColor = vbRed
                    flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                    
                dblAmountReceiptTotal = 0
                
                For i = 1 To flxBankTransactions.Rows - 1
                     If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                        dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                        If dblAmountReceipt > 0 Then
                            dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                        End If
                      End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                     'flxBankTransactions.SetFocus
                     cmbTenant.Visible = True
                     cmbTenant.SetFocus
                     Exit Sub
               Else
                flxBankTransactions.CellForeColor = vbBlack
             End If
            flxBankTransactions.col = 8
            'flxBankTransactions.SetFocus
            flxBankTransactions_DblClick
        End If
'     ElseIf KeyAscii = 13 And flxSPayment.row >= flxSPayment.Rows - 2 Then
'        flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = cmbTenant.text
'        cmbTenant.Visible = False
'        cmdOK.SetFocus
    End If
End Sub

Private Sub cmbTenant_KeyUp(KeyCode As Integer, Shift As Integer)
    Call FindComboString(cmbTenant, KeyCode)
    
End Sub

Private Sub cmbTenant_LostFocus()
    Dim dblAmountReceiptTotal As Double
    Dim dblAmountReceipt As Double
    Dim i As Integer
     Dim bolvalid As Boolean
    flxBankTransactions.TextMatrix(iCurRow, 7) = cmbTenant.text
        flxBankTransactions.col = 7
        
        If cmbTenant.text <> "" And ValidtenantID(cmbTenant.text) = False Then
                    
                    MsgBox "Please select valid tenant ID to proceed", vbInformation, "Sorry"
                    flxBankTransactions.CellForeColor = vbRed
                    flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                    
        Else
                flxBankTransactions.CellForeColor = vbBlack
                 bolvalid = True
        End If
        If flxBankTransactions.TextMatrix(iCurRow, 7) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 8)) = "" Or Trim(flxBankTransactions.TextMatrix(iCurRow, 9)) = "" Or bolvalid = False Then
            flxBankTransactions.TextMatrix(iCurRow, 10) = "No"
        Else
             flxBankTransactions.TextMatrix(iCurRow, 10) = "Yes"
        End If
        dblAmountReceiptTotal = 0
           
           For i = 1 To flxBankTransactions.Rows - 1
                If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                   dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                   If dblAmountReceipt > 0 Then
                       dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                   End If
                 End If
           Next i
           txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
    cmbTenant.Visible = False
    flxBankTransactions.ScrollBars = flexScrollBarBoth
'
End Sub

Private Sub cmdAll_Click()
    Dim iRow As Integer
    cmdAll.Caption = "^"
    txtTenantName.text = ""
    chkMatchedRec.Value = 0
    For iRow = 1 To flxSPayment.Rows - 1
        flxSPayment.RowHeight(iRow) = 240
    Next iRow
    cmdAll.Caption = "All"
    picDmdLeaseList.Visible = False
End Sub
Private Sub loadsettingReg()
    On Error GoTo ERR
    Dim iFound As Integer
    Dim szChoice As String
    Dim szaChoice() As String
    Dim i As Integer
    Dim STRs As String
    szChoice = GetSetting("PropertyManagement", "ChoosedOption", "UR-c")
    szaChoice = Split(szChoice, "#")
    
    If UBound(szaChoice) > 0 Then
       STRs = szaChoice(0)
       For i = 0 To cmbRptAmtType.ListCount - 1
            If cmbRptAmtType.List(i) = STRs Then
               cmbRptAmtType.ListIndex = i
               Exit For
            End If
       Next i
       STRs = szaChoice(1)
       For i = 0 To cmbFund.ListCount - 1
            If cmbFund.List(i) = STRs Then
               cmbFund.ListIndex = i
               Exit For
            End If
       Next i
       txtInputFile.text = szaChoice(2)
    End If
    Exit Sub
ERR:
End Sub
Public Function DropDownListPoint(conCombo As Control, szValue As String) As Integer
   Dim i As Integer
   DropDownListPoint = -1
   For i = 0 To conCombo.ListCount - 1
      If conCombo.Column(0, i) = szValue Then
         conCombo.ListIndex = i
         DropDownListPoint = i
         Exit For
      End If
   Next i
End Function

Private Sub savesettinginReg()
   Dim szChoice As String
   Dim szaChoice(3) As String
   szaChoice(0) = cmbRptAmtType.text
   szaChoice(1) = cmbFund.text
   szaChoice(2) = txtInputFile.text
   szChoice = Join(szaChoice, "#")
   SaveSetting "PropertyManagement", "ChoosedOption", "UR-c", szChoice
End Sub

Private Sub CreateTableTlbBatchReceiptUploadFile(adoconn As ADODB.Connection, adoRST As ADODB.Recordset)
    On Error GoTo CreateBatchReceiptUploadFile
           adoRST.Open "SELECT * FROM tlbBatchReceiptUploadFile;", adoconn, adOpenStatic, adLockReadOnly
           adoRST.Close
           GoTo LoadData
CreateBatchReceiptUploadFile:
           adoconn.Execute _
              "CREATE TABLE tlbBatchReceiptUploadFile " & _
                 "(" & _
                    "FileName      TEXT(200) NOT NULL, " & _
                    "UploadDate    DateTime  NOT NULL, " & _
                    "UploadedBy    TEXT(100), " & _
                    "UploadedFrom  TEXT(100), " & _
                    "PRIMARY KEY (FileName)" & _
                 ");"

LoadData:
End Sub
Private Sub cmdApply_Click()
    Dim adoconn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    Dim k As Integer
    Dim rsTenant  As New ADODB.Recordset
    Dim i As Integer
    Dim iCol As Integer
    Dim tempstr As String
    
    If cmbRptAmtType.text = "" Then
        MsgBox "Please select a valid amount type to proceed", vbInformation, "Select a amount type"
        Exit Sub
    End If
    If cmbFund.text = "" Then
        MsgBox "Please select a valid fund to proceed", vbInformation, "Select a fund"
        Exit Sub
    End If
    If txtInputFile.text <> "" Then
    
    If MsgBox("This will assign the receipt amount, transaction date and the reference from the bank statement that you have selected based on your matching criteria." & vbNewLine & vbNewLine & "WARNING: Please make sure bank statement file is closed before you proceed. Do you really want to continue", vbYesNo, "Batch Receipt from Bank Statement") = vbYes Then
        Call savesettinginReg
        txtSearchMemo.text = ""
       
        adoconn.Open getConnectionString
        Call CreateTableTlbBatchReceiptUploadFile(adoconn, adoRST) 'sub routine seperated on 20161125 by anol

        ConfigureFlxSPayment
        chkMatchedRec.Value = 0
        chkStatement.Value = 0
        
        FillBatchReceipts adoconn, txtInputFile.text 'Main procedure to run
        flxSPayment.ColWidth(2) = 1600
        flxSPayment.ColWidth(3) = 1800
        'For allocation  'Written by anol
        For i = 1 To flxSPayment.Rows - 1
                If flxSPayment.TextMatrix(i, 11) = "Yes" Then
                    dblAmountReceiptTotal = dblAmountReceiptTotal + Val(flxSPayment.TextMatrix(i, 8))
                End If
        Next i
        txtGrossTotal.text = Format(dblAmountReceiptTotal, "0.00")
        flxBankTransactions.SetFocus
        
        rsTenant.Open "Select TenantID,Ref from Tenants order by TenantID", adoconn, adOpenDynamic, adLockReadOnly
        For k = 1 To flxBankTransactions.Rows - 1
            If Trim(flxBankTransactions.TextMatrix(k, 6)) <> "" Then
                tempstr = Replace(flxBankTransactions.TextMatrix(k, 6), "'", "''")
                rsTenant.filter = " Ref = '" & tempstr & "'"
                    If Not rsTenant.EOF Then
                         flxBankTransactions.TextMatrix(k, 7) = rsTenant.Fields("TenantID").Value
                         flxBankTransactions.TextMatrix(k, 10) = "Yes"
                    End If
                End If
        Next k
                '**********************
                'Total Receipts on  Accounts
                dblAmountReceiptTotal = 0
                For i = 1 To flxBankTransactions.Rows - 1
                If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                    dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                    If dblAmountReceipt > 0 Then
                        dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                    End If
                End If
                Next i
                txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                '****************************
        rsTenant.Close
         adoconn.Close
        'End of addition by anol
        flxBankTransactions.ScrollBars = flexScrollBarBoth
    End If
     flxSPayment.ColWidth(13) = 0
     
       flxSPayment.row = 0
       flxBankTransactions.row = 0

    cmdUnmatchAll.Enabled = True
    cmdUnmatch.Enabled = True
Else
    MsgBox "Please select a bank statement file", vbInformation, "No Statement File Selected"
    cmdBrowseFile.SetFocus
End If
'Unload Me
End Sub
Private Sub LoadRptAmtType(szValue As String, adoconn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open SQLStr1, adoconn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   cmbRptAmtType.Clear
   cmbAmountTypeGrid.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      cmbAmountTypeGrid.AddItem adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   cmbRptAmtType.Column() = szaData()
  
End Sub
Private Sub cmdBrowseFile_Click()

Dim file As String
file = GetFileName
txtInputFile.text = file

End Sub

Private Function GetFileName() As String
   Dim ofn                    As OPENFILENAME
   Dim lHwnd                  As Long
   Const HKEY_LOCAL_MACHINE   As Long = &H80000002
   Dim szOldFile_PathName     As String
   Dim szNewFile_Path         As String
   Dim szNewFile_Name         As String
   Dim szNewFile_PathName     As String
   Dim fso                    As Object
   Dim szImportFile           As String
   Dim szNC                   As String
   Dim szFund                 As String

On Error GoTo FileError
   
   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
'   ofn.lpstrFilter = "MS Office Excel Workbook 2007-2010 (*.xlsx)" + Chr$(0) + "*.xlsx" + Chr$(0) + _
'                     "MS Office Excel Workbook 97-2003 (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + _
'                     "CSV Files (*.csv)" + Chr$(0) + "*.csv" + Chr$(0)

   ofn.lpstrFilter = "CSV Files (*.csv)" + Chr$(0) + "*.csv" + Chr$(0)
  
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   If txtInputFile = "" Then
    ofn.lpstrInitialDir = CurDir
   Else
    ofn.lpstrInitialDir = txtInputFile.text
   End If
   ofn.lpstrTitle = "Select File to Open"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Function

   szImportFile = ofn.lpstrFile

   If szImportFile = "" Then
      MsgBox "Please select an input file to import", vbInformation, "Input File"
      cmdBrowseFile.SetFocus
      Exit Function
   End If
   
   GetFileName = szImportFile
   Exit Function

FileError:
   GetFileName = ""
   Exit Function
End Function

Private Function CheckIfUploaded(adoconn As ADODB.Connection) As Boolean
On Error GoTo ERR
   Dim rsFile As New ADODB.Recordset
   Dim isFound As Boolean
   isFound = False
   
   rsFile.Open "SELECT * FROM tlbBatchReceiptUploadFile WHERE FileName = '" & Filename & "'", adoconn, adOpenStatic, adLockBatchOptimistic
   
   While Not rsFile.EOF
      isFound = True
      rsFile.MoveNext
   Wend
   
   rsFile.Close
   CheckIfUploaded = isFound
ERR:
End Function
Public Function DeleteLine(fName As String, LineNumber As Long) _
     As Boolean
'Purpose: Deletes a Line from a text file by anol

'Parameters: fName = FullPath to File
'            LineNumber = LineToDelete

'Returns:    True if Successful, false otherwise

'Requires:   Reference to Microsoft Scripting Runtime

'Example: DeleteLine("C:\Myfile.txt", 3)
'           Deletes third line of Myfile.txt

'______Introduced by anol 28 Jun 2015 Reason: The cSv file had extra first line and wthat was causing prblem when we were trying to read it for recrdset________________________________________________________
                       

  Dim oFSO As New FileSystemObject
  Dim oFSTR As Scripting.TextStream
  'Dim ret As Long
  Dim lCtr As Long
  Dim sTemp As String, sLine As String
  Dim bLineFound As Boolean
  
  On Error GoTo ErrorHandler
  If oFSO.FileExists(fName) Then
     Set oFSTR = oFSO.OpenTextFile(fName)
    lCtr = 1
'    If LineNumber = 1 Then
'        sLine = oFSTR.ReadLine
'        If InStr(1, sLine, "BIB") = 0 Then GoTo ErrorHandler
'    End If
    
     Do While Not oFSTR.AtEndOfStream
        sLine = oFSTR.ReadLine
        'If lCtr <> LineNumber Then
        If lCtr <> LineNumber Then
            sTemp = sTemp & sLine & vbCrLf
            
        Else
            If InStr(1, sLine, "BIB") = 0 Then 'if there is no BIB in the line 1 then dont erase it
                sTemp = sTemp & sLine & vbCrLf
            End If
            bLineFound = True
            
        End If
        lCtr = lCtr + 1
    Loop
   
     oFSTR.Close
     Set oFSTR = oFSO.CreateTextFile(fName, True) '2nd parameter is overwrite true
     oFSTR.Write sTemp
  
    DeleteLine = bLineFound
   End If
   
 
ErrorHandler:
On Error Resume Next
oFSTR.Close
Set oFSTR = Nothing
Set oFSO = Nothing

End Function

Public Function DeleteLineSantandar(fName As String) _
     As Boolean
'Purpose: Deletes a Line from a text file by anol

'Parameters: fName = FullPath to File
'            LineNumber = LineToDelete

'Returns:    True if Successful, false otherwise

'Requires:   Reference to Microsoft Scripting Runtime

'Example: DeleteLine("C:\Myfile.txt", 3)
'           Deletes third line of Myfile.txt

'______Introduced by anol 28 Jun 2015 Reason: The cSv file had extra first line and wthat was causing prblem when we were trying to read it for recrdset________________________________________________________
                       

  Dim oFSO As New FileSystemObject
  Dim oFSTR As Scripting.TextStream
  'Dim ret As Long
  Dim lCtr As Long
  Dim sTemp As String, sLine As String
  Dim bLineFound As Boolean
  
  On Error GoTo ErrorHandler
  If oFSO.FileExists(fName) Then
     Set oFSTR = oFSO.OpenTextFile(fName)
    lCtr = 1
    
    
     Do While Not oFSTR.AtEndOfStream
        sLine = oFSTR.ReadLine
        sLine = Replace(Replace(sLine, """", ""), "GBP", "")
        If Left(sLine, 14) = "Date,Narrative" Then
              bLineFound = True
        End If
        If bLineFound Then
            sTemp = sTemp & sLine & vbCrLf
        End If
        lCtr = lCtr + 1
    Loop
   
     oFSTR.Close
     Set oFSTR = oFSO.CreateTextFile(fName, True) '2nd parameter is overwrite true
     oFSTR.Write sTemp
  
    
   End If
   
 
ErrorHandler:
On Error Resume Next
oFSTR.Close
Set oFSTR = Nothing
Set oFSO = Nothing

End Function
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClearReference_Click()
If MsgBox("WARNING:" & vbNewLine & vbNewLine & "This will clear all the defaut reference saved for the tenants from the database. Do you really want to continue?", vbYesNo, "Clearing default transaction reference") = vbYes Then
   
   Dim i As Integer
   For i = 1 To flxSPayment.Rows - 1
      flxSPayment.TextMatrix(i, iFlxSPayCol) = ""
   Next i
   
   ClearReference
End If
End Sub

Private Sub cmdDmdGridUnitLookup_Click()
   
   picDmdLeaseList.Visible = False

End Sub

Private Sub cmdDmdTenantLookup_Click()
 On Error GoTo ERR
   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchName.text = ""
   txtDmdTenantSearchUnitName.text = ""
   txtTenantName.text = ""
   picDmdLeaseList.Top = txtTenantName.Top + txtTenantName.Height + 5
   picDmdLeaseList.Left = txtTenantName.Left + 5
   picDmdLeaseList.Visible = True
   picDmdLeaseList.ZOrder 0
   txtDmdTenantSearchID.SetFocus
   Exit Sub
ERR:
    ShowMsgInTaskBar ERR.description, "Y", "P"
End Sub
Private Sub RefreshGridSupp()
   Dim iRow As Integer

   For iRow = 1 To flxSPayment.Rows - 1
      flxSPayment.RowHeight(iRow) = 240
   Next iRow
    'flxSPayment.Refresh
  If flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) = "ALL" Then Exit Sub

   For iRow = 1 To flxSPayment.Rows - 1
   'Below line has been modified by anol 10 Apr 2015
   'Grid was not showing the correct value after search
   'issue 445
      If flxSPayment.TextMatrix(iRow, 2) <> flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) Then
         flxSPayment.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub
Public Sub HighLightRowsFlxGrid1(conFlxGrid As Control, iSelRow As Integer)
   Dim iRow As Integer, iCol As Integer

   conFlxGrid.row = iSelRow
   For iCol = 1 To conFlxGrid.Cols - 1
      conFlxGrid.col = iCol
      conFlxGrid.CellBackColor = &H8000000D
   Next iCol
End Sub
Private Sub cmdMatch_Click()
    Dim iCol As Integer
    Dim amount, balance, ostAmount As Double
    If flxSPayment.row = 0 Then
'        MsgBox "Please select a sales Receipt to match."
         MsgBox " Please select a sales invoice to match.", vbInformation
       
        Exit Sub
    End If
    If flxBankTransactions.row = 0 Then
        MsgBox "Please select a bank statement to match.", vbInformation
        Exit Sub
    End If
    If flxSPayment.TextMatrix(flxSPayment.row, 11) = "Yes" Then
        
        flxSPayment.TopRow = flxSPayment.row
        HighLightRowsFlxGrid1 flxSPayment, flxSPayment.row
        MsgBox "This sales receipt transaction is already matched.", vbInformation
        Exit Sub
    End If
    'if marking of transaction has some  value and tenant ID has some value then you cannot allocation, give warning
    If InStr(1, flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), "S") = 0 And Len(flxBankTransactions.TextMatrix(flxBankTransactions.row, 7)) > 0 Then
        MsgBox "This statement entry cannot be matched as it has been assigned to a Tenant or leaseholder.", vbInformation
        Exit Sub
    End If
    amount = flxBankTransactions.TextMatrix(flxBankTransactions.row, bankAmountColIndex)
    balance = Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, 0))
    ostAmount = flxSPayment.TextMatrix(flxSPayment.row, 7)
    
    If balance > 0 Then
       If ostAmount < balance Then
          flxSPayment.TextMatrix(flxSPayment.row, 8) = ostAmount
       Else
          flxSPayment.TextMatrix(flxSPayment.row, 8) = balance
       End If
    
       balance = balance - ostAmount
       
       If balance < 0 Then
           flxBankTransactions.TextMatrix(flxBankTransactions.row, 0) = "0.00"
       Else
           flxBankTransactions.TextMatrix(flxBankTransactions.row, 0) = Format(Round(balance, 2), "0.00")
           'Go for allocation putting Tenant ID while clicking match button
           
       End If
       flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = flxSPayment.TextMatrix(flxSPayment.row, 2) 'Tenant ID
       flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "Yes"
       'Puttng transaction Id on 2nd grid at last row
       If flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = "" Then
            flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = flxSPayment.TextMatrix(flxSPayment.row, 1)
       Else
            flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) + "," + flxSPayment.TextMatrix(flxSPayment.row, 1)
       End If
       flxSPayment.TextMatrix(flxSPayment.row, 9) = flxBankTransactions.TextMatrix(flxBankTransactions.row, postingDateColIndex)
       
       If flxSPayment.TextMatrix(flxSPayment.row, 10) = "" Then
           flxSPayment.TextMatrix(flxSPayment.row, 10) = Mid$(flxBankTransactions.TextMatrix(flxBankTransactions.row, bankReferenceColIndex + bankReferenceColIndexDev), 1, 18) 'Deviation is there because I have deleted one column in the array
       End If
       
       flxSPayment.TextMatrix(flxSPayment.row, 11) = "Yes"
       flxSPayment.TextMatrix(flxSPayment.row, 12) = Mid$(flxBankTransactions.TextMatrix(flxBankTransactions.row, bankReferenceColIndex + bankReferenceColIndexDev), 1, 200)
       ResizeRow flxSPayment.row
       'Written by anol 20161023
       'If  outstanding amount IS ZERO you need to move ROW from grid 2 to 3
        If Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, 0)) = 0 Then
           'copy the cell from current grid to the last row of grid 2
           For iCol = 0 To flxBankTransactions.Cols - 1
                 flxStatement2.TextMatrix(flxStatement2.Rows - 1, iCol) = flxBankTransactions.TextMatrix(flxBankTransactions.row, iCol)
                 flxBankTransactions.TextMatrix(flxBankTransactions.row, iCol) = ""
            Next iCol
            flxStatement2.TopRow = flxStatement2.Rows - 1
            flxBankTransactions.RowHeight(flxBankTransactions.row) = 0
            flxBankTransactions.row = flxBankTransactions.row + 1
            flxStatement2.AddItem ""
        End If
        
    Else
        MsgBox "The selected bank statement transaction is already allocated to a receipt transaction."
    End If
    Dim i As Integer
    Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    For i = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(i, 11) = "Yes" Then
            dblAmountReceipt = Val(flxSPayment.TextMatrix(i, 8))
            If dblAmountReceipt > 0 Then
                dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
            End If
        End If
    Next i
    txtGrossTotal.text = Format(dblAmountReceiptTotal, "0.00")
    
    dblAmountReceiptTotal = 0
    For i = 1 To flxBankTransactions.Rows - 1
        If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
            dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
            If dblAmountReceipt > 0 Then
                dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
            End If
        End If
    Next i
    txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
    'flxSPayment.TextMatrix(flxSPayment.row, 8) = flxSPayment.TextMatrix(flxSPayment.row, 7)

End Sub

Private Sub cmdOK_Click()
    If frmBatchRpt.bSavedPayment Then
        If MsgBox("Processing a new upload file will clear your current batch receipt selections. Are you sure you wish to do this?", vbYesNo, "Clear batch receipt?") = vbYes Then
            Call frmBatchRpt.cmdPaymentDiscard_Click
        End If
    End If
    cmdUpdateReference_Click
    If cmbFund.text = "" Then
        MsgBox "Please select a valid fund to proceed.", vbInformation, "Select a fund"
        cmbFund.SetFocus
        Exit Sub
    End If
        Dim i, j, k, currow As Integer
        Dim dblAmount As Double
        Dim tranID As String
        'ConfigureFlxallocation
        For i = 0 To frmBatchRpt.flxSPayment.Rows - 1
            For j = 1 To flxSPayment.Rows - 1
                If frmBatchRpt.flxSPayment.TextMatrix(i, 1) = flxSPayment.TextMatrix(j, 1) And flxSPayment.TextMatrix(j, 11) = "Yes" Then
                    frmBatchRpt.flxSPayment.TextMatrix(i, 11) = flxSPayment.TextMatrix(j, 8) 'ReceiptAmount
                    frmBatchRpt.flxSPayment.TextMatrix(i, 24) = flxSPayment.TextMatrix(j, 10) 'Reference
                    frmBatchRpt.flxSPayment.TextMatrix(i, 25) = flxSPayment.TextMatrix(j, 9) 'Posting Date
                    frmBatchRpt.flxSPayment.TextMatrix(i, 22) = flxSPayment.TextMatrix(j, 9) 'Reporting Date
                    HasRecforBatchReceipt = True 'this flag shall tell something has been posted to the batch receipt form and hence create the bank receipt file
                    GoTo NextTran:
                End If
            Next j
            
NextTran:
        'added by anol
        
        Next i
    frmBatchRpt.ConfigureFlxallocation
    frmBatchRpt.cmbTenantID.Visible = False
    For k = 1 To flxBankTransactions.Rows - 1
        If Trim(flxBankTransactions.TextMatrix(k, 7)) <> "" And Val(flxBankTransactions.TextMatrix(k, 0)) > 0 Then
        '  szHeader$ = "|<No|<LesseID|<FundCode|<Amount|<Payment Date|>Reference|>Posting date"
             
                currow = currow + 1
                frmBatchRpt.flxAllocation.AddItem ""
                frmBatchRpt.flxAllocation.TextMatrix(currow, 0) = ""
                frmBatchRpt.flxAllocation.TextMatrix(currow, 1) = currow '1.No
                frmBatchRpt.flxAllocation.TextMatrix(currow, 2) = flxBankTransactions.TextMatrix(k, 7) '2.LesseID--combo
                frmBatchRpt.flxAllocation.TextMatrix(currow, 3) = flxBankTransactions.TextMatrix(k, 8) '3.FundCode--combo
                dblAmount = dblAmount + Val(flxBankTransactions.TextMatrix(k, 0))
                frmBatchRpt.flxAllocation.TextMatrix(currow, 5) = Format(Val(flxBankTransactions.TextMatrix(k, 0)), "0.00") '4.Amount
                frmBatchRpt.flxAllocation.TextMatrix(currow, 6) = flxBankTransactions.TextMatrix(k, postingDateColIndex)  '5.Payment Date
                frmBatchRpt.flxAllocation.TextMatrix(currow, 7) = Left(flxBankTransactions.TextMatrix(k, bankReferenceColIndex), 20) '6.Reference 'left added on 20161129 by anol
                frmBatchRpt.flxAllocation.TextMatrix(currow, 8) = flxBankTransactions.TextMatrix(k, postingDateColIndex)  '7.Posting date
                frmBatchRpt.flxAllocation.TextMatrix(currow, 4) = flxBankTransactions.TextMatrix(k, 9) '4.Amount type
                frmBatchRpt.flxAllocation.TextMatrix(currow, 9) = "B" '0.B for automatic bank transaction means this line is from this this form. M shall be used on manual from batch receipt procedure
                HasRecforBatchReceipt = True 'this flag shall tell something has been posted to the batch receipt form and hence create the bank receipt file
                
        End If
    Next k
    

   frmBatchRpt.txtAlloctotal.text = Format(dblAmount, "0.00")
   
    ResizeallocaGrid
    'End of addition

SumUpTotal

frmBatchRpt.BankStatementFile = Filename


    Me.Hide
    frmBatchRpt.flxAllocation.SetFocus
End Sub
Private Function ResizeallocaGrid()
    Dim intRow As Integer
    Dim intCol As Integer
   
    With frmBatchRpt.flxAllocation
        For intCol = 0 To .Cols - 1
           '.ColWidth(intCol) = Me.TextWidth(.TextMatrix(intRow, intCol)) + 100
           If intCol = 0 Then
               .ColWidth(intCol) = 350
           Else
              .ColWidth(intCol) = frmBatchRpt.lblAllocation(intCol).Left - frmBatchRpt.lblAllocation(intCol - 1).Left
           End If
        Next
    End With
End Function


'Private Function ReturnReceiptAmount(strReference As String, strDate As String, strTenantID As String) As Double
'    Dim i As Integer
'            flxBankTransactions.ColWidth(flxBankTransactions.Cols - 2) = 1500
'            flxBankTransactions.ColWidth(flxBankTransactions.Cols - 1) = 1500
'            flxBankTransactions.ColWidth(flxBankTransactions.Cols - 3) = 1800
'            flxBankTransactions.ColWidth(flxBankTransactions.Cols - 4) = 1700
'    For i = 1 To flxBankTransactions.Rows - 1
'
'        If flxBankTransactions.TextMatrix(i, bankReferenceColIndex) = strReference And flxBankTransactions.TextMatrix(i, postingDateColIndex) = strDate Then
'            ReturnReceiptAmount = Val(flxBankTransactions.TextMatrix(i, 0))
'
'            flxBankTransactions.TextMatrix(i, flxBankTransactions.Cols - 4) = strTenantID
'            flxBankTransactions.TextMatrix(i, flxBankTransactions.Cols - 1) = "Yes"
'            Exit For
'        End If
'    Next i
'End Function
Private Function FindRowNumber(InvNo As String) As Integer
    Dim i As Integer
    For i = 1 To flxSPayment.Rows - 1
           If flxSPayment.TextMatrix(i, 1) = InvNo Then
                FindRowNumber = i
                Exit Function
           End If
    Next i

End Function
Private Sub cmdUnmatch_Click()
    Dim temp
    Dim i, k As Integer
    Dim tempstr As String
    Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    Dim iCol As Integer
    If LastSelGridnumber = 1 Then
        MsgBox "Please select a matched Bank Statement line to unmatch a transaction", vbInformation
        Exit Sub
    End If
        'sometimes there is some extra comma in the marking on the receipt Iam going to remove that
    If InStr(1, flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), "S") = 0 And LastSelGridnumber = 2 Then
        flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = Replace(flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), ",", "")
    End If
    If InStr(1, flxStatement2.TextMatrix(flxStatement2.row, 11), "S") = 0 And LastSelGridnumber = 3 Then
        flxStatement2.TextMatrix(flxStatement2.row, 11) = Replace(flxStatement2.TextMatrix(flxStatement2.row, 11), ",", "")
    End If
    
    If LastSelGridnumber = 2 Then
                            If flxBankTransactions.row <= 0 Then  'Or Trim(flxSPayment.TextMatrix(flxSPayment.row, 11)) = ""
                                MsgBox "Please select a matched receipt transaction.", vbInformation
                            Else
                                If flxBankTransactions.row > 0 Then
                                'flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = Replace(flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), ",,", ",")
                                'If flxSPayment.row <= 1 Then 'If no row selected in upper grid
                                'As we had written trnsaction ID's  in the flxBankTransactions column 11 from which receipts are made.
                                'we need to make a loop with this one row multuple value data
                                            temp = Split(flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), ",")
                                            For i = 0 To UBound(temp)
                                                If Len(temp(i)) > 0 Then
                                                     tempstr = temp(i)
                                                     k = FindRowNumber(tempstr) 'Finding that transaction ID in the first grid and empty the second grid return row number
                                                     If k > 0 Then
                                                         flxSPayment.TextMatrix(k, 8) = "0.00"
                                                         flxSPayment.TextMatrix(k, 9) = ""
                                                         flxSPayment.TextMatrix(k, 10) = ""
                                                         flxSPayment.TextMatrix(k, 11) = "No"
                                                         flxSPayment.TextMatrix(k, 12) = ""
                                                         'Making receipt amount is equal to outstanding amount
                                                         flxBankTransactions.TextMatrix(flxBankTransactions.row, 0) = Format(Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, bankAmountColIndex)), "0.00")
                                                         'making the mark of transaction number blank
                                                         flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = ""
                                                    End If
                                                End If
                                            Next i
                                            'Removing assignment if no trans ID assocaiated
                                            If flxBankTransactions.TextMatrix(flxBankTransactions.row, 0) = Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, bankAmountColIndex)) Then
                                                flxBankTransactions.TextMatrix(flxBankTransactions.row, 11) = ""
                                                flxBankTransactions.TextMatrix(flxBankTransactions.row, 10) = "No"
                                                flxBankTransactions.TextMatrix(flxBankTransactions.row, 7) = ""
                                            End If
                                              
                                        ' Finding the total receipt amount 1st grid and total allocated amount in 2nd grid
                                       
                                        For i = 1 To flxSPayment.Rows - 1
                                            If flxSPayment.TextMatrix(i, 11) = "Yes" Then
                                                dblAmountReceipt = Val(flxSPayment.TextMatrix(i, 8))
                                                If dblAmountReceipt > 0 Then
                                                    dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                                                End If
                                            End If
                                        Next i
                                        txtGrossTotal.text = Format(dblAmountReceiptTotal, "0.00")
                                        dblAmountReceiptTotal = 0
                                        For i = 1 To flxBankTransactions.Rows - 1
                                            If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
                                                dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
                                                If dblAmountReceipt > 0 Then
                                                    dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                                                End If
                                            End If
                                        Next i
                                        txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                                      '*********************
                                  Else
                                     MsgBox "Please select a matched receipt transaction!"
                                  End If
                              End If
                              
                              'Now you need to move some row in between the grid 2 and 3
    ElseIf LastSelGridnumber = 3 Then
                        If flxStatement2.row <= 0 Then  'Or Trim(flxSPayment.TextMatrix(flxSPayment.row, 11)) = ""
                                MsgBox "Please select a matched receipt transaction."
                            Else
                                If flxStatement2.row > 0 Then
                                'flxStatement2.TextMatrix(flxStatement2.row, 11) = Replace(flxStatement2.TextMatrix(flxStatement2.row, 11), ",,", ",")
                                'If flxSPayment.row <= 1 Then 'If no row selected in upper grid
                                'As we had written trnsaction ID's  in the flxStatement2 column 11 from which receipts are made.
                                'we need to make a loop with this one row multuple value data
                                            temp = Split(flxStatement2.TextMatrix(flxStatement2.row, 11), ",")
                                            For i = 0 To UBound(temp)
                                                If Len(temp(i)) > 0 Then
                                                     tempstr = temp(i)
                                                     k = FindRowNumber(tempstr) 'Finding that transaction ID in the first grid and empty the second grid return row number
                                                     If k > 0 Then
                                                         flxSPayment.TextMatrix(k, 8) = "0.00"
                                                         flxSPayment.TextMatrix(k, 9) = ""
                                                         flxSPayment.TextMatrix(k, 10) = ""
                                                         flxSPayment.TextMatrix(k, 11) = "No"
                                                         flxSPayment.TextMatrix(k, 12) = ""
                                                         'Making receipt amount is equal to outstanding amount
                                                         flxStatement2.TextMatrix(flxStatement2.row, 0) = Format(Val(flxStatement2.TextMatrix(flxStatement2.row, bankAmountColIndex)), "0.00")
                                                         'making the mark of transaction number blank
                                                         flxStatement2.TextMatrix(flxStatement2.row, 11) = ""
                                                    End If
                                                End If
                                            Next i
                                            'Removing assignment if no trans ID assocaiated
                                            If flxStatement2.TextMatrix(flxStatement2.row, 0) = Val(flxStatement2.TextMatrix(flxStatement2.row, bankAmountColIndex)) Then
                                                flxStatement2.TextMatrix(flxStatement2.row, 11) = ""
                                                flxStatement2.TextMatrix(flxStatement2.row, 10) = "No"
                                                flxStatement2.TextMatrix(flxStatement2.row, 7) = ""
                                            End If
                                              
                                        ' Finding the total receipt amount 1st grid and total allocated amount in 2nd grid
                                        
                                        For i = 1 To flxSPayment.Rows - 1
                                            If flxSPayment.TextMatrix(i, 11) = "Yes" Then
                                                dblAmountReceipt = Val(flxSPayment.TextMatrix(i, 8))
                                                If dblAmountReceipt > 0 Then
                                                    dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                                                End If
                                            End If
                                        Next i
                                        txtGrossTotal.text = Format(dblAmountReceiptTotal, "0.00")
                                        dblAmountReceiptTotal = 0
                                        For i = 1 To flxStatement2.Rows - 1
                                            If flxStatement2.TextMatrix(i, 10) = "Yes" Then
                                                dblAmountReceipt = Val(flxStatement2.TextMatrix(i, 0))
                                                If dblAmountReceipt > 0 Then
                                                    dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
                                                End If
                                            End If
                                        Next i
                                        txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
                                      '*********************
                                        'If receipt amount is equal to outstanding amount you need to move some row in grid 2 from 3
                                      If flxStatement2.TextMatrix(flxStatement2.row, 0) = Format(Val(flxStatement2.TextMatrix(flxStatement2.row, bankAmountColIndex)), "0.00") Then
                                         'copy the cell from current grid to the last row of grid 2
                                         For iCol = 0 To flxStatement2.Cols - 1
                                               flxBankTransactions.TextMatrix(flxBankTransactions.Rows - 1, iCol) = flxStatement2.TextMatrix(flxStatement2.row, iCol)
                                               flxStatement2.TextMatrix(flxStatement2.row, iCol) = ""
                                          Next iCol
                                          flxBankTransactions.TextMatrix(flxBankTransactions.Rows - 1, 0) = flxBankTransactions.TextMatrix(flxBankTransactions.Rows - 1, bankAmountColIndex)
                                          flxBankTransactions.TopRow = flxBankTransactions.Rows - 1
                                          flxStatement2.RowHeight(flxStatement2.row) = 0
                                          flxStatement2.row = flxStatement2.row + 1
                                          flxBankTransactions.AddItem ""
                                      End If
                                  Else
                                     MsgBox "Please select a matched receipt transaction!"
                                  End If
                              End If
                 
     End If


End Sub

Private Sub cmdUnmatchAll_Click()
    Dim con As New ADODB.Connection
    Dim i As Integer
    Dim directory As String
    Dim rsStatementRecords As New ADODB.Recordset
    Dim szHeader As String
    Dim lCtr As Integer
    Dim sLine As String
    Dim sColValue() As String
    If MsgBox("This will reset all the batch receipt and bank statement transactions. Do you really want to continue", vbYesNo, "Unmatching all transactions") = vbYes Then
       For i = 1 To flxSPayment.Rows - 1
          UnMatchReceiptTransactions i
       Next i
    
'       For i = 0 To flxBankTransactions.Rows - 1
'          flxBankTransactions.TextMatrix(i, 0) = flxBankTransactions.TextMatrix(i, bankAmountColIndex)
'          flxBankTransactions.TextMatrix(i, 7) = "" 'tenant iD
'          flxBankTransactions.TextMatrix(i, 10) = "No"
'          flxBankTransactions.TextMatrix(i, 11) = "" 'tranID
'       Next i
    End If
    flxBankTransactions.Clear
    flxBankTransactions.Rows = 2
    flxStatement2.Clear
    flxStatement2.Rows = 2
   
   
    If Dir(txtInputFile.text, vbDirectory) = vbNullString Then
        MsgBox "The file does not exist. Please select a valid file.", vbInformation, "Invalid Directory"
        Exit Sub
    End If
    Dim oFSTR As Scripting.TextStream
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set oFSTR = fso.OpenTextFile(txtInputFile.text)  'Reading the CSV file into a text stream.
    
    If cmbBank.text = "TSB" Then
            flxBankTransactions.Cols = 5 + 4
            postingDateColIndex = 1 'location in the grid is 1 but the staement file that is 0
            bankReferenceColIndex = 4
            bankAmountColIndex = 5 ' THIS 5 IS THE LOCATION OF GRID AMOUNT, IN THE FILE IT WILL BE 4
            szHeader$ = "O/S Amt £|<Transaction Date|< Transaction Type|< Account Number|<  Transaction Description|< Credit Amount|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "Barclays" Then
           'Date,Memo,Amount
           'Number  Date    Account Amount  Subcategory Memo
           'Date    Account Amount  Subcategory Memo
           flxBankTransactions.Cols = 5 + 4
           postingDateColIndex = 1 ''location in the grid is 1 but the statement file that is 0
           bankReferenceColIndex = 5
           bankAmountColIndex = 3
           szHeader$ = "O/S Amt £|<Date|<Account|<Amount|<Subcategory|<Memo|<StatementRowNo|<Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "Lloyds" Then
            'In File Header "Transaction Date    Transaction Type    Sort Code   Account Number  Transaction Description Debit Amount    Credit Amount   Balance"
             szHeader$ = "O/S Amt £|<Transaction Date|<Transaction Type|<Account Number|<Transaction Description|<Credit Amount|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 1 'location in the grid is 1 but the statement file that is 0
             bankReferenceColIndex = 4 '356
             bankAmountColIndex = 5 ' THIS 5 IS THE LOCATION OF GRID AMOUNT, IN THE FILE IT WILL BE 4
    ElseIf cmbBank.text = "BIB" Then
            postingDateColIndex = 2
            bankReferenceColIndex = 5
            bankAmountColIndex = 3
'            Exit Function
            'GROUP   ACC ID  ACCOUNT NO  TYPE    BANK CODE   CURR    ENTRY DATE  AS AT   AMOUNT  TLA CODE    CHEQUE NO   STATUS  DESCRIPTION
            szHeader$ = "O/S Amt £|<ACCOUNT NO|< ENTRY DATE|< AMOUNT|< TLA CODE|< DESCRIPTION|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
    ElseIf cmbBank.text = "Santander" Then
             'In File Header "Date    Narrative   Transaction Type    Debit   Credit  Current Balance
             szHeader$ = "O/S Amt £|< Date|<Narrative |<Transaction Type|<Credit Amount|<Current Balance|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 1 'location in the grid is 1 but the statement file that is 0
             bankReferenceColIndex = 2 '356
             bankAmountColIndex = 4 ' THIS 5 IS THE LOCATION OF GR
    ElseIf cmbBank.text = "HSBC" Then
             'In File Header "Date    Type    Description Paid out    Paid in Balance
             szHeader$ = "O/S Amt £|< Date|<Type |<Description|<Paid in|<Balance|< StatementRowNo|< Tenant ID |< Fund Code|< Amount Type|< Assigned|<TransID"
             postingDateColIndex = 0 'I have taken this index from the receipt file
             bankReferenceColIndex = 2 'I have taken this index from the receipt file
             bankAmountColIndex = 4 ' I have taken this index from the receipt file
    End If
    flxBankTransactions.FormatString = szHeader$
'    flxStatement2.FixedCols = 1
    flxStatement2.Cols = flxBankTransactions.Cols
    flxBankTransactions.ColWidth(6) = 0 '7 Account
    flxBankTransactions.ColWidth(7) = 1700 '7 Tenant ID
    flxBankTransactions.ColWidth(8) = 1800 '8 Fund Code
    flxBankTransactions.ColWidth(9) = 1500 '9 Amount Type
    flxBankTransactions.ColWidth(10) = 1500 '10 Assigned
    flxBankTransactions.ColWidth(11) = 0 '11 TransID
    
    'added by anol 20161021
    
    flxStatement2.FormatString = szHeader$
    flxStatement2.ColWidth(6) = 0 '7 Account
    flxStatement2.ColWidth(7) = 1700 '7 Tenant ID
    flxStatement2.ColWidth(8) = 1800 '8 Fund Code
    flxStatement2.ColWidth(9) = 1500 '9 Amount Type
    flxStatement2.ColWidth(10) = 1500 '10 Assigned
    flxStatement2.ColWidth(11) = 0 '11 TransID
    
    
    Dim referenceValue, record, unit, lessee, filter As String
    Dim iBankTranRow As Integer
    iBankTranRow = 1
   
    
    Dim receiptBalance, receiptAmount As Double
    receiptBalance = 0
    receiptAmount = 0
    lCtr = 1
  
    
    'End of addition
'    Dim record, unit, lessee, Filter As String
'    Dim iBankTranRow As Integer
    iBankTranRow = 1
   
   

    Do While Not oFSTR.AtEndOfStream
            sLine = oFSTR.ReadLine
           If lCtr = 1 Then
                ' Excluding the header
                'Compare the header for diffrent banks
                If cmbBank.text = "TSB" Then
                    If sLine = "Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance" Then
                        'so Header format is fine here
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                ElseIf cmbBank.text = "Barclays" Then
                    If sLine = "Number,Date,Account,Amount,Subcategory,Memo" Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                 ElseIf cmbBank.text = "Lloyds" Then
                    If sLine = "Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance" Or _
                    sLine = "Transaction Date,Transaction Type,Sort Code,Account Number,Transaction Description,Debit Amount,Credit Amount,Balance," Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                 ElseIf cmbBank.text = "BIB" Then
                    If sLine = "GROUP,ACC ID,ACCOUNT NO,TYPE,BANK CODE,CURR,ENTRY DATE,AS AT,AMOUNT,TLA CODE,CHEQUE NO,STATUS,DESCRIPTION" Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                 ElseIf cmbBank.text = "Santander" Then
                    If sLine = "Date,Narrative,Transaction Type,Debit,Credit,Current Balance" Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                 ElseIf cmbBank.text = "HSBC" Then
                    If sLine = "Date,Type,Description,Paid out,Paid in,Balance" Then
                    Else
                        oFSTR.Close
                        MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                    End If
                 Else
                       MsgBox "You may have selected the wrong bank statement for the selected bank.", vbExclamation, "Invalid Statement File"
                        Exit Sub
                 End If
           Else
                 sColValue = Split(sLine, ",")
                If cmbBank.text = "TSB" Then
                    'making the array as same format as grid
                    DeleteArrayItem sColValue, 2
                    DeleteArrayItem sColValue, 4
                    DeleteArrayItem sColValue, 5
                ElseIf cmbBank.text = "Barclays" Then
                    DeleteArrayItem sColValue, 0 'suppresing/deleitng null column
                ElseIf cmbBank.text = "Lloyds" Then
                    DeleteArrayItem sColValue, 3
                    DeleteArrayItem sColValue, 5
                ElseIf cmbBank.text = "BIB" Then
                    DeleteArrayItem sColValue, 0
                    DeleteArrayItem sColValue, 0
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 1
                    DeleteArrayItem sColValue, 2
                    DeleteArrayItem sColValue, 4
                    DeleteArrayItem sColValue, 4
                ElseIf cmbBank.text = "Santander" Then
                    DeleteArrayItem sColValue, 3
                End If
                 referenceValue = sColValue(bankReferenceColIndex - 1)
                
                If IsNull(sColValue(bankReferenceColIndex - 1)) Or sColValue(bankReferenceColIndex - 1) = "" Then
                    GoTo NextReceipt
                End If
                If IsNull(sColValue(bankAmountColIndex - 1)) Or sColValue(bankAmountColIndex - 1) = "" Then
                    GoTo NextReceipt
                End If
                receiptAmount = sColValue(bankAmountColIndex - 1)
                
                
                If receiptAmount < 0 Then
                    GoTo NextReceipt
                End If
                
                receiptBalance = receiptAmount
                
                    flxBankTransactions.RowHeight(iBankTranRow) = 280

                    For i = 0 To 4 'used generally -1, now 4+2=5
                       flxBankTransactions.TextMatrix(iBankTranRow, i + 1) = IIf(IsNull(sColValue(i)), "", sColValue(i)) 'rsStatementRecords.Fields(i).Value
                    Next i
                    
                    flxBankTransactions.TextMatrix(iBankTranRow, 0) = Format(sColValue(bankAmountColIndex - 1), "0.00") 'rsStatementRecords.Fields(amountCol).Value 'IsNull ommited
                    flxBankTransactions.TextMatrix(iBankTranRow, 6) = lCtr ' Bank statement Row number
                    flxBankTransactions.TextMatrix(iBankTranRow, 7) = "" 'TenantID
                    flxBankTransactions.TextMatrix(iBankTranRow, 8) = cmbFund.text
                    flxBankTransactions.TextMatrix(iBankTranRow, 9) = cmbRptAmtType.text
                    If flxBankTransactions.TextMatrix(iBankTranRow, 10) = "" Then
                        flxBankTransactions.TextMatrix(iBankTranRow, 10) = "No"
                    End If
                    flxBankTransactions.AddItem ""
                    iBankTranRow = iBankTranRow + 1
NextReceipt:
         End If
         lCtr = lCtr + 1
    Loop
    oFSTR.Close
'    flxStatement2.Clear
    flxStatement2.row = 0
    ResizeGrid
    Dim dblAmountReceipt, dblAmountReceiptTotal As Double
    For i = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(i, 11) = "Yes" Then
            dblAmountReceipt = Val(flxSPayment.TextMatrix(i, 8))
            If dblAmountReceipt > 0 Then
                dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
            End If
        End If
    Next i
    txtGrossTotal.text = Format(dblAmountReceiptTotal, "0.00")
    dblAmountReceiptTotal = 0
    For i = 1 To flxBankTransactions.Rows - 1
         If flxBankTransactions.TextMatrix(i, 10) = "Yes" Then
            dblAmountReceipt = Val(flxBankTransactions.TextMatrix(i, 0))
            If dblAmountReceipt > 0 Then
                dblAmountReceiptTotal = dblAmountReceiptTotal + dblAmountReceipt
            End If
          End If
    Next i
    txtReceiptsOnAccount.text = Format(dblAmountReceiptTotal, "0.00")
End Sub

Private Sub UnMatchReceiptTransactions(rowIndex As Integer)

    flxSPayment.TextMatrix(rowIndex, 8) = "0.00"
    flxSPayment.TextMatrix(rowIndex, 9) = ""
    flxSPayment.TextMatrix(rowIndex, 10) = ""
    flxSPayment.TextMatrix(rowIndex, 11) = "No"
    flxSPayment.TextMatrix(rowIndex, 12) = ""

End Sub

Private Sub cmdUpdateReference_Click()

If MsgBox("Do you wish to save the references of the matched and assigned transactions selected as the default reference for tenant searching ?", vbYesNo, "Saving default transaction reference") = vbYes Then
    
    SaveReference
        
    MsgBox "The default tenant search references selected have been successfully saved.", vbOKOnly, "Transaction Reference Saved"
End If
End Sub

Private Sub SaveReference()

Dim adoconn As New ADODB.Connection
adoconn.Open getConnectionString

Dim i As Integer
Dim ref, TenantID As String

On Error GoTo ERR:

For i = 1 To flxSPayment.Rows - 1
    
    ref = flxSPayment.TextMatrix(i, iFlxSPayCol)
    
    If flxSPayment.TextMatrix(i, 11) = "Yes" And ref <> "" Then
        TenantID = flxSPayment.TextMatrix(i, 2)
        adoconn.Execute "UPDATE Tenants AS N " & _
                      "SET    N.Ref = '" & ref & "' " & _
                      "WHERE  N.TenantID = '" & TenantID & "';"
    End If
Next i
For i = 1 To flxBankTransactions.Rows - 1
    ref = flxBankTransactions.TextMatrix(i, 2)
    If flxBankTransactions.TextMatrix(i, 10) = "Yes" And ref <> "" Then
        TenantID = flxBankTransactions.TextMatrix(i, 7)
        adoconn.Execute "UPDATE Tenants AS N " & _
                      "SET    N.Ref = '" & ref & "' " & _
                      "WHERE  N.TenantID = '" & TenantID & "';"
    End If
Next i

adoconn.Close
Exit Sub

ERR:
    MsgBox ERR.description & ":  :" & ERR.Number, , "N"
    adoconn.Close

End Sub

Private Sub ClearReference()

Dim adoconn As New ADODB.Connection
adoconn.Open getConnectionString

Dim i As Integer
Dim ref, TenantID As String

On Error GoTo ERR:

adoconn.Execute "UPDATE Tenants AS N SET N.Ref = '';"

adoconn.Close
Exit Sub

ERR:
    MsgBox ERR.description & ":  :" & ERR.Number, , "N"
    adoconn.Close

End Sub

Private Sub Command1_Click()
    flxSPayment.row = 1
End Sub





Private Sub CommandButton1_Click()
    txtSearchMemo.text = ""
    chkStatement.Value = 0
End Sub

Private Sub flxBankTransactions_Click()
    cmbTenant.Visible = False
    cmbFundGrid.Visible = False
    cmbAmountTypeGrid.Visible = False
    LastSelGridnumber = 2 'This variable shall hold last selected grid number
    'THIS SHALL TRY TO SELECT THE ROW IN A WHOLE IF IT IS CLICKED LEFT CELLS THAT CELL NUMBER LESS THAN 7
    Dim iSel As Integer
    iSel = flxBankTransactions.col
    If flxBankTransactions.col < 7 Then
        With flxBankTransactions
            .col = 0
            .ColSel = .Cols - 1
        End With
        flxBankTransactions.ScrollBars = flexScrollBarBoth
    Else
        Exit Sub
    End If
    flxBankTransactions.col = iSel
    'written by anol 20161022
    'This function shall highlight related invoices that is associated with receipt
    Dim temp
    Dim tempstr As String
    Dim i, iCol As Integer
    Dim k As Integer
    Dim Flag As Integer
    Dim ilastCol As Integer
    LastSelGridnumber = 2 'This variable shall hold last selected grid number
'    MsgBox flxStatement2.TextMatrix(flxStatement2.row, 11)
    'HighLightRowsFlxGrid
'    flxStatement2.row = 0
    Dim gridRows(25) As Integer
    For i = LBound(gridRows) To UBound(gridRows)
         gridRows(i) = "200"
    Next
    
    k = flxBankTransactions.row
    ilastCol = flxBankTransactions.col
    For i = 1 To flxBankTransactions.Rows - 1
        flxBankTransactions.col = 1
       
        flxBankTransactions.row = i
        If flxBankTransactions.CellBackColor = RGB(233, 232, 155) Or flxBankTransactions.CellBackColor = RGB(233, 232, 228) Or flxBankTransactions.CellBackColor = &H8000000D Then
             For iCol = 1 To flxBankTransactions.Cols - 1
               flxBankTransactions.col = iCol
               flxBankTransactions.CellBackColor = vbWhite
            Next iCol
        End If
    Next
    flxBankTransactions.row = k
    
    
    For iCol = 1 To flxBankTransactions.Cols - 1
        flxBankTransactions.col = iCol
        flxBankTransactions.CellBackColor = &H8000000D
    Next iCol
    flxBankTransactions.col = ilastCol
    '#
    k = flxSPayment.row
   
     For i = 1 To flxSPayment.Rows - 1
        flxSPayment.col = 1
        flxSPayment.row = i
        If flxSPayment.CellBackColor = RGB(233, 232, 155) Or flxSPayment.CellBackColor = RGB(233, 232, 228) Or flxBankTransactions.CellBackColor = &H8000000D Then
             For iCol = 1 To flxSPayment.Cols - 1
               flxSPayment.col = iCol
               flxSPayment.CellBackColor = vbWhite
            Next iCol
        End If
    Next
    
    flxSPayment.row = k
    flxStatement2.row = 0
    If Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, 0)) = 0 Then Exit Sub
    If Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, 0)) = Val(flxBankTransactions.TextMatrix(flxBankTransactions.row, bankAmountColIndex)) Then Exit Sub
   

    temp = Split(flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), ",")
    For i = 0 To UBound(temp)
        If Len(temp(i)) > 0 Then
             tempstr = temp(i)
             k = FindRowNumber(tempstr) 'Finding that transaction ID in the first grid and empty the second grid return row number
             If k > 0 Then
                HighLightRowsFlxGrid flxSPayment, k
                InsertArrayItem gridRows(), 0, k
'                flxSPayment.TopRow = k
'                If Flag > k Then
'                    flxSPayment.TopRow = k
'                    Flag = k
'                End If
             End If
        End If
    Next
    SortIntegerArray gridRows()
    If gridRows(0) = 200 Then Exit Sub
    flxSPayment.TopRow = gridRows(0)
End Sub

Private Sub flxBankTransactions_DblClick()
    Dim i As Integer
    cmbTenant.Visible = False
    cmbFundGrid.Visible = False
    cmbAmountTypeGrid.Visible = False
     If flxBankTransactions.row < flxBankTransactions.Rows - 1 Then
        If flxBankTransactions.col = 7 Then
            If InStr(1, flxBankTransactions.TextMatrix(flxBankTransactions.row, 11), "S") > 0 Then
                MsgBox "This Tenant ID cannot be changed as it has been matched with some transactions." & vbCrLf & " For assign this receipt to a tenant please unmatch it first.", vbInformation
                Exit Sub
            End If
            cmbTenant.Top = flxBankTransactions.CellTop + flxBankTransactions.Top
            cmbTenant.Left = flxBankTransactions.CellLeft + flxBankTransactions.Left
            cmbTenant.Width = flxBankTransactions.ColWidth(flxBankTransactions.col)
            cmbTenant.text = flxBankTransactions.TextMatrix(flxBankTransactions.row, flxBankTransactions.col)
            cmbTenant.Visible = True
            cmbTenant.SetFocus
            SelTxtInCtrl cmbTenant
            'SelTxtInCtrl cmbTenant
            'SendKeys "{vbKeyDown}"
'            SendKeys vbKeyUp
         End If
         If flxBankTransactions.col = 8 Then
            cmbFundGrid.Top = flxBankTransactions.CellTop + flxBankTransactions.Top
            cmbFundGrid.Left = flxBankTransactions.CellLeft + flxBankTransactions.Left
            cmbFundGrid.Width = flxBankTransactions.ColWidth(flxBankTransactions.col)
            cmbFundGrid.text = flxBankTransactions.TextMatrix(flxBankTransactions.row, flxBankTransactions.col)
            cmbFundGrid.Visible = True
            cmbFundGrid.SetFocus
            SelTxtInCtrl cmbFundGrid
         End If
         If flxBankTransactions.col = 9 Then
            cmbAmountTypeGrid.Top = flxBankTransactions.CellTop + flxBankTransactions.Top
            cmbAmountTypeGrid.Left = flxBankTransactions.CellLeft + flxBankTransactions.Left
            cmbAmountTypeGrid.Width = flxBankTransactions.ColWidth(flxBankTransactions.col)
            cmbAmountTypeGrid.text = flxBankTransactions.TextMatrix(flxBankTransactions.row, flxBankTransactions.col)
            cmbAmountTypeGrid.Visible = True
            cmbAmountTypeGrid.SetFocus
            SelTxtInCtrl cmbAmountTypeGrid
         End If
     End If
      iCurRow = flxBankTransactions.row
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
Private Sub LoadcmbTenant()
    cmbTenant.Clear
    Dim adoconn As New ADODB.Connection
    Dim rsTenantID As New ADODB.Recordset
    adoconn.Open getConnectionString
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
    rsTenantID.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    While Not rsTenantID.EOF
        cmbTenant.AddItem rsTenantID("SageAccountNumber").Value
    rsTenantID.MoveNext
    Wend
    adoconn.Close
End Sub

Private Sub flxBankTransactions_KeyPress(KeyAscii As Integer)
       If KeyAscii = 13 Then flxBankTransactions_DblClick
End Sub

Private Sub flxBankTransactions_RowColChange()
    cmbTenant.Visible = False
End Sub

Private Sub flxDmdLeaseList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtDmdTenantSearchID.SetFocus
    End If
End Sub

Private Sub flxDmdLeaseList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDmdLeaseList_Click
    End If
End Sub

Private Sub flxSPayment_Click()
'un select the colored one's and make blue selected one
    Dim i As Integer
    Dim k As Integer
    Dim iCol As Integer
    Dim ilastCol As Integer
    LastSelGridnumber = 1
     k = flxSPayment.row
      ilastCol = flxSPayment.col
    For i = 1 To flxSPayment.Rows - 1
        flxSPayment.col = 1
       
        flxSPayment.row = i
        If flxSPayment.CellBackColor = RGB(233, 232, 155) Or flxSPayment.CellBackColor = RGB(233, 232, 228) Or flxSPayment.CellBackColor = &H8000000D Then
             For iCol = 1 To flxSPayment.Cols - 1
               flxSPayment.col = iCol
               flxSPayment.CellBackColor = vbWhite
            Next iCol
        End If
    Next
    flxSPayment.row = k
      
    For iCol = 1 To flxSPayment.Cols - 1
        flxSPayment.col = iCol
        flxSPayment.CellBackColor = &H8000000D
    Next iCol
    flxSPayment.col = ilastCol
    'we are not doing anything with un assighned one because we need to maek it free for normal operation
    If flxSPayment.TextMatrix(flxSPayment.row, 11) = "Yes" Then
        flxBankTransactions.row = 0
        flxStatement2.row = 0
        For i = 1 To flxBankTransactions.Rows - 1
            If InStr(1, flxBankTransactions.TextMatrix(i, 11), flxSPayment.TextMatrix(flxSPayment.row, 1)) > 0 And Len(flxBankTransactions.TextMatrix(i, 11)) > 0 Then
                flxBankTransactions.row = i
                For iCol = 1 To flxBankTransactions.Cols - 1
                   flxBankTransactions.col = iCol
                   flxBankTransactions.CellBackColor = RGB(233, 232, 155)
                Next iCol
            Else
               flxBankTransactions.row = i
               If flxBankTransactions.CellBackColor = RGB(233, 232, 155) Or flxBankTransactions.CellBackColor = RGB(233, 232, 228) Or flxBankTransactions.CellBackColor = &H8000000D Then
                     For iCol = 1 To flxBankTransactions.Cols - 1
                       flxBankTransactions.col = iCol
                       flxBankTransactions.CellBackColor = vbWhite
                    Next iCol
                End If
            End If
        Next
'        select the assigned one
'       Highlight the third grid based on the selection of first grid
        For i = 1 To flxStatement2.Rows - 1
            If InStr(1, flxStatement2.TextMatrix(i, 11), flxSPayment.TextMatrix(flxSPayment.row, 1)) > 0 And Len(flxStatement2.TextMatrix(i, 11)) > 0 Then
                flxStatement2.row = i
                flxStatement2.TopRow = i
                For iCol = 1 To flxStatement2.Cols - 1
                   flxStatement2.col = iCol
                   flxStatement2.CellBackColor = RGB(233, 232, 155)
                Next iCol
            Else
               flxStatement2.row = i
               If flxStatement2.CellBackColor = RGB(233, 232, 155) Or flxStatement2.CellBackColor = RGB(233, 232, 228) Or flxStatement2.CellBackColor = &H8000000D Then
                     For iCol = 1 To flxStatement2.Cols - 1
                       flxStatement2.col = iCol
                       flxStatement2.CellBackColor = vbWhite
                    Next iCol
                End If
            End If
        Next
        
    Else
        'Do not color anything in third and second
'        flxBankTransactions.row = 0
'        flxStatement2.row = 0
'       if I clear this I wouldnt be able to select for match'ex as requirement
        k = flxBankTransactions.row
        For i = 1 To flxBankTransactions.Rows - 1
            flxBankTransactions.col = 1
            flxBankTransactions.row = i
            If flxBankTransactions.CellBackColor = RGB(233, 232, 155) Or flxBankTransactions.CellBackColor = RGB(233, 232, 228) Then
                 For iCol = 1 To flxBankTransactions.Cols - 1
                   flxBankTransactions.col = iCol
                   flxBankTransactions.CellBackColor = vbWhite
                Next iCol
            End If
        Next
       flxBankTransactions.row = k
       
       k = flxStatement2.row
        For i = 1 To flxStatement2.Rows - 1
            flxStatement2.col = 1
            flxStatement2.row = i
            If flxStatement2.CellBackColor = RGB(233, 232, 155) Or flxStatement2.CellBackColor = RGB(233, 232, 228) Then
                 For iCol = 1 To flxStatement2.Cols - 1
                   flxStatement2.col = iCol
                   flxStatement2.CellBackColor = vbWhite
                Next iCol
            End If
        Next
       flxStatement2.row = k
    End If
End Sub

Private Sub flxSPayment_dblClick()
Dim i As Integer
If flxSPayment.RowHeight(flxSPayment.row) = 0 Then Exit Sub
flxSPayment.col = iFlxSPayCol

txtRefInput.Top = flxSPayment.CellTop + flxSPayment.Top
'iTop = txtRefInput.Top
txtRefInput.Left = flxSPayment.CellLeft + flxSPayment.Left
'iLeft = flxSPayment.CellLeft + flxSPayment.Left

txtRefInput.Width = flxSPayment.ColWidth(iFlxSPayCol)
txtRefInput.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
txtRefInput.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
txtRefInput.Visible = True
'flxSPayment.ScrollBars = flexScrollBarNone
txtRefInput.SetFocus

'bSavedPayment = False
'iCurRow = flxSPayment.row
txtRefInput.Visible = True

SelTxtInCtrl txtRefInput
End Sub

Public Function PopulateDmdTenantLookup(adoconn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Me.MousePointer = vbHourglass
   
   Dim adoRST As New ADODB.Recordset
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
   adoRST.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   flxDmdLeaseList.TextMatrix(1, 1) = "ALL"
   flxDmdLeaseList.TextMatrix(1, 2) = "All Lessees"
   

   iRow = 2

   While Not adoRST.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRST!SageAccountNumber
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRST!Name
      flxDmdLeaseList.TextMatrix(iRow, 3) = adoRST!UnitNumber

      iRow = iRow + 1
      adoRST.MoveNext

      If Not adoRST.EOF Then flxDmdLeaseList.AddItem ""
   Wend
   adoRST.Close
   Set adoRST = Nothing

   txtTenantName.text = "ALL / All Lessees"

   Me.MousePointer = vbArrow
End Function

Private Sub flxStatement2_Click()
    'written by anol 20161022
    'This function shall highlight related invoices that is associated with receipt
    Dim temp
    Dim tempstr As String
    Dim i, iCol As Integer
    Dim k As Integer
    Dim Flag As Integer
    LastSelGridnumber = 3 'This variable shall hold last selected grid number
'    MsgBox flxStatement2.TextMatrix(flxStatement2.row, 11)
    'HighLightRowsFlxGrid
    '&H8000000D&
    For i = 1 To flxSPayment.Rows - 1
        flxSPayment.col = 1
        flxSPayment.row = i
        If flxSPayment.CellBackColor = RGB(233, 232, 155) Or flxSPayment.CellBackColor = RGB(233, 232, 228) Or flxSPayment.CellBackColor = &H8000000D Then
             For iCol = 1 To flxSPayment.Cols - 1
               flxSPayment.col = iCol
               flxSPayment.CellBackColor = vbWhite
            Next iCol
        End If
    Next
    'do not highlight second grid when you select third grid
    For i = 1 To flxBankTransactions.Rows - 1
        flxBankTransactions.col = 1
        flxBankTransactions.row = i
        If flxBankTransactions.CellBackColor = RGB(233, 232, 155) Or flxBankTransactions.CellBackColor = RGB(233, 232, 228) Or flxBankTransactions.CellBackColor = &H8000000D Then
             For iCol = 1 To flxBankTransactions.Cols - 1
               flxBankTransactions.col = iCol
               flxBankTransactions.CellBackColor = vbWhite
            Next iCol
        End If
    Next
    flxBankTransactions.row = 0
    flxSPayment.row = 0
'    K = flxStatement2.row
'    For i = 1 To flxStatement2.Rows - 1
'        flxStatement2.col = 1
'        flxStatement2.row = i
'        If flxStatement2.CellBackColor = RGB(233, 232, 155) Then
'             For iCol = 1 To flxStatement2.Cols - 1
'               flxStatement2.col = iCol
'               flxStatement2.CellBackColor = vbWhite
'            Next iCol
'        End If
'    Next
'    HighLightRowsFlxGrid flxStatement2, K

    Dim gridRows(25) As Integer
    For i = LBound(gridRows) To UBound(gridRows)
         gridRows(i) = "200"
    Next
    temp = Split(flxStatement2.TextMatrix(flxStatement2.row, 11), ",")
    For i = 0 To UBound(temp)
        If Len(temp(i)) > 0 Then
             tempstr = temp(i)
             k = FindRowNumber(tempstr) 'Finding that transaction ID in the first grid and empty the second grid return row number
             If k > 0 Then
                HighLightRowsFlxGrid flxSPayment, k
'                If Flag = False Then
'                    flxSPayment.TopRow = K
'                    Flag = True
'                End If
                InsertArrayItem gridRows(), 0, k
'                flxSPayment.TopRow = k
'                If Flag > k Then
'                    flxSPayment.TopRow = k
'                    Flag = k
'                End If
             End If
        End If
    Next
    SortIntegerArray gridRows()
    If gridRows(0) = 200 Then Exit Sub
     flxSPayment.TopRow = gridRows(0)
End Sub
Private Sub SortIntegerArray(paintArray() As Integer)
    '------------------------------------------------------------------------
     
    ' This sub uses the Bubble Sort algorithm to sort an array of integers.
     
    Dim lngX As Long
    Dim lngY As Long
    Dim intTemp As Integer
    For lngX = LBound(paintArray) To (UBound(paintArray) - 1)
     
        For lngY = LBound(paintArray) To (UBound(paintArray) - 1)
         
        If paintArray(lngY) > paintArray(lngY + 1) Then
        ' exchange the items
        intTemp = paintArray(lngY)
        paintArray(lngY) = paintArray(lngY + 1)
        paintArray(lngY + 1) = intTemp
        End If
         
        Next
     
    Next
     
    
 
End Sub
Sub InsertArrayItem(arr As Variant, Index As Long, newValue As Variant)
    Dim i As Long
    For i = UBound(arr) - 1 To Index Step -1
    arr(i + 1) = arr(i)
    Next
    arr(Index) = newValue
End Sub
Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
    Dim adoconn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    Dim STRs As String
    Dim i As Integer
    
    cmdUnmatchAll.Enabled = False
    cmdUnmatch.Enabled = False
    adoconn.Open getConnectionString
   
    PopulateDmdTenantLookup adoconn, ""
    ' we shall not need this any more rem by anol 20161201
'    LoadMappinginCollections
    
    LoadDataDropDowns
    ConfigureFlxSPayment
    LoadcmbTenant
    LoadFund
    LoadRptAmtType "RECEIPT AMOUNT TYPE", adoconn
    adoconn.Close
    STRs = GetSetting("PropertyManagement", "ChoosedOption", "ULR", cmbBank.text)
    
    For i = 0 To cmbBank.ListCount - 1
        If cmbBank.List(i) = STRs Then
           cmbBank.ListIndex = i
           Exit For
        End If
    Next i
   Call loadsettingReg
   Call WheelHook(Me.hWnd)
End Sub
Private Sub txtDmdTenantSearchID_Change()

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 09 Aug 2014
   FilterTenantsList
   
End Sub
'Resolved by BOSL
'Issue No: 0000445.
'The function generates the expression of matching string pattern by using SQL LIKE operation and
'uses the in-built Filter function of the ADODB recordset to filter the records that match with the
'expression and finally bind the filtered records to the grid.
'Modified By: Asif. 09 Aug 2014
Private Function FilterTenantsList() As String
   
   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString
   
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset
   Dim tempstr As String
   Dim filter As String
   
   If Len(txtDmdTenantSearchID.text) > 0 Then
      txtDmdTenantSearchName.text = ""
      txtDmdTenantSearchUnitName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchID.text), "'", "''")
      filter = " SageAccountNumber LIKE '%" + tempstr + "*'"
      
   End If
  'Issue No: 0000445. note 933 Wild card searching has been implemented by anol 23 Feb 2015
   If Len(txtDmdTenantSearchName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchUnitName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchName.text), "'", "''")
      filter = " Name LIKE '%" + tempstr + "*'"
   End If

   If Len(txtDmdTenantSearchUnitName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchName.text = ""
      tempstr = Replace(UCase(txtDmdTenantSearchUnitName.text), "'", "''")
      filter = " UnitNumber LIKE '%" + tempstr + "*'"
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
   adoRST.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
 
   adoRST.filter = filter
        
   flxDmdLeaseList.Clear
   'Resolved by BOSL
   ' Issue 530/992 On selection of lessees in batch receipts user not able to go back to full list.
   'Modified by anol 23 Mar 2015
   flxDmdLeaseList.Rows = adoRST.RecordCount + 2
   
   flxDmdLeaseList.TextMatrix(1, 1) = "ALL"
   flxDmdLeaseList.TextMatrix(1, 2) = "All Lessees"

   Dim iRow As Integer
   iRow = 2

   While Not adoRST.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRST!SageAccountNumber
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRST!Name
      flxDmdLeaseList.TextMatrix(iRow, 3) = adoRST!UnitNumber

      iRow = iRow + 1
      adoRST.MoveNext

'      If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   Wend
   
   adoRST.Close
   Set adoRST = Nothing

   adoconn.Close
   Set adoconn = Nothing

End Function

Private Sub txtDmdTenantSearchID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        flxDmdLeaseList.SetFocus
    End If
End Sub

Private Sub txtDmdTenantSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDmdTenantSearchName.SetFocus
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
    txtDmdTenantSearchUnitName.SetFocus
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
   Dim adoconn As New ADODB.Connection

   txtTenantName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) & " / " & flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)

   

   RefreshGridSupp

   picDmdLeaseList.Visible = False
End Sub
Private Sub ConfigFlxDmdLeaseList()
   Dim szHeader As String

   flxDmdLeaseList.Clear
   flxDmdLeaseList.Cols = 4
   flxDmdLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Tenant ID|<Tenant Name|<Unit Name"
   flxDmdLeaseList.FormatString = szHeader$
   flxDmdLeaseList.ColWidth(0) = Label20(9).Left - flxDmdLeaseList.Left    '240        Solid column
   flxDmdLeaseList.ColWidth(1) = Label20(8).Left - Label20(9).Left - 20    '1400       'Tenant ID
   flxDmdLeaseList.ColWidth(2) = Label20(7).Left - Label20(8).Left - 20    'Tenant Name
   flxDmdLeaseList.ColWidth(3) = flxDmdLeaseList.Left + flxDmdLeaseList.Width - Label20(7).Left - 300 'Unit Name
   flxDmdLeaseList.Rows = 3
End Sub
Private Sub LoadFund()
    Dim adoconn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND;", adoconn, adOpenStatic, adLockReadOnly
    cmbFund.Clear
    cmbFundGrid.Clear
    While Not rsFund.EOF
        cmbFund.AddItem rsFund.Fields("FundCode").Value
        cmbFundGrid.AddItem rsFund.Fields("FundCode").Value
        rsFund.MoveNext
   Wend
   rsFund.Close
   adoconn.Close
    
End Sub
Private Sub ResizeRow(intRow As Integer)
   Dim intCol As Integer
   With flxSPayment
      For intCol = 0 To .Cols - 1
          If .ColWidth(intCol) < Me.TextWidth(.TextMatrix(intRow, intCol)) + 75 Then
             .ColWidth(intCol) = Me.TextWidth(.TextMatrix(intRow, intCol)) + 75
          End If
      Next
   End With
   
End Sub

Private Sub ResizeGrid()
    Dim intRow As Integer
    Dim intCol As Integer
   
    With flxSPayment
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .ColWidth(intCol) < Me.TextWidth(.TextMatrix(intRow, intCol)) + 75 Then
                   .ColWidth(intCol) = Me.TextWidth(.TextMatrix(intRow, intCol)) + 75
                End If
            Next
        Next
    End With
    
       
    With flxBankTransactions
        For intCol = 0 To .Cols - 7
            For intRow = 0 To .Rows - 1
                'formating os amount as decimal anol 20161029
                If intRow > 0 And intCol = 0 Then
                    .TextMatrix(intRow, intCol) = Format(.TextMatrix(intRow, intCol), "0.00")
                End If
                If .ColWidth(intCol) < Me.TextWidth(.TextMatrix(intRow, intCol)) + 100 Then
                   .ColWidth(intCol) = Me.TextWidth(.TextMatrix(intRow, intCol)) + 100
                End If
            Next
        Next
    End With
End Sub

Private Sub txtDmdTenantSearchUnitName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDmdLeaseList.SetFocus
    End If
End Sub

Private Sub txtGrossTotal_Change()
        txtGrandTotal.text = Format(Val(txtReceiptsOnAccount.text) + Val(txtGrossTotal.text), "0.00")
End Sub

Private Sub txtReceiptsOnAccount_Change()
        txtGrandTotal.text = Format(Val(txtReceiptsOnAccount.text) + Val(txtGrossTotal.text), "0.00")
End Sub

Private Sub txtRefInput_Change()
    On Error GoTo ERR
    If Len(txtRefInput.text) > 50 Then
       MsgBox "The transaction reference must not be more than 50 characters."
       txtRefInput.text = Mid$(txtRefInput.text, 1, 50)
       txtRefInput.SetFocus
    End If
ERR:
End Sub

Private Sub txtRefInput_GotFocus()
'txtRefInput.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
'txtRefInput.text = ""
'txtRefInput.Visible = False
End Sub

Private Sub txtRefInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      flxSPayment.SetFocus
      txtRefInput = ""
      txtRefInput.Visible = False
'      flxSPayment.ScrollBars = flexScrollBarBoth
   End If

   If KeyAscii = 13 Then
   
      If flxSPayment.TextMatrix(flxSPayment.row, 11) <> "Yes" Then
         MsgBox "The transaction is not matched. Please match the transaction before you set the reference for the tenant"
         Exit Sub
      End If

      flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol) = txtRefInput.text
      txtRefInput.Visible = False
'      flxSPayment.ScrollBars = flexScrollBarBoth
      
'      If Len(txtRefInput.text) > 0 Then
'
'         MsgBox "This will apply this reference for all batch receipts for this tenant"
'
'         Dim tenant As String
'         Dim i As Integer
'
'         tenant = flxSPayment.TextMatrix(flxSPayment.row, 2)
'         For i = 1 To flxSPayment.Rows - 1
'            If flxSPayment.TextMatrix(i, 2) = tenant Then
'               flxSPayment.TextMatrix(i, iFlxSPayCol) = txtRefInput.text
'            End If
'         Next i
'      End If

   End If
End Sub

Private Sub txtRefInput_LostFocus()
'txtRefInput_KeyPress 13
'txtRefInput.text = ""
txtRefInput.Visible = False
End Sub

Private Sub SumUpTotal()
   Dim i As Integer, cGT As Currency

   For i = 1 To frmBatchRpt.flxSPayment.Rows - 1
      If Left(frmBatchRpt.flxSPayment.TextMatrix(i, 1), 2) = "SI" Then
         cGT = cGT + Val(frmBatchRpt.flxSPayment.TextMatrix(i, 11))
      Else
         cGT = cGT - Val(frmBatchRpt.flxSPayment.TextMatrix(i, 11))
      End If
   Next i

   frmBatchRpt.txtGrossTotal.text = Format(cGT, "0.00")
End Sub




Private Sub txtSearchMemo_Change()
    Dim i As Integer
    If Len(txtSearchMemo.text) > 0 Then
        For i = 1 To flxBankTransactions.Rows - 1
           flxBankTransactions.RowHeight(i) = 0
           If InStr(1, UCase(flxBankTransactions.TextMatrix(i, bankReferenceColIndex)), UCase(txtSearchMemo.text)) > 0 Then
              flxBankTransactions.RowHeight(i) = 250
           End If
        Next i
  Else
        For i = 1 To flxBankTransactions.Rows - 1
           flxBankTransactions.RowHeight(i) = 250
        Next i
   End If
End Sub


