VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTenant2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tenants"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "frmTenant2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11670
   Begin VB.Frame fmeTenant 
      Caption         =   "Tenant Information"
      Height          =   2835
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "0.00"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtDeposite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   167
         Text            =   "0.00"
         Top             =   1695
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Tenant"
         Height          =   345
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Tenant"
         Height          =   345
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel Tenant"
         Height          =   345
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Tenant"
         Height          =   345
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1590
         TabIndex        =   13
         Top             =   150
         Width           =   3060
         Begin VB.OptionButton optCurrentTenant 
            Caption         =   "Current"
            Height          =   195
            Left            =   1020
            TabIndex        =   16
            Top             =   80
            Width           =   885
         End
         Begin VB.OptionButton optExTenant 
            Caption         =   "Ex-Tenant"
            Height          =   195
            Left            =   1950
            TabIndex        =   15
            Top             =   80
            Width           =   1035
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   80
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
         Height          =   1275
         Left            =   120
         TabIndex        =   166
         Top             =   2280
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2249
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.CommandButton cmdBankCodeLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   10
         Top             =   1350
         Width           =   255
         VariousPropertyBits=   25
         Caption         =   """"
         Size            =   "450;450"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txtBankCode 
         Height          =   315
         Left            =   8040
         TabIndex        =   165
         Top             =   1320
         Width           =   3015
         VariousPropertyBits=   746604575
         BackColor       =   15858158
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Bank:"
         Height          =   195
         Left            =   6540
         TabIndex        =   164
         Top             =   1320
         Width           =   1005
      End
      Begin MSForms.TextBox txtUnit 
         Height          =   315
         Left            =   8040
         TabIndex        =   25
         Top             =   960
         Width           =   3015
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   315
         Left            =   8040
         TabIndex        =   24
         Top             =   570
         Width           =   3015
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClient 
         Height          =   315
         Left            =   8040
         TabIndex        =   23
         Top             =   180
         Width           =   3015
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance (£):"
         Height          =   195
         Left            =   6540
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Held (£):"
         Height          =   195
         Left            =   6540
         TabIndex        =   21
         Top             =   1695
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Left            =   6540
         TabIndex        =   20
         Top             =   570
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   6540
         TabIndex        =   19
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant:"
         Height          =   195
         Left            =   450
         TabIndex        =   18
         Top             =   600
         Width           =   555
      End
      Begin MSForms.TextBox txtCompanyName 
         Height          =   315
         Left            =   1620
         TabIndex        =   8
         Top             =   1350
         Width           =   2985
         VariousPropertyBits=   746604571
         Size            =   "5265;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboSageAccountNumber 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   1740
         Width           =   3000
         VariousPropertyBits=   1820346395
         DisplayStyle    =   3
         Size            =   "5292;556"
         BoundColumn     =   0
         TextColumn      =   1
         ColumnCount     =   3
         ListRows        =   20
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtName 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   930
         Width           =   2985
         VariousPropertyBits=   746604571
         Size            =   "5265;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdTenantLookup 
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   570
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sage A/C:"
         Height          =   195
         Left            =   450
         TabIndex        =   4
         Top             =   1785
         Width           =   750
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   195
         Left            =   450
         TabIndex        =   3
         Top             =   1395
         Width           =   705
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   450
         TabIndex        =   2
         Top             =   990
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Left            =   6540
         TabIndex        =   1
         Top             =   960
         Width           =   360
      End
      Begin MSForms.TextBox txtTenantID 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   540
         Width           =   2985
         VariousPropertyBits=   746604575
         BackColor       =   15858158
         Size            =   "5265;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4508
      ScaleHeight     =   315
      ScaleWidth      =   2655
      TabIndex        =   121
      Top             =   4178
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   122
         Top             =   60
         Width           =   2475
      End
   End
   Begin TabDlg.SSTab tabTenant 
      Height          =   5625
      Left            =   75
      TabIndex        =   17
      Top             =   3000
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "&Tenant Details"
      TabPicture(0)   =   "frmTenant2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeTenantAddress"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Lease Agreement"
      TabPicture(1)   =   "frmTenant2.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeTenancyDetails"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Bank Payment Details"
      TabPicture(2)   =   "frmTenant2.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeBankPaymentDetails"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Event History"
      TabPicture(3)   =   "frmTenant2.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeEventHistory"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Memo/Attachments"
      TabPicture(4)   =   "frmTenant2.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).Control(1)=   "Frame8"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74760
         TabIndex        =   159
         Top             =   4560
         Width           =   11175
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   345
            Left            =   8670
            Style           =   1  'Graphical
            TabIndex        =   162
            Top             =   360
            Width           =   1110
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   345
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   360
            Width           =   1110
         End
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   345
            Left            =   9840
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   360
            Width           =   1110
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   180
            TabIndex        =   163
            Top             =   360
            Width           =   4890
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8625;503"
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Memo"
         Height          =   4065
         Left            =   -74760
         TabIndex        =   111
         Top             =   360
         Width           =   11175
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Height          =   345
            Left            =   9870
            TabIndex        =   115
            Top             =   3630
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Height          =   345
            Left            =   8685
            TabIndex        =   114
            Top             =   3630
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   345
            Left            =   7500
            TabIndex        =   113
            Top             =   3630
            Width           =   1125
         End
         Begin VB.TextBox txtUnitMemo 
            Height          =   3075
            Left            =   210
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   112
            Top             =   450
            Width           =   10785
         End
      End
      Begin VB.Frame fmeEventHistory 
         Caption         =   "Property Maintenance History"
         Height          =   5175
         Left            =   -74820
         TabIndex        =   88
         Top             =   360
         Width           =   11265
         Begin VB.CommandButton cmdMType 
            Caption         =   "..."
            Height          =   315
            Left            =   1680
            TabIndex        =   95
            Top             =   810
            Width           =   255
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10650
            TabIndex        =   102
            Top             =   810
            Width           =   345
         End
         Begin VB.TextBox txtEventHistoryID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   150
            TabIndex        =   93
            Top             =   270
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CommandButton cmdNewEvent 
            Caption         =   "&New"
            Height          =   315
            Left            =   7290
            TabIndex        =   92
            Top             =   4770
            Width           =   915
         End
         Begin VB.CommandButton cmdEditEvent 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   8265
            TabIndex        =   91
            Top             =   4770
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelEvent 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10215
            TabIndex        =   90
            Top             =   4770
            Width           =   915
         End
         Begin VB.CommandButton cmdSaveEvent 
            Caption         =   "&Save"
            Height          =   315
            Left            =   9240
            TabIndex        =   89
            Top             =   4770
            Width           =   915
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridEventHistory 
            Height          =   3525
            Left            =   120
            TabIndex        =   169
            Top             =   1200
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   6218
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   12632256
            BackColorSel    =   -2147483638
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.TextBox txtEventTenantID 
            Height          =   315
            Left            =   3240
            TabIndex        =   158
            Top             =   120
            Visible         =   0   'False
            Width           =   1875
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3307;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox dtpReportedDate 
            Height          =   315
            Left            =   1920
            TabIndex        =   96
            Top             =   810
            Width           =   1035
            VariousPropertyBits=   746604571
            Size            =   "1826;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   315
            Left            =   2940
            TabIndex        =   97
            Top             =   810
            Width           =   2415
            VariousPropertyBits=   746604571
            Size            =   "4260;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox dtpDateCompleted 
            Height          =   315
            Left            =   5370
            TabIndex        =   98
            Top             =   810
            Width           =   1245
            VariousPropertyBits=   746604571
            Size            =   "2196;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTaskOwner 
            Height          =   315
            Left            =   6630
            TabIndex        =   99
            Top             =   810
            Width           =   1395
            VariousPropertyBits=   746604571
            Size            =   "2461;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtContact 
            Height          =   315
            Left            =   8040
            TabIndex        =   100
            Top             =   810
            Width           =   1395
            VariousPropertyBits=   746604571
            Size            =   "2461;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox dtpRemindDate 
            Height          =   315
            Left            =   9450
            TabIndex        =   101
            Top             =   810
            Width           =   1125
            VariousPropertyBits=   746604571
            Size            =   "1984;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboEventType 
            Height          =   315
            Left            =   150
            TabIndex        =   94
            Top             =   810
            Width           =   1575
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "2778;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label46 
            Caption         =   "Remind   Date:"
            Height          =   435
            Left            =   9450
            TabIndex        =   110
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label45 
            Caption         =   "Contact:"
            Height          =   255
            Left            =   8040
            TabIndex        =   109
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label Label66 
            Caption         =   "Task Owner:"
            Height          =   255
            Left            =   6630
            TabIndex        =   108
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label Label44 
            Caption         =   "Date  Actioned:"
            Height          =   435
            Left            =   5400
            TabIndex        =   107
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label64 
            Caption         =   "Alarm"
            Height          =   195
            Left            =   10560
            TabIndex        =   106
            Top             =   570
            Width           =   405
         End
         Begin VB.Label Label43 
            Caption         =   "Reported   Date:"
            Height          =   435
            Left            =   1920
            TabIndex        =   105
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label Label42 
            Caption         =   "Event Type:"
            Height          =   255
            Left            =   150
            TabIndex        =   104
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label59 
            Caption         =   "Description:"
            Height          =   255
            Left            =   2940
            TabIndex        =   103
            Top             =   570
            Width           =   1275
         End
      End
      Begin VB.Frame fmeBankPaymentDetails 
         Caption         =   "Bank Payment Details"
         Height          =   5205
         Left            =   -74760
         TabIndex        =   66
         Top             =   360
         Width           =   10875
         Begin VB.CheckBox chkIsDefaultAC 
            Caption         =   "Yes"
            Height          =   315
            Left            =   7410
            TabIndex        =   116
            Top             =   2100
            Width           =   795
         End
         Begin VB.CommandButton cmdGetPaymentMethods 
            Caption         =   "..."
            Height          =   315
            Left            =   10380
            TabIndex        =   77
            Top             =   330
            Width           =   285
         End
         Begin VB.CommandButton cmdNewBank 
            Caption         =   "&New"
            Height          =   375
            Left            =   6840
            TabIndex        =   70
            Top             =   4700
            Width           =   915
         End
         Begin VB.CommandButton cmdEditBank 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   7815
            TabIndex        =   69
            Top             =   4700
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   9765
            TabIndex        =   68
            Top             =   4700
            Width           =   915
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "&Save"
            Height          =   375
            Left            =   8790
            TabIndex        =   67
            Top             =   4700
            Width           =   915
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBank 
            Height          =   2025
            Left            =   180
            TabIndex        =   170
            Top             =   2550
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   3572
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   12632256
            BackColorSel    =   -2147483638
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.TextBox txtBankTenantID 
            Height          =   315
            Left            =   4905
            TabIndex        =   157
            Top             =   30
            Visible         =   0   'False
            Width           =   1875
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3307;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label51 
            Caption         =   "Branch:"
            Height          =   225
            Left            =   210
            TabIndex        =   156
            Top             =   735
            Width           =   825
         End
         Begin MSForms.TextBox txtBranchName 
            Height          =   315
            Left            =   1380
            TabIndex        =   155
            Top             =   690
            Width           =   3285
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5794;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankAddress3 
            Height          =   315
            Left            =   1380
            TabIndex        =   154
            Top             =   1680
            Width           =   3285
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5794;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label49 
            Caption         =   "Default Account:"
            Height          =   225
            Left            =   5700
            TabIndex        =   153
            Top             =   2160
            Width           =   1425
         End
         Begin MSForms.ComboBox cboBankId 
            Height          =   315
            Left            =   1380
            TabIndex        =   152
            Top             =   330
            Width           =   3285
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5794;556"
            BoundColumn     =   0
            TextColumn      =   1
            ColumnCount     =   8
            ListRows        =   20
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label41 
            Caption         =   "BACS Ref:"
            Height          =   225
            Left            =   5700
            TabIndex        =   87
            Top             =   1740
            Width           =   1035
         End
         Begin MSForms.TextBox txtBACSRef 
            Height          =   315
            Left            =   7410
            TabIndex        =   86
            Top             =   1740
            Width           =   3255
            VariousPropertyBits=   746604571
            Size            =   "5741;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label40 
            Caption         =   "A/C Number:"
            Height          =   225
            Left            =   5700
            TabIndex        =   85
            Top             =   1034
            Width           =   1035
         End
         Begin VB.Label Label38 
            Caption         =   "Sort Code:"
            Height          =   225
            Left            =   5700
            TabIndex        =   84
            Top             =   1386
            Width           =   1035
         End
         Begin MSForms.TextBox txtBankSortCode 
            Height          =   315
            Left            =   7410
            TabIndex        =   83
            Top             =   1386
            Width           =   1545
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "2725;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankACNumber 
            Height          =   315
            Left            =   7410
            TabIndex        =   82
            Top             =   1034
            Width           =   3255
            VariousPropertyBits=   746604571
            Size            =   "5741;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label37 
            Caption         =   "Payment Method:"
            Height          =   225
            Left            =   5700
            TabIndex        =   81
            Top             =   330
            Width           =   1305
         End
         Begin VB.Label Label36 
            Caption         =   "A/C Name:"
            Height          =   225
            Left            =   5700
            TabIndex        =   80
            Top             =   682
            Width           =   1035
         End
         Begin MSForms.ComboBox cboPaymentMethod 
            Height          =   315
            Left            =   7410
            TabIndex        =   79
            Top             =   330
            Width           =   2985
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5265;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankACName 
            Height          =   315
            Left            =   7410
            TabIndex        =   78
            Top             =   682
            Width           =   3255
            VariousPropertyBits=   746604571
            Size            =   "5741;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label39 
            Caption         =   "Post Code:"
            Height          =   225
            Left            =   210
            TabIndex        =   76
            Top             =   2000
            Width           =   825
         End
         Begin MSForms.TextBox txtBankPostCode 
            Height          =   315
            Left            =   1380
            TabIndex        =   75
            Top             =   2000
            Width           =   1515
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "2672;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankAddress2 
            Height          =   315
            Left            =   1380
            TabIndex        =   74
            Top             =   1365
            Width           =   3285
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5794;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankAddress1 
            Height          =   315
            Left            =   1380
            TabIndex        =   73
            Top             =   1050
            Width           =   3285
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5794;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label62 
            Caption         =   "Address:"
            Height          =   225
            Left            =   210
            TabIndex        =   72
            Top             =   1110
            Width           =   825
         End
         Begin VB.Label Label61 
            Caption         =   "Bank:"
            Height          =   225
            Left            =   210
            TabIndex        =   71
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame fmeTenancyDetails 
         Caption         =   "Tenancy Details"
         Height          =   3315
         Left            =   -74820
         TabIndex        =   45
         Top             =   1170
         Width           =   11115
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Ref:"
            Height          =   195
            Left            =   315
            TabIndex        =   124
            Top             =   750
            Width           =   780
         End
         Begin MSForms.TextBox txtLeaseId 
            Height          =   315
            Left            =   2355
            TabIndex        =   123
            Top             =   690
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Rent Review Date:"
            Height          =   195
            Left            =   6000
            TabIndex        =   65
            Top             =   1530
            Width           =   1740
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            Height          =   195
            Left            =   315
            TabIndex        =   64
            Top             =   2310
            Width           =   1080
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Holding Over:"
            Height          =   195
            Left            =   315
            TabIndex        =   63
            Top             =   1920
            Width           =   975
         End
         Begin MSForms.TextBox TextBox15 
            Height          =   315
            Left            =   2340
            TabIndex        =   62
            Top             =   2640
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtRentReviewDate 
            Height          =   315
            Left            =   7980
            TabIndex        =   61
            Top             =   1470
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TextBox13 
            Height          =   315
            Left            =   2340
            TabIndex        =   60
            Top             =   2250
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHoldingOver 
            Height          =   315
            Left            =   2340
            TabIndex        =   59
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEndDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   58
            Top             =   1470
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtStartDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   57
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Left            =   315
            TabIndex        =   56
            Top             =   1140
            Width           =   765
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Termination Date:"
            Height          =   195
            Left            =   315
            TabIndex        =   55
            Top             =   2700
            Width           =   1935
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            Height          =   195
            Left            =   315
            TabIndex        =   54
            Top             =   1530
            Width           =   720
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Charge:"
            Height          =   195
            Left            =   6000
            TabIndex        =   53
            Top             =   1920
            Width           =   1140
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/C Frequency:"
            Height          =   195
            Left            =   6000
            TabIndex        =   52
            Top             =   2310
            Width           =   1125
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Frequency:"
            Height          =   195
            Left            =   6000
            TabIndex        =   51
            Top             =   1140
            Width           =   1185
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent:"
            Height          =   195
            Left            =   6000
            TabIndex        =   50
            Top             =   750
            Width           =   390
         End
         Begin MSForms.TextBox txtBRPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   49
            Top             =   690
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtBASERENTFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   48
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSCPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   47
            Top             =   1860
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtSERVICECHARGEFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   46
            Top             =   2250
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fmeTenantAddress 
         Height          =   4995
         Left            =   330
         TabIndex        =   27
         Top             =   480
         Width           =   10995
         Begin VB.CommandButton cmdCancelTenantAddress 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   9840
            TabIndex        =   148
            Top             =   4590
            Width           =   1035
         End
         Begin VB.CommandButton cmdSaveTenantAddress 
            Caption         =   "&Save"
            Height          =   315
            Left            =   8775
            TabIndex        =   147
            Top             =   4590
            Width           =   1035
         End
         Begin VB.CommandButton cmdEditTenantAddress 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7710
            TabIndex        =   146
            Top             =   4590
            Width           =   1035
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alternative Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6360
            TabIndex        =   151
            Top             =   480
            Width           =   1725
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   150
            Top             =   480
            Width           =   1410
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice To:"
            Height          =   195
            Left            =   3030
            TabIndex        =   149
            Top             =   150
            Width           =   810
         End
         Begin MSForms.ComboBox cboInvoiceTo 
            Height          =   315
            Left            =   4080
            TabIndex        =   145
            Top             =   90
            Width           =   3000
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5292;556"
            BoundColumn     =   0
            TextColumn      =   1
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "0"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   6570
            TabIndex        =   144
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   6570
            TabIndex        =   143
            Top             =   2670
            Width           =   780
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Left            =   6570
            TabIndex        =   142
            Top             =   840
            Width           =   600
         End
         Begin MSForms.TextBox txtContact2 
            Height          =   315
            Left            =   7680
            TabIndex        =   141
            Top             =   780
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine1 
            Height          =   315
            Left            =   7680
            TabIndex        =   140
            Top             =   1170
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   139
            Top             =   1500
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine3 
            Height          =   315
            Left            =   7680
            TabIndex        =   138
            Top             =   1830
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine4 
            Height          =   315
            Left            =   7680
            TabIndex        =   137
            Top             =   2160
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillPostCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   136
            Top             =   2610
            Width           =   1815
            VariousPropertyBits=   746604571
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillFax 
            Height          =   315
            Left            =   7680
            TabIndex        =   135
            Top             =   4170
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillTelephone 
            Height          =   315
            Left            =   7680
            TabIndex        =   134
            Top             =   3780
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   133
            Top             =   3390
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail2 
            Height          =   315
            Left            =   7680
            TabIndex        =   132
            Top             =   3000
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Left            =   6570
            TabIndex        =   131
            Top             =   3090
            Width           =   420
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Left            =   6570
            TabIndex        =   130
            Top             =   3465
            Width           =   810
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Left            =   6570
            TabIndex        =   129
            Top             =   4200
            Width           =   300
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   6570
            TabIndex        =   128
            Top             =   3825
            Width           =   510
         End
         Begin MSForms.TextBox txtHOFax 
            Height          =   315
            Left            =   1440
            TabIndex        =   44
            Top             =   4200
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOTelephone 
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   3810
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   42
            Top             =   3420
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail1 
            Height          =   315
            Left            =   1440
            TabIndex        =   41
            Top             =   3030
            Width           =   2445
            VariousPropertyBits=   746604571
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Left            =   360
            TabIndex        =   40
            Top             =   3060
            Width           =   420
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Left            =   360
            TabIndex        =   39
            Top             =   3435
            Width           =   810
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   4170
            Width           =   300
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   3795
            Width           =   510
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   360
            TabIndex        =   36
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   2640
            Width           =   780
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   600
         End
         Begin MSForms.TextBox txtContact1 
            Height          =   315
            Left            =   1440
            TabIndex        =   33
            Top             =   810
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   32
            Top             =   1200
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine2 
            Height          =   315
            Left            =   1440
            TabIndex        =   31
            Top             =   1530
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine3 
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine4 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            Top             =   2190
            Width           =   2985
            VariousPropertyBits=   746604571
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOPostCode 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   2640
            Width           =   1815
            VariousPropertyBits=   746604571
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00B3C0C6&
            BackStyle       =   1  'Opaque
            Height          =   4065
            Left            =   120
            Top             =   480
            Width           =   4665
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00B3C0C6&
            BackStyle       =   1  'Opaque
            Height          =   4065
            Left            =   6120
            Top             =   480
            Width           =   4665
         End
      End
   End
   Begin VB.PictureBox fmeTenantLookup 
      BackColor       =   &H00B3C0C6&
      Height          =   2025
      Left            =   1800
      ScaleHeight     =   1965
      ScaleWidth      =   7905
      TabIndex        =   117
      Top             =   1200
      Visible         =   0   'False
      Width           =   7965
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
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTenantLookup 
         Height          =   1605
         Left            =   0
         TabIndex        =   118
         Top             =   330
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   2831
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   8687183
         ForeColorFixed  =   16777215
         BackColorSel    =   13884353
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridColorFixed  =   8421376
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox txtSearchTenant 
         Height          =   315
         Left            =   0
         TabIndex        =   120
         Top             =   0
         Width           =   1785
         VariousPropertyBits=   746604571
         Size            =   "3149;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label Label72 
      BackColor       =   &H000000C0&
      Height          =   75
      Left            =   0
      TabIndex        =   26
      Top             =   2840
      Width           =   12000
   End
End
Attribute VB_Name = "frmTenant2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LOAD_TENANT_TENANTID As String

Dim NEWMODE_ As Boolean
Dim SEARCHTenantMODE_ As Boolean
'' Tenant DETAILS ENTRY FLAG
Dim M_HISTORY_NEW_ENTRY_ As Boolean
Dim BANK_PAYMENT_NEW_ENTRY_ As Boolean
Dim IMAGE_FILE_NAME_ As String

Dim DSN_ALARM_ As String

Private Sub cboBankId_Change()

If Not IsNull(cboBankId) And cboBankId <> "" Then
    txtBankSortCode.text = cboBankId.Column(2)
    txtBranchName.text = cboBankId.Column(3)
    txtBankAddress1.text = cboBankId.Column(4)
    txtBankAddress2.text = cboBankId.Column(5)
    txtBankAddress3.text = cboBankId.Column(6)
    txtBankPostCode.text = cboBankId.Column(7)
End If

End Sub

Private Sub cboBankId_Click()
'MsgBox cboBankId.ListCount
'txtBankSortCode.text = cboBankId.Column(2)
'txtBranchName.text = cboBankId.Column(3)
'txtBankAddress1.text = cboBankId.Column(4)
'txtBankAddress2.text = cboBankId.Column(5)
'txtBankAddress3.text = cboBankId.Column(6)
''txtBankPostCode.text = cboBankId.Column(7)
End Sub

Private Sub cboSageAccountNumber_LostFocus()
   If NEWMODE_ And cboSageAccountNumber.text <> "" Then
      txtTenantID.text = cboSageAccountNumber.Value
   End If
End Sub

Private Sub cmdBankCodeLookup_Click()
   MousePointer = vbHourglass
   gridBankCode.Top = 1740
   gridBankCode.Left = 6740
   gridBankCode.Visible = True
   gridBankCode.ZOrder 0
'
   BankAccount
'
   MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
ComponentInFrameEnableMode frmTenant2, fmeTenant, DefaultMode

NEWMODE_ = False
SEARCHTenantMODE_ = True
'txtTenantID.Enabled = True
txtName.Enabled = True
cmdTenantLookup.Enabled = True

If txtTenantID.text = "" Then
    Exit Sub
End If
tabTenant.Enabled = True
cmdBankCodeLookup.Enabled = False
End Sub

Private Sub cmdCancelBank_Click()
ComponentInFrameEnableMode frmTenant2, fmeBankPaymentDetails, DefaultMode
End Sub

Private Sub cmdCancelEvent_Click()
ComponentInFrameEnableMode frmTenant2, fmeEventHistory, DefaultMode
End Sub

Private Sub cmdCancelTenantAddress_Click()
ComponentInFrameEnableMode frmTenant2, fmeTenantAddress, DefaultMode
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachment cmbFiles, "Tenants", txtTenantID.text
   MsgBox "File has been saved successfull, Thanks"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachment cmbFiles, cmbFiles.Column(2), txtTenantID.text, "Tenants"
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdEdit_Click()
If txtTenantID.text = "" Then
    MsgBox "Please select a Tenant to continue.", vbInformation, "Edit Tenant"
    Exit Sub
End If
NEWMODE_ = False
SEARCHTenantMODE_ = False
ComponentInFrameEnableMode frmTenant2, fmeTenant, EditMode

'cboCurrentTenant.Enabled = False
'txtTenantID.Locked = True
txtName.SetFocus
tabTenant.Enabled = False
cmdTenantLookup.Enabled = False
cmdBankCodeLookup.Enabled = True
End Sub

Private Sub cmdEditBank_Click()
BANK_PAYMENT_NEW_ENTRY_ = False
ComponentInFrameEnableMode frmTenant2, fmeBankPaymentDetails, EditMode
cboBankId.Locked = True
txtBankACNumber.Locked = True
End Sub

Private Sub cmdEditEvent_Click()
M_HISTORY_NEW_ENTRY_ = False
ComponentInFrameEnableMode frmTenant2, fmeEventHistory, EditMode
End Sub

Private Sub cmdEditTenantAddress_Click()
ComponentInFrameEnableMode frmTenant2, fmeTenantAddress, EditMode
End Sub

Private Sub cmdGridTenantLookup_Click()
fmeTenantLookup.Visible = False
End Sub

Private Sub cmdNew_Click()
   NEWMODE_ = True
   SEARCHTenantMODE_ = False

   tabTenant.Enabled = False
   ComponentInFrameEnableMode frmTenant2, fmeTenant, NewEntryMode
   cmdBankCodeLookup.Enabled = True
   txtName.SetFocus
End Sub

Private Sub cmdNewBank_Click()
   BANK_PAYMENT_NEW_ENTRY_ = True
   ComponentInFrameEnableMode frmTenant2, fmeBankPaymentDetails, NewEntryMode
   cboBankId.Locked = False
   txtBankACNumber.Locked = False
End Sub

Private Sub cmdNewEvent_Click()
   M_HISTORY_NEW_ENTRY_ = True
   ComponentInFrameEnableMode frmTenant2, fmeEventHistory, NewEntryMode
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a tenant to continue.", vbExclamation, "No Tenant Selected"
      Exit Sub

   ElseIf txtName.text = "" Then
      MsgBox "Please enter a Tenant Name to continue.", vbExclamation, "No Tenant Name"
      txtName.SetFocus
      txtName.text = ""
      Exit Sub
   End If
   
   If txtDeposite.text = "" Then txtDeposite.text = "0.00"
   If txtBalance.text = "" Then txtBalance.text = "0.00"

   If SaveTenantInformation Then
       MsgBox "The record is saved successfully", vbInformation
   End If

   NEWMODE_ = False
   ComponentInFrameEnableMode frmTenant2, fmeTenant, DefaultMode
   SEARCHTenantMODE_ = True

   txtName.Enabled = True
   tabTenant.Enabled = True
   cmdTenantLookup.Enabled = True
   cmdBankCodeLookup.Enabled = False
End Sub

Public Function PopulateTenantLookup(ByVal strFilter_ As String)
   
  'cmdClientID.Default = True
   Dim adoConn As New ADODB.Connection
   Dim sSQLQuery_ As String
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
      sSQLQuery_ = "SELECT " _
    & " SageAccountNumber, Name, " _
    & " iif(isnull(HOAddressLine1),'',HOAddressLine1) + ' ' + iif(isnull(HOAddressLine2),'',HOAddressLine2) + ' ' +  iif(isnull(HOAddressLine3),'',HOAddressLine3) as Address, " _
    & " HOPostCode , HOTelephone " _
    & " From " _
    & " Tenants " & strFilter_
                       
   Dim iRow As Integer
   iRow = 1

   gridTenantLookup.Clear
   gridTenantLookup.Rows = 2
   gridTenantLookup.Cols = 5
   ConfigurFlexGrid
   
'   On Error Resume Next
'   While Not rstTenant_.EOF
'      gridTenantLookup.TextMatrix(iRow, 0) = IIf(rstTenant_!TenantId = Null, "", rstTenant_!TenantId)
'      gridTenantLookup.TextMatrix(iRow, 1) = IIf(rstTenant_!TenantName = Null, "", rstTenant_!TenantName)
'      gridTenantLookup.TextMatrix(iRow, 2) = IIf(rstTenant_!Address = Null, "", rstTenant_!Address)
'      gridTenantLookup.TextMatrix(iRow, 3) = IIf(rstTenant_!ProPOSTCODE = Null, "", rstTenant_!ProPOSTCODE)
'      gridTenantLookup.TextMatrix(iRow, 4) = IIf(rstTenant_!TotalArea = Null, "", rstTenant_!TotalArea)
'      rstTenant_.MoveNext
'      If Not rstTenant_.EOF Then gridTenantLookup.AddItem ""
'      iRow = iRow + 1
'   Wend

 populateGrid adoConn, sSQLQuery_, gridTenantLookup
 
   adoConn.Close
   Set adoConn = Nothing
   
   'cmdSelected.Enabled = True
End Function

Private Sub cmdSaveBank_Click()
   Dim sSQLQuery As String, sWhere As String
   Dim adoConn As New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
     
   txtBankTenantID.text = txtTenantID.text

   sSQLQuery = "SELECT * " & _
               "FROM TenantBankDetails"
   sWhere = " WHERE BankTenantID = '" & txtTenantID.text & "' AND " & _
               "BankID = '" & cboBankId.Value & "' AND " & _
               "BankACNumber = '" & txtBankACNumber.text & "'"

    If BANK_PAYMENT_NEW_ENTRY_ Then
        If PostToDBUsingADODB(frmTenant2, fmeBankPaymentDetails, adoConn, sSQLQuery & sWhere, True) Then
            MsgBox "The bank payment methods has been saved successfully.", vbInformation
        Else
            MsgBox "Error occured while saving the bank payment method.", vbInformation
        End If
    Else
        If PostToDBUsingADODB(frmTenant2, fmeBankPaymentDetails, adoConn, sSQLQuery & sWhere, False) Then
            MsgBox "The bank payment method has been updated successfully", vbInformation
        Else
            MsgBox "Error occured while updating the contact information", vbInformation
        End If
    End If
    
    populateGrid adoConn, "SELECT * FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'", gridBank
    adoConn.Close
    Set adoConn = Nothing


ComponentInFrameEnableMode frmTenant2, fmeBankPaymentDetails, DefaultMode
End Sub

Private Sub cmdSaveEvent_Click()
   Dim adoConn As New ADODB.Connection
   Dim szHeader As String, sSQLQuery_ As String

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   ' Event Type
   txtEventHistoryID.text = txtTenantID.text & "-" & cboEventType.Value & "-" & dtpReportedDate.Value
   txtEventTenantID.text = txtTenantID.text
   Dim sSQLQuery As String
   sSQLQuery = "SELECT * " & _
                 "FROM TenantEventHistory WHERE " & _
                 " EventHistoryID = '" & txtEventHistoryID.text & "'"

    If M_HISTORY_NEW_ENTRY_ Then
        If PostToDBUsingADODB(frmTenant2, fmeEventHistory, adoConn, sSQLQuery, True) Then
            MsgBox "The event history has been saved successfully.", vbInformation
        Else
            MsgBox "Error occured while saving the event history.", vbInformation
        End If
    Else
        If PostToDBUsingADODB(frmTenant2, fmeEventHistory, adoConn, sSQLQuery, False) Then
            MsgBox "The event history has been updated successfully.", vbInformation
        Else
            MsgBox "Error occured while updating the event history.", vbInformation
        End If
    End If

   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
   sSQLQuery_ = "SELECT * FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
   populateGridSimply adoConn, sSQLQuery_, gridEventHistory, szHeader
   
   adoConn.Close
   Set adoConn = Nothing
   ComponentInFrameEnableMode frmTenant2, fmeEventHistory, DefaultMode
End Sub

Private Sub cmdSaveTenantAddress_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   ' Event Type
   Dim sSQLQuery As String
   sSQLQuery = "SELECT * " & _
                "FROM TENANTS " & _
                "WHERE TenantID = '" & txtTenantID.text & "'"

   If PostToDBUsingADODB(frmTenant2, fmeTenantAddress, adoConn, sSQLQuery, False) Then
       MsgBox "The contact details of the tenant has been updated successfully", vbInformation
   Else
       MsgBox "Error occured while updating the contact information", vbInformation
   End If

   adoConn.Close
   Set adoConn = Nothing

   ComponentInFrameEnableMode frmTenant2, fmeTenantAddress, DefaultMode
End Sub

Private Sub cmdTenantLookup_Click()
   fmeTenantLookup.Visible = True
   fmeTenantLookup.ZOrder 0
   gridTenantLookup.Visible = True
   txtSearchTenant.SetFocus
   txtSearchTenant.text = ""
   PopulateTenantLookup ""
End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtUnitMemo.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveMemo("Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo) Then
      MsgBox "Memo has been saved successfully.", vbInformation + vbOKOnly, "Memo"
   End If
   MemoButtonEnable False
End Sub

Private Sub dtpDateCompleted_Change()
   TextBoxChangeDate dtpDateCompleted
End Sub

Private Sub dtpDateCompleted_KeyPress(KeyAscii As MSForms.ReturnInteger)
   TextBoxKeyPrsDate dtpDateCompleted, KeyAscii
End Sub

Private Sub dtpDateCompleted_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpDateCompleted
End Sub

Private Sub dtpRemindDate_Change()
   TextBoxChangeDate dtpRemindDate
End Sub

Private Sub dtpRemindDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
   TextBoxKeyPrsDate dtpRemindDate, KeyAscii
End Sub

Private Sub dtpRemindDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpRemindDate
End Sub

Private Sub dtpReportedDate_Change()
   TextBoxChangeDate dtpReportedDate
End Sub

Private Sub dtpReportedDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
   TextBoxKeyPrsDate dtpReportedDate, KeyAscii
End Sub

Private Sub dtpReportedDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpReportedDate
End Sub

Private Sub Form_Activate()
   If LOAD_TENANT_TENANTID <> "" Then
      LoadTenantByTenantID
   End If
End Sub

Private Sub Form_Load()
MousePointer = vbHourglass

Me.Top = 50
Me.Left = 50
Me.Caption = "Tenants"
DSN_ALARM_ = "WD_ALARM"
tabTenant.Tab = 0
ComponentInFrameEnableMode frmTenant2, fmeTenant, DefaultMode

txtSearchTenant.Enabled = True
NEWMODE_ = False
SEARCHTenantMODE_ = True

'' Populate the codes
PopulateCodes
SageCustomerAccCombo cboSageAccountNumber
TenantTabEnabled False

MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub ConfigurFlexGrid()

   'gridTenantLookup.Left = txtTenantID.Height
      
   fmeTenantLookup.Visible = True
   gridTenantLookup.Visible = True
   
   gridTenantLookup.RowHeight(0) = 350
   gridTenantLookup.Row = 0
   Dim i As Integer
   For i = 0 To gridTenantLookup.Cols - 1
        gridTenantLookup.Col = i
        gridTenantLookup.CellFontBold = True
   Next i
   
   
   gridTenantLookup.ColWidth(0) = 900
   gridTenantLookup.TextMatrix(0, 0) = "Sage A/C"
   
   gridTenantLookup.ColWidth(1) = 2100
   gridTenantLookup.TextMatrix(0, 1) = "Name"
   
   gridTenantLookup.ColWidth(2) = 2500
   gridTenantLookup.TextMatrix(0, 2) = "Address"
   
   gridTenantLookup.ColWidth(3) = 1000
   gridTenantLookup.TextMatrix(0, 3) = "Post Code"
   
   gridTenantLookup.ColWidth(4) = 1000
   gridTenantLookup.TextMatrix(0, 4) = "Telephone"
   
End Sub

Private Sub gridBank_Click()
populateControl frmTenant2, gridBank
End Sub

Private Sub gridBankCode_Click()
   txtBankCode.text = gridBankCode.TextMatrix(gridBankCode.Row, 0)
   gridBankCode.Visible = False
End Sub

Private Sub gridEventHistory_Click()
populateControl frmTenant2, gridEventHistory
End Sub

Private Sub LoadTenantByTenantID()
   Dim sSQLQuery_ As String, szHeader As String
   
   SEARCHTenantMODE_ = False
   fmeTenantLookup.Visible = False
   
   '' LOAD MAIN Tenant INFORMATION
   
   fmeLoading.Visible = True
   fmeLoading.Refresh
   
   Dim adoConn As New ADODB.Connection
   
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   'Populate the Tenant Header
   PopulateTenantInformation adoConn, LOAD_TENANT_TENANTID
   ' Populate Bank Details
   populateGrid adoConn, "SELECT * FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'", gridBank
   ' Populate Event History
   SetGridEventHistory
   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
   sSQLQuery_ = "SELECT * FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
   populateGridSimply adoConn, sSQLQuery_, gridEventHistory, szHeader
   
   'populateGrid adoConn, "SELECT * FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'", gridEventHistory
   
   '' LOAD Tenant DETAIL INFORMATION
   RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo
   
   fmeLoading.Visible = False
   adoConn.Close
   Set adoConn = Nothing
   
   ' SET OTHERS
   'fmeTenantLookup.Visible = False
   SEARCHTenantMODE_ = True
   TenantTabEnabled True
End Sub

Private Sub gridTenantLookup_Click()
Dim sSQLQuery_ As String, szHeader As String

SEARCHTenantMODE_ = False
fmeTenantLookup.Visible = False

'' LOAD MAIN Tenant INFORMATION

fmeLoading.Visible = True
fmeLoading.Refresh

Dim adoConn As New ADODB.Connection

adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'Populate the Tenant Header
PopulateTenantInformation adoConn, gridTenantLookup.TextMatrix(gridTenantLookup.Row, 0)
' Populate Bank Details
populateGrid adoConn, "SELECT * FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'", gridBank
' Populate Event History
SetGridEventHistory
szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
sSQLQuery_ = "SELECT * FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
populateGridSimply adoConn, sSQLQuery_, gridEventHistory, szHeader

'populateGrid adoConn, "SELECT * FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'", gridEventHistory

'' LOAD Tenant DETAIL INFORMATION
RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo

fmeLoading.Visible = False
adoConn.Close
Set adoConn = Nothing

' SET OTHERS
'fmeTenantLookup.Visible = False
SEARCHTenantMODE_ = True
TenantTabEnabled True

End Sub




Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub tabTenant_Click(PreviousTab As Integer)
   Select Case tabTenant.Tab
   Case 4:
      If txtTenantID.text <> "" Then _
            Call LoadAttachmentFiles(cmbFiles, txtTenantID.text, "Tenants")
   End Select
End Sub

Private Sub txtBankAddress1_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankAddress2_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankAddress3_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankPostCode_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankSortCode_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBranchName_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtSearchTenant_Change()
If Not SEARCHTenantMODE_ Then
    Exit Sub
End If
Dim sFilter_ As String
sFilter_ = "WHERE TenantID LIKE '" & Trim(txtSearchTenant.text) & "%' " & _
              "ORDER BY TenantID;"
PopulateTenantLookup sFilter_
End Sub

Public Function PopulateTenantInformation(ByVal adoConn As ADODB.Connection, ByVal sTenantSageAC As String) As Boolean

   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT TENANTS.*, LEASEINFO.* " _
        & " FROM TENANTS LEFT JOIN (" _
        & " SELECT " _
        & " CLIENT.CLIENTNAME AS CLIENT, " _
        & " PROPERTY.PROPERTYID + '-' + PROPERTY.PROPERTYNAME AS PROPERTY, " _
        & " UNITS.UNITNUMBER + '-'+ UNITS.UNITNAME AS UNIT, " _
        & " LeaseDetails.LeaseID,LeaseDetails.SageAccountNumber as LeaseSAGEAC, " _
        & " LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.RentReviewDate, " _
        & " LeaseDetails.BRPayable, LeaseDetails.BRFrequency, " _
        & " LeaseDetails.SCPayable, LeaseDetails.SCFrequency " _
        & " From " _
        & " LEASEDETAILS, " _
        & " UNITS , CLIENT, TENANTS, PROPERTY " _
        & " Where " _
        & " LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " _
        & " UNITS.PROPERTYID = PROPERTY.PROPERTYID AND " _
        & " Property.CLIENTID = CLIENT.CLIENTID " _
        & " )AS LEASEINFO ON TENANTS.SAGEACCOUNTNUMBER = LEASEINFO.LeaseSAGEAC " _
        & " WHERE SageAccountNumber = '" & sTenantSageAC & "'"

'    Debug.Print sSQLQuery_
   
    If Not FillFormUsingADODB(frmTenant2, adoConn, sSQLQuery_) Then
         MsgBox "WARNING !! No information found for the specified Tenant.", vbExclamation
    End If
    
    If txtLeaseId.text = "" Then
        MsgBox "WARNING !! There is no Lease setup for this Tenant.", vbExclamation
    End If
    
End Function


Public Sub PopulateCodes()
          
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
     
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
     
   ' Event Type
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'MTYP'"

   populateCombo adoConn, sSQLQuery, cboEventType
   
   ' Invoice Address
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'INVADD'"

   populateCombo adoConn, sSQLQuery, cboInvoiceTo
   
   ' Banks
   sSQLQuery = "SELECT BANK_ID, BANK_NAME, SORT_CODE, BANK_BRANCH, BANK_ADDRESS1, BANK_ADDRESS2, BANK_ADDRESS3, BANK_POST_CODE " & _
                 "FROM tlbBank "

   populateCombo adoConn, sSQLQuery, cboBankId
   
   ' Payment Method
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'PM'"

   populateCombo adoConn, sSQLQuery, cboPaymentMethod
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub EventHistoryButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
    Case ComponentMode.DefaultMode
        cmdNewEvent.Enabled = True
        cmdEditEvent.Enabled = False
        cmdSaveEvent.Enabled = False
        cmdCancelEvent.Enabled = False
        gridEventHistory.Enabled = True
    
        cboEventType.Enabled = False
        dtpReportedDate.Enabled = False
        txtDescription.Enabled = False
        dtpDateCompleted.Enabled = False
        txtTaskOwner.Enabled = False
        txtContact.Enabled = False
        dtpRemindDate.Enabled = False
        chkAlarm.Enabled = False

    Case ComponentMode.GridRowOnSelection
        cmdNewEvent.Enabled = True
        cmdEditEvent.Enabled = True
        cmdSaveEvent.Enabled = False
        cmdCancelEvent.Enabled = False
        gridEventHistory.Enabled = True
    
    Case ComponentMode.NewEntryMode
        cmdNewEvent.Enabled = False
        cmdEditEvent.Enabled = False
        cmdSaveEvent.Enabled = True
        cmdCancelEvent.Enabled = True
        gridEventHistory.Enabled = False
    
        cboEventType.Enabled = True
        dtpReportedDate.Enabled = True
        txtDescription.Enabled = True
        txtDescription.text = ""
        dtpDateCompleted.Enabled = True
        txtTaskOwner.Enabled = True
        txtTaskOwner.text = ""
        txtContact.Enabled = True
        txtContact.text = ""
        dtpRemindDate.Enabled = True
        chkAlarm.Enabled = True
        chkAlarm.Value = 0

        Case ComponentMode.EditMode
            cmdNewEvent.Enabled = False
            cmdEditEvent.Enabled = False
            cmdSaveEvent.Enabled = True
            cmdCancelEvent.Enabled = True
            gridEventHistory.Enabled = False
        
            cboEventType.Enabled = True
            dtpReportedDate.Enabled = True
            txtDescription.Enabled = True
            dtpDateCompleted.Enabled = True
            txtTaskOwner.Enabled = True
            txtContact.Enabled = True
            dtpRemindDate.Enabled = True
            chkAlarm.Enabled = True
    End Select
End Sub

Public Sub SetGridEventHistory()
   gridEventHistory.Clear
   gridEventHistory.Rows = 2
   gridEventHistory.Cols = 10

   gridEventHistory.ColWidth(0) = 0
   gridEventHistory.ColWidth(1) = 0
   gridEventHistory.ColWidth(2) = cboEventType.Width + cmdMType.Width
   gridEventHistory.ColWidth(3) = dtpReportedDate.Width
   gridEventHistory.ColWidth(4) = txtDescription.Width
   gridEventHistory.ColWidth(5) = dtpDateCompleted.Width
   gridEventHistory.ColWidth(6) = txtTaskOwner.Width
   gridEventHistory.ColWidth(7) = txtContact.Width
   gridEventHistory.ColWidth(8) = dtpRemindDate.Width
   gridEventHistory.ColWidth(9) = chkAlarm.Width
End Sub

Public Function SaveTenantInformation() As Boolean
   Dim adoConn As New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   ' Event Type
   Dim sSQLQuery As String
   sSQLQuery = "SELECT * " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & txtTenantID.text & "'"

   If Not NEWMODE_ Then
       If PostToDBUsingADODB(frmTenant2, fmeTenant, adoConn, sSQLQuery, False) Then
           SaveTenantInformation = True
       Else
           SaveTenantInformation = False
       End If
   Else
       If PostToDBUsingADODB(frmTenant2, fmeTenant, adoConn, sSQLQuery, True) Then
           SaveTenantInformation = True
       Else
           SaveTenantInformation = False
       End If
   End If

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub SageCustomerAccCombo(ByVal cboSage As Control)
   cboSage.Clear
   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oSalesRecord As SageDataObject120.SalesRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
'   oSDO.Workspaces.Clear
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then
      Set oSalesRecord = oWS.CreateObject("SalesRecord")

      ' Added by Asif. 18/01/06
      Dim TotalRow, TotalCol As Long
      Dim data() As String

      TotalRow = oSalesRecord.Count
      TotalCol = 2 - 1

      ReDim data(TotalCol, TotalRow) As String

      ' Move to the First Record
      oSalesRecord.MoveFirst
      Dim rRow As Integer
      For rRow = 0 To oSalesRecord.Count - 1
         data(0, rRow) = oSalesRecord.Fields.Item("ACCOUNT_REF").Value
         data(1, rRow) = oSalesRecord.Fields.Item("NAME").Value
         oSalesRecord.MoveNext
      Next rRow

      cboSage.Clear
      cboSage.ColumnCount = 2
      cboSage.Column() = data()
      cboSage.BoundColumn = 1
      cboSage.TextColumn = 1

      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

Error_Handler:

   MsgBox "(pcm_002) The SDO generated the following error: " & oSDO.LastError.text
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub TenantTabEnabled(ByVal IsEnabled As Boolean)
    tabTenant.Enabled = IsEnabled
    
    If IsEnabled Then
        ComponentInFrameEnableMode frmTenant2, fmeTenantAddress, DefaultMode
        ComponentInFrameEnableMode frmTenant2, fmeTenantAddress, DefaultMode
        ComponentInFrameEnableMode frmTenant2, fmeTenancyDetails, EditMode
        ComponentInFrameEnableMode frmTenant2, fmeBankPaymentDetails, DefaultMode
        ComponentInFrameEnableMode frmTenant2, fmeEventHistory, DefaultMode
    End If
End Sub

Private Sub BankAccount()
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim clsBankAC As clsArray
   Dim iBankAc As Integer
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oBankRecord As SageDataObject120.BankRecord
   Dim oNominalRecord As SageDataObject120.NominalRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then
   
      Set oBankRecord = oWS.CreateObject("BankRecord")
   
      ' Move to the First Record
      oBankRecord.MoveFirst
      Set clsBankAC = New clsArray
      For iBankAc = 1 To oBankRecord.Count
         clsBankAC.AddItem oBankRecord.Fields.Item("ACCOUNT_REF").Value
         oBankRecord.MoveNext
      Next iBankAc
   
      Set oBankRecord = Nothing
   
      Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
      oNominalRecord.MoveFirst
   
      Dim rRow As Integer
      Dim iRec As Integer
      rRow = 1
      
      gridBankCode.TextMatrix(0, 0) = "Reference"
      gridBankCode.TextMatrix(0, 1) = "Name"
      gridBankCode.ColWidth(0) = 1200
      gridBankCode.ColWidth(1) = 2600
      
      For iRec = 1 To oNominalRecord.Count
         If clsBankAC.IsItem(CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)) Then
            gridBankCode.TextMatrix(rRow, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
            gridBankCode.TextMatrix(rRow, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
            gridBankCode.AddItem ""
            rRow = rRow + 1
         End If
         oNominalRecord.MoveNext
      Next iRec
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text

   Set oBankRecord = Nothing
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

