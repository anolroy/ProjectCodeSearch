VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLease2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lease Information"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   12300
   Icon            =   "Lease2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12300
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   330
      Left            =   6600
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Main"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoBreaches 
      Height          =   330
      Left            =   4560
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Breaches"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraMain 
      Caption         =   "Select Lease"
      Height          =   1665
      Left            =   80
      TabIndex        =   31
      Top             =   60
      Width           =   12135
      Begin VB.TextBox txtUnitName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   3075
      End
      Begin VB.TextBox txtTenant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   660
         Width           =   3075
      End
      Begin VB.CommandButton cmdTenants 
         Caption         =   "V"
         Enabled         =   0   'False
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
         Left            =   4400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   660
         Width           =   255
      End
      Begin VB.CheckBox chkExpLease 
         Caption         =   "Expired Leases only"
         Height          =   315
         Left            =   4800
         TabIndex        =   181
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdLease 
         Caption         =   "V"
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
         Left            =   4400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtLeaseID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   3075
      End
      Begin VB.ComboBox cboTenant 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSForms.ComboBox cboUnit 
         Height          =   315
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   3075
         VariousPropertyBits=   746604575
         DisplayStyle    =   3
         Size            =   "5424;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unit Name:"
         Height          =   195
         Left            =   7320
         TabIndex        =   183
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   240
         TabIndex        =   149
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Left            =   7230
         TabIndex        =   148
         Top             =   1140
         Width           =   660
      End
      Begin MSForms.TextBox txtClient 
         Height          =   315
         Left            =   1320
         TabIndex        =   147
         Top             =   1140
         Width           =   3075
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5424;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   315
         Left            =   8400
         TabIndex        =   146
         Top             =   1140
         Width           =   3075
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5424;556"
         SpecialEffect   =   6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lease ID:"
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tenant: "
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unit Number:"
         Height          =   195
         Left            =   7230
         TabIndex        =   32
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00959595&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   555
      Left            =   80
      TabIndex        =   7
      Top             =   6360
      Width           =   12135
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Lease"
         Height          =   375
         Left            =   5451
         TabIndex        =   104
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel Changes"
         Height          =   375
         Left            =   8525
         TabIndex        =   103
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel New Lease"
         Height          =   375
         Left            =   3794
         TabIndex        =   102
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Terminate Lease"
         Height          =   375
         Left            =   10065
         TabIndex        =   101
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Lease"
         Height          =   375
         Left            =   720
         TabIndex        =   100
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "&Save Changes"
         Height          =   375
         Left            =   6988
         TabIndex        =   99
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New Lease"
         Height          =   375
         Left            =   2257
         TabIndex        =   98
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin TabDlg.SSTab tabLease 
      Height          =   4395
      Left            =   80
      TabIndex        =   24
      Top             =   1890
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7752
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   1
      TabsPerRow      =   11
      TabHeight       =   520
      TabCaption(0)   =   "&Lease Details"
      TabPicture(0)   =   "Lease2.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Rent Charges"
      TabPicture(1)   =   "Lease2.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label36"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboRentPayable"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Rent Re&view"
      TabPicture(2)   =   "Lease2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSerial"
      Tab(2).Control(1)=   "txtRentReviewDate"
      Tab(2).Control(2)=   "txtRentIncreaseDate"
      Tab(2).Control(3)=   "txtRentIncreaseAmount"
      Tab(2).Control(4)=   "cmdEditRentAnalysis"
      Tab(2).Control(5)=   "cmdSageRentAnalysis"
      Tab(2).Control(6)=   "cmdCancelRentAnalysis"
      Tab(2).Control(7)=   "cmdNewRentAnalysis"
      Tab(2).Control(8)=   "txtRentIncAmt"
      Tab(2).Control(9)=   "txtRentIncDt"
      Tab(2).Control(10)=   "txtRentReviewDt"
      Tab(2).Control(11)=   "flxRentAnalysis"
      Tab(2).Control(12)=   "Label37"
      Tab(2).Control(13)=   "Label59"
      Tab(2).Control(14)=   "Label58"
      Tab(2).Control(15)=   "Label52"
      Tab(2).Control(16)=   "Label23"
      Tab(2).Control(17)=   "Label25"
      Tab(2).Control(18)=   "Label24"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Brea&ks"
      TabPicture(3)   =   "Lease2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(1)=   "cboBreakClause"
      Tab(3).Control(2)=   "Frame1(0)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Service &Charges"
      TabPicture(4)   =   "Lease2.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(1)=   "Frame1(2)"
      Tab(4).Control(2)=   "cboSCPayable"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Interest Charges"
      TabPicture(5)   =   "Lease2.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label49"
      Tab(5).Control(1)=   "Label9"
      Tab(5).Control(2)=   "Frame3"
      Tab(5).Control(3)=   "cboIntCrgable"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "&Breaches"
      TabPicture(6)   =   "Lease2.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraBreaches"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Assignment"
      TabPicture(7)   =   "Lease2.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label54"
      Tab(7).Control(1)=   "Label53"
      Tab(7).Control(2)=   "Label19"
      Tab(7).Control(3)=   "gridAssignment"
      Tab(7).Control(4)=   "txtAssignment_Date"
      Tab(7).Control(5)=   "txtAssignmentID"
      Tab(7).Control(6)=   "cmdAssignmentNew"
      Tab(7).Control(7)=   "cmdAssignmentCancel"
      Tab(7).Control(8)=   "cmdAssignmentSave"
      Tab(7).Control(9)=   "cmdAssignmentEdit"
      Tab(7).Control(10)=   "txtAssignee"
      Tab(7).Control(11)=   "txtDescription"
      Tab(7).ControlCount=   12
      TabCaption(8)   =   "I&nsurance"
      TabPicture(8)   =   "Lease2.frx":09AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cboInsurancePayable"
      Tab(8).Control(1)=   "fmeInsurance"
      Tab(8).Control(2)=   "Label65"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "&Supplementary"
      TabPicture(9)   =   "Lease2.frx":09C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame6"
      Tab(9).Control(1)=   "Frame9"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "&Memo"
      TabPicture(10)  =   "Lease2.frx":09E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1(5)"
      Tab(10).ControlCount=   1
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -69360
         ScrollBars      =   2  'Vertical
         TabIndex        =   212
         Top             =   720
         Width           =   4275
      End
      Begin VB.TextBox txtAssignee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72240
         ScrollBars      =   2  'Vertical
         TabIndex        =   211
         Top             =   720
         Width           =   2715
      End
      Begin VB.Frame fraBreaches 
         Height          =   3735
         Left            =   -74520
         TabIndex        =   189
         Top             =   480
         Width           =   11175
         Begin VB.TextBox txtBreachID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9600
            TabIndex        =   210
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdBreachEdit 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   9720
            TabIndex        =   201
            Top             =   1665
            Width           =   1300
         End
         Begin VB.CommandButton cmdBreachSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   9720
            TabIndex        =   198
            Top             =   2415
            Width           =   1300
         End
         Begin VB.CommandButton cmdBreachCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   9720
            TabIndex        =   200
            Top             =   3180
            Width           =   1300
         End
         Begin VB.CommandButton cmdBreachNew 
            Caption         =   "&New"
            Height          =   375
            Left            =   9720
            TabIndex        =   199
            Top             =   900
            Width           =   1300
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   315
            Left            =   9120
            TabIndex        =   197
            Top             =   660
            Width           =   375
         End
         Begin VB.TextBox txtDateReceived 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6000
            TabIndex        =   195
            Top             =   660
            Width           =   1275
         End
         Begin VB.CheckBox chkResolved 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5340
            TabIndex        =   194
            Top             =   720
            Width           =   665
         End
         Begin VB.TextBox txtInitiatedBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3300
            TabIndex        =   193
            Top             =   660
            Width           =   1935
         End
         Begin VB.TextBox txtCommenceDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2160
            ScrollBars      =   2  'Vertical
            TabIndex        =   192
            Top             =   660
            Width           =   1155
         End
         Begin VB.TextBox txtReceivedBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7260
            TabIndex        =   196
            Top             =   660
            Width           =   1815
         End
         Begin VB.CommandButton cmdSetBreachType 
            Caption         =   "..."
            Height          =   315
            Left            =   1920
            TabIndex        =   191
            Top             =   660
            Width           =   255
         End
         Begin MSDataListLib.DataCombo cboBreachType 
            Bindings        =   "Lease2.frx":09FE
            DataSource      =   "adoBreaches"
            Height          =   315
            Left            =   240
            TabIndex        =   190
            Top             =   660
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBreach 
            Height          =   2595
            Left            =   240
            TabIndex        =   202
            Top             =   960
            Width           =   9315
            _ExtentX        =   16431
            _ExtentY        =   4577
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   12632256
            BackColorSel    =   4210752
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            BandDisplay     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label47 
            Caption         =   "Memo:"
            Height          =   255
            Left            =   9060
            TabIndex        =   209
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label46 
            Caption         =   "Date Received:"
            Height          =   255
            Left            =   6000
            TabIndex        =   208
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label43 
            Caption         =   "Resolved"
            Height          =   195
            Left            =   5160
            TabIndex        =   207
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label42 
            Caption         =   "Initiated By:"
            Height          =   255
            Left            =   3300
            TabIndex        =   206
            Top             =   420
            Width           =   1515
         End
         Begin VB.Label Label40 
            Caption         =   "Received By:"
            Height          =   255
            Left            =   7260
            TabIndex        =   205
            Top             =   420
            Width           =   1515
         End
         Begin VB.Label Label45 
            Caption         =   "Commence Date:"
            Height          =   435
            Left            =   2160
            TabIndex        =   204
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label44 
            Caption         =   "Breach Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   203
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Index           =   0
         Left            =   -71520
         TabIndex        =   184
         Top             =   1440
         Width           =   5055
         Begin VB.ComboBox cboBreak 
            Height          =   315
            Left            =   1740
            TabIndex        =   186
            Top             =   1380
            Width           =   2000
         End
         Begin VB.TextBox txtBreakDate 
            Height          =   315
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   185
            Top             =   720
            Width           =   1960
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Break Date:"
            Height          =   195
            Left            =   600
            TabIndex        =   188
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Break Type:"
            Height          =   195
            Left            =   600
            TabIndex        =   187
            Top             =   1440
            Width           =   870
         End
      End
      Begin VB.CommandButton cmdAssignmentEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -64560
         TabIndex        =   174
         Top             =   1880
         Width           =   1300
      End
      Begin VB.CommandButton cmdAssignmentSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   -64560
         TabIndex        =   173
         Top             =   2680
         Width           =   1300
      End
      Begin VB.CommandButton cmdAssignmentCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -64560
         TabIndex        =   172
         Top             =   3480
         Width           =   1300
      End
      Begin VB.CommandButton cmdAssignmentNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   -64560
         TabIndex        =   171
         Top             =   1080
         Width           =   1300
      End
      Begin VB.TextBox txtAssignmentID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   170
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAssignment_Date 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73950
         ScrollBars      =   2  'Vertical
         TabIndex        =   169
         Top             =   720
         Width           =   1635
      End
      Begin VB.ComboBox cboIntCrgable 
         Height          =   315
         Left            =   -71640
         TabIndex        =   155
         Text            =   "No"
         Top             =   480
         Width           =   915
      End
      Begin VB.Frame Frame9 
         Caption         =   "Date Flag"
         Height          =   1515
         Left            =   -74700
         TabIndex        =   141
         Top             =   360
         Width           =   11355
         Begin VB.TextBox txtDtFlgDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Height          =   285
            Left            =   2340
            MaxLength       =   10
            TabIndex        =   143
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtDtFlgDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000006&
            Height          =   645
            Left            =   2340
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   142
            Top             =   780
            Width           =   8655
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Flag Date:"
            Height          =   195
            Left            =   780
            TabIndex        =   145
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Left            =   780
            TabIndex        =   144
            Top             =   840
            Width           =   840
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Define Supplementary Fields"
         Height          =   2415
         Left            =   -74700
         TabIndex        =   129
         Top             =   1860
         Width           =   11415
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   132
            Top             =   1200
            Width           =   6375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   131
            Top             =   780
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   130
            Top             =   1620
            Width           =   6375
         End
         Begin VB.Label lblSupplementary3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplementary 3:"
            Height          =   195
            Left            =   780
            TabIndex        =   140
            Top             =   1680
            Width           =   1770
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Double Click to define the caption for supplementary fields"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   139
            Top             =   240
            Width           =   7455
         End
         Begin MSForms.TextBox txtSuppCaption3 
            Height          =   315
            Left            =   8880
            TabIndex        =   138
            Top             =   1620
            Visible         =   0   'False
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   -2147483633
            Size            =   "3201;556"
            SpecialEffect   =   3
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSuppCaption2 
            Height          =   315
            Left            =   8880
            TabIndex        =   137
            Top             =   1200
            Visible         =   0   'False
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   -2147483633
            Size            =   "3201;556"
            SpecialEffect   =   3
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSuppCaption1 
            Height          =   315
            Left            =   8880
            TabIndex        =   136
            Top             =   780
            Visible         =   0   'False
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   -2147483633
            Size            =   "3201;556"
            SpecialEffect   =   3
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblSupplementary2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplementary 2:"
            Height          =   195
            Left            =   780
            TabIndex        =   135
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label lblSupplementary1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplementary 1:"
            Height          =   195
            Left            =   780
            TabIndex        =   134
            Top             =   840
            Width           =   1770
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Height          =   1515
            Left            =   360
            TabIndex        =   133
            Top             =   600
            Width           =   10635
         End
      End
      Begin VB.ComboBox cboInsurancePayable 
         Height          =   315
         Left            =   -72720
         TabIndex        =   127
         Text            =   "No"
         Top             =   480
         Width           =   840
      End
      Begin VB.Frame fmeInsurance 
         Height          =   3255
         Left            =   -74580
         TabIndex        =   107
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txtInsuranceEndDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   121
            Top             =   1230
            Width           =   2415
         End
         Begin VB.Frame Frame10 
            Caption         =   "Charging Methods"
            Height          =   2115
            Left            =   6120
            TabIndex        =   112
            Top             =   420
            Width           =   4695
            Begin VB.TextBox txtAnnualInsuranceCharge 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   118
               Top             =   720
               Width           =   1500
            End
            Begin VB.TextBox txtInsurancePercentage 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   117
               Top             =   360
               Width           =   1500
            End
            Begin VB.TextBox txtInsuranceEachPeriod 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E6EDFB&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   1
               EndProperty
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
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   116
               Top             =   1620
               Width           =   1500
            End
            Begin VB.TextBox txtTotalYearlyInsurance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E6EDFB&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   115
               Text            =   "0.00"
               Top             =   1260
               Width           =   1500
            End
            Begin VB.OptionButton optAnnualInsuranceCharge 
               Caption         =   "Annual Insurance Charge Amount"
               Height          =   255
               Left            =   180
               TabIndex        =   114
               Top             =   720
               Width           =   2775
            End
            Begin VB.OptionButton optInsurancePercentage 
               Caption         =   "Insurance Charge Percentage (%)"
               Height          =   255
               Left            =   180
               TabIndex        =   113
               Top             =   360
               Width           =   2715
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Total Insurance Charge for Year:"
               Height          =   195
               Left            =   180
               TabIndex        =   120
               Top             =   1260
               Width           =   2310
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "Insurance Charge Due Each Period:"
               Height          =   195
               Left            =   180
               TabIndex        =   119
               Top             =   1620
               Width           =   2565
            End
         End
         Begin VB.ComboBox cboInsuranceDemandType 
            Height          =   315
            Left            =   2040
            TabIndex        =   111
            Text            =   "1"
            Top             =   2115
            Width           =   2415
         End
         Begin VB.ComboBox cboInsuranceFrequency 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2040
            TabIndex        =   110
            Text            =   "cboInsuranceFrequency"
            Top             =   1665
            Width           =   2415
         End
         Begin VB.TextBox txtInsuranceStartDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   109
            Top             =   780
            Width           =   2415
         End
         Begin VB.TextBox txtInsuranceNextDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F9EE&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   108
            Top             =   2550
            Width           =   1320
         End
         Begin MSForms.ComboBox cboInsuranceDept 
            Height          =   315
            Left            =   2040
            TabIndex        =   167
            Top             =   360
            Width           =   2415
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;35277"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   153
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance End Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   126
            Top             =   1275
            Width           =   1470
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   240
            TabIndex        =   125
            Top             =   2130
            Width           =   1050
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance Start Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   124
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance Frequency:"
            Height          =   195
            Left            =   240
            TabIndex        =   123
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   122
            Top             =   2565
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Notes"
         Height          =   3795
         Index           =   5
         Left            =   -74640
         TabIndex        =   105
         Top             =   360
         Width           =   11595
         Begin VB.TextBox txtMemo 
            Height          =   3315
            Left            =   240
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   106
            Top             =   360
            Width           =   11055
         End
      End
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   80
         Top             =   780
         Width           =   1500
      End
      Begin VB.ComboBox cboBreakClause 
         Height          =   315
         Left            =   -70365
         TabIndex        =   87
         Text            =   "No"
         Top             =   660
         Width           =   780
      End
      Begin VB.TextBox txtRentReviewDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72907
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   81
         Top             =   780
         Width           =   2500
      End
      Begin VB.TextBox txtRentIncreaseDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69934
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   82
         Top             =   780
         Width           =   2500
      End
      Begin VB.TextBox txtRentIncreaseAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
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
         Left            =   -66960
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   780
         Width           =   2500
      End
      Begin VB.CommandButton cmdEditRentAnalysis 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -64320
         TabIndex        =   76
         Top             =   1800
         Width           =   1300
      End
      Begin VB.CommandButton cmdSageRentAnalysis 
         Caption         =   "&Save"
         Height          =   375
         Left            =   -64320
         TabIndex        =   85
         Top             =   2520
         Width           =   1300
      End
      Begin VB.CommandButton cmdCancelRentAnalysis 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -64320
         TabIndex        =   77
         Top             =   3240
         Width           =   1300
      End
      Begin VB.CommandButton cmdNewRentAnalysis 
         Caption         =   "&New"
         Height          =   375
         Left            =   -64320
         TabIndex        =   75
         Top             =   1080
         Width           =   1300
      End
      Begin VB.TextBox txtRentIncAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -65880
         TabIndex        =   74
         Top             =   3780
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtRentIncDt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69840
         MaxLength       =   10
         TabIndex        =   73
         Top             =   3780
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtRentReviewDt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   72
         Top             =   3780
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cboSCPayable 
         Height          =   315
         Left            =   -72360
         TabIndex        =   70
         Text            =   "No"
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   3555
         Index           =   2
         Left            =   -74160
         TabIndex        =   46
         Top             =   720
         Width           =   10455
         Begin VB.TextBox txt10a 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E6EDFB&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
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
            Left            =   7890
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            Caption         =   "Charging Methods"
            Height          =   2775
            Left            =   5280
            TabIndex        =   94
            Top             =   240
            Width           =   4815
            Begin VB.TextBox txtGlobalAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               TabIndex        =   62
               Top             =   1440
               Width           =   1500
            End
            Begin VB.TextBox txtAnnualService 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox txtPPSqFoot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   720
               Width           =   1500
            End
            Begin VB.TextBox txtSCPercentage 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   360
               Width           =   1500
            End
            Begin VB.TextBox txtFinalAmout 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E6EDFB&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   1
               EndProperty
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
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   2400
               Width           =   1500
            End
            Begin VB.TextBox txtAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E6EDFB&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   63
               Text            =   "0.00"
               Top             =   2040
               Width           =   1500
            End
            Begin VB.OptionButton optGlobalData 
               Caption         =   "Global Data"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   1440
               Width           =   1215
            End
            Begin VB.OptionButton optFixedTotal 
               Caption         =   "Annual Service Charge Amount"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   1080
               Width           =   2535
            End
            Begin VB.OptionButton optSqFoot 
               Caption         =   "Price Per Sq. Foot/Metre"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   720
               Width           =   2175
            End
            Begin VB.OptionButton optPercentage 
               Caption         =   "Service Charge Percentage (%)"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Service Charge Total for Year:"
               Height          =   195
               Left            =   240
               TabIndex        =   96
               Top             =   2040
               Width           =   2145
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Service Charge Due Each Period:"
               Height          =   195
               Left            =   240
               TabIndex        =   95
               Top             =   2400
               Width           =   2400
            End
         End
         Begin VB.TextBox txt10c 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E6EDFB&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
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
            Left            =   2850
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   3132
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtSCNextDueDt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1470
            TabIndex        =   50
            Top             =   1701
            Width           =   2715
         End
         Begin VB.ComboBox cboFreqSC 
            Height          =   315
            ItemData        =   "Lease2.frx":0A18
            Left            =   1470
            List            =   "Lease2.frx":0A1A
            TabIndex        =   49
            Top             =   1204
            Width           =   2715
         End
         Begin VB.TextBox txtTOLimit 
            Height          =   285
            Left            =   2850
            TabIndex        =   51
            Top             =   2665
            Width           =   1335
         End
         Begin VB.TextBox txtPayableFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   48
            Top             =   737
            Width           =   2715
         End
         Begin VB.ComboBox cboSCDemandType 
            Height          =   315
            Left            =   1470
            TabIndex        =   15
            Text            =   "cboSCDemandType"
            Top             =   2168
            Width           =   2715
         End
         Begin MSForms.ComboBox cboServiceChargeDept 
            Height          =   315
            Left            =   1470
            TabIndex        =   47
            Top             =   240
            Width           =   2715
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4789;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;35277"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "xService Charge Total for Year:"
            Height          =   195
            Left            =   5085
            TabIndex        =   97
            Top             =   3120
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "xService Charge Due Each Period:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   3120
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Width           =   1110
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Payable From:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "T/O Limit:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   2640
            Width           =   705
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   2160
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -73440
         TabIndex        =   45
         Top             =   840
         Width           =   9225
         Begin VB.TextBox txtAmtCrgIntOn 
            Height          =   315
            Left            =   2400
            TabIndex        =   158
            Top             =   1920
            Width           =   2475
         End
         Begin VB.TextBox txtAdditionalIntRate 
            Height          =   315
            Left            =   2400
            TabIndex        =   157
            Top             =   1320
            Width           =   2475
         End
         Begin VB.TextBox txtInt2bChrg 
            Height          =   315
            Left            =   6960
            TabIndex        =   159
            Top             =   720
            Width           =   1800
         End
         Begin VB.TextBox txtIntPayableAfterDays 
            Height          =   315
            Left            =   6960
            TabIndex        =   160
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox cboIntDemandType 
            Height          =   315
            Left            =   6960
            TabIndex        =   161
            Text            =   "3"
            Top             =   1920
            Width           =   1800
         End
         Begin MSForms.ComboBox cboIntChargeDept 
            Height          =   315
            Left            =   2400
            TabIndex        =   156
            Top             =   720
            Width           =   2475
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4366;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;35277"
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount to charge Interest on:"
            Height          =   195
            Left            =   240
            TabIndex        =   166
            Top             =   1995
            Width           =   2100
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Additional Interest Rate:"
            Height          =   195
            Left            =   240
            TabIndex        =   165
            Top             =   1350
            Width           =   1695
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Interest to be charged:"
            Height          =   195
            Left            =   5160
            TabIndex        =   164
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Interest Payable After"
            Height          =   195
            Left            =   5160
            TabIndex        =   163
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   5160
            TabIndex        =   162
            Top             =   1935
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            Height          =   195
            Left            =   8160
            TabIndex        =   154
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   152
            Top             =   720
            Width           =   405
         End
      End
      Begin VB.ComboBox cboRentPayable 
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Text            =   "No"
         Top             =   720
         Width           =   840
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   -73267
         TabIndex        =   36
         Top             =   690
         Width           =   8835
         Begin VB.CheckBox chkSubLease 
            Caption         =   "Yes"
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdLeaseType 
            Caption         =   "..."
            Height          =   300
            Left            =   3840
            TabIndex        =   11
            Top             =   2040
            Width           =   405
         End
         Begin VB.ComboBox cboHeadLease 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtLeaseEndDate 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   14
            Top             =   2040
            Width           =   1995
         End
         Begin VB.TextBox txtYearEnd 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   12
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtLeaseStDt 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1260
            Width           =   1995
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "Lease2.frx":0A1C
            Left            =   1800
            List            =   "Lease2.frx":0A1E
            TabIndex        =   10
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Lease:"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1260
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Head Lease:"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Type"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   2040
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease End Date:"
            Height          =   195
            Left            =   4680
            TabIndex        =   39
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Year End:"
            Height          =   195
            Left            =   4680
            TabIndex        =   38
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Start Date:"
            Height          =   195
            Left            =   4680
            TabIndex        =   37
            Top             =   1260
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Top             =   1080
         Width           =   9135
         Begin VB.ComboBox cboBRDemandType 
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   2280
            Width           =   2700
         End
         Begin VB.ComboBox cboFreqBR 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            TabIndex        =   18
            Top             =   990
            Width           =   2700
         End
         Begin VB.TextBox txtRentStartDate 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   1680
            Width           =   2700
         End
         Begin VB.TextBox txtTotalRentYear 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   21
            Top             =   300
            Width           =   1500
         End
         Begin VB.TextBox txtNextDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F9EE&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   23
            Top             =   1680
            Width           =   1500
         End
         Begin VB.TextBox txtRentDueEachPeriod 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F9EE&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   22
            Top             =   990
            Width           =   1500
         End
         Begin MSForms.ComboBox cboRentChargeDept 
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            Top             =   360
            Width           =   2700
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4762;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;35277"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   150
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   2280
            Width           =   1050
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent Start Date:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   990
            Width           =   795
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total Rent for Year:"
            Height          =   195
            Index           =   4
            Left            =   4800
            TabIndex        =   28
            Top             =   300
            Width           =   1425
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due Date:"
            Height          =   195
            Index           =   6
            Left            =   4800
            TabIndex        =   27
            Top             =   1680
            Width           =   1110
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent Due Each Period:"
            Height          =   195
            Index           =   5
            Left            =   4800
            TabIndex        =   26
            Top             =   990
            Width           =   1650
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentAnalysis 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   86
         Top             =   1080
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAssignment 
         Height          =   2835
         Left            =   -73950
         TabIndex        =   175
         Top             =   1020
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   4210752
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label19 
         Caption         =   "Description:"
         Height          =   255
         Left            =   -69360
         TabIndex        =   213
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label53 
         Caption         =   "Assignee:"
         Height          =   255
         Left            =   -72240
         TabIndex        =   177
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label54 
         Caption         =   "Assignment Date:"
         Height          =   255
         Left            =   -73950
         TabIndex        =   176
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Chargeable:"
         Height          =   195
         Left            =   -73440
         TabIndex        =   168
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Payable:"
         Height          =   195
         Left            =   -74580
         TabIndex        =   128
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Review No.:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   93
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Review Date:"
         Height          =   195
         Left            =   -72907
         TabIndex        =   91
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Increase Date:"
         Height          =   195
         Left            =   -69934
         TabIndex        =   90
         Top             =   540
         Width           =   1440
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Increase Amount:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -66960
         TabIndex        =   89
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Break Clause:"
         Height          =   195
         Left            =   -71505
         TabIndex        =   88
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Increase Amount:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -67680
         TabIndex        =   84
         Top             =   3780
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Increase Date:"
         Height          =   195
         Left            =   -71400
         TabIndex        =   79
         Top             =   3780
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Review Date:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   78
         Top             =   3780
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Service Charge Payable:"
         Height          =   195
         Left            =   -74160
         TabIndex        =   71
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rent Payable:"
         Height          =   195
         Left            =   1920
         TabIndex        =   44
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label49 
         Height          =   255
         Left            =   -72900
         TabIndex        =   35
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.PictureBox picLeaseList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   -720
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   178
      Top             =   840
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdGridUnitLookup 
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
         Height          =   2175
         Left            =   45
         TabIndex        =   180
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C000C0&
      Caption         =   "Label17"
      Height          =   15
      Left            =   0
      TabIndex        =   182
      Top             =   1800
      Width           =   12375
   End
   Begin VB.Label Label41 
      Caption         =   "Description:"
      Height          =   255
      Left            =   4950
      TabIndex        =   34
      Top             =   2880
      Width           =   1515
   End
End
Attribute VB_Name = "frmLease2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BREACH_NEW_ENTRY_ As Boolean
Dim ASSIGNMENT_NEW_ENTRY_ As Boolean

Dim szaDemandtype() As String

Dim Conn1 As New RDO.rdoConnection
Dim Env1 As rdoEnvironment
Dim Envs1 As rdoEnvironments
Dim Rst1 As rdoResultset
Dim Conn2 As New RDO.rdoConnection
Dim Env2 As rdoEnvironment
Dim Envs2 As rdoEnvironments
Dim Rst2 As rdoResultset
Dim SQLStr1 As String
Dim SQLStr2 As String

Public FormLoad As Boolean
    
Private Sub cboBreak_LostFocus()
Dim i, j, match As Integer

If cboBreak.text <> "" Then
    match = 0
    j = cboBreak.ListCount - 1
    For i = 0 To j
        If cboBreak.List(i) = cboBreak.text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Break Type is invalid.", vbOKOnly + vbCritical, "Invalid Break Type"
        cboBreak.text = ""
        Exit Sub
    End If
End If

End Sub

Private Sub cboBreakClause_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboFreqBR_LostFocus()
Dim i, j, match As Integer

match = 0
j = cboFreqBR.ListCount - 1
For i = 0 To j
    If cboFreqBR.List(i) = cboFreqBR.text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    MsgBox "Rent frequency is invalid.", vbOKOnly + vbCritical, "Invalid Frequency"
    Exit Sub
End If

If txtRentStartDate.text <> "" Then Call CalculateBR
If cboFreqBR.ListIndex < 6 Then txtNextDueDate.Enabled = True
End Sub

Private Sub cboFreqSC_GotFocus()
   If txtPayableFrom.text = "" Then
      MsgBox "You must enter the Payable from date before enter frequency.", vbInformation + vbOKOnly, "Payable from date missing"
      txtPayableFrom.SetFocus
   End If
End Sub

Private Sub cboFreqSC_LostFocus()
If cboFreqSC.text = "" Then Exit Sub

Dim i, j, match As Integer

match = 0
j = cboFreqSC.ListCount - 1
For i = 0 To j
    If cboFreqSC.List(i) = cboFreqSC.text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    MsgBox "Service Charge frequency is invalid.", vbOKOnly + vbCritical, "Invalid Frequency"
    Exit Sub
End If

If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC
End Sub

Private Sub cboInsuranceFrequency_GotFocus()
   If txtInsuranceStartDate.text = "" Then
      MsgBox "You must enter the Insurance Start Date before enter frequency.", vbInformation + vbOKOnly, "Insurance start date missing"
      txtInsuranceStartDate.SetFocus
   End If
End Sub

Private Sub cboInsuranceFrequency_LostFocus()
    If cboInsuranceFrequency.text = "" Then Exit Sub
    
    Dim i, j, match As Integer
    
    match = 0
    j = cboInsuranceFrequency.ListCount - 1
    For i = 0 To j
        If cboInsuranceFrequency.List(i) = cboInsuranceFrequency.text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Insurance Charge frequency is invalid.", vbOKOnly + vbCritical, "Invalid Frequency"
        Exit Sub
    End If
    
    If cboInsurancePayable.text = "Yes" Then Call CalculateInsuranceCharge
End Sub

Private Sub cboInsurancePayable_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboIntCrgable_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboRentPayable_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboSCPayable_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboTenant_Click()
   txtTenant.text = cboTenant.text
End Sub

Private Sub cboTenant_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboTenant.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboUnit_Click()
   If txtTenant.text = "" Then
      MsgBox "Please select a tenant.", vbCritical + vbOKOnly, "Lease"
      cboUnit.text = ""
      cmdtenants_Click
      Exit Sub
   End If
   
   Dim szaUnits() As String

   szaUnits = Split(cboUnit.text, " - ")
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT ClientName, PropertyName " & _
             "FROM Client, Property, Units " & _
             "WHERE Client.ClientID = Property.ClientID And " & _
                 "Property.PropertyID = Units.PropertyID And " & _
                 "Units.UnitNumber = '" & szaUnits(0) & "';"

   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   txtClient.text = Rst1!ClientName
   txtProperty.text = Rst1!PropertyName
   txtUnitName.text = szaUnits(1)
   
   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
'***************************************************************************
'CREATE THE LEASE ID AUTOMATICALLY
'***************************************************************************
   If cmdAddNew.Visible Then Exit Sub

   Dim szaTenant() As String, szaUnit() As String

   If cboTenant.text <> "" Then szaTenant = Split(cboTenant.text, " / ")
   szaUnit = Split(cboUnit.text, " - ")

   txtLeaseID.text = OnlyNumericString(szaTenant(0)) & OnlyNumericString(szaUnit(0) & Format(Now, "d-m-yy"))
End Sub

Private Sub chkSubLease_Click()

If chkSubLease.Value = 1 Then
    cboHeadLease.Enabled = True
Else
    cboHeadLease.Enabled = False
End If
End Sub

Private Sub cmdAddNew_Click()
   If MsgBox("Do you want to add new lease?", vbQuestion + vbYesNo, "Lease - New") = vbNo Then Exit Sub

   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst As rdoResultset

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   Set rdoRst = rdoConn.OpenResultset("SELECT * FROM GlobalData", rdOpenStatic, rdConcurReadOnly)

   If rdoRst.RowCount = 0 Then
      MsgBox "You Need to Enter the Global Data before you can add a lease record.", vbOKOnly + vbInformation, "Global Data"
      rdoRst.Close
      rdoConn.Close
      Set rdoConn = Nothing
      Exit Sub
   Else
      rdoRst.Close
      rdoConn.Close
      Set rdoConn = Nothing
   End If

   Call EmptyBoxes
   Call GetTenantsWithoutLease
   Call GetUnitWithoutLease
   Call EnableBoxes

   Call FillcboType(szaDemandtype)

   cmdAddNew.Visible = False
   cmdDelete.Visible = False
   cmdEdit.Visible = False
   cmdCancelNew.Visible = True
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
End Sub

Private Sub GetUnitWithoutLease()
   Dim temp As String, i As Integer

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT UnitNumber, UnitName " & _
             "FROM Units " & _
             "WHERE Occupied = 'N';"

   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   i = 0
   If Rst1.EOF = False Then
       While Rst1.EOF = False
           cboUnit.AddItem Rst1!UnitNumber & " - " & Rst1!UNITNAME, i
           i = i + 1
           Rst1.MoveNext
       Wend
   End If
   Rst1.Close
   Conn1.Close
End Sub

Private Sub cmdAssignmentCancel_Click()
   AssignmentButtonMode DefaultMode
End Sub

Private Sub cmdAssignmentEdit_Click()
ASSIGNMENT_NEW_ENTRY_ = False
AssignmentButtonMode EditMode

End Sub

Private Sub cmdAssignmentNew_Click()
ASSIGNMENT_NEW_ENTRY_ = True
AssignmentButtonMode NewEntryMode
End Sub

Private Sub cmdAssignmentSave_Click()
If SaveAssignment Then
    MsgBox "The assignment information saved successfully", vbInformation, "Save Assignment"
Else
    MsgBox "Could not save assignment information", vbInformation, "Save Assignment"
End If
AssignmentButtonMode DefaultMode
End Sub

Private Sub cmdBreachCancel_Click()
BreachButtonMode DefaultMode
End Sub

Private Sub cmdBreachEdit_Click()
BREACH_NEW_ENTRY_ = False
BreachButtonMode EditMode

End Sub

Private Sub cmdBreachNew_Click()
   BREACH_NEW_ENTRY_ = True
   BreachButtonMode NewEntryMode
End Sub

Private Sub cmdBreachSave_Click()

If SaveBreaches Then
    MsgBox "The breach information saved successfully", vbInformation, "Save Breaches"
Else
    MsgBox "Could not save breach information", vbInformation, "Save Breaches"
End If
BreachButtonMode DefaultMode

End Sub

Private Sub cmdCancelEdit_Click()
   Call EmptyBoxes
   Call GetTenantsWithLease
   Call DisableBoxes
End Sub

Private Sub cmdCancelNew_Click()
   If MsgBox("Do you want to cancel the adding new lease?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      Call EmptyBoxes
      Call GetTenantsWithLease
      Call DisableBoxes
   End If
End Sub

Private Sub cmdCancelRentAnalysis_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   UnlockTextBoxes False
   RentReviewButtonMode DefaultMode
End Sub

Private Sub cmdDelete_Click()
   If txtLeaseID.text = "" Then
       MsgBox "You must select a lease to terminate", vbOKOnly + vbCritical, "Lease"
       cmdLease.SetFocus
       Exit Sub
   Else
      If MsgBox("Are you sure you want to terminate the lease for tenant: " & txtTenant.text, vbYesNo + vbQuestion, "Delete Lease") = vbYes Then
         Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
         Conn1.CursorDriver = rdUseOdbc
         Conn1.EstablishConnection rdDriverNoPrompt

         SQLStr1 = "UPDATE LeaseDetails " & _
                   "SET Status = False " & _
                   "WHERE LeaseID = '" & txtLeaseID.text & "'"

         Conn1.Execute SQLStr1
         
         SQLStr1 = "UPDATE Units " & _
                   "SET Occupied = 'N' " & _
                   "WHERE UnitNumber = '" & cboUnit.text & "'"

         Conn1.Execute SQLStr1
         Conn1.Close
         Set Conn1 = Nothing

         Call EmptyBoxes
         Call GetTenantsWithLease

         MsgBox "The lease have been terninated successfully.", vbInformation + vbOKOnly, "Lease Terminate"
       End If
   End If
End Sub

Private Sub cmdEdit_Click()
   If txtLeaseID.text = "" Then
       MsgBox "You must select a Lease to edit.", vbOKOnly + vbCritical, "No Lease Selected"
       cmdLease.SetFocus
       Exit Sub
   End If

   If MsgBox("Do you want to edit the lease?", vbQuestion + vbYesNo, "Lease - Edit") = vbNo Then Exit Sub

   Dim szaTenant() As String
   Dim szText As String, iCboIndex As Integer

   Call EnableBoxes
   '******************
   'EnableBoxes method unlock the cboUnit combo,
   'but user will not able to change the unit number once the lease created
   cboUnit.Locked = True
   cmdTenants.Enabled = False
   '********************
'   szText = cboBRDemandType.text
'   iCboIndex = cboBRDemandType.ListIndex
'   Call FillcboType(szaDemandtype)
'   cboBRDemandType.ListIndex = iCboIndex
'
'   iCboIndex = cboSCDemandType.ListIndex
'   cboSCDemandType.ListIndex = iCboIndex
'
'   iCboIndex = cboIntDemandType.ListIndex
'   cboIntDemandType.ListIndex = iCboIndex

   cboTenant.Enabled = False

   cmdAddNew.Visible = False
   cmdDelete.Visible = False
   cmdEdit.Visible = False
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = True
   cmdSaveEdit.TabIndex = 25
   cmdCancelEdit.Visible = True
   cmdCancelEdit.TabIndex = 26
End Sub

Private Sub cmdEditRentAnalysis_Click()
   If MsgBox("Do you want to Edit data?", vbQuestion + vbYesNo, "Edit Data") = vbNo Then Exit Sub
   flxRentAnalysis.TextMatrix(flxRentAnalysis.Row, 0) = "X"
   UnlockTextBoxes True
   cmdEditRentAnalysis.Enabled = False
   flxRentAnalysis.Enabled = False
   RentReviewButtonMode EditMode
End Sub

Private Sub UnlockTextBoxes(bState As Boolean)
   txtRentReviewDate.Locked = Not bState
   txtRentIncreaseDate.Locked = Not bState
   txtRentIncreaseAmount.Locked = Not bState
   txtSerial.Locked = Not bState
   
   If Not bState Then
      txtRentReviewDate.text = ""
      txtRentIncreaseDate.text = ""
      txtRentIncreaseAmount.text = ""
      txtSerial.text = ""
   End If
End Sub

Private Sub cmdGridUnitLookup_Click()
   picLeaseList.Visible = False
End Sub

Private Sub cmdLease_Click()
   Call PrepareList

   picLeaseList.Top = fraMain.Top + txtLeaseID.Top + txtLeaseID.Height + 5
   picLeaseList.Left = fraMain.Left + txtLeaseID.Left + 5
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxLeaseList
   LoadAllLeaseFlxGrd
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 8
   conFlxGrid.Clear
   szHeader$ = "|<LeaseID|<Tenant ID|<Tenant Name|<Unit Name"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 300        'Solid column
   conFlxGrid.ColWidth(1) = 0          'Lease ID
   conFlxGrid.ColWidth(2) = 900        'Client ID
   conFlxGrid.ColWidth(3) = 2000       'Client Name
   conFlxGrid.ColWidth(4) = 1800       'Unit Name
   conFlxGrid.ColWidth(5) = 0          'Unit Number
   conFlxGrid.ColWidth(6) = 0          'Client Name
   conFlxGrid.ColWidth(7) = 0          'Property
   conFlxGrid.Rows = 2
'
   conFlxGrid.RowHeightMin = 285
End Sub

Private Sub LoadAllLeaseFlxGrd()
   Dim conLease As New RDO.rdoConnection
   Dim rdoLease As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conLease.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conLease.CursorDriver = rdUseIfNeeded
   conLease.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT LeaseID, LeaseDetails.SageAccountNumber, " & _
               "CompanyName, UnitName, LeaseDetails.UnitNumber, " & _
               "ClientName, PropertyName " & _
           "FROM LeaseDetails, Units, Property, Client " & _
           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
               "LeaseDetails.Status = " & IIf(chkExpLease.Value = 0, "True", "False") & " And " & _
               "Units.PropertyId = Property.PropertyID And " & _
               "Property.ClientID = Client.ClientID " & _
           "ORDER BY CompanyName;"

   Set rdoLease = conLease.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rdoLease.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   While Not rdoLease.EOF
      flxLeaseList.TextMatrix(iRow, 1) = rdoLease!LeaseId
      flxLeaseList.TextMatrix(iRow, 2) = rdoLease!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 3) = rdoLease!CompanyName
      flxLeaseList.TextMatrix(iRow, 4) = rdoLease!UNITNAME
      flxLeaseList.TextMatrix(iRow, 5) = rdoLease!UnitNumber
      flxLeaseList.TextMatrix(iRow, 6) = rdoLease!ClientName
      flxLeaseList.TextMatrix(iRow, 7) = rdoLease!PropertyName
      rdoLease.MoveNext
      If Not rdoLease.EOF Then flxLeaseList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rdoLease.Close
   conLease.Close
   Set rdoLease = Nothing
   Set conLease = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
   
   conLease.Close
   Set rdoLease = Nothing
   Set conLease = Nothing
End Sub

Private Sub cmdLease_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And picLeaseList.Visible = True Then
      picLeaseList.Visible = False
   End If
End Sub

Private Sub cmdLeaseType_Click()
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   LoadType "LEASE TYPE", cboType
End Sub

Private Sub cmdNewRentAnalysis_Click()
   If MsgBox("Do you want to Add new data?", vbQuestion + vbYesNo, "Add new Data") = vbNo Then Exit Sub
   UnlockTextBoxes True
   cmdNewRentAnalysis.Enabled = False
   flxRentAnalysis.Enabled = False
   RentReviewButtonMode NewEntryMode
End Sub

Private Sub cmdSageRentAnalysis_Click()
   If cmdEditRentAnalysis.Enabled = True And cmdNewRentAnalysis.Enabled = True Then Exit Sub

   Dim lID As Long

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT * " & _
             "FROM RentAnalysis"

   If cmdEditRentAnalysis.Enabled = False Then
      lID = FindXID
      SQLStr1 = SQLStr1 + " WHERE ID = " & lID & ";"

      Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
      Rst1.Edit
      flxRentAnalysis.TextMatrix(flxRentAnalysis.Row, 0) = "X"
   End If
   If cmdNewRentAnalysis.Enabled = False Then
      SQLStr1 = SQLStr1 + ";"
      
      Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
      Rst1.AddNew
   End If

   Rst1!SageAccountNumber = flxLeaseList.TextMatrix(flxLeaseList.Row, 2)
   Rst1!SerialNumber = txtSerial.text
   Rst1!RentReviewDate = CDate(Format(txtRentReviewDate.text, "dd/mm/yyyy"))
   Rst1!RentIncreaseDate = CDate(Format(txtRentIncreaseDate.text, "dd/mm/yyyy"))
   Rst1!RentIncreaseAmount = CCur(Format(txtRentIncreaseAmount.text, "0.00"))
   Rst1.Update
'
   Rst1.Close
   Conn1.Close
'
   Set Rst1 = Nothing
   Set Conn1 = Nothing
'
   MsgBox "Data has been updated.", vbInformation, "Rent"
'
   cmdEditRentAnalysis.Enabled = True
   cmdNewRentAnalysis.Enabled = True
   flxRentAnalysis.Enabled = True
   UnlockTextBoxes False
   ConfigureFlxGrid flxRentAnalysis
   LoadFlxGrid flxRentAnalysis
   RentReviewButtonMode DefaultMode
End Sub

Private Function FindXID() As Long
   Dim iRow As Integer
   
   For iRow = 1 To flxRentAnalysis.Rows - 1
      If flxRentAnalysis.TextMatrix(iRow, 0) = "X" Then
         FindXID = CLng(flxRentAnalysis.TextMatrix(iRow, 5))
         Exit Function
      End If
   Next iRow
   FindXID = -1
End Function

Private Sub cmdSaveEdit_Click()
   If MsgBox("Do you want to update the lease?", vbQuestion + vbYesNo, "Lease - Update") = vbNo Then Exit Sub
   
   If SaveUpdateLease(False) Then
      MsgBox "The new lease record has been saved", vbOKOnly + vbInformation, "Updated"
'   Else
      Conn1.Close
      Set Conn1 = Nothing
   End If
End Sub

Private Sub cmdSaveNew_Click()
   If MsgBox("Do you want to save the lease?", vbQuestion + vbYesNo, "Lease - Save") = vbNo Then Exit Sub
   
   If SaveUpdateLease(True) Then
      MsgBox "The new lease record has been saved", vbOKOnly + vbInformation, "Saved"
   Else
      Conn1.Close
      Set Conn1 = Nothing
   End If
End Sub

Private Function SaveUpdateLease(bSaveUpdate As Boolean) As Boolean
   Dim i As Integer

   If txtLeaseID.text = "" Then
      MsgBox "You must enter a Lease reference to continue!", vbOKOnly + vbCritical, "Reference Required"
      txtLeaseID.SetFocus
      SaveUpdateLease = False
      Exit Function
   End If
   
   If cboType.text = "" Then
      MsgBox "You must enter the Lease Type to continue!", vbOKOnly + vbCritical, "Lease Type missing"
      tabLease.Tab = 0
      cboType.SetFocus
      SaveUpdateLease = False
      Exit Function
   End If
   
   If txtYearEnd.text = "" Then
      MsgBox "You must enter the Year End date to continue!", vbOKOnly + vbCritical, "Year End date missing"
      tabLease.Tab = 0
      txtYearEnd.SetFocus
      SaveUpdateLease = False
      Exit Function
   End If
   
   If txtLeaseStDt.text = "" Then
      MsgBox "You must enter a Lease Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 0
      txtLeaseStDt.SetFocus
      SaveUpdateLease = False
      Exit Function
   End If
   
   If txtLeaseEndDate.text = "" Then
      MsgBox "You must enter a Lease End Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 0
      txtLeaseEndDate.SetFocus
      SaveUpdateLease = False
      Exit Function
   End If
   
   If cboRentPayable.text = "Yes" Then
      If cboRentChargeDept.text = "" Then
         MsgBox "You must select a Department of rent.", vbOKOnly + vbCritical, "Rent - Department"
         tabLease.Tab = 1
         cboRentChargeDept.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If cboFreqBR.text = "" Then
         MsgBox "You must select a Rent Frequency!", vbOKOnly + vbCritical, "Frequency Required"
         tabLease.Tab = 1
         cboFreqBR.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If txtRentStartDate.text = "" Then
         MsgBox "You must enter a Rent Start Date!", vbOKOnly + vbCritical, "Date Required"
         tabLease.Tab = 1
         txtRentStartDate.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If cboBRDemandType.text = "" Then
         MsgBox "You must choose Demand type from the dropdown menu.", vbCritical + vbOKOnly, "Data Required"
         tabLease.Tab = 1
         cboBRDemandType.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If txtTotalRentYear.text = "" Then
         MsgBox "You must enter a Total rent for the year.!", vbOKOnly + vbCritical, "Date Required"
         tabLease.Tab = 1
         txtTotalRentYear.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
   End If
   
   If cboSCPayable.text = "Yes" Then
      If cboServiceChargeDept.text = "" Then
         MsgBox "You must select a department for the service charge!", vbOKOnly + vbCritical, "Service Charge - Department"
         tabLease.Tab = 4
         cboServiceChargeDept.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
       If txtPayableFrom.text = "" Then
           MsgBox "You must enter a Payable From Date for Service Charge!", vbOKOnly + vbCritical, "Date Required"
           SaveUpdateLease = False
           Exit Function
       End If
       If cboFreqSC.text = "" Then
           MsgBox "You must select a Service Charge Frequency!", vbOKOnly + vbCritical, "Frequency Required"
           SaveUpdateLease = False
           Exit Function
       End If
   End If
   
   If cboIntCrgable.text = "Yes" Then
      If cboIntChargeDept.text = "" Then
         MsgBox "You must select a department for the interest charge!", vbOKOnly + vbCritical, "Interest Charge - Department"
         tabLease.Tab = 5
         cboIntChargeDept.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If txtIntPayableAfterDays.text = "" Then
         MsgBox "You must enter number of days interest will charge after!", vbOKOnly + vbCritical, "Interest Charge"
         tabLease.Tab = 5
         txtIntPayableAfterDays.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If cboIntDemandType.text = "Yes" Then
         MsgBox "You must select interest demand type!", vbOKOnly + vbCritical, "Demand Type"
         tabLease.Tab = 5
         cboIntDemandType.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
   End If
   
   If cboInsurancePayable.text = "Yes" Then
      If cboInsuranceDept.text = "" Then
         MsgBox "You must select department of insurance!", vbOKOnly + vbCritical, "Insurance"
         tabLease.Tab = 9
         cboInsuranceDept.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If txtInsuranceStartDate.text = "" Then
         MsgBox "You must enter insurance start date!", vbOKOnly + vbCritical, "Insurance"
         tabLease.Tab = 9
         txtInsuranceStartDate.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If txtInsuranceEndDate.text = "" Then
         MsgBox "You must enter insurance end date!", vbOKOnly + vbCritical, "Insurance"
         tabLease.Tab = 9
         txtInsuranceEndDate.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If cboInsuranceFrequency.text = "" Then
         MsgBox "You must select insurance frequency!", vbOKOnly + vbCritical, "Insurance"
         tabLease.Tab = 9
         cboInsuranceFrequency.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
      If cboInsuranceDemandType.text = "" Then
         MsgBox "You must select insurance demand type!", vbOKOnly + vbCritical, "Insurance"
         tabLease.Tab = 9
         cboInsuranceDemandType.SetFocus
         SaveUpdateLease = False
         Exit Function
      End If
   End If
   
   If cboUnit.text = "" Then
       MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
       SaveUpdateLease = False
       Exit Function
   End If
   
   Dim szaUnit() As String
   
   szaUnit = Split(cboUnit.text, " - ")
   GetGlobalDataForProperty (szaUnit(0))
   
   'save the details to a new record
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt
   
   SQLStr1 = "SELECT * FROM LeaseDetails " & _
             "WHERE LeaseID = '" & txtLeaseID.text & "' And " & _
                  "Status = True;"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
   
   Dim szaTemp() As String
   
   '****************************************************
   ' Adding data in the LeaseDetails Table
   '****************************************************
   Dim szaTenant() As String
   
   If bSaveUpdate Then
      If Not Rst1.EOF Or Not Rst1.BOF Then
          MsgBox "Cannot save the lease. The lease reference already exist.", vbInformation, "Save Lease"
          txtLeaseID.text = ""
          txtLeaseID.SetFocus
          SaveUpdateLease = False
          Exit Function
      End If
      
      Rst1.AddNew
      Rst1!LeaseId = txtLeaseID.text
      szaTenant = Split(txtTenant.text, " / ")
   Else
      ReDim szaTenant(2) As String
      szaTenant(0) = flxLeaseList.TextMatrix(flxLeaseList.Row, 2)
      szaTenant(1) = flxLeaseList.TextMatrix(flxLeaseList.Row, 3)
      Rst1.Edit
   End If

   Rst1!SageAccountNumber = szaTenant(0)
   Rst1!CompanyName = szaTenant(1)
   
   If chkSubLease.Value Then
       If cboHeadLease.text <> "" Then Rst1!HeadLease = cboHeadLease.text
   End If
   
   Rst1!UnitNumber = szaUnit(0)
   Rst1!TYPEOFSTORE = cboType.text
   
   Rst1!StartDate = CDate(Format(txtLeaseStDt.text, "dd mmmm yyyy"))
   Rst1!EndDate = CDate(Format(txtLeaseEndDate.text, "dd mmmm yyyy"))
   Rst1!YearEnd = CDate(Format(txtYearEnd.text, "dd mmmm yyyy"))

   '****************************************************
   ' Adding data from Rent Charge tab
   '****************************************************
   If cboRentPayable.text = "Yes" Then
      Rst1!BRPayable = "Y"

      szaTemp = Split(cboFreqBR.text, "-")
      Rst1!BRStartDate = CDate(Format(txtRentStartDate.text, "dd mmmm yyyy"))
      Rst1!BRfrequency = CInt(szaTemp(0))
      If txtTotalRentYear.text <> "" Then Rst1!BRTotal = CDbl(txtTotalRentYear.text)
      If txtNextDueDate.text <> "" Then Rst1!BRNextDueDate = txtNextDueDate.text
      If txtRentDueEachPeriod.text <> "" Then Rst1!BRAmount = CDbl(txtRentDueEachPeriod.text)
      Rst1!BRDemandType = CByte(IIf(cboBRDemandType.text <> "", cboBRDemandType.ListIndex + 1, 0))
      Rst1!RentChargeDept = cboRentChargeDept.BoundColumn
   Else
      Rst1!BRPayable = "N"
   End If

   '****************************************************
   ' Adding data from Service Charge tab
   '****************************************************
   If cboSCPayable.text = "Yes" Then
      Rst1!SCPayable = "Y"

      szaTemp = Split(cboFreqSC.text, "-")
      Rst1!SCfrequency = CInt(szaTemp(0))
      Rst1!SCPayableFrom = CDate(Format(txtPayableFrom.text, "dd mmmm yyyy"))
      If txtSCNextDueDt.text <> "" Then Rst1!SCNextDueDate = CDate(Format(txtSCNextDueDt.text, "dd mmmm yyyy"))

      If optPercentage.Value Then
         Rst1!SCPercentage = CDbl(txtSCPercentage.text)
         Rst1!SCPricePerSqFoot = Null
         Rst1!SCAnnual = Null
         Rst1!SCGlobal = Null
      End If

      If optSqFoot.Value Then
         Rst1!SCPricePerSqFoot = CDbl(txtPPSqFoot.text)
         Rst1!SCPercentage = Null
         Rst1!SCAnnual = Null
         Rst1!SCGlobal = Null
      End If

      If optFixedTotal.Value Then
         Rst1!SCAnnual = CDbl(txtAnnualService.text)
         Rst1!SCPricePerSqFoot = Null
         Rst1!SCPercentage = Null
         Rst1!SCGlobal = Null
      End If

      If optGlobalData.Value Then
         Rst1!SCGlobal = CDbl(txtGlobalAmount.text)
         Rst1!SCPricePerSqFoot = Null
         Rst1!SCPercentage = Null
         Rst1!SCAnnual = Null
      End If

      If cboSCPayable.text = "Yes" Then
         Rst1!SCAmount = CDbl(txtFinalAmout.text)
         Rst1!SCTotal = CDbl(txtAmount.text)
      End If

      If txtTOLimit.text <> "" Then Rst1!SCTOLimit = CDbl(txtTOLimit.text)
      Rst1!SCDemandType = CByte(IIf(cboSCDemandType.text <> "", cboSCDemandType.ListIndex + 1, 0))
      Rst1!ServiceChargeDept = cboServiceChargeDept.BoundColumn
   Else
      Rst1!SCPayable = "N"
   End If
   
   '****************************************************
   ' Adding data from Interest Charge tab
   '****************************************************
   If cboIntCrgable.text = "Yes" Then
      Rst1!InterestChargeable = "Y"
      
      If txtIntPayableAfterDays.text <> "" Then Rst1!DaysAfterInterestPayable = CInt(txtIntPayableAfterDays.text)
      If txtAdditionalIntRate.text <> "" Then Rst1!AdditionalInterest = CDbl(txtAdditionalIntRate.text)
      If txtAmtCrgIntOn.text <> "" Then Rst1!InterestChargedOn = CDbl(txtAmtCrgIntOn.text)
      If txtInt2bChrg.text <> "" Then Rst1!InterestAmount = CDbl(txtInt2bChrg.text)
      Rst1!IntDemandType = CByte(IIf(cboIntDemandType.text <> "", cboIntDemandType.ListIndex + 1, 0))
      Rst1!IntChargeDept = cboIntChargeDept.BoundColumn
   Else
      Rst1!InterestChargeable = "N"
   End If
   
   '****************************************************
   ' Adding data from Break Clause tab
   '****************************************************
   If cboBreakClause.text = "Yes" Then
      Rst1!BreakClause = "Y"
   
      If txtBreakDate.text <> "" Then Rst1!BreakDate = txtBreakDate.text
      If cboBreak.text <> "" Then Rst1!BreakType = cboBreak.text
   Else
      Rst1!BreakClause = "N"
   End If
   
   '****************************************************
   ' Adding data from Rent Review tab
   '****************************************************
   If txtRentReviewDt.text <> "" Then Rst1!RentReviewDate = txtRentReviewDt.text
   If txtRentIncDt.text <> "" Then Rst1!RentIncreaseDate = txtRentIncDt.text
   If txtRentIncAmt.text <> "" Then Rst1!RentIncreaseAmount = CDbl(txtRentIncAmt.text)
   
   '****************************************************
   ' Adding data from Supplementary tab
   '****************************************************
   If txtDtFlgDate.text <> "" Then Rst1!DateFlagDate = txtDtFlgDate.text
   If txtDtFlgDesc.text <> "" Then Rst1!DateFlagDescription = txtDtFlgDesc.text
   If txtMemo.text <> "" Then Rst1!Notes = txtMemo.text
   If Text1.text <> "" Then Rst1!Text1 = Text1.text
   If Text2.text <> "" Then Rst1!Text2 = Text2.text
   If Text3.text <> "" Then Rst1!Text3 = Text3.text
   If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
   If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
   If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption

'   ****************************************************
'    Adding data from Insurance tab
'   ****************************************************
   If cboInsurancePayable.text = "Yes" Then
      Rst1!InsurancePayable = "Y"

      szaTemp = Split(cboInsuranceFrequency.text, "-")
      Rst1!InsuranceFrequency = CInt(szaTemp(0))
      Rst1!InsuranceStartDate = IIf(txtInsuranceStartDate.text = "", "", txtInsuranceStartDate.text)
      Rst1!InsuranceEndDate = IIf(txtInsuranceEndDate.text = "", "", txtInsuranceEndDate.text)
      Rst1!InsuranceDemandType = CByte(IIf(cboInsuranceDemandType.text <> "", cboInsuranceDemandType.ListIndex + 1, 0))
      If txtInsuranceNextDueDate.text <> "" Then Rst1!InsuranceNextDueDate = txtInsuranceNextDueDate.text
      If optInsurancePercentage.Value Then
         Rst1!InsurancePercentage = CDbl(IIf(txtInsurancePercentage.text = "", 0, txtInsurancePercentage.text))
         Rst1!AnnualInsuranceCharge = 0
      End If
      If optAnnualInsuranceCharge.Value Then
         Rst1!AnnualInsuranceCharge = CDbl(IIf(txtAnnualInsuranceCharge.text = "", 0, txtAnnualInsuranceCharge.text))
         Rst1!InsurancePercentage = 0
      End If
      Rst1!TotalYearlyInsurance = CDbl(IIf(txtTotalYearlyInsurance.text = "", 0, txtTotalYearlyInsurance.text))
      Rst1!InsuranceEachPeriod = CDbl(IIf(txtInsuranceEachPeriod.text = "", 0, txtInsuranceEachPeriod.text))
      Rst1!InsuranceDept = cboInsuranceDept.BoundColumn
   Else
      Rst1!InsurancePayable = "N"
   End If

'   ********************************************************
'    Mark the Lease as live. Expired Lease status is false.
'   ********************************************************
   Rst1!Status = True
'   ********************************************************

   Rst1.Update
   Rst1.Close

   SQLStr1 = "UPDATE Units " & _
             "SET OCCUPIED = 'Y' " & _
             "WHERE UNITNUMBER ='" & szaUnit(0) & "';"
   Conn1.Execute SQLStr1

   Conn1.Close
   Set Conn1 = Nothing

   Call DisableBoxes

   cmdAddNew.Visible = True
   cmdAddNew.TabIndex = 25
   cmdDelete.Visible = True
   cmdDelete.TabIndex = 26
   cmdEdit.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False

   Call EmptyBoxes

   SaveUpdateLease = True
End Function

Private Sub cmdtenants_Click()
   cboTenant.Left = txtTenant.Left
   cboTenant.Top = txtTenant.Top
   cboTenant.Visible = True
   cboTenant.SetFocus
End Sub

Private Sub flxLeaseList_Click()
   Call EmptyBoxes
   cmdGridUnitLookup_Click
   
   txtLeaseID.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 1)
   txtTenant.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 3)
   txtUnitName.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 4)
   cboUnit.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 5)
   txtClient.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 6)
   txtProperty.text = flxLeaseList.TextMatrix(flxLeaseList.Row, 7)
   
   GetRecord
   
   If chkExpLease.Value = 1 Then
      cmdAddNew.Visible = False
      cmdEdit.Visible = False
      cmdDelete.Visible = False
   Else
      cmdAddNew.Visible = True
      cmdEdit.Visible = True
      cmdDelete.Visible = True
   End If
End Sub

Private Sub flxRentAnalysis_RowColChange()
'   If cmdEditRentAnalysis.Enabled Then Exit Sub
   populateControl frmLease2, flxRentAnalysis
   RentReviewButtonMode GridRowOnSelection
End Sub

Private Sub Form_Load()
   If Not AllDemandType(szaDemandtype) Then
      MsgBox "You have not defined any demand types. Please create demand types within Global Data.", vbInformation + vbOKOnly, "Demand Type"
      FormLoad = False
      Exit Sub
   End If
   
    Me.Top = 50
    Me.Left = 50

   tabLease.Tab = 0
   
On Error GoTo ErrorTrap

ConfigureFlxGrid flxRentAnalysis

Call EmptyBoxes
Call DisableBoxes
Call FillCbos

BreachButtonMode DefaultMode
AssignmentButtonMode DefaultMode
RentReviewButtonMode DefaultMode

Call FillcboType(szaDemandtype)

FormLoad = True
LoadDept

Exit Sub
ErrorTrap:
    If ERR.Number > 0 Then
        If ERR.Number = 40002 Then
            If MsgBox("DSN - " & Adsn & " not found. Please check with your system administrator.", vbRetryCancel + vbCritical, "DSN Set Up Error") = vbRetry Then
                Resume
            Else
                Exit Sub
            End If
        Else
            MsgBox ERR.Number & " - " & ERR.description
            Exit Sub
        End If
    End If
End Sub

Private Sub LoadFlxGrid(conFlxGrid As Control)
'   Dim szaTenant() As String
   Dim iRow As Integer
   
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
   
   'get all sage account numbers and company names from tenants.
   SQLStr2 = "SELECT * " & _
             "FROM RentAnalysis " & _
             "WHERE SAGEACCOUNTNUMBER = '" & flxLeaseList.TextMatrix(flxLeaseList.Row, 2) & "' " & _
             "ORDER BY ID ASC"
   Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)
   
   If Not Rst2.EOF Then
      iRow = 1
      While Not Rst2.EOF
         conFlxGrid.TextMatrix(iRow, 1) = IIf(IsNull(Rst2!SerialNumber), "", Rst2!SerialNumber)
         conFlxGrid.TextMatrix(iRow, 2) = IIf(IsNull(Rst2!RentReviewDate), "", Rst2!RentReviewDate)
         conFlxGrid.TextMatrix(iRow, 3) = IIf(IsNull(Rst2!RentIncreaseDate), "", Rst2!RentReviewDate)
         conFlxGrid.TextMatrix(iRow, 4) = IIf(IsNull(Rst2!RentIncreaseAmount), "", Rst2!RentIncreaseAmount)
         conFlxGrid.TextMatrix(iRow, 5) = Rst2!ID
         Rst2.MoveNext
         If Not Rst2.EOF Then conFlxGrid.AddItem ""
         iRow = iRow + 1
      Wend
   End If
   
   Rst2.Close
   Conn2.Close
   Set Rst2 = Nothing
   Set Conn2 = Nothing
End Sub

Private Sub ConfigureFlxGrid(conFlxGrid As Control)
   Dim szFlxHeader As String
   
   conFlxGrid.RowHeight(0) = 150
   conFlxGrid.Clear
   conFlxGrid.Cols = 6
   szFlxHeader$ = "|<Serial|<RentReviewDate|<RentIncreaseDate|>RentIncreaseAmount|ID"
   conFlxGrid.FormatString = szFlxHeader$
   
   conFlxGrid.ColWidth(0) = 0
   conFlxGrid.ColWidth(1) = txtRentReviewDate.Left - txtSerial.Left
   conFlxGrid.ColWidth(2) = txtRentIncreaseDate.Left - txtRentReviewDate.Left
   conFlxGrid.ColWidth(3) = txtRentIncreaseAmount.Left - txtRentIncreaseDate.Left
   conFlxGrid.ColWidth(4) = txtRentIncreaseAmount.Width
   conFlxGrid.ColWidth(5) = 0       'ID
End Sub

Public Sub GetTenantsWithLease()
   Dim temp As String

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT SageAccountNumber, CompanyName " & _
             "FROM LeaseDetails " & _
             "ORDER BY SageAccountNumber"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   cboTenant.Clear
   If Rst1.EOF = False Then
       While Rst1.EOF = False
           cboTenant.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
           Rst1.MoveNext
       Wend
   End If
   Rst1.Close
   Conn1.Close
'
   cmdAddNew.Visible = True
   cmdAddNew.TabIndex = 25
   cmdDelete.Visible = True
   cmdDelete.TabIndex = 26
   cmdEdit.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
End Sub

Public Sub GetTenantsWithoutLease()
   Dim temp As String
   
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt

   'get all sage account numbers and company names from tenants.
   SQLStr2 = "SELECT SageAccountNumber, CompanyName " & _
             "FROM Tenants " & _
             "WHERE Tenants.SageAccountNumber NOT IN " & _
                 "(SELECT LeaseDetails.SageAccountNumber " & _
                 "FROM LeaseDetails " & _
                 "WHERE Status=True) " & _
             "ORDER BY SageAccountNumber"
   Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)
   cboTenant.Clear
   While Rst2.EOF = False
       cboTenant.AddItem Rst2!SageAccountNumber & " / " & Rst2!CompanyName
       Rst2.MoveNext
   Wend
   
   Rst2.Close
   Conn2.Close
   
   cmdAddNew.Visible = False
   cmdEdit.Visible = False
   cmdDelete.Visible = False
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
   cmdSaveNew.Visible = True
   cmdSaveNew.TabIndex = 25
   cmdCancelNew.Visible = True
   cmdCancelNew.TabIndex = 26
End Sub

Public Sub FillCbos()
   Dim i As Integer

   'Fill the yes / no cbos
   cboRentPayable.AddItem "No", 0
   cboRentPayable.AddItem "Yes", 1
   cboSCPayable.AddItem "No", 0
   cboSCPayable.AddItem "Yes", 1
   cboIntCrgable.AddItem "No", 0
   cboIntCrgable.AddItem "Yes", 1
   cboBreakClause.AddItem "No", 0
   cboBreakClause.AddItem "Yes", 1
   '' Added by Asif 09-01-2006
   cboInsurancePayable.AddItem "No", 0
   cboInsurancePayable.AddItem "Yes", 1

   'fill the frequencies
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
'
   SQLStr1 = "SELECT * FROM Frequencies"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   i = 0
   If Rst1.EOF = False Then
       While Rst1.EOF = False
           cboFreqBR.AddItem Rst1!ID & "-" & Rst1!Frequency, i
           cboFreqSC.AddItem Rst1!ID & "-" & Rst1!Frequency, i
           ' Insurance. By Asif 09/01/2006
           cboInsuranceFrequency.AddItem Rst1!ID & "-" & Rst1!Frequency, i
           i = i + 1
           Rst1.MoveNext
       Wend
       cboFreqSC.ListIndex = 0
   End If
   
   '' Fill the Head leases
   Rst1.Close
   SQLStr1 = "SELECT LeaseID FROM LeaseDetails"

   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   i = 0
   If Rst1.EOF = False Then
       While Rst1.EOF = False
           cboHeadLease.AddItem Rst1!LeaseId, i
           i = i + 1
           Rst1.MoveNext
       Wend
   End If

   Rst1.Close

   'fill the type of store cbo.
   LoadType "LEASE TYPE", cboType

   'fill the break type cbo.
   cboBreak.AddItem "Landlord", 0
   cboBreak.AddItem "Tenant", 1
   cboBreak.AddItem "Mutual", 2
      
   'Set the RDO Connections to the dataset
   adoBreaches.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
     
   SQLStr1 = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'BTYP'"

   adoBreaches.RecordSource = SQLStr1
   adoBreaches.CommandType = adCmdText
   adoBreaches.Refresh
   
   Conn1.Close
End Sub

Private Sub LoadType(szValue As String, conCombo As Control)
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
'
   SQLStr1 = "SELECT SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"
   Set Rst2 = Conn2.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   conCombo.Clear
   While Not Rst2.EOF
      conCombo.AddItem Rst2!V
      Rst2.MoveNext
   Wend
'
   Rst2.Close
   Conn2.Close
   Set Rst2 = Nothing
   Set Conn2 = Nothing
End Sub

Public Sub GetRecord()
   Dim i As Integer

   'Set the RDO Connection to the dataset
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   
   'Get record for selected Tenant.
   SQLStr1 = "SELECT LeaseDetails.*, Property.PropertyName, Client.ClientName " & _
             "FROM LeaseDetails, Property, Client, Units " & _
             "WHERE LeaseDetails.LeaseID = '" & txtLeaseID.text & "' AND " & _
             "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
             "Units.PropertyID = Property.PropertyID AND " & _
             "Client.ClientId = Property.ClientId"

   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   'Check for sub lease.
   Dim HeadLease As String
   If IsNull(Rst1!HeadLease) = False Then
      HeadLease = Rst1!HeadLease
      chkSubLease.Value = 1
   Else
      HeadLease = ""
      chkSubLease.Value = 0
   End If

   'Fill text boxes with lease details.
   cboType.text = IIf(IsNull(Rst1!TYPEOFSTORE), "", Rst1!TYPEOFSTORE)
   txtLeaseStDt.text = IIf(IsNull(Rst1!StartDate), "", Rst1!StartDate)
   txtLeaseEndDate.text = IIf(IsNull(Rst1!EndDate), "", Rst1!EndDate)
   txtYearEnd.text = IIf(IsNull(Rst1!YearEnd), "", Rst1!YearEnd)
   
   '****************************************
   '  Rent Charge tab
   '****************************************
   If Rst1!BRPayable = "Y" Then
      cboRentPayable.text = "Yes"
      
      txtRentStartDate.text = Rst1!BRStartDate
      txtTotalRentYear.text = Rst1!BRTotal
      txtNextDueDate.text = Rst1!BRNextDueDate
      txtRentDueEachPeriod.text = Rst1!BRAmount
      cboBRDemandType.ListIndex = (CByte(Rst1!BRDemandType) - 1)
      cboRentChargeDept.text = IIf(IsNull(Rst1!RentChargeDept), "", DeptName(IIf(IsNull(Rst1!RentChargeDept) Or Rst1!RentChargeDept = "", 1, Rst1!RentChargeDept)))
      If Rst1!BRfrequency > 0 Then
          cboFreqBR.text = cboFreqBR.List(Rst1!BRfrequency - 1)
      End If
   Else
      cboRentPayable.text = "No"
   End If
   
   '****************************************
   '  Service Charge tab
   '****************************************
   If Rst1!SCPayable = "Y" Then
      cboSCPayable.text = "Yes"

      cboServiceChargeDept.text = IIf(IsNull(Rst1!ServiceChargeDept), "", DeptName(IIf(IsNull(Rst1!ServiceChargeDept) Or Rst1!ServiceChargeDept = "", 1, Rst1!ServiceChargeDept)))
      txtPayableFrom.text = Rst1!SCPayableFrom
      If IsNull(Rst1!SCTOLimit) = False Then txtTOLimit.text = Rst1!SCTOLimit
      cboSCDemandType.ListIndex = CByte(Rst1!SCDemandType) - 1

      If Val(IIf(IsNull(Rst1!SCPercentage), 0, Rst1!SCPercentage)) > 0 Then
         txtSCPercentage.text = Format(Rst1!SCPercentage, "0.00")
         optPercentage.Value = True
      Else
         If Val(IIf(IsNull(Rst1!SCPricePerSqFoot), 0, Rst1!SCPricePerSqFoot)) > 0 Then
            txtPPSqFoot.text = Format(Rst1!SCPricePerSqFoot, "0.00")
            optSqFoot.Value = True
         Else
            If Val(IIf(IsNull(Rst1!SCAnnual), 0, Rst1!SCAnnual)) > 0 Then
               txtAnnualService.text = Format(Rst1!SCAnnual, "0.00")
               optFixedTotal.Value = True
            Else
               txtGlobalAmount.text = Format(Rst1!SCGlobal, "0.00")
               optGlobalData.Value = True
            End If
         End If
      End If

      txtAmount.text = Format(Rst1!SCTotal, "0.00")
      txtFinalAmout.text = IIf(IsNull(Rst1!SCAmount), "0.00", Rst1!SCAmount)
      If IsNull(Rst1!SCNextDueDate) = False Then txtSCNextDueDt.text = Rst1!SCNextDueDate

      If Not IsNull(Rst1!SCfrequency) Then
          If Rst1!SCfrequency > 0 Then
              cboFreqSC.text = cboFreqSC.List(Rst1!SCfrequency - 1)
          End If
      End If
   Else
      cboSCPayable.text = "No"
   End If

   '****************************************
   '  Interest Charge tab
   '****************************************
   If Rst1!InterestChargeable = "Y" Then
      cboIntCrgable.text = "Yes"
      If IsNull(Rst1!DaysAfterInterestPayable) = False Then txtIntPayableAfterDays.text = Rst1!DaysAfterInterestPayable
      cboIntChargeDept.text = IIf(IsNull(Rst1!IntChargeDept), "", DeptName(IIf(IsNull(Rst1!IntChargeDept) Or Rst1!IntChargeDept = "", 1, Rst1!IntChargeDept)))
      If IsNull(Rst1!AdditionalInterest) = False Then txtAdditionalIntRate.text = Rst1!AdditionalInterest
      If IsNull(Rst1!InterestChargedOn) = False Then txtAmtCrgIntOn.text = Rst1!InterestChargedOn
      If IsNull(Rst1!InterestAmount) = False Then txtInt2bChrg.text = Rst1!InterestAmount
      If Not IsNull(Rst1!IntDemandType) Then cboIntDemandType.ListIndex = CByte(Rst1!IntDemandType) - 1
   Else
      cboIntCrgable.text = "No"
   End If

   '****************************************
   '  Break Clause tab
   '****************************************
   If Rst1!BreakClause = "Y" Then
      cboBreakClause.text = "Yes"
      If IsNull(Rst1!BreakType) = False Then cboBreak.text = Rst1!BreakType
      If IsNull(Rst1!BreakDate) = False Then txtBreakDate.text = Rst1!BreakDate
   Else
      cboBreakClause.text = "No"
   End If
   
   '****************************************
   '  Rent Review tab
   '****************************************
   If IsNull(Rst1!RentReviewDate) = False Then txtRentReviewDt.text = Rst1!RentReviewDate
   If IsNull(Rst1!RentIncreaseDate) = False Then txtRentIncDt.text = Rst1!RentIncreaseDate
   If IsNull(Rst1!RentIncreaseAmount) = False Then txtRentIncAmt.text = Rst1!RentIncreaseAmount
   LoadFlxGrid flxRentAnalysis
   
   '****************************************
   '  Supplementary tab
   '****************************************
   If IsNull(Rst1!DateFlagDate) = False Then txtDtFlgDate.text = Rst1!DateFlagDate
   If IsNull(Rst1!DateFlagDescription) = False Then txtDtFlgDesc.text = Rst1!DateFlagDescription
   If IsNull(Rst1!Notes) = False Then txtMemo.text = Rst1!Notes
   
   If IsNull(Rst1!Text1) = False Then Text1.text = Rst1!Text1
   If IsNull(Rst1!Text2) = False Then Text2.text = Rst1!Text2
   If IsNull(Rst1!Text3) = False Then Text3.text = Rst1!Text3
   
   If IsNull(Rst1!SuppCaption1) = False Then lblSupplementary1.Caption = Rst1!SuppCaption1
   If IsNull(Rst1!SuppCaption2) = False Then lblSupplementary2.Caption = Rst1!SuppCaption2
   If IsNull(Rst1!SuppCaption3) = False Then lblSupplementary3.Caption = Rst1!SuppCaption3
   
   '****************************************
   '  Insurance tab
   '****************************************
   If Rst1!InsurancePayable = "Y" Then
      cboInsurancePayable.text = "Yes"
      If Not IsNull(Rst1!InsuranceFrequency) Then
          If Rst1!InsuranceFrequency > 0 Then
              cboInsuranceFrequency.text = cboInsuranceFrequency.List(Rst1!InsuranceFrequency - 1)
          End If
      End If
      cboInsuranceDept.text = IIf(IsNull(Rst1!InsuranceDept), "", DeptName(IIf(IsNull(Rst1!InsuranceDept) Or Rst1!InsuranceDept = "", 1, Rst1!InsuranceDept)))
      txtInsuranceStartDate.text = IIf(IsNull(Rst1!InsuranceStartDate), "", Rst1!InsuranceStartDate)
      txtInsuranceEndDate.text = IIf(IsNull(Rst1!InsuranceEndDate), "", Rst1!InsuranceEndDate)
      cboInsuranceDemandType.ListIndex = IIf(IsNull(Rst1!InsuranceDemandType), "", CByte(Rst1!InsuranceDemandType) - 1)
      txtInsuranceNextDueDate.text = IIf(IsNull(Rst1!InsuranceNextDueDate), "", Rst1!InsuranceNextDueDate)

      If Not IsNull(Rst1!InsurancePercentage) Then
         If Rst1!InsurancePercentage > 0 Then
            txtInsurancePercentage.text = Format(Rst1!InsurancePercentage, "0.0000")
            optInsurancePercentage.Value = True
            txtAnnualInsuranceCharge.text = ""
         End If
      End If
      If Not IsNull(Rst1!AnnualInsuranceCharge) Then
         If Rst1!AnnualInsuranceCharge > 0 And Rst1!TotalYearlyInsurance = Rst1!AnnualInsuranceCharge Then
            txtAnnualInsuranceCharge.text = Format(Rst1!AnnualInsuranceCharge, "0.00")
            optAnnualInsuranceCharge.Value = True
            txtInsurancePercentage.text = ""
         End If
      End If
      txtTotalYearlyInsurance.text = Format(IIf(IsNull(Rst1!TotalYearlyInsurance), "", Rst1!TotalYearlyInsurance), "0.00")
      txtInsuranceEachPeriod.text = Format(IIf(IsNull(Rst1!InsuranceEachPeriod), "", Rst1!InsuranceEachPeriod), "0.00")
   Else
      cboInsurancePayable.text = "No"
   End If

   Rst1.Close
   Conn1.Close
   Set Conn1 = Nothing

   PopulateBreaches
   PopulateAssignments
End Sub

Public Sub PopulateBreaches()
   'Set the RDO Connections to the dataset
   Dim sSQLQuery_ As String

   Dim adoConn As New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT LeaseBreaches.BreachID, " & _
         "SecondaryCode.value as BreachType, " & _
         "LeaseBreaches.CommenceDate, " & _
         "LeaseBreaches.InitiatedBy, LeaseBreaches.Resolved, " & _
         "LeaseBreaches.DateReceived, LeaseBreaches.ReceivedBy " & _
         "FROM LeaseBreaches, SecondaryCode " & _
         "WHERE LeaseBreaches.LeaseID = '" & txtLeaseID.text & "' " & _
         "AND SecondaryCode.Code = LeaseBreaches.BreachType " & _
         "AND SecondaryCode.PrimaryCode = 'BTYP'"
   
   populateGrid adoConn, sSQLQuery_, gridBreach
   SetBreachGrid
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub PopulateAssignments()
   'Set the RDO Connections to the dataset
   Dim sSQLQuery_ As String, szHeader As String

   Dim adoConn As New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT LeaseAssignments.AssignmentID, " & _
                  "LeaseAssignments.AssignDate, " & _
                  "LeaseAssignments.Assignee, " & _
                  "LeaseAssignments.Decp " & _
                "FROM LeaseAssignments " & _
                "WHERE LeaseAssignments.LeaseID = '" & txtLeaseID.text & "' "

   SetAssignmentGrid
   szHeader$ = "<AssignmentID|<Assignment_Date|<Assignee|<Description"
   populateGridSimply adoConn, sSQLQuery_, gridAssignment, szHeader

   adoConn.Close
   Set adoConn = Nothing
End Sub


Public Sub SetBreachGrid()
   
   Dim conBreach As New RDO.rdoConnection
   Dim rstBreach As rdoResultset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
   conBreach.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBreach.CursorDriver = rdUseIfNeeded
   conBreach.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT LeaseBreaches.BreachID, " & _
      "LeaseBreaches.BreachType, " & _
      "LeaseBreaches.CommenceDate, " & _
      "LeaseBreaches.InitiatedBy, LeaseBreaches.Resolved, " & _
      "LeaseBreaches.DateReceived, LeaseBreaches.ReceivedBy " & _
      "FROM LeaseBreaches, SecondaryCode " & _
      "WHERE LeaseBreaches.LeaseID = '" & txtLeaseID.text & "' "

   Set rstBreach = conBreach.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

'   gridBreach.Clear
'   gridBreach.Rows = 2
   gridBreach.Cols = 7

   gridBreach.ColWidth(0) = 0
   gridBreach.ColWidth(1) = cboBreachType.Width + cmdSetBreachType.Width + 5
   gridBreach.ColWidth(2) = txtCommenceDate.Width + 5
   gridBreach.ColWidth(3) = txtInitiatedBy.Width + 5
   gridBreach.ColWidth(4) = txtDateReceived.Left - (txtInitiatedBy.Left + txtInitiatedBy.Width) + 5
   gridBreach.ColWidth(5) = txtDateReceived.Width + 5
   gridBreach.ColWidth(6) = txtReceivedBy.Width + 5

   Dim oColumn As rdoColumn
   Dim iColumn As Integer
   iColumn = 0

   gridBreach.Cols = rstBreach.rdoColumns.Count
   For Each oColumn In rstBreach.rdoColumns
        gridBreach.TextMatrix(0, iColumn) = oColumn.Name
        iColumn = iColumn + 1
   Next oColumn

   rstBreach.Close
   conBreach.Close
   Set rstBreach = Nothing
   Set conBreach = Nothing
End Sub

Public Sub SetAssignmentGrid()
   
'   Dim conAssignment As New RDO.rdoConnection
'   Dim rstAssignment As rdoResultset
'   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
'   conAssignment.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conAssignment.CursorDriver = rdUseIfNeeded
'   conAssignment.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'   sSQLQuery_ = "SELECT LeaseAssignments.AssignmentID, " & _
'                  "LeaseAssignments.AssignDate, " & _
'                  "LeaseAssignments.Assignee, " & _
'                  "LeaseAssignments.Decp " & _
'                "FROM LeaseAssignments " & _
'                "WHERE LeaseAssignments.LeaseID = '" & txtLeaseID.text & "' "
'
'   Set rstAssignment = conAssignment.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

   gridAssignment.Clear
   gridAssignment.Rows = 2
   gridAssignment.Cols = 4

   gridAssignment.ColWidth(0) = 0
   gridAssignment.ColWidth(1) = txtAssignee.Left - txtAssignment_Date.Left
   gridAssignment.ColWidth(2) = txtDescription.Left - txtAssignee.Left
   gridAssignment.ColWidth(3) = txtDescription.Width

'   Dim oColumn As rdoColumn
'   Dim iColumn As Integer
'   iColumn = 0
'
'   gridAssignment.Cols = rstAssignment.rdoColumns.Count
'   For Each oColumn In rstAssignment.rdoColumns
'        gridAssignment.TextMatrix(0, iColumn) = oColumn.Name
'        iColumn = iColumn + 1
'   Next oColumn
'
'   rstAssignment.Close
'   conAssignment.Close
'   Set rstAssignment = Nothing
'   Set conAssignment = Nothing
End Sub

'Public Sub GetUnits()
'    Dim i As Integer
'
'    'Get all the unoccupied units and put in cbounit
'    cboUnit.Clear
'
'    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'    Conn1.CursorDriver = rdUseIfNeeded
'    Conn1.EstablishConnection rdDriverNoPrompt
'
'    SQLStr1 = "SELECT UnitNumber FROM Units WHERE Occupied = 'N' ORDER BY UnitNumber"
'    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'    If Rst1.EOF = False Then
'        While Rst1.EOF = False
'            cboUnit.AddItem Rst1!UnitNumber
'            Rst1.MoveNext
'        Wend
'    End If
'
'    Rst1.Close
'
'    'Get unit of current tenant and put in cbounit.text
'    If TenantCode <> "" Then
'        SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
'        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'        If IsNull(Rst1!CurrentRental) Then OldUnit = "" Else OldUnit = Rst1!CurrentRental
'
'        Rst1.Close
'        Conn1.Close
'        If OldUnit = "" Then
'            cboUnit.text = ""
'        Else
'            cboUnit.AddItem OldUnit, 0
'            cboUnit.text = cboUnit.List(0)
'        End If
'    Else
'        Conn1.Close
'    End If
'
'End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub mnuDemands_Click()

Load frmDemands
Unload Me
frmDemands.Show

End Sub

'Private Sub mnuEdit_Click()
'
'Call Edit
'
'End Sub

Private Sub mnuExit_Click()
   Unload frmMMain
End Sub

Private Sub mnuGlobal_Click()

Call EmptyBoxes

Load frmGlobal
Unload Me
frmGlobal.Show

End Sub

Private Sub mnuMain_Click()
   Call EmptyBoxes
   
   Unload Me
End Sub

Private Sub mnuShopCentre_Click()

Call EmptyBoxes

Load frmShoppingCentre
Unload Me
frmShoppingCentre.Show

End Sub

Private Sub mnuTenants_Click()
   Call EmptyBoxes
End Sub

Private Sub mnuUnits_Click()
   Call EmptyBoxes
   Unload Me
End Sub

Private Sub gridAssignment_RowColChange()
   populateControl frmLease2, gridAssignment
   AssignmentButtonMode GridRowOnSelection
End Sub

Private Sub gridBreach_Click()
   BreachButtonMode GridRowOnSelection
End Sub

Private Sub gridBreach_RowColChange()
   populateControl frmLease2, gridBreach
End Sub

Private Sub lblSupplementary1_Click()
   txtSuppCaption1.Visible = True
   txtSuppCaption1.Left = lblSupplementary1.Left
   txtSuppCaption1.text = lblSupplementary1.Caption
   txtSuppCaption1.SetFocus
End Sub

Private Sub lblSupplementary2_Click()
   txtSuppCaption2.Visible = True
   txtSuppCaption2.Left = lblSupplementary2.Left
   txtSuppCaption2.text = lblSupplementary2.Caption
   txtSuppCaption2.SetFocus
End Sub

Private Sub lblSupplementary3_Click()
txtSuppCaption3.Visible = True
txtSuppCaption3.Left = lblSupplementary3.Left
txtSuppCaption3.text = lblSupplementary3.Caption

txtSuppCaption3.SetFocus

End Sub

Private Sub optAnnualInsuranceCharge_Click()
   txtInsurancePercentage.Locked = True
   txtAnnualInsuranceCharge.Locked = False
   If cmdSaveEdit.Visible = True Then txtAnnualInsuranceCharge.SetFocus
   txtInsurancePercentage.text = "0.00"
End Sub

Private Sub optFixedTotal_Click()
   txtSCPercentage.text = ""
   txtSCPercentage.Locked = True
   txtPPSqFoot.text = ""
   txtPPSqFoot.Locked = True
   txtAnnualService.Locked = False
   txtGlobalAmount.text = ""
   txtGlobalAmount.Locked = True

   If cmdSaveNew.Visible Or cmdSaveEdit.Visible Then txtAnnualService.SetFocus
End Sub

Private Sub optGlobalData_Click()
   If Not cmdSaveNew.Visible And Not cmdSaveEdit.Visible Then Exit Sub
   If cboFreqSC.text = "" Then
      MsgBox "Please choose the frequency.", vbCritical + vbOKOnly, "Charging Method"
      cboFreqSC.SetFocus
      Exit Sub
   End If

   MousePointer = vbHourglass

   txtSCPercentage.text = ""
   txtSCPercentage.Locked = True
   txtPPSqFoot.text = ""
   txtPPSqFoot.Locked = True
   txtAnnualService.text = ""
   txtAnnualService.Locked = True
   txtGlobalAmount.Locked = False

   If cmdSaveNew.Visible Or cmdSaveEdit.Visible Then txtAnnualService.SetFocus

   Dim rdoConn As New RDO.rdoConnection
   Dim rstRst As rdoResultset
   Dim szSQL As String, szaUnit() As String

   szaUnit = Split(cboUnit.text, " - ")
   txtGlobalAmount.text = Format(GetPPSF(szaUnit(0)) * GetUnitTA(szaUnit(0)), "0.00")
   txtAmount.text = txtGlobalAmount.text

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   If cboFreqSC.ListIndex > -1 Then
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
   Else
      Dim temp() As String

      temp = Split(cboFreqSC.text, "-")
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & Val(temp(0)) & ";"
   End If
   Set rstRst = rdoConn.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   txtFinalAmout.text = Format((CDbl(txtAmount.text) / CInt(rstRst!PARTOFYEAR)), "0.00")

   rstRst.Close
   rdoConn.Close

   If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC

   Set rstRst = Nothing
   Set rdoConn = Nothing

   MousePointer = vbDefault
End Sub

Private Sub Option4_Click()
   txtSCPercentage.Locked = False
   txtPPSqFoot.Locked = True
   txtAnnualService.Locked = True
   txtSCPercentage.SetFocus
End Sub

Private Sub optInsurancePercentage_Click()
   txtInsurancePercentage.Locked = False
   txtAnnualInsuranceCharge.Locked = True
   If cmdSaveEdit.Visible = True Then txtInsurancePercentage.SetFocus
   txtAnnualInsuranceCharge.text = "0.00"
End Sub

Private Sub optPercentage_Click()
   txtSCPercentage.Locked = False
   txtPPSqFoot.text = ""
   txtPPSqFoot.Locked = True
   txtAnnualService.text = ""
   txtAnnualService.Locked = True
   txtGlobalAmount.text = ""
   txtGlobalAmount.Locked = True

   If cmdSaveNew.Visible Or cmdSaveEdit.Visible Then txtSCPercentage.SetFocus
End Sub

Private Sub optSqFoot_Click()
   txtSCPercentage.text = ""
   txtSCPercentage.Locked = True
   txtPPSqFoot.Locked = False
   txtAnnualService.text = ""
   txtAnnualService.Locked = True
   txtGlobalAmount.text = ""
   txtGlobalAmount.Locked = True

   If cmdSaveNew.Visible Or cmdSaveEdit.Visible Then txtPPSqFoot.SetFocus
End Sub

Private Sub picLeaseList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then picLeaseList.Visible = False
End Sub

Private Sub txtAssignment_Date_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtAssignment_Date
End Sub

Private Sub txtAssignment_Date_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtAssignment_Date, KeyAscii
End Sub

Private Sub txtDateReceived_Change()
   'Added by Samrat. 16.01.2006
   TextBoxChangeDate txtDateReceived
End Sub

Private Sub txtDateReceived_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtDateReceived, KeyAscii
End Sub

Private Sub txtPayableFrom_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtPayableFrom
End Sub

Private Sub txtPayableFrom_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtPayableFrom, KeyAscii
End Sub

Private Sub txtSCNextDueDt_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtSCNextDueDt
End Sub

Private Sub txtSCNextDueDt_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtSCNextDueDt, KeyAscii
End Sub

Private Sub txtTOLimit_LostFocus()
   If txtTOLimit.text <> "" Then
       If NumberCheck2(txtTOLimit.text) = False Then
           txtTOLimit.text = ""
       Else
           txtTOLimit.text = Round(CDbl(txtTOLimit.text), 2)
       End If
   End If
End Sub

Private Sub txtSCNextDueDt_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtSCNextDueDt
End Sub

Private Sub txtBreakDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtBreakDate
End Sub

Private Sub txtBreakDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtBreakDate, KeyAscii
End Sub

Private Sub txtLeaseEndDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtLeaseEndDate
End Sub

Private Sub txtLeaseEndDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtLeaseEndDate, KeyAscii
End Sub

Private Sub txtYearEnd_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtYearEnd
End Sub

Private Sub txtYearEnd_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtYearEnd, KeyAscii
End Sub

Private Sub txtRentStartDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtRentStartDate
End Sub

Private Sub txtRentStartDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtRentStartDate, KeyAscii
End Sub

Private Sub txtPayableFrom_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtPayableFrom
End Sub

Private Sub txtAnnualInsuranceCharge_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtAnnualInsuranceCharge_LostFocus()
   If txtAnnualInsuranceCharge.Locked Then Exit Sub
   If cboInsuranceFrequency.text = "" Then
      MsgBox "Please select the Insurance Frequency.", vbCritical + vbInformation, "Insurance Frequency"
      Exit Sub
   End If

   Dim Area As String, Total As Double

   MousePointer = vbHourglass

   txtAnnualInsuranceCharge.text = Format(IIf(txtAnnualInsuranceCharge.text = "", 0, txtAnnualInsuranceCharge.text), "0.00")

   Total = CDbl(txtAnnualInsuranceCharge.text)
   txtTotalYearlyInsurance.text = Format(Total, "0.00")

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   
   Dim temp() As String

   temp = Split(cboInsuranceFrequency.text, "-")

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & Val(temp(0)) & ";"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   txtInsuranceEachPeriod.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")

   Rst1.Close
   Conn1.Close

   If cboInsurancePayable.text = "Yes" Then Call CalculateInsuranceCharge

   MousePointer = vbDefault
End Sub

Private Sub txtAnnualService_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtAnnualService_LostFocus()
   If txtAnnualService.Locked Then Exit Sub
   Dim Area As String, Total As Double
   
   MousePointer = vbHourglass
'
   txtAnnualService.text = Format(IIf(txtAnnualService.text = "", 0, txtAnnualService.text), "0.00")
'
   Total = CDbl(txtAnnualService.text)
   txtAmount.text = Format(Total, "0.00")
'
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   If cboFreqSC.ListIndex > -1 Then
      SQLStr1 = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
   Else
      Dim temp() As String
      
      temp = Split(cboFreqSC.text, "-")
      SQLStr1 = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & Val(temp(0)) & ";"
   End If
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   txtFinalAmout.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")
   
   Rst1.Close
   Conn1.Close
'
   If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC

   MousePointer = vbDefault
End Sub


Private Sub txtAssignment_Date_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtAssignment_Date
End Sub

Private Sub txtCommenceDate_Change()
   'Added by Samrat. 16.01.2006
   TextBoxChangeDate txtCommenceDate
End Sub

Private Sub txtCommenceDate_KeyPress(KeyAscii As Integer)
   'Added by Samrat. 16.01.2006
   TextBoxKeyPrsDate txtCommenceDate, KeyAscii
End Sub

Private Sub txtCommenceDate_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtCommenceDate
End Sub


Private Sub txtDateReceived_LostFocus()
   'Added By Asif. 13/01/2006
   TextBoxFormatDate txtDateReceived
End Sub

Private Sub txtInsuranceEndDate_Change()
   'Added by Samrat. 16.01.2006
   TextBoxChangeDate txtInsuranceEndDate
End Sub

Private Sub txtInsuranceEndDate_KeyPress(KeyAscii As Integer)
   'Added by Samrat. 16.01.2006
   TextBoxKeyPrsDate txtInsuranceEndDate, KeyAscii
End Sub

Private Sub txtInsuranceEndDate_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtInsuranceEndDate
End Sub

Private Sub txtInsurancePercentage_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtInsurancePercentage_LostFocus()
   If txtInsurancePercentage.Locked Then Exit Sub
   If cboInsuranceFrequency.text = "" Then
      MsgBox "Please select the Insurance Frequency.", vbCritical + vbInformation, "Insurance Frequency"
      Exit Sub
   End If

   Dim TotalInsurance As Double

   txtInsurancePercentage.text = Format(IIf(txtInsurancePercentage.text = "", 0, txtInsurancePercentage.text), "0.0000")

   txtTotalYearlyInsurance.text = Format(txtInsurancePercentage.text, "0.00")

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT GlobalData.YearlyInsurance " & _
             "FROM GlobalData, Units " & _
             "WHERE Units.PropertyID = GlobalData.PropertyID " & _
               "AND Units.UnitNumber = '" & Left(cboUnit, 8) & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   TotalInsurance = CDbl(Rst1!YearlyInsurance)
   Rst1.Close

   txtTotalYearlyInsurance.text = Format(TotalInsurance * (CDbl(txtInsurancePercentage.text) / 100), "0.00")

   Dim temp() As String

   temp = Split(cboInsuranceFrequency.text, "-")
   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & Val(temp(0)) & ";"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   txtInsuranceEachPeriod.text = Format((CDbl(txtTotalYearlyInsurance.text) / CInt(Rst1!PARTOFYEAR)), "0.00")

   Rst1.Close
   Conn1.Close
'
'   If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC
End Sub

Private Sub txtInsuranceStartDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtInsuranceStartDate
End Sub

Private Sub txtInsuranceStartDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtInsuranceStartDate, KeyAscii
End Sub

Private Sub txtInsuranceStartDate_LostFocus()
   TextBoxFormatDate txtInsuranceStartDate
End Sub

Private Function OnlyNumericString(szString As String) As String
   Dim i As Integer, X As Integer
   
   For i = 1 To Len(szString)
      X = Asc(Mid(szString, i, 1))
'      If (X > 47 And X < 58) Or (X > 64 And X < 91) Or (X > 96 And X < 123) Then
      If (X > 47 And X < 58) Then
         OnlyNumericString = OnlyNumericString & Mid(szString, i, 1)
      End If
   Next i
End Function

Private Sub txtLeaseStDt_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtLeaseStDt
End Sub

Private Sub txtLeaseStDt_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtLeaseStDt, KeyAscii
End Sub

Private Sub txtPPSqFoot_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtPPSqFoot_LostFocus()
   If txtPPSqFoot.Locked Then Exit Sub

   Dim Area As String, Total As Double
   
   txtPPSqFoot.text = Format(IIf(txtPPSqFoot.text = "", 0, txtPPSqFoot.text), "0.0000")

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT TotalArea " & _
             "FROM Units " & _
             "WHERE UnitNumber = '" & Left(cboUnit.text, 8) & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   If Not IsNull(Rst1!TotalArea) Then
      Area = Rst1!TotalArea
   Else
      Area = 0
   End If

   If Area = 0 Then MsgBox "            The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
                         "Please enter the area of the unit in the Unit Analysis Screen.", vbInformation, "Unit Total Area"

   Total = Area * CDbl(txtPPSqFoot.text)
   txtAmount.text = Format(Total, "0.00")

   Rst1.Close

   Dim temp() As String

   temp = Split(cboFreqSC.text, "-")
'   If cboFreqSC.ListIndex > -1 Then
'      SQLStr1 = "SELECT PARTOFYEAR " & _
'                "FROM FREQUENCIES " & _
'                "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
'   Else
   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & Val(temp(0)) & ";"
'   End If
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   txtFinalAmout.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")
   
   Rst1.Close
   Conn1.Close
'
   If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC

   MousePointer = vbDefault
End Sub

Private Sub txtIntPayableAfterDays_LostFocus()

If txtIntPayableAfterDays.text <> "" Then
    If NumberCheck(txtIntPayableAfterDays.text) = False Then
        txtIntPayableAfterDays.text = ""
    Else
        If txtIntPayableAfterDays.text <> "" And txtAdditionalIntRate.text <> "" And txtAmtCrgIntOn.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txtAdditionalIntRate_LostFocus()

If txtAdditionalIntRate.text <> "" Then
    If NumberCheck2(txtAdditionalIntRate.text) = False Then
        txtAdditionalIntRate.text = ""
    Else
        txtAdditionalIntRate.text = Round(CDbl(txtAdditionalIntRate.text), 2)
        If txtIntPayableAfterDays.text <> "" And txtAdditionalIntRate.text <> "" And txtAmtCrgIntOn.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txtAmtCrgIntOn_LostFocus()

If txtAmtCrgIntOn.text <> "" Then
    If NumberCheck2(txtAmtCrgIntOn.text) = False Then
        txtAmtCrgIntOn.text = ""
    Else
        If txtIntPayableAfterDays.text <> "" And txtAdditionalIntRate.text <> "" And txtAmtCrgIntOn.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txtBreakDate_LostFocus()

'If txtBreakDate.text <> "" Then If CheckDate(txtBreakDate.text) = False Then txtBreakDate.text = ""
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txtBreakDate.text <> "" Then TextBoxFormatDate txtBreakDate

End Sub

Private Sub txtRentReviewDt_LostFocus()

If txtRentReviewDt.text <> "" Then If CheckDate(txtRentReviewDt.text) = False Then txtRentReviewDt.text = ""

End Sub

Private Sub txtRentIncDt_LostFocus()

If txtRentIncDt.text <> "" Then If CheckDate(txtRentIncDt.text) = False Then txtRentIncDt.text = ""

End Sub

Private Sub txtRentIncAmt_LostFocus()

If txtRentIncAmt.text <> "" Then
    If NumberCheck2(txtRentIncAmt.text) = False Then
        txtRentIncAmt.text = ""
    Else
        txtRentIncAmt.text = Round(CDbl(txtRentIncAmt.text), 2)
    End If
End If

End Sub

Private Sub txtDtFlgDate_LostFocus()

If txtDtFlgDate.text <> "" Then If CheckDate(txtDtFlgDate.text) = False Then txtDtFlgDate.text = ""

End Sub

Private Sub txtAmount_GotFocus()
   If Not optPercentage.Value And Not optSqFoot.Value And Not optFixedTotal.Value And Not optGlobalData.Value Then
      MsgBox "Please choose any of the above charge option.", vbInformation + vbOKOnly, "Service Charge Option"
   End If
   txtAmount.SelStart = 0
   txtAmount.SelLength = Len(txtAmount.text)
End Sub

Private Sub txtLeaseStDt_LostFocus()
' Modified by Samrat 08/02/2006
If txtLeaseStDt.text <> "" Then TextBoxFormatDate txtLeaseStDt
End Sub

Private Sub txtLeaseEndDate_LostFocus()
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txtLeaseEndDate.text <> "" Then TextBoxFormatDate txtLeaseEndDate
End Sub

Private Sub txtYearEnd_LostFocus()
' Added by Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txtYearEnd.text <> "" Then TextBoxFormatDate txtYearEnd
End Sub

Private Sub txtRentStartDate_LostFocus()
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txtRentStartDate.text <> "" Then TextBoxFormatDate txtRentStartDate
End Sub

Private Sub txtTotalRentYear_LostFocus()

If txtTotalRentYear.text <> "" Then
    If NumberCheck2(txtTotalRentYear.text) = False Then
        txtTotalRentYear.text = ""
        Exit Sub
    Else
        txtTotalRentYear.text = Round(CDbl(txtTotalRentYear.text), 2)
    End If
    If txtRentStartDate.text <> "" Then Call CalculateBR
End If

End Sub

Public Sub CalculateBR()
Dim xy As Integer
Dim xm As Integer
Dim xd As Integer
Dim qy1 As Integer
Dim qy2 As Integer
Dim qy3 As Integer
Dim qy4 As Integer
Dim qm1 As Integer
Dim qm2 As Integer
Dim qm3 As Integer
Dim qm4 As Integer
Dim qd1 As Integer
Dim qd2 As Integer
Dim qd3 As Integer
Dim qd4 As Integer
Dim hy1 As Integer
Dim hy2 As Integer
Dim hm1 As Integer
Dim hm2 As Integer
Dim hd1 As Integer
Dim hd2 As Integer
Dim yy As Integer
Dim ym As Integer
Dim yd As Integer
Dim b As Integer
' xm = base rent start month; xd = base rent start day
' qmi = quarterly payemnt months; qdi = quarterly payment day
' hmi = half yearly payment months; hdi = half yearly payment day
' ym = yearly payment month; yd = yearly payment day
xy = CInt(Right(txtRentStartDate.text, 2))
xm = CInt(Mid(txtRentStartDate.text, 4, 2))
xd = CInt(Left(txtRentStartDate.text, 2))
qy1 = Right(Year(Date), 2)
qy2 = Right(Year(Date), 2)
qy3 = Right(Year(Date), 2)
qy4 = Right(Year(Date), 2)
qm1 = CInt(Mid(quarterly1, 4, 2))
qm2 = CInt(Mid(quarterly2, 4, 2))
qm3 = CInt(Mid(quarterly3, 4, 2))
qm4 = CInt(Mid(quarterly4, 4, 2))
qd1 = CInt(Left(quarterly1, 2))
qd2 = CInt(Left(quarterly2, 2))
qd3 = CInt(Left(quarterly3, 2))
qd4 = CInt(Left(quarterly4, 2))
hy1 = Right(Year(Date), 2)
hy2 = Right(Year(Date), 2)
hm1 = CInt(Mid(halfyearly1, 4, 2))
hm2 = CInt(Mid(halfyearly2, 4, 2))
hd1 = CInt(Left(halfyearly1, 2))
hd2 = CInt(Left(halfyearly2, 2))
yy = Right(Year(Date), 2)
ym = CInt(Mid(yearly, 4, 2))
yd = CInt(Left(yearly, 2))

'Have to ensure that if the frequency is changed i.e. what calls CalculateBR then the base rate is changed. Txt5 or other will be based on txtNextDueDate. txtNextDueDate has to change
'KOut next 2

'txtNextDueDate.Enabled = True
'txtRentDueEachPeriod.Enabled = True


'In conditional below remmed all set equal to txtRentStartDate, 0, 2, 4, 10. Undone now

'Added And txtNextDueDate.Text = "", to the conditional below not as of 22:11 20/11/2002

If txtTotalRentYear.text <> "" Then 'If there is a Base Rate for Year Figure
    Select Case cboFreqBR.text
        Case cboFreqBR.List(0): ' Weekly in advance

            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 52), 2)

            'Line below, why is it there, remmed 10:42, clearly unremmed

            txtNextDueDate.text = txtRentStartDate.text

        Case cboFreqBR.List(1): ' Weekly in arrears
            
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 52), 2)
            txtNextDueDate.text = DateAdd("d", 7, txtRentStartDate.text)
        Case cboFreqBR.List(2): ' Fortnightly in advance
        
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 26), 2)
            txtNextDueDate.text = txtRentStartDate.text
        Case cboFreqBR.List(3): ' Fortnightly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 26), 2)
            txtNextDueDate.text = DateAdd("d", 7, txtRentStartDate.text)
        Case cboFreqBR.List(4): ' Monthly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 12), 2)
            txtNextDueDate.text = txtRentStartDate.text
        Case cboFreqBR.List(5): ' Monthly in arrears
        
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 12), 2)
            txtNextDueDate.text = DateAdd("m", 1, txtRentStartDate.text)
            
        Case cboFreqBR.List(6): ' Quarterly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 4), 2)
            Select Case xm
                Case Is < qm2:
                    Select Case xm
                        Case Is < qm1:
                            txtNextDueDate.text = quarterly1
                        Case qm1:
                            Select Case xd
                                Case Is < qd1:
                                    txtNextDueDate.text = quarterly1
                                Case qd1:
                                    txtNextDueDate.text = quarterly1
                                Case Is > qd1:
                                    txtNextDueDate.text = quarterly2
                            End Select
                        Case Is > qm1:
                            txtNextDueDate.text = quarterly2
                    End Select
                Case qm2:
                    Select Case xd
                        Case Is < qd2:
                            txtNextDueDate.text = quarterly2
                        Case qd2:
                            txtNextDueDate.text = quarterly2
                        Case Is > qd2:
                            txtNextDueDate.text = quarterly3
                    End Select
                Case Is > qm2:
                    Select Case xm
                        Case Is < qm3:
                            txtNextDueDate.text = quarterly3
                        Case qm3:
                            Select Case xd
                                Case Is < qd3:
                                    txtNextDueDate.text = quarterly3
                                Case qd3:
                                    txtNextDueDate.text = quarterly3
                                Case Is > qd3:
                                    txtNextDueDate.text = quarterly4
                            End Select
                        Case Is > qm3:
                            Select Case xm
                                Case Is < qm4:
                                    txtNextDueDate.text = quarterly4
                                Case qm4:
                                    Select Case xd
                                        Case Is < qd4:
                                            txtNextDueDate.text = quarterly4
                                        Case qd4:
                                            txtNextDueDate.text = quarterly4
                                        Case Is > qd4:
                                            txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                    End Select
                                Case Is > qm4:
                                   txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                            End Select
                    End Select
            End Select
        Case cboFreqBR.List(7): ' Quarterly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 4), 2)
            Select Case xm
                Case Is < qm2:
                    Select Case xm
                        Case Is < qm1:
                            txtNextDueDate.text = quarterly2
                        Case qm1:
                            Select Case xd
                                Case Is < qd1:
                                    txtNextDueDate.text = quarterly2
                                Case qd1:
                                    txtNextDueDate.text = quarterly2
                                Case Is > qd1:
                                    txtNextDueDate.text = quarterly3
                            End Select
                        Case Is > qm1:
                            txtNextDueDate.text = quarterly3
                    End Select
                Case qm2:
                    Select Case xd:
                        Case Is < qd2:
                            txtNextDueDate.text = quarterly3
                        Case qd2:
                            txtNextDueDate.text = quarterly3
                        Case Is > qd2:
                            txtNextDueDate.text = quarterly4
                    End Select
                Case Is > qm2:
                        Select Case xm
                            Case Is < qm3:
                                txtNextDueDate.text = quarterly4
                            Case qm3:
                                Select Case xd
                                    Case Is < qd3:
                                        txtNextDueDate.text = quarterly4
                                    Case qd3:
                                        txtNextDueDate.text = quarterly4
                                    Case Is > qd3:
                                        'quarterly1 (next year)
                                        txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                End Select
                            Case Is > qm3:
                                Select Case xm
                                    Case Is < qm4:
                                        txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                    Case qm4:
                                        Select Case xd
                                            Case Is < qd4:
                                                txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                            Case qd4:
                                                txtNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                            Case Is > qd4:
                                                txtNextDueDate.text = DateAdd("yyyy", 1, quarterly2)
                                        End Select
                                    Case Is > qm4:
                                        txtNextDueDate.text = DateAdd("yyyy", 1, quarterly2)
                                End Select
                        End Select
            End Select
        Case cboFreqBR.List(8): ' Half yearly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 2), 2)
            Select Case xm
                Case Is < hm1:
                    txtNextDueDate.text = halfyearly1
                Case hm1:
                    Select Case xd
                        Case Is < hd1:
                            txtNextDueDate.text = halfyearly1
                        Case hd1:
                            txtNextDueDate.text = halfyearly1
                        Case Is > hd1:
                            txtNextDueDate.text = halfyearly2
                    End Select
                Case Is > hm1:
                    Select Case xm:
                        Case Is < hm2:
                            txtNextDueDate.text = halfyearly2
                        Case hm2:
                            Select Case xd
                                Case Is < hd2:
                                    txtNextDueDate.text = halfyearly2
                                Case hd2:
                                    txtNextDueDate.text = halfyearly2
                                Case Is > hd2:
                                    txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm2:
                            txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
            End Select
        Case cboFreqBR.List(9): ' Half yearly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 2), 2)
            Select Case xm
                Case Is < hm2:
                    Select Case xm
                        Case Is < hm1:
                            txtNextDueDate.text = halfyearly2
                        Case hm1:
                            Select Case xd
                                Case Is < hd1:
                                    txtNextDueDate.text = halfyearly2
                                Case hd1:
                                    txtNextDueDate.text = halfyearly2
                                Case Is > hd1:
                                    txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm1:
                            txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
                Case hm2:
                    Select Case xd
                        Case Is < hd2:
                            txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                        Case hd2:
                            txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                        Case Is > hd2:
                            txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly2)
                    End Select
                Case Is > hm2:
                    txtNextDueDate.text = DateAdd("yyyy", 1, halfyearly2)
            End Select
        Case cboFreqBR.List(10): ' Yearly in advance
            txtRentDueEachPeriod.text = Round(CDbl(txtTotalRentYear.text), 2)
            'remmed below at 13:19
            txtNextDueDate.text = txtRentStartDate.text

'            Select Case xm
 '               Case Is < ym:
  '                  txtNextDueDate.Text = yearly
   '             Case ym:
    '                Select Case xd
     '                   Case Is < yd:
      '                      txtNextDueDate.Text = yearly
       '                 Case yd:
        '                    txtNextDueDate.Text = yearly
         '               Case Is > yd:
          '                  txtNextDueDate.Text = DateAdd("yyyy", 1, yearly)
           '         End Select
            '    Case Is > ym:
             '       txtNextDueDate.Text = DateAdd("yyyy", 1, yearly)
            'End Select
        Case cboFreqBR.List(11): ' Yearly in arrears
            txtRentDueEachPeriod.text = Round(CDbl(txtTotalRentYear.text), 2)
            txtNextDueDate.text = DateAdd("yyyy", 1, txtRentStartDate.text)
        '    Select Case xm
         '       Case Is < ym:
          '          txtNextDueDate.Text = DateAdd("yyyy", 1, yearly)
           '     Case ym:
            '        Select Case xd
             '           Case Is < yd:
              '              txtNextDueDate.Text = DateAdd("yyyy", 1, yearly)
               '         Case yd:
                '            txtNextDueDate.Text = DateAdd("yyyy", 1, yearly)
                 '       Case Is > yd:
                  '          txtNextDueDate.Text = DateAdd("yyyy", 2, yearly)
                   ' End Select
                'Case Is > ym:
                 '   txtNextDueDate.Text = DateAdd("yyyy", 2, yearly)
            
            'End Select
    
    End Select





'kc remmed next 7 lines 11:27 9/10/2002 and moved it above the End If, as in SC
'why do we want to do this?? Also think they should be part of the primary conditional
'Undone as of 22:11 20/11/2002

If txtNextDueDate.text <> "" And txtRentStartDate.text <> "" Then
    b = DateDiff("yyyy", txtNextDueDate.text, txtRentStartDate.text)
    'txtRentReviewDt.Text = b
    
    If b <> 0 Then
        txtNextDueDate.text = DateAdd("yyyy", b, txtNextDueDate.text)
    End If
End If



End If



txtNextDueDate.Enabled = False
txtRentDueEachPeriod.Enabled = False

End Sub

Public Sub SetNextDueDtSC()

Dim xy As Integer
Dim xm  As Integer
Dim xd As Integer
Dim qy1 As Integer
Dim qy2 As Integer
Dim qy3 As Integer
Dim qy4 As Integer
Dim qm1 As Integer
Dim qm2 As Integer
Dim qm3 As Integer
Dim qm4 As Integer
Dim qd1 As Integer
Dim qd2 As Integer
Dim qd3 As Integer
Dim qd4 As Integer
Dim hy1 As Integer
Dim hy2 As Integer
Dim hm1 As Integer
Dim hm2 As Integer
Dim hd1 As Integer
Dim hd2 As Integer
Dim yy As Integer
Dim ym As Integer
Dim yd As Integer
Dim b As Integer

xy = CInt(Right(txtPayableFrom.text, 2))
xm = CInt(Mid(txtPayableFrom.text, 4, 2))
xd = CInt(Left(txtPayableFrom.text, 2))
qy1 = Right(Year(Date), 2)
qy2 = Right(Year(Date), 2)
qy3 = Right(Year(Date), 2)
qy4 = Right(Year(Date), 2)
qm1 = CInt(Mid(quarterly1, 4, 2))
qm2 = CInt(Mid(quarterly2, 4, 2))
qm3 = CInt(Mid(quarterly3, 4, 2))
qm4 = CInt(Mid(quarterly4, 4, 2))
qd1 = CInt(Left(quarterly1, 2))
qd2 = CInt(Left(quarterly2, 2))
qd3 = CInt(Left(quarterly3, 2))
qd4 = CInt(Left(quarterly4, 2))
hy1 = Right(Year(Date), 2)
hy2 = Right(Year(Date), 2)
hm1 = CInt(Mid(halfyearly1, 4, 2))
hm2 = CInt(Mid(halfyearly2, 4, 2))
hd1 = CInt(Left(halfyearly1, 2))
hd2 = CInt(Left(halfyearly2, 2))
yy = Right(Year(Date), 2)
ym = CInt(Mid(yearly, 4, 2))
yd = CInt(Left(yearly, 2))

txtSCNextDueDt.Enabled = True

If cboUnit.text = "" Then
    MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
    Exit Sub
End If

Select Case cboFreqSC.text
    Case cboFreqSC.List(0): 'Weekly in advance
        txtSCNextDueDt.text = txtPayableFrom.text
    Case cboFreqSC.List(1): 'Weekly in arrears
        txtSCNextDueDt.text = DateAdd("d", 7, txtPayableFrom.text)
    Case cboFreqSC.List(2): 'Fortnightly in advance
        txtSCNextDueDt.text = txtPayableFrom.text
    Case cboFreqSC.List(3): 'Fortnightly in arrears
        txtSCNextDueDt.text = DateAdd("d", 14, txtPayableFrom.text)
    Case cboFreqSC.List(4): 'Monthly in advance
        txtSCNextDueDt.text = txtPayableFrom.text
    Case cboFreqSC.List(5): 'Monthly in arrears
        txtSCNextDueDt.text = DateAdd("m", 1, txtPayableFrom.text)
    Case cboFreqSC.List(6): 'Quarterly in advance
        Select Case xm
            Case Is < qm2:
                Select Case xm
                    Case Is < qm1:
                        txtSCNextDueDt.text = quarterly1
                    Case qm1:
                        Select Case xd
                            Case Is < qd1:
                                txtSCNextDueDt.text = quarterly1
                            Case qd1:
                                txtSCNextDueDt.text = quarterly1
                            Case Is > qd1:
                                txtSCNextDueDt.text = quarterly2
                        End Select
                    Case Is > qm1:
                        txtSCNextDueDt.text = quarterly2
                End Select
            Case qm2:
                Select Case xd
                    Case Is < qd2:
                        txtSCNextDueDt.text = quarterly2
                    Case qd2:
                        txtSCNextDueDt.text = quarterly2
                    Case Is > qd2:
                        txtSCNextDueDt.text = quarterly3
                End Select
            Case Is > qm2:
                Select Case xm
                    Case Is < qm3:
                        txtSCNextDueDt.text = quarterly3
                    Case qm3:
                        Select Case xd
                            Case Is < qd3:
                                txtSCNextDueDt.text = quarterly3
                            Case qd3:
                                txtSCNextDueDt.text = quarterly3
                            Case Is > qd3:
                                txtSCNextDueDt.text = quarterly4
                        End Select
                    Case Is > qm3:
                        Select Case xm
                            Case Is < qm4:
                                txtSCNextDueDt.text = quarterly4 'qm4
                            Case qm4:
                                Select Case xd
                                    Case Is < qd4:
                                        txtSCNextDueDt.text = quarterly4
                                    Case qd4:
                                        txtSCNextDueDt.text = quarterly4
                                    Case Is > qd4:
                                        txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                                End Select
                            Case Is > qm4:
                               txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                        End Select
                End Select
            End Select
    Case cboFreqSC.List(7): 'Quarterly in arrears
        Select Case xm
            Case Is < qm2:
                Select Case xm
                    Case Is < qm1:
                        txtSCNextDueDt.text = quarterly2
                    Case qm1:
                        Select Case xd
                            Case Is < qd1:
                                txtSCNextDueDt.text = quarterly2
                            Case qd1:
                                txtSCNextDueDt.text = quarterly2
                            Case Is > qd1:
                                txtSCNextDueDt.text = quarterly3
                        End Select
                    Case Is > qm1:
                        txtSCNextDueDt.text = quarterly3
                End Select
            Case qm2:
                Select Case xd:
                    Case Is < qd2:
                        txtSCNextDueDt.text = quarterly3
                    Case qd2:
                        txtSCNextDueDt.text = quarterly3
                    Case Is > qd2:
                        txtSCNextDueDt.text = quarterly4
                End Select
            Case Is > qm2:
                    Select Case xm
                        Case Is < qm3:
                            txtSCNextDueDt.text = quarterly4
                        Case qm3:
                            Select Case xd
                                Case Is < qd3:
                                    txtSCNextDueDt.text = quarterly4
                                Case qd3:
                                    txtSCNextDueDt.text = quarterly4
                                Case Is > qd3:
                                    'quarterly1 (next year)
                                    txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                            End Select
                        Case Is > qm3:
                            Select Case xm
                                Case Is < qm4:
                                    txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                                Case qm4:
                                    Select Case xd
                                        Case Is < qd4:
                                            txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                                        Case qd4:
                                            txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly1)
                                        Case Is > qd4:
                                            txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly2)
                                    End Select
                                Case Is > qm4:
                                    txtSCNextDueDt.text = DateAdd("yyyy", 1, quarterly2)
                            End Select
                    End Select
        End Select
    Case cboFreqSC.List(8): 'Half yearly in advance
        Select Case xm
                Case Is < hm1:
                    txtSCNextDueDt.text = halfyearly1
                Case hm1:
                    Select Case xd
                        Case Is < hd1:
                            txtSCNextDueDt.text = halfyearly1
                        Case hd1:
                            txtSCNextDueDt.text = halfyearly1
                        Case Is > hd1:
                            txtSCNextDueDt.text = halfyearly2
                    End Select
                Case Is > hm1:
                    Select Case xm:
                        Case Is < hm2:
                            txtSCNextDueDt.text = halfyearly2
                        Case hm2:
                            Select Case xd
                                Case Is < hd2:
                                    txtSCNextDueDt.text = halfyearly2
                                Case hd2:
                                    txtSCNextDueDt.text = halfyearly2
                                Case Is > hd2:
                                    txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm2:
                            txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
        End Select
    Case cboFreqSC.List(9): 'Half yearly in arrears
        Select Case xm
                Case Is < hm2:
                    Select Case xm
                        Case Is < hm1:
                            txtSCNextDueDt.text = halfyearly2
                        Case hm1:
                            Select Case xd
                                Case Is < hd1:
                                    txtSCNextDueDt.text = halfyearly2
                                Case hd1:
                                    txtSCNextDueDt.text = halfyearly2
                                Case Is > hd1:
                                    txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm1:
                            txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
                Case hm2:
                    Select Case xd
                        Case Is < hd2:
                            txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                        Case hd2:
                            txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly1)
                        Case Is > hd2:
                            txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly2)
                    End Select
                Case Is > hm2:
                    txtSCNextDueDt.text = DateAdd("yyyy", 1, halfyearly2)
            End Select
    Case cboFreqSC.List(10): ' yearly in advance
        txtSCNextDueDt.text = txtPayableFrom.text
    Case cboFreqSC.List(11): ' yearly in arrears
        txtSCNextDueDt.text = DateAdd("yyyy", 1, txtPayableFrom.text)
End Select

If txtSCNextDueDt.text <> "" And txtPayableFrom.text <> "" Then
    b = DateDiff("yyyy", txtSCNextDueDt.text, txtPayableFrom.text)
    If b <> 0 Then
        txtSCNextDueDt.text = DateAdd("yyyy", b, txtSCNextDueDt.text)
    End If
End If

txtSCNextDueDt.Enabled = False
End Sub

Public Sub CalculateInsuranceCharge()
   If cboUnit.text = "" Then
       MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
       Exit Sub
   End If

   Dim xy As Integer
   Dim xm  As Integer
   Dim xd As Integer
   Dim qy1 As Integer
   Dim qy2 As Integer
   Dim qy3 As Integer
   Dim qy4 As Integer
   Dim qm1 As Integer
   Dim qm2 As Integer
   Dim qm3 As Integer
   Dim qm4 As Integer
   Dim qd1 As Integer
   Dim qd2 As Integer
   Dim qd3 As Integer
   Dim qd4 As Integer
   Dim hy1 As Integer
   Dim hy2 As Integer
   Dim hm1 As Integer
   Dim hm2 As Integer
   Dim hd1 As Integer
   Dim hd2 As Integer
   Dim yy As Integer
   Dim ym As Integer
   Dim yd As Integer
   Dim b As Integer

   xy = CInt(Right(txtInsuranceStartDate.text, 2))
   xm = Format(txtInsuranceStartDate.text, "mm")
   xd = Format(txtInsuranceStartDate.text, "dd")
   qy1 = Right(Year(Date), 2)
   qy2 = Right(Year(Date), 2)
   qy3 = Right(Year(Date), 2)
   qy4 = Right(Year(Date), 2)
   qm1 = CInt(Mid(quarterly1, 4, 2))
   qm2 = CInt(Mid(quarterly2, 4, 2))
   qm3 = CInt(Mid(quarterly3, 4, 2))
   qm4 = CInt(Mid(quarterly4, 4, 2))
   qd1 = CInt(Left(quarterly1, 2))
   qd2 = CInt(Left(quarterly2, 2))
   qd3 = CInt(Left(quarterly3, 2))
   qd4 = CInt(Left(quarterly4, 2))
   hy1 = Right(Year(Date), 2)
   hy2 = Right(Year(Date), 2)
   hm1 = CInt(Mid(halfyearly1, 4, 2))
   hm2 = CInt(Mid(halfyearly2, 4, 2))
   hd1 = CInt(Left(halfyearly1, 2))
   hd2 = CInt(Left(halfyearly2, 2))
   yy = Right(Year(Date), 2)
   ym = CInt(Mid(yearly, 4, 2))
   yd = CInt(Left(yearly, 2))

   txtInsuranceNextDueDate.Enabled = True

   Select Case cboInsuranceFrequency.text
       Case cboInsuranceFrequency.List(0): 'Weekly in advance
           txtInsuranceNextDueDate.text = txtInsuranceStartDate.text
       Case cboInsuranceFrequency.List(1): 'Weekly in arrears
           txtInsuranceNextDueDate.text = DateAdd("d", 7, txtInsuranceStartDate.text)
       Case cboInsuranceFrequency.List(2): 'Fortnightly in advance
           txtInsuranceNextDueDate.text = txtInsuranceStartDate.text
       Case cboInsuranceFrequency.List(3): 'Fortnightly in arrears
           txtInsuranceNextDueDate.text = DateAdd("d", 14, txtInsuranceStartDate.text)
       Case cboInsuranceFrequency.List(4): 'Monthly in advance
           txtInsuranceNextDueDate.text = txtInsuranceStartDate.text
       Case cboInsuranceFrequency.List(5): 'Monthly in arrears
           txtInsuranceNextDueDate.text = DateAdd("m", 1, txtInsuranceStartDate.text)
       Case cboInsuranceFrequency.List(6): 'Quarterly in advance
           Select Case xm
               Case Is < qm2:
                   Select Case xm
                       Case Is < qm1:
                           txtInsuranceNextDueDate.text = quarterly1
                       Case qm1:
                           Select Case xd
                               Case Is < qd1:
                                   txtInsuranceNextDueDate.text = quarterly1
                               Case qd1:
                                   txtInsuranceNextDueDate.text = quarterly1
                               Case Is > qd1:
                                   txtInsuranceNextDueDate.text = quarterly2
                           End Select
                       Case Is > qm1:
                           txtInsuranceNextDueDate.text = quarterly2
                   End Select
               Case qm2:
                   Select Case xd
                       Case Is < qd2:
                           txtInsuranceNextDueDate.text = quarterly2
                       Case qd2:
                           txtInsuranceNextDueDate.text = quarterly2
                       Case Is > qd2:
                           txtInsuranceNextDueDate.text = quarterly3
                   End Select
               Case Is > qm2:
                   Select Case xm
                       Case Is < qm3:
                           txtInsuranceNextDueDate.text = quarterly3
                       Case qm3:
                           Select Case xd
                               Case Is < qd3:
                                   txtInsuranceNextDueDate.text = quarterly3
                               Case qd3:
                                   txtInsuranceNextDueDate.text = quarterly3
                               Case Is > qd3:
                                   txtInsuranceNextDueDate.text = quarterly4
                           End Select
                       Case Is > qm3:
                           Select Case xm
                               Case Is < qm4:
                                   txtInsuranceNextDueDate.text = quarterly4 'qm4
                               Case qm4:
                                   Select Case xd
                                       Case Is < qd4:
                                           txtInsuranceNextDueDate.text = quarterly4
                                       Case qd4:
                                           txtInsuranceNextDueDate.text = quarterly4
                                       Case Is > qd4:
                                           txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                   End Select
                               Case Is > qm4:
                                  txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                           End Select
                   End Select
               End Select
       Case cboInsuranceFrequency.List(7): 'Quarterly in arrears
           Select Case xm
               Case Is < qm2:
                   Select Case xm
                       Case Is < qm1:
                           txtInsuranceNextDueDate.text = quarterly2
                       Case qm1:
                           Select Case xd
                               Case Is < qd1:
                                   txtInsuranceNextDueDate.text = quarterly2
                               Case qd1:
                                   txtInsuranceNextDueDate.text = quarterly2
                               Case Is > qd1:
                                   txtInsuranceNextDueDate.text = quarterly3
                           End Select
                       Case Is > qm1:
                           txtInsuranceNextDueDate.text = quarterly3
                   End Select
               Case qm2:
                   Select Case xd:
                       Case Is < qd2:
                           txtInsuranceNextDueDate.text = quarterly3
                       Case qd2:
                           txtInsuranceNextDueDate.text = quarterly3
                       Case Is > qd2:
                           txtInsuranceNextDueDate.text = quarterly4
                   End Select
               Case Is > qm2:
                       Select Case xm
                           Case Is < qm3:
                               txtInsuranceNextDueDate.text = quarterly4
                           Case qm3:
                               Select Case xd
                                   Case Is < qd3:
                                       txtInsuranceNextDueDate.text = quarterly4
                                   Case qd3:
                                       txtInsuranceNextDueDate.text = quarterly4
                                   Case Is > qd3:
                                       'quarterly1 (next year)
                                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                               End Select
                           Case Is > qm3:
                               Select Case xm
                                   Case Is < qm4:
                                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                   Case qm4:
                                       Select Case xd
                                           Case Is < qd4:
                                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                           Case qd4:
                                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly1)
                                           Case Is > qd4:
                                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly2)
                                       End Select
                                   Case Is > qm4:
                                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, quarterly2)
                               End Select
                       End Select
           End Select
       Case cboInsuranceFrequency.List(8): 'Half yearly in advance
           Select Case xm
                   Case Is < hm1:
                       txtInsuranceNextDueDate.text = halfyearly1
                   Case hm1:
                       Select Case xd
                           Case Is < hd1:
                               txtInsuranceNextDueDate.text = halfyearly1
                           Case hd1:
                               txtInsuranceNextDueDate.text = halfyearly1
                           Case Is > hd1:
                               txtInsuranceNextDueDate.text = halfyearly2
                       End Select
                   Case Is > hm1:
                       Select Case xm:
                           Case Is < hm2:
                               txtInsuranceNextDueDate.text = halfyearly2
                           Case hm2:
                               Select Case xd
                                   Case Is < hd2:
                                       txtInsuranceNextDueDate.text = halfyearly2
                                   Case hd2:
                                       txtInsuranceNextDueDate.text = halfyearly2
                                   Case Is > hd2:
                                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                               End Select
                           Case Is > hm2:
                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                       End Select
           End Select
       Case cboInsuranceFrequency.List(9): 'Half yearly in arrears
           Select Case xm
                   Case Is < hm2:
                       Select Case xm
                           Case Is < hm1:
                               txtInsuranceNextDueDate.text = halfyearly2
                           Case hm1:
                               Select Case xd
                                   Case Is < hd1:
                                       txtInsuranceNextDueDate.text = halfyearly2
                                   Case hd1:
                                       txtInsuranceNextDueDate.text = halfyearly2
                                   Case Is > hd1:
                                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                               End Select
                           Case Is > hm1:
                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                       End Select
                   Case hm2:
                       Select Case xd
                           Case Is < hd2:
                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                           Case hd2:
                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly1)
                           Case Is > hd2:
                               txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly2)
                       End Select
                   Case Is > hm2:
                       txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, halfyearly2)
               End Select
       Case cboInsuranceFrequency.List(10): ' yearly in advance
           txtInsuranceNextDueDate.text = txtInsuranceStartDate.text
       Case cboInsuranceFrequency.List(11): ' yearly in arrears
           txtInsuranceNextDueDate.text = DateAdd("yyyy", 1, txtInsuranceStartDate.text)
   End Select

   If txtInsuranceNextDueDate.text <> "" And txtInsuranceStartDate.text <> "" Then
       b = DateDiff("yyyy", txtInsuranceNextDueDate.text, txtInsuranceStartDate.text)
       If b <> 0 Then
           txtInsuranceNextDueDate.text = DateAdd("yyyy", b, txtInsuranceNextDueDate.text)
       End If
   End If

   txtInsuranceNextDueDate.Enabled = False
End Sub


Public Sub CalculateInterest()

Dim a As Double
Dim r As Double
Dim d As Integer

a = CDbl(txtAmtCrgIntOn.text)
r = (BaseRate + CDbl(txtAdditionalIntRate.text)) / 100
d = CInt(txtIntPayableAfterDays.text)

txtInt2bChrg.text = Round(a * r * d / 365, 2)

End Sub



Private Sub txtRentIncreaseDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtRentIncreaseDate
End Sub

Private Sub txtRentIncreaseDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtRentIncreaseDate, KeyAscii
End Sub

Private Sub txtRentIncreaseDate_LostFocus()
'Added By Asif. 13/01/2006
'Modified By Samrat. 06/02/2006
If Not txtRentIncreaseDate.Locked Then TextBoxFormatDate txtRentIncreaseDate
End Sub

Private Sub txtRentReviewDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtRentReviewDate
End Sub

Private Sub txtRentReviewDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtRentReviewDate, KeyAscii
End Sub

Private Sub txtRentReviewDate_LostFocus()
'Added By Asif. 13/01/2006
'Modified By Samrat. 06/02/2006
If Not txtRentReviewDate.Locked Then TextBoxFormatDate txtRentReviewDate
End Sub

Private Sub txtSCPercentage_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtSCPercentage_LostFocus()
   If txtSCPercentage.Locked Then Exit Sub
   If optPercentage.Value = 0 Then
      txtSCPercentage.text = ""
      Exit Sub
   End If

   Dim Total As Double, TotalServiceCharge As String

   On Error GoTo ErrorHander

   MousePointer = vbHourglass

   txtSCPercentage.text = Format(IIf(txtSCPercentage.text = "", 0, txtSCPercentage.text), "0.0000")

   ' The following code is to calculate SC payable according to percentage
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT GlobalData.TotalSC " & _
             "FROM GlobalData,Units " & _
             "WHERE Units.PropertyID = GlobalData.PropertyID " & _
               "AND Units.UnitNumber = '" & Left(cboUnit, 8) & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   TotalServiceCharge = Rst1!TotalSC

   Rst1.Close

   Total = CDbl(TotalServiceCharge) * (CDbl(txtSCPercentage.text) / 100)
   txtAmount.text = Format(Total, "0.00")

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   txtFinalAmout.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")

   Rst1.Close
   Conn1.Close

   If cboSCPayable.text = "Yes" Then Call SetNextDueDtSC

   MousePointer = vbDefault

   Exit Sub
ErrorHander:
   Rst1.Close
   Conn1.Close
   MousePointer = vbDefault
End Sub

Private Sub txtSuppCaption1_LostFocus()
   txtSuppCaption1.Visible = False
   lblSupplementary1.Caption = txtSuppCaption1.text
End Sub

Private Sub txtSuppCaption2_LostFocus()
'txtSuppCaption2.Visible = False
'lblSupplementary2.Caption = txtSuppCaption2.text
End Sub

Private Sub txtSuppCaption3_LostFocus()
   txtSuppCaption3.Visible = False
   lblSupplementary3.Caption = txtSuppCaption3.text
End Sub

Public Function SaveBreaches() As Boolean
    Dim conBreach As New RDO.rdoConnection
    Dim rstBreach As rdoResultset
    Dim sSQLQuery_ As String
    Dim sSQLDelete As String
    Dim sSQLFilter As String
    Dim iRowIndex As Integer

    sSQLFilter = ""

    On Error GoTo Exception
    'Set the RDO Connections to the dataset
    conBreach.Connect = "DSN=" & Adsn & ";UID=;PWD="
    conBreach.CursorDriver = rdUseIfNeeded
    conBreach.EstablishConnection rdDriverNoPrompt


    If Not BREACH_NEW_ENTRY_ Then
        sSQLFilter = "WHERE LeaseId = '" & txtLeaseID.text & "' AND BreachID = " & txtBreachID.text & ""
    Else
        sSQLFilter = ""
    End If

    sSQLQuery_ = "SELECT * " & _
    "FROM LeaseBreaches " & sSQLFilter


    Set rstBreach = conBreach.OpenResultset(sSQLQuery_, rdOpenDynamic, rdConcurRowVer)

    'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
    If BREACH_NEW_ENTRY_ Then
        rstBreach.AddNew
    Else
        rstBreach.Edit
    End If

    rstBreach!LeaseId = txtLeaseID.text
    rstBreach!BreachType = cboBreachType.BoundText
    rstBreach!CommenceDate = IIf(txtCommenceDate.text = "", Null, txtCommenceDate.text)
    rstBreach!InitiatedBy = txtInitiatedBy.text
    If chkResolved.Value = 1 Then
        rstBreach!Resolved = True
    Else
        rstBreach!Resolved = False
    End If
    rstBreach!DateReceived = IIf(txtDateReceived.text = "", Null, txtDateReceived.text)
    rstBreach!ReceivedBy = txtReceivedBy.text
    rstBreach.Update

    'Next iRowIndex


    rstBreach.Close
    conBreach.Close
    Set rstBreach = Nothing
    Set conBreach = Nothing
    SaveBreaches = True
    PopulateBreaches
    Exit Function
    
Exception:
    
    MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
    SaveBreaches = False
End Function

Public Function SaveAssignment() As Boolean
    Dim conAssignment As New RDO.rdoConnection
    Dim rstAssignment As rdoResultset
    Dim sSQLQuery_ As String
    Dim sSQLDelete As String
    Dim sSQLFilter As String
    Dim iRowIndex As Integer

    sSQLFilter = ""

    On Error GoTo Exception
    'Set the RDO Connections to the dataset
    conAssignment.Connect = "DSN=" & Adsn & ";UID=;PWD="
    conAssignment.CursorDriver = rdUseIfNeeded
    conAssignment.EstablishConnection rdDriverNoPrompt

    If Not ASSIGNMENT_NEW_ENTRY_ Then
        sSQLFilter = "WHERE LeaseId = '" & txtLeaseID.text & "' AND AssignmentID = " & txtAssignmentID.text & ""
    Else
        sSQLFilter = ""
    End If

    sSQLQuery_ = "SELECT * " & _
    "FROM LeaseAssignments " & sSQLFilter

    Set rstAssignment = conAssignment.OpenResultset(sSQLQuery_, rdOpenDynamic, rdConcurRowVer)

    'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
    If ASSIGNMENT_NEW_ENTRY_ Then
        rstAssignment.AddNew
    Else
        rstAssignment.Edit
    End If

    rstAssignment!LeaseId = txtLeaseID.text
    rstAssignment!AssignDate = txtAssignment_Date.text
    rstAssignment!Assignee = txtAssignee.text
    rstAssignment!Decp = txtDescription.text
    rstAssignment.Update

    rstAssignment.Close
    conAssignment.Close
    Set rstAssignment = Nothing
    Set conAssignment = Nothing
    SaveAssignment = True
    PopulateAssignments
    Exit Function

Exception:
    
    MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
    SaveAssignment = False
End Function

Public Sub BreachButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
        Case ComponentMode.DefaultMode
            cmdBreachNew.Enabled = True
            cmdBreachEdit.Enabled = False
            cmdBreachSave.Enabled = False
            cmdBreachCancel.Enabled = False
            
            gridBreach.Enabled = True
        
            cboBreachType.Enabled = False
            cmdSetBreachType.Enabled = False
            txtCommenceDate.Locked = True
            txtInitiatedBy.Locked = True
            chkResolved.Enabled = False
            txtDateReceived.Locked = True
            txtReceivedBy.Locked = True
        
        Case ComponentMode.GridRowOnSelection
            cmdBreachNew.Enabled = True
            cmdBreachEdit.Enabled = True
            cmdBreachSave.Enabled = False
            cmdBreachCancel.Enabled = False
            
            gridBreach.Enabled = True
        
        Case ComponentMode.NewEntryMode
            cmdBreachNew.Enabled = False
            cmdBreachEdit.Enabled = False
            cmdBreachSave.Enabled = True
            cmdBreachCancel.Enabled = True
            
            gridBreach.Enabled = False
        
            cboBreachType.Enabled = True
            cmdSetBreachType.Enabled = True
            txtCommenceDate.Locked = False
            txtCommenceDate.text = ""
            txtInitiatedBy.Locked = False
            txtInitiatedBy.text = ""
            chkResolved.Enabled = True
            txtDateReceived.Locked = False
            txtDateReceived.text = ""
            txtReceivedBy.Locked = False
            txtReceivedBy.text = ""
                    
        Case ComponentMode.EditMode
            cmdBreachNew.Enabled = False
            cmdBreachEdit.Enabled = False
            cmdBreachSave.Enabled = True
            cmdBreachCancel.Enabled = True
            
            gridBreach.Enabled = False
        
            cboBreachType.Enabled = True
            cmdSetBreachType.Enabled = True
            txtCommenceDate.Locked = False
            txtInitiatedBy.Locked = False
            chkResolved.Enabled = True
            txtDateReceived.Locked = False
            txtReceivedBy.Locked = False
            
    End Select
End Sub

Public Sub AssignmentButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
        Case ComponentMode.DefaultMode
            cmdAssignmentNew.Enabled = True
            cmdAssignmentEdit.Enabled = False
            cmdAssignmentSave.Enabled = False
            cmdAssignmentCancel.Enabled = False
            
            gridAssignment.Enabled = True
        
            txtAssignment_Date.Locked = True
            txtTenant.Locked = True
        
        Case ComponentMode.GridRowOnSelection
            cmdAssignmentNew.Enabled = True
            cmdAssignmentEdit.Enabled = True
            cmdAssignmentSave.Enabled = False
            cmdAssignmentCancel.Enabled = False
            
            gridAssignment.Enabled = True
        
        Case ComponentMode.NewEntryMode
            cmdAssignmentNew.Enabled = False
            cmdAssignmentEdit.Enabled = False
            cmdAssignmentSave.Enabled = True
            cmdAssignmentCancel.Enabled = True

            gridAssignment.Enabled = False

            txtAssignment_Date.Locked = False
            txtAssignment_Date.text = ""
            txtTenant.Locked = False
            txtTenant = ""

        Case ComponentMode.EditMode
            cmdAssignmentNew.Enabled = False
            cmdAssignmentEdit.Enabled = False
            cmdAssignmentSave.Enabled = True
            cmdAssignmentCancel.Enabled = True

            gridAssignment.Enabled = False

            txtAssignment_Date.Locked = False
            txtTenant.Locked = False
    End Select
End Sub

Public Sub RentReviewButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
        Case ComponentMode.DefaultMode
            cmdNewRentAnalysis.Enabled = True
            cmdEditRentAnalysis.Enabled = False
            cmdSageRentAnalysis.Enabled = False
            cmdCancelRentAnalysis.Enabled = False
            
            flxRentAnalysis.Enabled = True
        
'            txtAssignment_Date.Locked = True
'            txtTenant.Locked = True
        
        Case ComponentMode.GridRowOnSelection
            cmdNewRentAnalysis.Enabled = True
            cmdEditRentAnalysis.Enabled = True
            cmdSageRentAnalysis.Enabled = False
            cmdCancelRentAnalysis.Enabled = False
            
            flxRentAnalysis.Enabled = True
        
        Case ComponentMode.NewEntryMode
            cmdNewRentAnalysis.Enabled = False
            cmdEditRentAnalysis.Enabled = False
            cmdSageRentAnalysis.Enabled = True
            cmdCancelRentAnalysis.Enabled = True

            flxRentAnalysis.Enabled = False

'            txtAssignment_Date.Locked = False
'            txtAssignment_Date.text = ""
'            txtTenant.Locked = False
'            txtTenant = ""

        Case ComponentMode.EditMode
            cmdNewRentAnalysis.Enabled = False
            cmdEditRentAnalysis.Enabled = False
            cmdSageRentAnalysis.Enabled = True
            cmdCancelRentAnalysis.Enabled = True

            flxRentAnalysis.Enabled = False

'            txtAssignment_Date.Locked = False
'            txtTenant.Locked = False
    End Select
End Sub

Private Function AllDemandType(ByRef szaDemandtype) As Boolean
   On Error GoTo ErrorHandler
   AllDemandType = True
   If UBound(szaDemandtype) > 0 Then Exit Function
   AllDemandType = True
   Exit Function

ErrorHandler:
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String, i As Integer

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT COUNT(*) AS C_I FROM DemandTypes"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   If rdoRst1!C_I = 0 Then
      AllDemandType = False
      Exit Function
   End If
   ReDim szaDemandtype(rdoRst1!C_I) As String

   rdoRst1.Close

   SQLStr1 = "SELECT ID, Type FROM DemandTypes"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   i = 0
   If rdoRst1.EOF = False Then
       While rdoRst1.EOF = False
          szaDemandtype(i) = rdoRst1!ID & " / " & rdoRst1!Type
          i = i + 1
          rdoRst1.MoveNext
       Wend
   End If

   AllDemandType = True
   rdoRst1.Close
   rdoConn.Close
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Function

Private Sub FillcboType(szaDemandtype() As String)
   Dim SQLStr1 As String, i As Integer

   cboBRDemandType.Clear
   cboSCDemandType.Clear
   cboIntDemandType.Clear
   cboInsuranceDemandType.Clear

   i = 0
   While szaDemandtype(i) <> ""
      cboBRDemandType.AddItem szaDemandtype(i)
      cboSCDemandType.AddItem szaDemandtype(i)
      cboIntDemandType.AddItem szaDemandtype(i)
      cboInsuranceDemandType.AddItem szaDemandtype(i)
      i = i + 1
   Wend
End Sub

Private Sub LoadDept()
   Dim data() As String
   Dim rRow As Integer
   
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   '  Set oSDO = New SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   '  Set oWS = oSDO.Workspaces.Add("WkpsSupplier")
   Dim oDepartmentData As SageDataObject120.DepartmentData

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'   szDataPath = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & Sdsn & "", "DataPathname")
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oDepartmentData = oWS.CreateObject("DepartmentData")
         
         ReDim data(2, oDepartmentData.Count) As String
         
         For rRow = 2 To oDepartmentData.Count
            oDepartmentData.Read (rRow)
            data(0, rRow - 2) = CStr(rRow - 1)
            data(1, rRow - 2) = CStr(oDepartmentData.Fields.Item("NAME").Value)
         Next rRow
         'Disconnect
         oWS.Disconnect
      End If
   End If

   cboRentChargeDept.Clear
   cboRentChargeDept.Column() = data()
   cboServiceChargeDept.Clear
   cboServiceChargeDept.Column() = data()
   cboIntChargeDept.Clear
   cboIntChargeDept.Column() = data()
   cboInsuranceDept.Clear
   cboInsuranceDept.Column() = data()

   ' Destroy Objects
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

      MsgBox "(pcm_009) The SDO generated the following error: " & oSDO.LastError.text
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Public Sub DisableBoxes()
   cmdLease.Enabled = True
   cmdTenants.Enabled = False
   cboTenant.Enabled = False
   cboUnit.Locked = True
   cboUnit.Enabled = True
   txtLeaseID.Enabled = False

   'Lease details
   Frame2.Enabled = False
'   chkSubLease.Enabled = True
'   cboType.Enabled = True
'   cmdLeaseType.Enabled = True
'   txtYearEnd.Enabled = True
'   txtLeaseStDt.Enabled = True
'   txtLeaseEndDate.Enabled = True
   
   'Rent Charges
   cboRentPayable.Enabled = False
   Frame1(1).Enabled = False
'   cboRentChargeDept.Enabled = True
'   cboFreqBR.Enabled = True
'   txtRentStartDate.Enabled = True
'   cboBRDemandType.Enabled = True
'   txtTotalRentYear.Enabled = True
   
   'Service Charges
   Frame1(2).Enabled = False
   cboSCPayable.Enabled = False
'   cboServiceChargeDept.Enabled = True
'   txtPayableFrom.Enabled = True
'   cboFreqSC.Enabled = True
'   txtTOLimit.Enabled = True

   'Interest Charge
   cboIntCrgable.Enabled = False
   Frame3.Enabled = False
'   txtIntPayableAfterDays.Enabled = True
'   txtAdditionalIntRate.Enabled = True
'   txtAmtCrgIntOn.Enabled = True
   
   'Break Clause
   cboBreakClause.Enabled = False
   Frame1(0).Enabled = False
'   cboBreak.Enabled = True
'   txtBreakDate.Enabled = True

   'Rent Review
   txtRentReviewDt.Enabled = False
   txtRentIncDt.Enabled = False
   txtRentIncAmt.Enabled = False
   
   'Insurance
   cboInsurancePayable.Enabled = False
   fmeInsurance.Enabled = False
   
   'Supplementary
   txtDtFlgDate.Enabled = False
   txtDtFlgDesc.Enabled = False
   txtMemo.Enabled = False
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   
   'Breaches
   fraBreaches.Enabled = False
End Sub

Public Sub EnableBoxes()
   cmdLease.Enabled = False
   cmdTenants.Enabled = True
   cboTenant.Enabled = True
   cboUnit.Locked = False
   cboUnit.Enabled = True
   txtLeaseID.Enabled = True

   'Lease details
   '-------------
   Frame2.Enabled = True
   
   'Rent Charges
   '------------
   cboRentPayable.Enabled = True
   Frame1(1).Enabled = True
   
   'Service Charges
   '---------------
   Frame1(2).Enabled = True
   cboSCPayable.Enabled = True

   'Interest Charge
   '---------------
   cboIntCrgable.Enabled = True
   Frame3.Enabled = True
   
   'Break Clause
   '------------
   cboBreakClause.Enabled = True
   Frame1(0).Enabled = True

   'Rent Review
   txtRentReviewDt.Enabled = True
   txtRentIncDt.Enabled = True
   txtRentIncAmt.Enabled = True
   
   'Insurance
   cboInsurancePayable.Enabled = True
   fmeInsurance.Enabled = True
   
   'Supplementary
   txtDtFlgDate.Enabled = True
   txtDtFlgDesc.Enabled = True
   txtMemo.Enabled = True
   Text1.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True

   'Breaches
   fraBreaches.Enabled = True
End Sub

Public Sub EmptyBoxes()
   txtLeaseID.text = ""
   txtTenant.text = ""
   txtUnitName.text = ""
   cboUnit.text = ""
   txtClient.text = ""
   txtProperty.text = ""

   'Lease Details
   cboHeadLease.text = ""
   chkSubLease.Value = 0
   cboType.text = ""
   txtYearEnd.text = ""
   txtLeaseStDt.text = ""
   txtLeaseEndDate.text = ""
   
   'Rent Charges
   cboRentPayable.text = "No"
   cboRentChargeDept.text = ""
   cboFreqBR.text = ""
   txtRentStartDate.text = ""
   cboBRDemandType.text = ""
   txtTotalRentYear.text = ""
   txtRentDueEachPeriod.text = ""
   txtNextDueDate.text = ""
   
   'Rent Review
   txtRentReviewDt.text = ""
   txtRentIncDt.text = ""
   txtRentIncAmt.text = ""
   
   'Breakes
   cboBreakClause.text = "No"
   txtBreakDate.text = ""
   cboBreak.text = ""
   
   'Service Charges
   cboSCPayable.text = "No"
   cboServiceChargeDept.text = ""
   txtPayableFrom.text = ""
   cboFreqSC.text = ""
   txtSCNextDueDt.text = ""
   cboSCDemandType.text = ""
   txtTOLimit.text = ""
   txtSCPercentage.text = ""
   txtPPSqFoot.text = ""
   txtAnnualService.text = ""
   txtGlobalAmount.text = ""
   txtAmount.text = ""
   txtFinalAmout.text = ""
   
   'Interest Charges
   cboIntCrgable.text = "No"
   cboIntChargeDept.text = ""
   txtAdditionalIntRate.text = ""
   txtAmtCrgIntOn.text = ""
   txtInt2bChrg.text = ""
   txtIntPayableAfterDays.text = ""
   cboIntDemandType.text = ""
   
   'Breaches
   cboBreachType.text = ""
   txtCommenceDate.text = ""
   txtInitiatedBy.text = ""
   chkResolved.Value = 0
   txtDateReceived.text = ""
   txtReceivedBy.text = ""
   
   'Assignment
   
   'Insurance
   cboInsurancePayable.text = "No"
   cboInsuranceDept.text = ""
   txtInsuranceStartDate.text = ""
   txtInsuranceEndDate.text = ""
   cboInsuranceFrequency.text = ""
   cboInsuranceDemandType.text = ""
   txtInsuranceNextDueDate.text = ""
   txtInsurancePercentage.text = ""
   txtAnnualInsuranceCharge.text = ""
   txtTotalYearlyInsurance.text = "0.00"
   txtInsuranceEachPeriod.text = "0.00"

   'Supplementary
   txtDtFlgDate.text = ""
   txtDtFlgDesc.text = ""
   Text1.text = ""
   Text2.text = ""
   Text3.text = ""
   txtSuppCaption1.text = ""
   txtSuppCaption2.text = ""
   txtSuppCaption3.text = ""

   'Memo
   txtMemo.text = ""
End Sub
