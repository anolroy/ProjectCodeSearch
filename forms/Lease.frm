VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLease 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lease"
   ClientHeight    =   7020
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   12300
   Icon            =   "Lease.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   12300
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   330
      Left            =   6480
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
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
      Left            =   4920
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
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
   Begin VB.Frame Frame5 
      Caption         =   "Select Tenant "
      Height          =   1425
      Left            =   120
      TabIndex        =   27
      Top             =   60
      Width           =   12075
      Begin VB.TextBox txtLeaseID 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   720
         Width           =   3195
      End
      Begin VB.ComboBox cboTenant1 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cboTenant2 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   7260
         TabIndex        =   171
         Top             =   750
         Width           =   465
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Left            =   7230
         TabIndex        =   170
         Top             =   1140
         Width           =   660
      End
      Begin MSForms.TextBox txtClient 
         Height          =   315
         Left            =   8400
         TabIndex        =   169
         Top             =   660
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
         Left            =   8400
         TabIndex        =   168
         Top             =   1050
         Width           =   3015
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5318;556"
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
         Left            =   360
         TabIndex        =   114
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tenant: "
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unit Number:"
         Height          =   195
         Left            =   7260
         TabIndex        =   29
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00959595&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   6420
      Width           =   12165
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Lease"
         Height          =   375
         Left            =   5451
         TabIndex        =   126
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel Changes"
         Height          =   375
         Left            =   8525
         TabIndex        =   125
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel New Lease"
         Height          =   375
         Left            =   3794
         TabIndex        =   124
         Top             =   120
         Width           =   1515
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Lease"
         Height          =   375
         Left            =   10065
         TabIndex        =   123
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New Lease"
         Height          =   375
         Left            =   720
         TabIndex        =   122
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "&Save Changes"
         Height          =   375
         Left            =   6988
         TabIndex        =   121
         Top             =   120
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New Lease"
         Height          =   375
         Left            =   2257
         TabIndex        =   120
         Top             =   120
         Width           =   1395
      End
   End
   Begin TabDlg.SSTab tabLease 
      Height          =   4875
      Left            =   120
      TabIndex        =   20
      Top             =   1470
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   1
      TabsPerRow      =   11
      TabHeight       =   520
      TabCaption(0)   =   "&Lease Details"
      TabPicture(0)   =   "Lease.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Rent Charges"
      TabPicture(1)   =   "Lease.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label36"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbo0"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Rent Re&view"
      TabPicture(2)   =   "Lease.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label24"
      Tab(2).Control(1)=   "Label25"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "Label52"
      Tab(2).Control(4)=   "Label58"
      Tab(2).Control(5)=   "Label59"
      Tab(2).Control(6)=   "Label37"
      Tab(2).Control(7)=   "flxRentAnalysis"
      Tab(2).Control(8)=   "txt15"
      Tab(2).Control(9)=   "txt16"
      Tab(2).Control(10)=   "txt17"
      Tab(2).Control(11)=   "cmdNewRentAnalysis"
      Tab(2).Control(12)=   "cmdCancelRentAnalysis"
      Tab(2).Control(13)=   "cmdSageRentAnalysis"
      Tab(2).Control(14)=   "cmdEditRentAnalysis"
      Tab(2).Control(15)=   "txtRentIncreateAmount"
      Tab(2).Control(16)=   "txtRentIncreaseDate"
      Tab(2).Control(17)=   "txtRentReviewDate"
      Tab(2).Control(18)=   "txtSerial"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Brea&ks"
      TabPicture(3)   =   "Lease.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label15"
      Tab(3).Control(1)=   "Label14"
      Tab(3).Control(2)=   "Label13"
      Tab(3).Control(3)=   "cboBreak"
      Tab(3).Control(4)=   "txt14"
      Tab(3).Control(5)=   "cbo3"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Service &Charges"
      TabPicture(4)   =   "Lease.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(1)=   "Frame1(2)"
      Tab(4).Control(2)=   "cbo1"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Interest Charges"
      TabPicture(5)   =   "Lease.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label49"
      Tab(5).Control(1)=   "Label9"
      Tab(5).Control(2)=   "Frame3"
      Tab(5).Control(3)=   "cbo2"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "&Breaches"
      TabPicture(6)   =   "Lease.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label47"
      Tab(6).Control(1)=   "Label46"
      Tab(6).Control(2)=   "Label43"
      Tab(6).Control(3)=   "Label42"
      Tab(6).Control(4)=   "Label40"
      Tab(6).Control(5)=   "Label45"
      Tab(6).Control(6)=   "Label44"
      Tab(6).Control(7)=   "gridBreach"
      Tab(6).Control(8)=   "cboBreachType"
      Tab(6).Control(9)=   "cmdBreachEdit"
      Tab(6).Control(10)=   "cmdBreachSave"
      Tab(6).Control(11)=   "cmdBreachCancel"
      Tab(6).Control(12)=   "cmdBreachNew"
      Tab(6).Control(13)=   "txtBreachID"
      Tab(6).Control(14)=   "Command1"
      Tab(6).Control(15)=   "txtDateReceived"
      Tab(6).Control(16)=   "chkResolved"
      Tab(6).Control(17)=   "txtInitiatedBy"
      Tab(6).Control(18)=   "txtCommenceDate"
      Tab(6).Control(19)=   "txtReceivedBy"
      Tab(6).Control(20)=   "cmdSetBreachType"
      Tab(6).ControlCount=   21
      TabCaption(7)   =   "&Assignment"
      TabPicture(7)   =   "Lease.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label54"
      Tab(7).Control(1)=   "Label53"
      Tab(7).Control(2)=   "gridAssignment"
      Tab(7).Control(3)=   "txtAssignDate"
      Tab(7).Control(4)=   "txtTenant"
      Tab(7).Control(5)=   "txtAssignmentID"
      Tab(7).Control(6)=   "cmdAssignmentNew"
      Tab(7).Control(7)=   "cmdAssignmentCancel"
      Tab(7).Control(8)=   "cmdAssignmentSave"
      Tab(7).Control(9)=   "cmdAssignmentEdit"
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "I&nsurance"
      TabPicture(8)   =   "Lease.frx":09AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label65"
      Tab(8).Control(1)=   "fmeInsurance"
      Tab(8).Control(2)=   "cboInsurancePayable"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "&Supplementary"
      TabPicture(9)   =   "Lease.frx":09C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame6"
      Tab(9).Control(1)=   "Frame9"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "&Memo"
      TabPicture(10)  =   "Lease.frx":09E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1(5)"
      Tab(10).ControlCount=   1
      Begin VB.CommandButton cmdAssignmentEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -66802
         TabIndex        =   197
         Top             =   4020
         Width           =   795
      End
      Begin VB.CommandButton cmdAssignmentSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -65947
         TabIndex        =   196
         Top             =   4020
         Width           =   795
      End
      Begin VB.CommandButton cmdAssignmentCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -65077
         TabIndex        =   195
         Top             =   4020
         Width           =   795
      End
      Begin VB.CommandButton cmdAssignmentNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -67657
         TabIndex        =   194
         Top             =   4020
         Width           =   795
      End
      Begin VB.TextBox txtAssignmentID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65677
         TabIndex        =   193
         Top             =   780
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTenant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70717
         TabIndex        =   192
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtAssignDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73417
         ScrollBars      =   2  'Vertical
         TabIndex        =   191
         Top             =   600
         Width           =   2715
      End
      Begin VB.ComboBox cbo2 
         Height          =   315
         Left            =   -71640
         TabIndex        =   177
         Text            =   "No"
         Top             =   720
         Width           =   915
      End
      Begin VB.Frame Frame9 
         Caption         =   "Date Flag"
         Height          =   1755
         Left            =   -74700
         TabIndex        =   163
         Top             =   360
         Width           =   11355
         Begin VB.TextBox txt18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   285
            Left            =   2340
            MaxLength       =   10
            TabIndex        =   165
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txt19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   765
            Left            =   2340
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   164
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
            TabIndex        =   167
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
            TabIndex        =   166
            Top             =   840
            Width           =   840
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Define Supplementary Fields"
         Height          =   2415
         Left            =   -74700
         TabIndex        =   151
         Top             =   2220
         Width           =   11415
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   154
            Top             =   1200
            Width           =   6375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   153
            Top             =   780
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   152
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
            TabIndex        =   162
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
            TabIndex        =   161
            Top             =   240
            Width           =   7455
         End
         Begin MSForms.TextBox txtSuppCaption3 
            Height          =   315
            Left            =   8880
            TabIndex        =   160
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
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
            Top             =   840
            Width           =   1770
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Height          =   1515
            Left            =   360
            TabIndex        =   155
            Top             =   600
            Width           =   10635
         End
      End
      Begin VB.ComboBox cboInsurancePayable 
         Height          =   315
         Left            =   -72720
         TabIndex        =   149
         Text            =   "No"
         Top             =   480
         Width           =   840
      End
      Begin VB.Frame fmeInsurance 
         Height          =   3375
         Left            =   -74580
         TabIndex        =   129
         Top             =   960
         Width           =   11295
         Begin VB.TextBox txtInsuranceEndDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   143
            Top             =   1470
            Width           =   2415
         End
         Begin VB.Frame Frame10 
            Caption         =   "Charging Methods"
            Height          =   2115
            Left            =   6120
            TabIndex        =   134
            Top             =   660
            Width           =   4695
            Begin VB.TextBox txtAnnualInsuranceCharge 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   140
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
               TabIndex        =   139
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
               TabIndex        =   138
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
               TabIndex        =   137
               Text            =   "0.00"
               Top             =   1260
               Width           =   1500
            End
            Begin VB.OptionButton optAnnualInsuranceCharge 
               Caption         =   "Annual Insurance Charge"
               Height          =   255
               Left            =   180
               TabIndex        =   136
               Top             =   720
               Width           =   2535
            End
            Begin VB.OptionButton optInsurancePercentage 
               Caption         =   "Insurance Charge Percentage (%)"
               Height          =   255
               Left            =   180
               TabIndex        =   135
               Top             =   360
               Width           =   2715
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Total Insurance Charge for Year:"
               Height          =   195
               Left            =   180
               TabIndex        =   142
               Top             =   1260
               Width           =   2310
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "Insurance Charge Due Each Period:"
               Height          =   195
               Left            =   180
               TabIndex        =   141
               Top             =   1620
               Width           =   2565
            End
         End
         Begin VB.ComboBox cboInsuranceDemandType 
            Height          =   315
            Left            =   2040
            TabIndex        =   133
            Text            =   "1"
            Top             =   2355
            Width           =   2415
         End
         Begin VB.ComboBox cboInsuranceFrequency 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   1905
            Width           =   2415
         End
         Begin VB.TextBox txtInsuranceStartDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   131
            Top             =   1020
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
            TabIndex        =   130
            Top             =   2790
            Width           =   1320
         End
         Begin MSForms.ComboBox cboInsuranceDept 
            Height          =   315
            Left            =   2040
            TabIndex        =   189
            Top             =   600
            Width           =   2415
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   175
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance End Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   148
            Top             =   1515
            Width           =   1470
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   240
            TabIndex        =   147
            Top             =   2370
            Width           =   1050
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance Start Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   146
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance Frequency:"
            Height          =   195
            Left            =   240
            TabIndex        =   145
            Top             =   1950
            Width           =   1545
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   144
            Top             =   2805
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Notes"
         Height          =   4395
         Index           =   5
         Left            =   -74700
         TabIndex        =   127
         Top             =   360
         Width           =   11595
         Begin VB.TextBox txt20 
            Height          =   3915
            Left            =   240
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   128
            Top             =   360
            Width           =   11055
         End
      End
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73635
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   98
         Top             =   660
         Width           =   1500
      End
      Begin VB.ComboBox cbo3 
         Height          =   315
         Left            =   -69280
         TabIndex        =   105
         Text            =   "No"
         Top             =   1500
         Width           =   2000
      End
      Begin VB.TextBox txt14 
         Height          =   315
         Left            =   -69280
         MaxLength       =   10
         TabIndex        =   106
         Top             =   2160
         Width           =   1960
      End
      Begin VB.ComboBox cboBreak 
         Height          =   315
         Left            =   -69280
         TabIndex        =   107
         Top             =   2820
         Width           =   2000
      End
      Begin VB.TextBox txtRentReviewDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   99
         Top             =   660
         Width           =   2500
      End
      Begin VB.TextBox txtRentIncreaseDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   100
         Top             =   660
         Width           =   2500
      End
      Begin VB.TextBox txtRentIncreateAmount 
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
         Left            =   -67080
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   660
         Width           =   2500
      End
      Begin VB.CommandButton cmdEditRentAnalysis 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -68760
         TabIndex        =   94
         Top             =   3900
         Width           =   1300
      End
      Begin VB.CommandButton cmdSageRentAnalysis 
         Caption         =   "&Save"
         Height          =   375
         Left            =   -73920
         TabIndex        =   103
         Top             =   3900
         Width           =   1300
      End
      Begin VB.CommandButton cmdCancelRentAnalysis 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -65040
         TabIndex        =   95
         Top             =   3900
         Width           =   1300
      End
      Begin VB.CommandButton cmdNewRentAnalysis 
         Caption         =   "&New"
         Height          =   375
         Left            =   -70200
         TabIndex        =   93
         Top             =   3900
         Width           =   1300
      End
      Begin VB.TextBox txt17 
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
         Left            =   -65160
         TabIndex        =   92
         Top             =   4380
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txt16 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69120
         MaxLength       =   10
         TabIndex        =   91
         Top             =   4380
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txt15 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72720
         MaxLength       =   10
         TabIndex        =   90
         Top             =   4380
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cbo1 
         Height          =   315
         Left            =   -72360
         TabIndex        =   88
         Text            =   "No"
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   3555
         Index           =   2
         Left            =   -74160
         TabIndex        =   64
         Top             =   840
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
            TabIndex        =   83
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            Caption         =   "Charging Methods"
            Height          =   2775
            Left            =   5280
            TabIndex        =   116
            Top             =   240
            Width           =   4815
            Begin VB.TextBox txtGlobalAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   80
               Top             =   1440
               Width           =   1500
            End
            Begin VB.TextBox txtAnnualService 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   78
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox txtPPSqFoot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   75
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
               TabIndex        =   73
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
               TabIndex        =   82
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
               TabIndex        =   81
               Text            =   "0.00"
               Top             =   2040
               Width           =   1500
            End
            Begin VB.OptionButton optGlobalData 
               Caption         =   "Global Data"
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   1440
               Width           =   1335
            End
            Begin VB.OptionButton optFixedTotal 
               Caption         =   "Annual Service Charge Amout"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   1080
               Width           =   2535
            End
            Begin VB.OptionButton optSqFoot 
               Caption         =   "Price Per Sq. Foot/Metre"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   720
               Width           =   2175
            End
            Begin VB.OptionButton optPercentage 
               Caption         =   "Service Charge Percentage (%)"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Service Charge Total for Year:"
               Height          =   195
               Left            =   240
               TabIndex        =   118
               Top             =   2040
               Width           =   2145
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Service Charge Due Each Period:"
               Height          =   195
               Left            =   240
               TabIndex        =   117
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
            TabIndex        =   71
            Top             =   3132
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txt10b 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1470
            TabIndex        =   68
            Top             =   1701
            Width           =   2715
         End
         Begin VB.ComboBox cboFreqSC 
            Height          =   315
            ItemData        =   "Lease.frx":09FE
            Left            =   1470
            List            =   "Lease.frx":0A00
            TabIndex        =   67
            Top             =   1204
            Width           =   2715
         End
         Begin VB.TextBox txt10 
            Height          =   285
            Left            =   2850
            TabIndex        =   69
            Top             =   2665
            Width           =   1335
         End
         Begin VB.TextBox txt9 
            Height          =   285
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   66
            Top             =   737
            Width           =   2715
         End
         Begin VB.ComboBox cboSCDemandType 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2168
            Width           =   2715
         End
         Begin MSForms.ComboBox cboServiceChargeDept 
            Height          =   315
            Left            =   1470
            TabIndex        =   65
            Top             =   240
            Width           =   2715
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4789;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   173
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "xService Charge Total for Year:"
            Height          =   195
            Left            =   5085
            TabIndex        =   119
            Top             =   3120
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "xService Charge Due Each Period:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
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
            TabIndex        =   86
            Top             =   1680
            Width           =   1110
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Payable From:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "T/O Limit:"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   2640
            Width           =   705
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   2160
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3375
         Left            =   -73462
         TabIndex        =   63
         Top             =   1140
         Width           =   9225
         Begin VB.TextBox txt12a 
            Height          =   315
            Left            =   2400
            TabIndex        =   180
            Top             =   2040
            Width           =   2475
         End
         Begin VB.TextBox txt12 
            Height          =   315
            Left            =   2400
            TabIndex        =   179
            Top             =   1440
            Width           =   2475
         End
         Begin VB.TextBox txt13 
            Height          =   315
            Left            =   6960
            TabIndex        =   181
            Top             =   840
            Width           =   1800
         End
         Begin VB.TextBox txt11 
            Height          =   315
            Left            =   6960
            TabIndex        =   182
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox cboIntDemandType 
            Height          =   315
            Left            =   6960
            TabIndex        =   183
            Text            =   "3"
            Top             =   2040
            Width           =   1800
         End
         Begin MSForms.ComboBox cboIntChargeDept 
            Height          =   315
            Left            =   2400
            TabIndex        =   178
            Top             =   840
            Width           =   2475
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4366;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705"
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount to charge Interest on:"
            Height          =   195
            Left            =   240
            TabIndex        =   188
            Top             =   2115
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
            TabIndex        =   187
            Top             =   1470
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
            TabIndex        =   186
            Top             =   840
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
            TabIndex        =   185
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Left            =   5160
            TabIndex        =   184
            Top             =   2055
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            Height          =   195
            Left            =   8160
            TabIndex        =   176
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   174
            Top             =   840
            Width           =   405
         End
      End
      Begin VB.CommandButton cmdSetBreachType 
         Caption         =   "..."
         Height          =   315
         Left            =   -72120
         TabIndex        =   53
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtReceivedBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66780
         TabIndex        =   52
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtCommenceDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71880
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txtInitiatedBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70740
         TabIndex        =   50
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkResolved 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -68700
         TabIndex        =   49
         Top             =   1140
         Width           =   435
      End
      Begin VB.TextBox txtDateReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68040
         TabIndex        =   48
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   -64920
         TabIndex        =   47
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtBreachID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73800
         TabIndex        =   46
         Top             =   540
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBreachNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -67800
         TabIndex        =   45
         Top             =   4380
         Width           =   795
      End
      Begin VB.CommandButton cmdBreachCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -65235
         TabIndex        =   44
         Top             =   4380
         Width           =   795
      End
      Begin VB.CommandButton cmdBreachSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -66090
         TabIndex        =   43
         Top             =   4380
         Width           =   795
      End
      Begin VB.CommandButton cmdBreachEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -66945
         TabIndex        =   42
         Top             =   4380
         Width           =   795
      End
      Begin VB.ComboBox cbo0 
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Text            =   "No"
         Top             =   600
         Width           =   840
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   -73267
         TabIndex        =   33
         Top             =   1050
         Width           =   8835
         Begin VB.CommandButton cmdLeaseType 
            Caption         =   "..."
            Height          =   300
            Left            =   3840
            TabIndex        =   7
            Top             =   2040
            Width           =   405
         End
         Begin VB.CheckBox chkSubLease 
            Caption         =   "Yes"
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   1260
            Width           =   735
         End
         Begin VB.ComboBox cboHeadLease 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   4
            Text            =   "cboHeadLease"
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txt3 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2040
            Width           =   1995
         End
         Begin VB.TextBox txt4 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   8
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtLeaseStDt 
            Height          =   285
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1260
            Width           =   1995
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "Lease.frx":0A02
            Left            =   1800
            List            =   "Lease.frx":0A04
            TabIndex        =   6
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   1260
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Index           =   1
         Left            =   1523
         TabIndex        =   21
         Top             =   1080
         Width           =   9255
         Begin VB.ComboBox cboBRDemandType 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2640
            Width           =   2700
         End
         Begin VB.ComboBox cboFreqBR 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1560
            TabIndex        =   14
            Text            =   "cboFreqBR"
            Top             =   1230
            Width           =   2700
         End
         Begin VB.TextBox txt5 
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Top             =   1920
            Width           =   2700
         End
         Begin VB.TextBox txt6 
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
            Left            =   7080
            TabIndex        =   17
            Top             =   540
            Width           =   1500
         End
         Begin VB.TextBox txt7 
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
            Left            =   7080
            TabIndex        =   19
            Top             =   1920
            Width           =   1500
         End
         Begin VB.TextBox txt8 
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
            Left            =   7080
            TabIndex        =   18
            Top             =   1230
            Width           =   1500
         End
         Begin MSForms.ComboBox cboRentChargeDept 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   540
            Width           =   2700
            VariousPropertyBits=   1820346395
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4762;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705"
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   172
            Top             =   540
            Width           =   405
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   40
            Top             =   2640
            Width           =   1050
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent Start Date:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frequency:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   1230
            Width           =   795
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total Rent for Year:"
            Height          =   195
            Index           =   4
            Left            =   5280
            TabIndex        =   24
            Top             =   540
            Width           =   1425
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due Date:"
            Height          =   195
            Index           =   6
            Left            =   5280
            TabIndex        =   23
            Top             =   1920
            Width           =   1110
         End
         Begin VB.Label lblRentCharges 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent Due Each Period:"
            Height          =   195
            Index           =   5
            Left            =   5280
            TabIndex        =   22
            Top             =   1230
            Width           =   1650
         End
      End
      Begin MSDataListLib.DataCombo cboBreachType 
         Bindings        =   "Lease.frx":0A06
         DataSource      =   "adoBreaches"
         Height          =   315
         Left            =   -73800
         TabIndex        =   54
         Top             =   1080
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
         Height          =   2835
         Left            =   -73800
         TabIndex        =   55
         Top             =   1500
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentAnalysis 
         Height          =   2835
         Left            =   -74040
         TabIndex        =   104
         Top             =   1020
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
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
         Left            =   -73357
         TabIndex        =   198
         Top             =   1020
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
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
      Begin VB.Label Label53 
         Caption         =   "Tenant:"
         Height          =   255
         Left            =   -70717
         TabIndex        =   200
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label54 
         Caption         =   "Assignment Date:"
         Height          =   255
         Left            =   -73417
         TabIndex        =   199
         Top             =   360
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
         TabIndex        =   190
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Payable:"
         Height          =   195
         Left            =   -74580
         TabIndex        =   150
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Serial"
         Height          =   195
         Left            =   -73560
         TabIndex        =   115
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Review Date:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   113
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Increase Date:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   112
         Top             =   420
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
         Left            =   -67080
         TabIndex        =   111
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Break Clause:"
         Height          =   195
         Left            =   -70425
         TabIndex        =   110
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Break Type:"
         Height          =   195
         Left            =   -70425
         TabIndex        =   109
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Break Date:"
         Height          =   195
         Left            =   -70425
         TabIndex        =   108
         Top             =   2220
         Width           =   855
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
         Left            =   -66960
         TabIndex        =   102
         Top             =   4380
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
         Left            =   -70680
         TabIndex        =   97
         Top             =   4380
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
         Left            =   -74160
         TabIndex        =   96
         Top             =   4380
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Service Charge Payable:"
         Height          =   195
         Left            =   -74160
         TabIndex        =   89
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label44 
         Caption         =   "Breach Type:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   62
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label45 
         Caption         =   "Commence Date:"
         Height          =   435
         Left            =   -71880
         TabIndex        =   61
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Received By:"
         Height          =   255
         Left            =   -66780
         TabIndex        =   60
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label42 
         Caption         =   "Initiated By:"
         Height          =   255
         Left            =   -70740
         TabIndex        =   59
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label43 
         Caption         =   "Resolved"
         Height          =   195
         Left            =   -68880
         TabIndex        =   58
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label46 
         Caption         =   "Date Received:"
         Height          =   255
         Left            =   -68040
         TabIndex        =   57
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label47 
         Caption         =   "Memo:"
         Height          =   255
         Left            =   -64980
         TabIndex        =   56
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rent Payable:"
         Height          =   195
         Left            =   1560
         TabIndex        =   41
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label49 
         Height          =   255
         Left            =   -72900
         TabIndex        =   32
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Label Label41 
      Caption         =   "Description:"
      Height          =   255
      Left            =   4950
      TabIndex        =   31
      Top             =   3000
      Width           =   1515
   End
End
Attribute VB_Name = "frmLease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BREACH_NEW_ENTRY_ As Boolean
Dim ASSIGNMENT_NEW_ENTRY_ As Boolean

Dim szaDemandtype() As String

Dim TenantCode As String
Dim TenantName As String
Dim OldUnit As String
Dim NewUnit As String
Dim BRFreq As Integer
Dim SCFreq As Integer
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
Dim bf As Integer
Dim scf As Integer
Private mintCurFrame As Integer ' Current Frame visible

Public FormLoad As Boolean
    
Dim gszSageAccountNumber As String
Dim gszCurUnitNum As String

Private Sub cbo0_LostFocus()

If cbo0.text <> "No" And cbo0.text <> "Yes" Then MsgBox "Invalid Base Rent Payable status.", vbOKOnly + vbCritical, "Invalid Data"

End Sub

Private Sub cbo1_LostFocus()

If cbo1.text <> "No" And cbo1.text <> "Yes" Then MsgBox "Invalid Service Charge Payable status.", vbOKOnly + vbCritical, "Invalid Data"


End Sub
Private Sub cbo2_LostFocus()
'On Error Resume Next
If cbo2.text <> "No" And cbo2.text <> "Yes" Then MsgBox "Invalid Interest Chargeable status.", vbOKOnly + vbCritical, "Invalid Data"
If cbo2.text = "Yes" Then
    txt12a.text = CDbl(txt8.text) + CDbl(txtAmount.text)
    If txt11.text <> "" And txt12.text <> "" And txt12a.text <> "" Then CalculateInterest
End If

End Sub
Private Sub cbo3_LostFocus()
'On Error Resume Next
If cbo3.text <> "No" And cbo3.text <> "Yes" Then MsgBox "Invalid Break Clause status.", vbOKOnly + vbCritical, "Invalid Data"

End Sub

Private Sub cboBreak_LostFocus()
'On Error Resume Next
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

Private Sub cboFreqBR_LostFocus()
'On Error Resume Next
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
    MsgBox "Base Rent frequency is invalid.", vbOKOnly + vbCritical, "Invalid Frequency"
'    cboFreqBR.text = "Select"
    Exit Sub
End If

If txt5.text <> "" Then Call CalculateBR
If cboFreqBR.ListIndex < 6 Then txt7.Enabled = True
bf = cboFreqBR.ListIndex

End Sub

Private Sub cboFreqSC_GotFocus()
   If txt9.text = "" Then
      MsgBox "You must enter the Payable from date before enter frequency.", vbInformation + vbOKOnly, "Payable from date missing"
      txt9.SetFocus
   End If
End Sub

Private Sub cboFreqSC_LostFocus()
If cboFreqSC.text = "" Then Exit Sub
'On Error Resume Next
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
'    cboFreqSC.text = "-Select-"
    Exit Sub
End If

If cbo1.text = "Yes" Then Call CalculateSC

'If cboFreqSC.ListIndex < 6 Then txt10b.Enabled = True

scf = cboFreqSC.ListIndex

End Sub

Private Sub cboInsuranceFrequency_GotFocus()
   If txtInsuranceStartDate.text = "" Then
      MsgBox "You must enter the Insurance Start Date before enter frequency.", vbInformation + vbOKOnly, "Insurance start date missing"
      txtInsuranceStartDate.SetFocus
   End If
End Sub

Private Sub cboInsuranceFrequency_LostFocus()
    
    If cboInsuranceFrequency.text = "" Then Exit Sub
    'On Error Resume Next
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
    
    'If cboFreqSC.ListIndex < 6 Then txt10b.Enabled = True
    
    'scf = cboFreqSC.ListIndex
   
End Sub

Private Sub cboInsuranceDept_Click()
   If cboInsuranceDept.text = "" Then
'      MsgBox "Please set the department name into SAGE.", vbCritical + vbOKOnly, "Department Name"
   End If
End Sub

Private Sub cboIntChargeDept_Click()
   If cboIntChargeDept.text = "" Then
'      MsgBox "Please set the department name into SAGE.", vbCritical + vbOKOnly, "Department Name"
   End If
End Sub

Private Sub cboRentChargeDept_Click()
   If cboRentChargeDept.text = "" Then
'      MsgBox "Please set the department name into SAGE.", vbCritical + vbOKOnly, "Department Name"
   End If
End Sub

Private Sub cboServiceChargeDept_Click()
   If cboServiceChargeDept.text = "" Then
'      MsgBox "Please set the department name into SAGE.", vbCritical + vbOKOnly, "Department Name"
   End If
End Sub

Private Sub cboTenant1_Click()
   Dim szaTenant() As String
   
   szaTenant = Split(cboTenant1.text, " / ")
   gszSageAccountNumber = szaTenant(0)

'On Error Resume Next
If cboTenant1.text = "" Then
    MsgBox "You must select the tenant whose lease you want to view.", vbOKOnly + vbCritical, "No tenant selected"
Else
   Call EmptyBoxes
   
   cmdAddNew.SetFocus
   
   Call GetRecord
   
   ConfigureFlxGrid flxRentAnalysis
   LoadFlxGrid flxRentAnalysis

   cmdAddNew.SetFocus
End If

End Sub

Private Sub cboTenant1_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboTenant1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboTenant2_Click()
   Dim szaTemp() As String

   szaTemp = Split(cboTenant2.text, " / ")
   TenantCode = szaTemp(0)
   TenantName = szaTemp(1)

   cboTenant2.Enabled = False
   Call GetUnits
   cboUnit.Enabled = True
   cboTenant1.Enabled = False

   cmdSaveNew.Visible = True
   cmdSaveNew.TabIndex = 25

   cbo0.text = "No"
   cbo1.text = "No"
   cbo2.text = "No"
   cbo3.text = "No"
End Sub

Private Sub cboUnit_Click()
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT ClientName, PropertyName " & _
              "FROM Client, Property, Units " & _
              "WHERE Client.ClientID = Property.ClientID And " & _
                  "Property.PropertyID = Units.PropertyID And " & _
                  "Units.UnitNumber = '" & cboUnit.text & "';"

    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    txtClient.text = Rst1!ClientName
    txtProperty.text = Rst1!PropertyName
    
    Rst1.Close
    Conn1.Close
    Set Rst1 = Nothing
    Set Conn1 = Nothing
End Sub

Private Sub cboUnit_LostFocus()

   'On Error Resume Next
   Dim i, j, match As Integer
   
   match = 0
   j = cboUnit.ListCount - 1
   For i = 0 To j
      If cboUnit.List(i) = cboUnit.text Then
         match = 1
         Exit For
      End If
   Next i
   If match = 0 Then
      'MsgBox "Tenant selected is invalid.", vbOKOnly + vbCritical, "Invalid Tenant"
      cboUnit.text = OldUnit
      Exit Sub
   End If
    
   'MsgBox Left(cboUnit.text, 8)
   GetGlobalDataForProperty (Left(cboUnit.text, 8))
   
   txtLeaseID_GotFocus
End Sub

Private Sub chkSubLease_Click()
'On Error Resume Next
If chkSubLease.Value = 1 Then
    cboHeadLease.Enabled = True
Else
    cboHeadLease.Enabled = False
End If
End Sub

Private Sub cmdAddNew_Click()
   Call AddNew
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
   'On Error Resume Next
'   Call GetTenantsWithLease
   Call DisableBoxes
   cboTenant2.Visible = False
   cboTenant1.Visible = True
   cboTenant1.Enabled = True
End Sub

Private Sub cmdCancelNew_Click()
Call EmptyBoxes
Call GetTenantsWithLease
Call DisableBoxes
cboTenant2.Visible = False
cboTenant1.Visible = True
cboTenant1.Enabled = True
End Sub

Private Sub cmdCancelRentAnalysis_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   If cmdEditRentAnalysis.Enabled = True And cmdNewRentAnalysis.Enabled = True Then Exit Sub
   
   UnlockTextBoxes False
   flxRentAnalysis.Enabled = False
End Sub

Private Sub cmdDelete_Click()

Call Delete

End Sub

Private Sub cmdEdit_Click()
   Dim szTenant() As String
   Dim szText As String, iCboIndex As Integer
   Dim TenantCode As String
   Dim TenantName As String

   If cboTenant1.text = "" Then
       MsgBox "You must select a Tenant to view the lease.", vbOKOnly + vbCritical, "No Tenant Selected"
       Exit Sub
   End If

   szTenant = Split(cboTenant1.text, " / ")
   TenantCode = szTenant(0)
   TenantName = szTenant(1)

   Call EnableBoxes
   Call GetUnits
   szText = cboBRDemandType.text
   iCboIndex = cboBRDemandType.ListIndex
   Call FillcboType(cboBRDemandType, szaDemandtype)
   cboBRDemandType.ListIndex = iCboIndex

   iCboIndex = cboSCDemandType.ListIndex
   Call FillcboType(cboSCDemandType, szaDemandtype)
   cboSCDemandType.ListIndex = iCboIndex

   'iCboIndex = cboIntDemandType.text
   iCboIndex = cboIntDemandType.ListIndex
   Call FillcboType(cboIntDemandType, szaDemandtype)
   cboIntDemandType.ListIndex = iCboIndex
   'bf = cboFreqBR.ListIndex
   'scf = cboFreqSC.ListIndex

   cboTenant1.Enabled = False

   cmdAddNew.Visible = False
   'mnuAdd.Enabled = False
   cmdDelete.Visible = False
   'mnuDelete.Enabled = False
   cmdEdit.Visible = False
   'mnuEdit.Enabled = False
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
End Sub

Private Sub UnlockTextBoxes(bState As Boolean)
   txtRentReviewDate.Locked = Not bState
   txtRentIncreaseDate.Locked = Not bState
   txtRentIncreateAmount.Locked = Not bState
   txtSerial.Locked = Not bState
   
   If Not bState Then
      txtRentReviewDate.text = ""
      txtRentIncreaseDate.text = ""
      txtRentIncreateAmount.text = ""
      txtSerial.text = ""
   End If
End Sub

Private Sub cmdLeaseType_Click()
   Load frmSecondaryCode
   frmSecondaryCode.Show
End Sub

Private Sub cmdNewRentAnalysis_Click()
   If MsgBox("Do you want to Add new data?", vbQuestion + vbYesNo, "Add new Data") = vbNo Then Exit Sub
   UnlockTextBoxes True
   cmdNewRentAnalysis.Enabled = False
   flxRentAnalysis.Enabled = False
End Sub

Private Sub cmdSageRentAnalysis_Click()
   If cmdEditRentAnalysis.Enabled = True And cmdNewRentAnalysis.Enabled = True Then Exit Sub
'
   Dim lID As Long
'
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt
'
   SQLStr1 = "SELECT * " & _
             "FROM RentAnalysis"
'
   If cmdEditRentAnalysis.Enabled = False Then
      lID = FindXID
      SQLStr1 = SQLStr1 + " WHERE ID = " & lID & ";"
'
      Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
      Rst1.Edit
      flxRentAnalysis.TextMatrix(flxRentAnalysis.Row, 0) = "X"
   End If
   If cmdNewRentAnalysis.Enabled = False Then
      SQLStr1 = SQLStr1 + ";"
      
      Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
      Rst1.AddNew
   End If
'
   Rst1!SageAccountNumber = gszSageAccountNumber
   Rst1!SerialNumber = txtSerial.text
   Rst1!RentReviewDate = CDate(Format(txtRentReviewDate.text, "dd/mm/yyyy"))
   Rst1!RentIncreaseDate = CDate(Format(txtRentIncreaseDate.text, "dd/mm/yyyy"))
   Rst1!RentIncreaseAmount = CCur(Format(txtRentIncreateAmount.text, "0.00"))
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
Dim i As Integer

On Error Resume Next

If txtLeaseID.text = "" Then
   MsgBox "You must enter a Lease reference to continue!", vbOKOnly + vbCritical, "Reference Required"
   txtLeaseID.SetFocus
   Exit Sub
End If
If cboType.text = "" Then
   MsgBox "You must enter the Lease Type to continue!", vbOKOnly + vbCritical, "Lease Type missing"
   tabLease.Tab = 0
   cboType.SetFocus
   Exit Sub
End If

If txt4.text = "" Then
   MsgBox "You must enter the Year End date to continue!", vbOKOnly + vbCritical, "Year End date missing"
   tabLease.Tab = 0
   txt4.SetFocus
   Exit Sub
End If

If txtLeaseStDt.text = "" Then
   MsgBox "You must enter a Lease Start Date!", vbOKOnly + vbCritical, "Date Required"
   tabLease.Tab = 0
   txtLeaseStDt.SetFocus
   Exit Sub
End If

If txt3.text = "" Then
   MsgBox "You must enter a Lease End Date!", vbOKOnly + vbCritical, "Date Required"
   tabLease.Tab = 0
   txt3.SetFocus
   Exit Sub
End If

If cbo0.text = "Yes" Then
   If cboRentChargeDept.Column(0) = "" Then
      MsgBox "You must select a Department of rent.", vbOKOnly + vbCritical, "Rent - Department"
      tabLease.Tab = 1
      cboRentChargeDept.SetFocus
      Exit Sub
   End If
   If cboFreqBR.text = "" Then
      MsgBox "You must select a Rent Frequency!", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 1
      cboFreqBR.SetFocus
      Exit Sub
   End If
   If txt5.text = "" Then
      MsgBox "You must enter a Rent Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 1
      txt5.SetFocus
      Exit Sub
   End If
End If

If cbo1.text = "Yes" Then
   If cboServiceChargeDept.text = "" Then
      MsgBox "You must select a department for the service charge!", vbOKOnly + vbCritical, "Service Charge - Department"
      tabLease.Tab = 4
      cboServiceChargeDept.SetFocus
      Exit Sub
   End If
    If txt9.text = "" Then
        MsgBox "You must enter a Payable From Date for Service Charge!", vbOKOnly + vbCritical, "Date Required"
        Exit Sub
    End If
    If cboFreqSC.text = "" Then
        MsgBox "You must select a Service Charge Frequency!", vbOKOnly + vbCritical, "Frequency Required"
        Exit Sub
    End If
End If

If cbo2.text = "Yes" Then
   If cboIntChargeDept.text = "" Then
      MsgBox "You must select a department for the interest charge!", vbOKOnly + vbCritical, "Interest Charge - Department"
      tabLease.Tab = 5
      cboIntChargeDept.SetFocus
      Exit Sub
   End If
   If txt11.text = "" Then
      MsgBox "You must enter number of days interest will charge after!", vbOKOnly + vbCritical, "Interest Charge"
      tabLease.Tab = 5
      txt11.SetFocus
      Exit Sub
   End If
   If cboIntDemandType.text = "Yes" Then
      MsgBox "You must select interest demand type!", vbOKOnly + vbCritical, "Demand Type"
      tabLease.Tab = 5
      cboIntDemandType.SetFocus
      Exit Sub
   End If
End If

If cboInsurancePayable.text = "Yes" Then
   If cboInsuranceDept.text = "" Then
      MsgBox "You must select department of insurance!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceDept.SetFocus
      Exit Sub
   End If
   If txtInsuranceStartDate.text = "" Then
      MsgBox "You must enter insurance start date!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      txtInsuranceStartDate.SetFocus
      Exit Sub
   End If
   If txtInsuranceEndDate.text = "" Then
      MsgBox "You must enter insurance end date!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      txtInsuranceEndDate.SetFocus
      Exit Sub
   End If
   If cboInsuranceFrequency.text = "" Then
      MsgBox "You must select insurance frequency!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceFrequency.SetFocus
      Exit Sub
   End If
   If cboInsuranceDemandType.text = "" Then
      MsgBox "You must select insurance demand type!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceDemandType.SetFocus
      Exit Sub
   End If
End If

If txtLeaseStDt.text = "" Then
    MsgBox "You must enter a Lease Start Date!", vbOKOnly + vbCritical, "Date Required"
     Exit Sub
End If
If txt3.text = "" Then
    MsgBox "You must enter a Lease End Date!", vbOKOnly + vbCritical, "Date Required"
    Exit Sub
End If

GetGlobalDataForProperty (Left(Trim(cboUnit.text), 8))
'validate rest
If cboUnit.text = "" Then
    MsgBox "You must select a unit!", vbOKOnly + vbExclamation, "No Unit Selected"
    Exit Sub
Else
    NewUnit = cboUnit.text
End If

If cboFreqBR.ListIndex <> -1 Then bf = cboFreqBR.ListIndex

If cboFreqSC.ListIndex <> -1 Then scf = cboFreqSC.ListIndex

'save the details to a new record
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseOdbc
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM LeaseDetails WHERE LeaseID = '" & txtLeaseID.text & "'"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If Rst1.EOF Or Rst1.BOF Then
    MsgBox "Cannot update the lease. No lease reference exist.", vbInformation, "Save Lease"
    Exit Sub
End If

Rst1.Edit
Rst1!SageAccountNumber = TenantCode
Rst1!LeaseID = txtLeaseID.text
Rst1!CompanyName = TenantName

If chkSubLease.Value Then
    If cboHeadLease.text <> "" Then Rst1!HeadLease = cboHeadLease.text
End If

Rst1!UnitNumber = cboUnit.text
Rst1!TYPEOFSTORE = cboType.text
Rst1!StartDate = txtLeaseStDt.text
Rst1!EndDate = txt3.text
Rst1!YearEnd = txt4.text
Rst1!BRStartDate = txt5.text
'MsgBox cboRentChargeDept.Column(1, 0)
Rst1!RentChargeDept = cboRentChargeDept.Value
For i = 2 To 3
    If Mid(cboFreqBR.text, i, 1) = "-" Then Rst1!BRfrequency = CInt(Left(cboFreqBR.text, i - 1))
Next i
'Rst1!BRfrequency = cboFreqBR.ListIndex + 1
If txt6.text = "" Then Rst1!BRTotal = Null Else Rst1!BRTotal = CDbl(txt6.text)
Rst1!BRNextDueDate = txt7.text
If txt8.text = "" Then Rst1!BRAmount = Null Else Rst1!BRAmount = CDbl(txt8.text)
Rst1!BRDemandType = CByte(IIf(cboBRDemandType.text <> "", cboBRDemandType.ListIndex + 1, 0))
Rst1!BRPayable = IIf(cbo0.text = "No", "N", "Y")
Rst1!SCPayable = IIf(cbo1.text = "No", "N", "Y")

If cbo1.text <> "No" Then
   For i = 2 To 3
       If Mid(cboFreqSC.text, i, 1) = "-" Then Rst1!SCfrequency = CInt(Left(cboFreqSC.text, i - 1))
   Next i
   'Rst1!SCFrequency = cboFreqSC.ListIndex + 1
   'Debug.Print cboServiceChargeDept.text
   'Rst1!ServiceChargeDept = DeptID(cboServiceChargeDept.text, cboRentChargeDept)
   Rst1!ServiceChargeDept = cboServiceChargeDept.Value
   Rst1!SCPayableFrom = txt9.text
   If txt10b.text = "" Then Rst1!SCNextDueDate = Null Else Rst1!SCNextDueDate = txt10b.text
   
   'If txtPPSqFoot.text <> "" Then Rst1!SCPricePerSqFoot = CDbl(txtPPSqFoot.text)
   If optSqFoot.Value Then
      Rst1!SCPricePerSqFoot = CDbl(IIf(txtPPSqFoot.text = "", 0, txtPPSqFoot.text))
      Rst1!SCPercentage = Null
      Rst1!SCTotal = Null
   End If
   
   'If txtSCPercentage.text <> "" Then Rst1!SCPercentage = CDbl(txtSCPercentage.text)
   If optPercentage.Value Then
      Rst1!SCPercentage = CDbl(IIf(txtSCPercentage.text = "", 0, txtSCPercentage.text))
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCTotal = Null
   End If
   
   'If txtAmount.text <> "" Then Rst1!SCTotal = CDbl(txtAmount.text)
   If optFixedTotal.Value Then
      Rst1!SCTotal = CDbl(IIf(txtAnnualService.text = "", 0, txtAnnualService.text))
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCPercentage = Null
   End If
   
   If optGlobalData.Value Then
      Rst1!SCTotal = Null
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCPercentage = Null
   End If
   
   'If txt10c.text = "" Then Rst1!SCAmount = Null Else Rst1!SCAmount = CDbl(txt10c.text)
   Rst1!SCAmount = CDbl(txtFinalAmout.text)
   
   If txt10.text = "" Then Rst1!SCTOLimit = Null Else Rst1!SCTOLimit = CDbl(txt10.text)
   Rst1!SCDemandType = CByte(IIf(cboSCDemandType.text <> "", cboSCDemandType.ListIndex + 1, 0))
End If

Rst1!InterestChargeable = IIf(cbo2.text = "No", "N", "Y")
If txt11.text = "" Then Rst1!DaysAfterInterestPayable = Null Else Rst1!DaysAfterInterestPayable = CInt(txt11.text)
'Rst1!IntChargeDept = DeptID(cboIntChargeDept.text, cboRentChargeDept)
Rst1!IntChargeDept = cboIntChargeDept.Value
If txt12.text = "" Then Rst1!AdditionalInterest = Null Else Rst1!AdditionalInterest = CDbl(txt12.text)
If txt12a.text = "" Then Rst1!InterestChargedOn = Null Else Rst1!InterestChargedOn = CDbl(txt12a.text)
If txt13.text = "" Then Rst1!InterestAmount = Null Else Rst1!InterestAmount = CDbl(txt13.text)
Rst1!IntDemandType = CByte(IIf(cboIntDemandType.text <> "", cboIntDemandType.ListIndex + 1, 0))
If cbo3.text = "Yes" Then Rst1!BreakClause = "Y" Else Rst1!BreakClause = "N"
Rst1!BreakClause = "Y"
If txt14.text = "" Then Rst1!BreakDate = Null Else Rst1!BreakDate = txt14.text
If cboBreak.text = "" Then Rst1!BreakType = Null Else Rst1!BreakType = cboBreak.text
Rst1!RentReviewDate = txt15.text
Rst1!RentIncreaseDate = txt16.text
If txt17.text = "" Then Rst1!RentIncreaseAmount = Null Else Rst1!RentIncreaseAmount = CDbl(txt17.text)
Rst1!DateFlagDate = txt18.text
Rst1!DateFlagDescription = txt19.text
Rst1!Notes = txt20.text

'Insurance
If cboInsurancePayable.text = "No" Then Rst1!InsurancePayable = "N"
If cboInsurancePayable.text = "Yes" Then Rst1!InsurancePayable = "Y"

For i = 2 To 3
    If Mid(cboInsuranceFrequency.text, i, 1) = "-" Then Rst1!InsuranceFrequency = CInt(Left(cboInsuranceFrequency.text, i - 1))
Next i

'Rst1!InsuranceDept = DeptID(cboInsuranceDept.text, cboRentChargeDept)
Rst1!InsuranceDept = cboInsuranceDept.Value
If txtInsuranceStartDate.text <> "" Then Rst1!InsuranceStartDate = txtInsuranceStartDate.text
If txtInsuranceEndDate.text <> "" Then Rst1!InsuranceEndDate = txtInsuranceEndDate.text
Rst1!InsuranceDemandType = CByte(IIf(cboInsuranceDemandType.text <> "", cboInsuranceDemandType.ListIndex + 1, 0))
If txtInsuranceEachPeriod.text <> "" Then Rst1!InsuranceEachPeriod = CDbl(txtInsuranceEachPeriod.text)
If txtInsuranceNextDueDate.text <> "" Then Rst1!InsuranceNextDueDate = txtInsuranceNextDueDate.text
If txtInsurancePercentage.text <> "" Then Rst1!InsurancePercentage = CDbl(txtInsurancePercentage.text)
If txtAnnualInsuranceCharge.text <> "" Then Rst1!AnnualInsuranceCharge = CDbl(txtAnnualInsuranceCharge.text)
If txtTotalYearlyInsurance.text <> "" Then Rst1!TotalYearlyInsurance = CDbl(txtTotalYearlyInsurance.text)
''

Rst1!Text1 = Text1.text
Rst1!Text2 = Text2.text
Rst1!Text3 = Text3.text

If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption

Rst1.Update
Rst1.Close
Conn1.Close

'check for new unit
If OldUnit <> cboUnit.text Then
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseOdbc
    Conn1.EstablishConnection rdDriverNoPrompt
        
    'update old unit record
    SQLStr1 = "SELECT Occupied, SageAccountNumber, TenantCompanyName FROM Units WHERE UnitNumber = '" & OldUnit & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
    
    Rst1.Edit
    Rst1!OCCUPIED = "N"
    Rst1!SageAccountNumber = ""
    Rst1!TenantCompanyName = ""
    Rst1.Update
    Rst1.Close
        
    'update new unit record
    SQLStr1 = "SELECT Occupied, SageAccountNumber, TenantCompanyName FROM Units WHERE UnitNumber = '" & NewUnit & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        
    Rst1.Edit
    Rst1!OCCUPIED = "Y"
    Rst1!SageAccountNumber = TenantCode
    Rst1!TenantCompanyName = TenantName
    Rst1.Update
    Rst1.Close
    
    'update tenant record
    SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
    
    Rst1.Edit
    Rst1!CurrentRental = NewUnit
    Rst1.Update
    Rst1.Close
    Conn1.Close
    
End If

Call DisableBoxes

Call GetTenantsWithLease

cboTenant1.text = TenantCode & " / " & TenantName

MsgBox "Your changes have been saved", vbOKOnly + vbInformation, "Saved"

cboTenant1.Enabled = True

cmdAddNew.Visible = True
cmdAddNew.TabIndex = 25
'mnuAdd.Enabled = True
cmdDelete.Visible = True
cmdDelete.TabIndex = 26
'mnuDelete.Enabled = True
cmdEdit.Visible = True
'mnuEdit.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False

End Sub

Private Sub cmdSaveNew_Click()
Dim i As Integer

If txtLeaseID.text = "" Then
   MsgBox "You must enter a Lease reference to continue!", vbOKOnly + vbCritical, "Reference Required"
   txtLeaseID.SetFocus
   Exit Sub
End If

If cboType.text = "" Then
   MsgBox "You must enter the Lease Type to continue!", vbOKOnly + vbCritical, "Lease Type missing"
   tabLease.Tab = 0
   cboType.SetFocus
   Exit Sub
End If

If txt4.text = "" Then
   MsgBox "You must enter the Year End date to continue!", vbOKOnly + vbCritical, "Year End date missing"
   tabLease.Tab = 0
   txt4.SetFocus
   Exit Sub
End If

If txtLeaseStDt.text = "" Then
   MsgBox "You must enter a Lease Start Date!", vbOKOnly + vbCritical, "Date Required"
   tabLease.Tab = 0
   txtLeaseStDt.SetFocus
   Exit Sub
End If

If txt3.text = "" Then
   MsgBox "You must enter a Lease End Date!", vbOKOnly + vbCritical, "Date Required"
   tabLease.Tab = 0
   txt3.SetFocus
   Exit Sub
End If

If cbo0.text = "Yes" Then
   If cboRentChargeDept.Column(0) = "" Then
      MsgBox "You must select a Department of rent.", vbOKOnly + vbCritical, "Rent - Department"
      tabLease.Tab = 1
      cboRentChargeDept.SetFocus
      Exit Sub
   End If
   If cboFreqBR.ListIndex = 12 Or cboFreqBR.ListIndex = -1 Then
      MsgBox "You must select a Rent Frequency!", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 1
      cboFreqBR.SetFocus
      Exit Sub
   End If
   If txt5.text = "" Then
      MsgBox "You must enter a Rent Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 1
      txt5.SetFocus
      Exit Sub
   End If
End If

If cbo1.text = "Yes" Then
   If cboServiceChargeDept.text = "" Then
      MsgBox "You must select a department for the service charge!", vbOKOnly + vbCritical, "Service Charge - Department"
      tabLease.Tab = 4
      cboServiceChargeDept.SetFocus
      Exit Sub
   End If
    If txt9.text = "" Then
        MsgBox "You must enter a Payable From Date for Service Charge!", vbOKOnly + vbCritical, "Date Required"
        Exit Sub
    End If
    If cboFreqSC.ListIndex = 12 Or cboFreqSC.ListIndex = -1 Then
        MsgBox "You must select a Service Charge Frequency!", vbOKOnly + vbCritical, "Frequency Required"
        Exit Sub
    End If
End If

If cbo2.text = "Yes" Then
   If cboIntChargeDept.text = "" Then
      MsgBox "You must select a department for the interest charge!", vbOKOnly + vbCritical, "Interest Charge - Department"
      tabLease.Tab = 5
      cboIntChargeDept.SetFocus
      Exit Sub
   End If
   If txt11.text = "" Then
      MsgBox "You must enter number of days interest will charge after!", vbOKOnly + vbCritical, "Interest Charge"
      tabLease.Tab = 5
      txt11.SetFocus
      Exit Sub
   End If
   If cboIntDemandType.text = "Yes" Then
      MsgBox "You must select interest demand type!", vbOKOnly + vbCritical, "Demand Type"
      tabLease.Tab = 5
      cboIntDemandType.SetFocus
      Exit Sub
   End If
End If

If cboInsurancePayable.text = "Yes" Then
   If cboInsuranceDept.text = "" Then
      MsgBox "You must select department of insurance!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceDept.SetFocus
      Exit Sub
   End If
   If txtInsuranceStartDate.text = "" Then
      MsgBox "You must enter insurance start date!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      txtInsuranceStartDate.SetFocus
      Exit Sub
   End If
   If txtInsuranceEndDate.text = "" Then
      MsgBox "You must enter insurance end date!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      txtInsuranceEndDate.SetFocus
      Exit Sub
   End If
   If cboInsuranceFrequency.text = "" Then
      MsgBox "You must select insurance frequency!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceFrequency.SetFocus
      Exit Sub
   End If
   If cboInsuranceDemandType.text = "" Then
      MsgBox "You must select insurance demand type!", vbOKOnly + vbCritical, "Insurance"
      tabLease.Tab = 9
      cboInsuranceDemandType.SetFocus
      Exit Sub
   End If
End If

If cboUnit.text = "" Then
    MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
    Exit Sub
Else
    NewUnit = cboUnit.text
End If

For i = 2 To 10
    If Mid(cboTenant2.text, i, 3) = " / " Then
        TenantCode = Left(cboTenant2.text, i - 1)
        TenantName = Mid(cboTenant2.text, i + 3, Len(cboTenant2.text))
    End If
Next i

GetGlobalDataForProperty (Left(Trim(cboUnit.text), 8))

'save the details to a new record
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseOdbc
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM LeaseDetails where LeaseID = '" & txtLeaseID.text & "'"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If Not Rst1.EOF Or Not Rst1.BOF Then
    MsgBox "Cannot save the lease. The lease reference already exist.", vbInformation, "Save Lease"
    txtLeaseID.text = ""
    txtLeaseID.SetFocus
    Exit Sub
End If

Rst1.AddNew
Rst1!LeaseID = txtLeaseID.text
Rst1!SageAccountNumber = TenantCode
Rst1!CompanyName = TenantName
If chkSubLease.Value Then
    If cboHeadLease.text <> "" Then Rst1!HeadLease = cboHeadLease.text
End If

Rst1!UnitNumber = Left(Trim(cboUnit.text), 8)
Rst1!TYPEOFSTORE = cboType.text
Rst1!StartDate = txtLeaseStDt.text
Rst1!EndDate = txt3.text
Rst1!YearEnd = txt4.text
If txt5.text <> "" Then Rst1!BRStartDate = txt5.text
For i = 2 To 3
    If Mid(cboFreqBR.text, i, 1) = "-" Then Rst1!BRfrequency = CInt(Left(cboFreqBR.text, i - 1))
Next i
'Rst1!BRfrequency = cboFreqBR.ListIndex + 1
If txt6.text <> "" Then Rst1!BRTotal = CDbl(txt6.text)
If txt7.text <> "" Then Rst1!BRNextDueDate = txt7.text
If txt8.text <> "" Then Rst1!BRAmount = CDbl(txt8.text)
Rst1!BRDemandType = CByte(IIf(cboBRDemandType.text <> "", cboBRDemandType.ListIndex + 1, 0))
Rst1!BRPayable = IIf(cbo0.text = "No", "N", "Y")
Rst1!SCPayable = IIf(cbo1.text = "No", "N", "Y")

If cbo1.text <> "No" Then
   For i = 2 To 3
       If Mid(cboFreqSC.text, i, 1) = "-" Then Rst1!SCfrequency = CInt(Left(cboFreqSC.text, i - 1))
   Next i
   'Rst1!SCfrequency = cboFreqSC.ListIndex + 1
   
   Rst1!SCPayableFrom = txt9.text
   If txt10b.text <> "" Then Rst1!SCNextDueDate = txt10b.text
   
   'If txtPPSqFoot.text <> "" Then Rst1!SCPricePerSqFoot = CDbl(txtPPSqFoot.text)
   If optSqFoot.Value Then
      Rst1!SCPricePerSqFoot = CDbl(txtPPSqFoot.text)
      Rst1!SCPercentage = Null
      Rst1!SCTotal = Null
   End If
   
   'If txtSCPercentage.text <> "" Then Rst1!SCPercentage = CDbl(txtSCPercentage.text)
   If optPercentage.Value Then
      Rst1!SCPercentage = CDbl(txtSCPercentage.text)
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCTotal = Null
   End If
   
   'If txtAmount.text <> "" Then Rst1!SCTotal = CDbl(txtAmount.text)
   If optFixedTotal.Value Then
      Rst1!SCTotal = CDbl(txtAnnualService.text)
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCPercentage = Null
   End If
   
   If optGlobalData.Value Then
      Rst1!SCTotal = Null
      Rst1!SCPricePerSqFoot = Null
      Rst1!SCPercentage = Null
   End If
   
   'If txt10c.text <> "" Then Rst1!SCAmount = CDbl(txt10c.text)
   If cbo1.text = "Yes" Then Rst1!SCAmount = CDbl(txtFinalAmout.text)
   
   If txt10.text <> "" Then Rst1!SCTOLimit = CDbl(txt10.text)
   Rst1!SCDemandType = CByte(IIf(cboSCDemandType.text <> "", cboSCDemandType.ListIndex + 1, 0))
   Rst1!ServiceChargeDept = cboServiceChargeDept.text
End If

Rst1!InterestChargeable = IIf(cbo2.text = "No", "N", "Y")
If txt11.text <> "" Then Rst1!DaysAfterInterestPayable = CInt(txt11.text)
If txt12.text <> "" Then Rst1!AdditionalInterest = CDbl(txt12.text)
If txt12a.text <> "" Then Rst1!InterestChargedOn = CDbl(txt12a.text)
If txt13.text <> "" Then Rst1!InterestAmount = CDbl(txt13.text)
Rst1!IntDemandType = CByte(IIf(cboIntDemandType.text <> "", cboIntDemandType.ListIndex + 1, 0))
Rst1!IntChargeDept = cboIntChargeDept.text
If cbo3.text = "No" Then Rst1!BreakClause = "N"
If cbo3.text = "Yes" Then Rst1!BreakClause = "Y"
If txt14.text <> "" Then Rst1!BreakDate = txt14.text
If cboBreak.text <> "" Then Rst1!BreakType = cboBreak.text
If txt15.text <> "" Then Rst1!RentReviewDate = txt15.text
If txt16.text <> "" Then Rst1!RentIncreaseDate = txt16.text
If txt17.text <> "" Then Rst1!RentIncreaseAmount = CDbl(txt17.text)
Rst1!RentChargeDept = cboRentChargeDept.BoundColumn
If txt18.text <> "" Then Rst1!DateFlagDate = txt18.text
If txt19.text <> "" Then Rst1!DateFlagDescription = txt19.text
If txt20.text <> "" Then Rst1!Notes = txt20.text

'Insurance
If cboInsurancePayable.text = "No" Then Rst1!InsurancePayable = "N"
If cboInsurancePayable.text = "Yes" Then Rst1!InsurancePayable = "Y"

For i = 2 To 3
    If Mid(cboInsuranceFrequency.text, i, 1) = "-" Then Rst1!InsuranceFrequency = CInt(Left(cboInsuranceFrequency.text, i - 1))
Next i

If txtInsuranceStartDate.text <> "" Then Rst1!InsuranceStartDate = txtInsuranceStartDate.text
If txtInsuranceEndDate.text <> "" Then Rst1!InsuranceEndDate = txtInsuranceEndDate.text
Rst1!InsuranceDemandType = CByte(IIf(cboInsuranceDemandType.text <> "", cboInsuranceDemandType.ListIndex + 1, 0))
If txtInsuranceEachPeriod.text <> "" Then Rst1!InsuranceEachPeriod = CDbl(txtInsuranceEachPeriod.text)
If txtInsuranceNextDueDate.text <> "" Then Rst1!InsuranceNextDueDate = txtInsuranceNextDueDate.text
If txtInsurancePercentage.text <> "" Then Rst1!InsurancePercentage = CDbl(txtInsurancePercentage.text)
If txtAnnualInsuranceCharge.text <> "" Then Rst1!AnnualInsuranceCharge = CDbl(txtAnnualInsuranceCharge.text)
If txtTotalYearlyInsurance.text <> "" Then Rst1!TotalYearlyInsurance = CDbl(txtTotalYearlyInsurance.text)
Rst1!InsuranceDept = cboInsuranceDept.text
''

If Text1.text <> "" Then Rst1!Text1 = Text1.text
If Text2.text <> "" Then Rst1!Text2 = Text2.text
If Text3.text <> "" Then Rst1!Text3 = Text3.text
If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption



Rst1.Update
Rst1.Close
Conn1.Close

'check for new unit
If OldUnit <> cboUnit.text Then
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseOdbc
    Conn1.EstablishConnection rdDriverNoPrompt

    If OldUnit = "" Then ' no old unit so only need to update new unit
        'update tenant record
        SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

        Rst1.Edit
        Rst1!CurrentRental = NewUnit
        Rst1.Update
        Rst1.Close

        SQLStr1 = "UPDATE Units " & _
                  "SET OCCUPIED = 'Y', " & _
                     "SageAccountNumber = '" & TenantCode & "', " & _
                     "TenantCompanyName = '" & TenantName & "' " & _
                  "WHERE UNITNUMBER ='" & Left(NewUnit, 8) & "';"
         Conn1.Execute SQLStr1

    Else 'old unit different to new
        'update tenant record
        SQLStr1 = "UPDATE Tenants " & _
                  "SET CurrentRental = '" & NewUnit & "' " & _
                  "WHERE SageAccountNumber = '" & TenantCode & "'"
        Conn1.Execute SQLStr1
         
        SQLStr1 = "UPDATE Units " & _
                  "SET OCCUPIED = 'N', " & _
                     "SageAccountNumber = '', " & _
                     "TenantCompanyName = '' " & _
                  "WHERE UNITNUMBER ='" & OldUnit & "';"
         Conn1.Execute SQLStr1
        
        'update new unit record
        
        SQLStr1 = "UPDATE Units " & _
                  "SET OCCUPIED = 'N', " & _
                     "SageAccountNumber = '" & TenantCode & "', " & _
                     "TenantCompanyName = '" & TenantName & "' " & _
                  "WHERE UnitNumber = '" & NewUnit & "'"
         Conn1.Execute SQLStr1
    End If
    Conn1.Close
    
End If

Call DisableBoxes
cboTenant2.Visible = False
cboTenant1.Visible = True
cboTenant1.Enabled = True

Call GetTenantsWithLease

cboTenant1.text = TenantCode & " / " & TenantName

MsgBox "The new lease record has been saved", vbOKOnly + vbInformation, "Saved"

cmdAddNew.Visible = True
cmdAddNew.TabIndex = 25
'mnuAdd.Enabled = True
cmdDelete.Visible = True
cmdDelete.TabIndex = 26
'mnuDelete.Enabled = True
cmdEdit.Visible = True
'mnuEdit.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub flxRentAnalysis_Click()
   If cmdEditRentAnalysis.Enabled = False Then Exit Sub
   populateControl frmLease, flxRentAnalysis
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

Me.Caption = "Lease Information"

ConfigureFlxGrid flxRentAnalysis

Call EmptyBoxes
Call GetTenantsWithLease
Call DisableBoxes
cboTenant2.Visible = False
cboTenant1.Visible = True
cboTenant1.Enabled = True
Call FillCbos
BreachButtonMode DefaultMode
AssignmentButtonMode DefaultMode

Call FillcboType(cboBRDemandType, szaDemandtype)
Call FillcboType(cboSCDemandType, szaDemandtype)
Call FillcboType(cboIntDemandType, szaDemandtype)
Call FillcboType(cboInsuranceDemandType, szaDemandtype)

    
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
   
'   szaTenant = Split(cboTenant1.text, " / ")
'   gszSageAccountNumber = szaTenant(0)

   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
   
   'get all sage account numbers and company names from tenants.
   SQLStr2 = "SELECT * " & _
             "FROM RentAnalysis " & _
             "WHERE SAGEACCOUNTNUMBER = '" & gszSageAccountNumber & "' " & _
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

   conFlxGrid.Clear
   conFlxGrid.Cols = 6
   szFlxHeader$ = "|<Serial|<RentReviewDate|<RentIncreaseDate|>RentIncreaseAmount|ID"
   conFlxGrid.FormatString = szFlxHeader$

   conFlxGrid.ColWidth(0) = 400
   conFlxGrid.ColWidth(1) = 1500
   conFlxGrid.ColWidth(2) = 2500
   conFlxGrid.ColWidth(3) = 2500
   conFlxGrid.ColWidth(4) = 2500
   conFlxGrid.ColWidth(5) = 0       'ID
End Sub

Public Sub GetTenantsWithLease()
   Dim temp As String

   cboTenant2.Visible = False
   cboTenant1.Visible = True
   cboTenant1.Clear

   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT SageAccountNumber, CompanyName FROM LeaseDetails ORDER BY SageAccountNumber"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

       If Rst1.EOF = False Then
           While Rst1.EOF = False
               cboTenant1.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
               Rst1.MoveNext
           Wend
       End If
   Rst1.Close
   Conn1.Close

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

'On Error Resume Next
Dim temp As String
Dim a, b, k, j As Integer

cboTenant1.Visible = False
cboTenant2.Visible = True
cboTenant2.Clear

Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn2.CursorDriver = rdUseIfNeeded
Conn2.EstablishConnection rdDriverNoPrompt

'get all sage account numbers and company names from tenants.
SQLStr2 = "SELECT SageAccountNumber, CompanyName " & _
          "FROM Tenants " & _
          "WHERE Tenants.SageAccountNumber NOT IN " & _
              "(SELECT LeaseDetails.SageAccountNumber " & _
              "FROM LeaseDetails) " & _
          "ORDER BY SageAccountNumber"
Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)

While Rst2.EOF = False
    cboTenant2.AddItem Rst2!SageAccountNumber & " / " & Rst2!CompanyName
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

Public Sub DisableBoxes()
txtLeaseID.Enabled = False
chkSubLease.Enabled = False
cboUnit.Enabled = False
cboType.Enabled = False
chkSubLease.Enabled = False
txtLeaseStDt.Enabled = False
txt3.Enabled = False
txt4.Enabled = False
txt5.Enabled = False
txt6.Enabled = False
txt7.Enabled = False
cboFreqBR.Enabled = False
cbo0.Enabled = False
cbo1.Enabled = False
cboFreqSC.Enabled = False
txt9.Enabled = False
txt10.Enabled = False
'txtAmount.Enabled = False
txt10b.Enabled = False
'txtPPSqFoot.Enabled = False
cbo2.Enabled = False
txt11.Enabled = False
txt12.Enabled = False
txt12a.Enabled = False
cbo3.Enabled = False
cboBreak.Enabled = False
txt14.Enabled = False
txt15.Enabled = False
txt16.Enabled = False
txt17.Enabled = False
txt18.Enabled = False
txt19.Enabled = False
txt20.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
'txtSCPercentage.Enabled = False
txt7.Enabled = True
txt10b.Enabled = True

' Insurance
cboInsurancePayable.Enabled = False
txtInsuranceStartDate.Enabled = False
txtInsuranceEndDate.Enabled = False
cboInsuranceDemandType.Enabled = False
txtInsuranceEachPeriod.Enabled = False
txtInsuranceNextDueDate.Enabled = False
txtInsurancePercentage.Enabled = False
txtAnnualInsuranceCharge.Enabled = False
txtTotalYearlyInsurance.Enabled = False
'fmeInsurance.Enabled = False

cmdAddNew.Visible = True
cmdAddNew.TabIndex = 25
'mnuAdd.Enabled = True
cmdDelete.Visible = True
cmdDelete.TabIndex = 26
'mnuDelete.Enabled = True
cmdEdit.Visible = True
'mnuEdit.Enabled = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False

cboBRDemandType.Enabled = False
cboSCDemandType.Enabled = False
cboIntDemandType.Enabled = False
End Sub

Public Sub EnableBoxes()

chkSubLease.Enabled = True
cboUnit.Enabled = True
cboType.Enabled = True
txtLeaseStDt.Enabled = True
txt3.Enabled = True
txt4.Enabled = True
txt5.Enabled = True
txt6.Enabled = True
cboFreqBR.Enabled = True
cbo0.Enabled = True
cbo1.Enabled = True
cboFreqSC.Enabled = True
txt9.Enabled = True
txt10.Enabled = True
'txtPPSqFoot.Enabled = True
cbo2.Enabled = True
txt11.Enabled = True
txt12.Enabled = True
txt12a.Enabled = True
cbo3.Enabled = True
cboBreak.Enabled = True
txt14.Enabled = True
txt15.Enabled = True
txt16.Enabled = True
txt17.Enabled = True
txt18.Enabled = True
txt19.Enabled = True
txt20.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
txt7.Enabled = True
txt10b.Enabled = True

cboBRDemandType.Enabled = True
cboSCDemandType.Enabled = True
cboIntDemandType.Enabled = True
End Sub


Public Sub EmptyBoxes()
   txtLeaseID.text = ""
   txt7.Enabled = True
   txt8.Enabled = True
   txt10b.Enabled = True
   txt10c.Enabled = True
   cboUnit.text = ""
   txtLeaseStDt.text = ""
   txt3.text = ""
   txt4.text = ""
   txt5.text = ""
   txt6.text = ""
   txt7.text = ""
   txt8.text = ""
   txt9.text = ""
   txt10.text = ""
   txtAmount.text = ""
   txt10b.text = ""
   txt10c.text = ""
   txtPPSqFoot.text = ""
   txt11.text = ""
   txt12.text = ""
   txt12a.text = ""
   txt13.text = ""
   txt14.text = ""
   txt15.text = ""
   txt16.text = ""
   txt17.text = ""
   txt18.text = ""
   txt19.text = ""
   txt20.text = ""
   Text1.text = ""
   Text2.text = ""
   Text3.text = ""
   txt7.Enabled = False
   txt8.Enabled = False
   txt10b.Enabled = False
   txt10c.Enabled = False
   txtSCPercentage.text = ""

   ' Insurance
   txtInsuranceNextDueDate.text = ""
   txtInsurancePercentage.text = ""
   txtInsuranceStartDate.text = ""
   txtTotalYearlyInsurance.text = ""
   txtAnnualInsuranceCharge.text = ""
   txtInsuranceEachPeriod.text = ""

   cboInsurancePayable.Enabled = True
   txtInsuranceStartDate.Enabled = True
   txtInsuranceEndDate.Enabled = True
   cboInsuranceDemandType.Enabled = True
   txtInsuranceEachPeriod.Enabled = False
   txtInsuranceNextDueDate.Enabled = False
   txtInsurancePercentage.Enabled = True
   txtAnnualInsuranceCharge.Enabled = True
   txtTotalYearlyInsurance.Enabled = False
   cboInsuranceDemandType.text = ""
End Sub

Public Sub FillCbos()
   Dim i As Integer
   
   'Fill the yes / no cbos
   cbo0.AddItem "No", 0
   cbo0.AddItem "Yes", 1
   cbo1.AddItem "No", 0
   cbo1.AddItem "Yes", 1
   cbo2.AddItem "No", 0
   cbo2.AddItem "Yes", 1
   cbo3.AddItem "No", 0
   cbo3.AddItem "Yes", 1
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
   
   'cboFreqBR.AddItem "<Not Selected>" & "-" & 0, 0
   'cboFreqSC.AddItem "<Not Selected>" & "-" & 0, 0
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
           cboHeadLease.AddItem Rst1!LeaseID, i
           i = i + 1
           Rst1.MoveNext
       Wend
   End If

   Rst1.Close

   SQLStr1 = "SELECT UnitNumber,UnitName FROM Units"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   i = 0
   If Rst1.EOF = False Then
       While Rst1.EOF = False
           cboUnit.AddItem Rst1!UnitNumber & " - " & Rst1!UnitName, i
           i = i + 1
           Rst1.MoveNext
       Wend
   End If
     
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
   
   Rst1.Close
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
'
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
          "WHERE LeaseDetails.SageAccountNumber = '" & gszSageAccountNumber & "' AND " & _
          "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
          "Units.PropertyID = Property.PropertyID AND " & _
          "Client.ClientId = Property.ClientId"
'Debug.Print SQLStr1
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

If IsNull(Rst1!LeaseID) = False Then
    txtLeaseID.text = Rst1!LeaseID
Else
    txtLeaseID.text = ""
End If

'Check for sub lease.'
Dim HeadLease As String
If IsNull(Rst1!HeadLease) = False Then
   HeadLease = Rst1!HeadLease
   If HeadLease = "" Then
       chkSubLease.Value = 0
   Else
       chkSubLease.Value = 1
   End If
End If

'Fill text boxes with lease details.
cboUnit.text = Rst1!UnitNumber
gszCurUnitNum = cboUnit.text
txtClient.text = IIf(IsNull(Rst1!ClientName), "", Rst1!ClientName)
txtProperty.text = IIf(IsNull(Rst1!PropertyName), "", Rst1!PropertyName)

cboType.text = IIf(IsNull(Rst1!TYPEOFSTORE), "", Rst1!TYPEOFSTORE)
txtLeaseStDt.text = IIf(IsNull(Rst1!StartDate), "", Rst1!StartDate)
txt3.text = IIf(IsNull(Rst1!EndDate), "", Rst1!EndDate)
txt4.text = IIf(IsNull(Rst1!YearEnd), "", Rst1!YearEnd)

If IsNull(Rst1!BRStartDate) = False Then txt5.text = Rst1!BRStartDate
If IsNull(Rst1!BRTotal) = False Then txt6.text = Rst1!BRTotal
If IsNull(Rst1!BRNextDueDate) = False Then txt7.text = Rst1!BRNextDueDate
If IsNull(Rst1!BRAmount) = False Then txt8.text = Rst1!BRAmount
If Not IsNull(Rst1!BRDemandType) Then cboBRDemandType.ListIndex = (CByte(Rst1!BRDemandType) - 1)
cboRentChargeDept.text = IIf(IsNull(Rst1!RentChargeDept), "", DeptName(IIf(IsNull(Rst1!RentChargeDept) Or Rst1!RentChargeDept = "", 1, Rst1!RentChargeDept)))
If Rst1!BRfrequency > 0 Then
   cboFreqBR.text = cboFreqBR.List(Rst1!BRfrequency - 1)
   bf = Rst1!BRfrequency - 1
Else
   bf = 0
End If
If Rst1!BRPayable = "N" Then cbo0.text = "No"
If Rst1!BRPayable = "Y" Then cbo0.text = "Yes"
If Rst1!SCPayable = "N" Then cbo1.text = "No"
If Rst1!SCPayable = "Y" Then cbo1.text = "Yes"
cboServiceChargeDept.text = IIf(IsNull(Rst1!ServiceChargeDept), "", DeptName(IIf(IsNull(Rst1!ServiceChargeDept) Or Rst1!ServiceChargeDept = "", 1, Rst1!ServiceChargeDept)))
If IsNull(Rst1!SCPayableFrom) = False Then txt9.text = Rst1!SCPayableFrom
If IsNull(Rst1!SCTOLimit) = False Then txt10.text = Rst1!SCTOLimit
'If Not IsNull(Rst1!SCDemandType) Then cboSCDemandType.text = DemandType(CByte(Rst1!SCDemandType))
If Not IsNull(Rst1!SCDemandType) Then cboSCDemandType.ListIndex = CByte(Rst1!SCDemandType) - 1

If Rst1!SCPayable = "Y" Then
   If Val(IIf(IsNull(Rst1!SCPercentage), 0, Rst1!SCPercentage)) > 0 Then
      txtSCPercentage.text = Rst1!SCPercentage
      optPercentage.Value = True
      txtAmount.text = Format(Rst1!SCPercentage, "0.00")
   Else
      If Val(IIf(IsNull(Rst1!SCPricePerSqFoot), 0, Rst1!SCPricePerSqFoot)) > 0 Then
         txtPPSqFoot.text = Rst1!SCPricePerSqFoot
         optSqFoot.Value = True
         txtAmount.text = Format(Rst1!SCPricePerSqFoot, "0.00")
      Else
         If Val(IIf(IsNull(Rst1!SCTotal), 0, Rst1!SCTotal)) > 0 Then
            txtAmount.text = Format(Rst1!SCTotal, "0.00")
            optFixedTotal.Value = True
            txtAmount.text = Format(Rst1!SCTotal, "0.00")
         Else
            optGlobalData.Value = True
         End If
      End If
   End If
End If
If IsNull(Rst1!SCAmount) = False Then txt10c.text = Rst1!SCAmount
txtFinalAmout.text = IIf(IsNull(Rst1!SCAmount), "0.00", Rst1!SCAmount)

If IsNull(Rst1!SCNextDueDate) = False Then txt10b.text = Rst1!SCNextDueDate

If Not IsNull(Rst1!SCfrequency) Then
    If Rst1!SCfrequency > 0 Then
        cboFreqSC.text = cboFreqSC.List(Rst1!SCfrequency - 1)
        scf = Rst1!SCfrequency - 1
    Else
        scf = 0
    End If
End If

If Rst1!InterestChargeable = "N" Then cbo2.text = "No"
If Rst1!InterestChargeable = "Y" Then cbo2.text = "Yes"
If IsNull(Rst1!DaysAfterInterestPayable) = False Then txt11.text = Rst1!DaysAfterInterestPayable
cboIntChargeDept.text = IIf(IsNull(Rst1!IntChargeDept), "", DeptName(IIf(IsNull(Rst1!IntChargeDept) Or Rst1!IntChargeDept = "", 1, Rst1!IntChargeDept)))
If IsNull(Rst1!AdditionalInterest) = False Then txt12.text = Rst1!AdditionalInterest
If IsNull(Rst1!InterestChargedOn) = False Then txt12a.text = Rst1!InterestChargedOn
If IsNull(Rst1!InterestAmount) = False Then txt13.text = Rst1!InterestAmount
'If Not IsNull(Rst1!IntDemandType) Then cboIntDemandType.text = DemandType(CByte(Rst1!IntDemandType))
If Not IsNull(Rst1!IntDemandType) Then cboIntDemandType.ListIndex = CByte(Rst1!IntDemandType) - 1
'                                       cboBRDemandType.ListIndex = CByte(Rst1!BRDemandType) - 1
If Rst1!BreakClause = "N" Then cbo3.text = "No"
If Rst1!BreakClause = "Y" Then cbo3.text = "Yes"
If IsNull(Rst1!BreakType) = False Then cboBreak.text = Rst1!BreakType
If IsNull(Rst1!BreakDate) = False Then txt14.text = Rst1!BreakDate
If IsNull(Rst1!RentReviewDate) = False Then txt15.text = Rst1!RentReviewDate
If IsNull(Rst1!RentIncreaseDate) = False Then txt16.text = Rst1!RentIncreaseDate
If IsNull(Rst1!RentIncreaseAmount) = False Then txt17.text = Rst1!RentIncreaseAmount
If IsNull(Rst1!DateFlagDate) = False Then txt18.text = Rst1!DateFlagDate
If IsNull(Rst1!DateFlagDescription) = False Then txt19.text = Rst1!DateFlagDescription
If IsNull(Rst1!Notes) = False Then txt20.text = Rst1!Notes

' Insurance
If Rst1!InsurancePayable = "N" Then cboInsurancePayable.text = "No"
If Rst1!InsurancePayable = "Y" Then cboInsurancePayable.text = "Yes"

If Not IsNull(Rst1!InsuranceFrequency) Then
    If Rst1!InsuranceFrequency > 0 Then
        cboInsuranceFrequency.text = cboInsuranceFrequency.List(Rst1!InsuranceFrequency - 1)
        'scf = Rst1!SCfrequency - 1
    End If
End If
cboInsuranceDept.text = IIf(IsNull(Rst1!InsuranceDept), "", DeptName(IIf(IsNull(Rst1!InsuranceDept) Or Rst1!InsuranceDept = "", 1, Rst1!InsuranceDept)))
If IsNull(Rst1!InsuranceStartDate) = False Then txtInsuranceStartDate.text = Rst1!InsuranceStartDate
If IsNull(Rst1!InsuranceEndDate) = False Then txtInsuranceEndDate.text = Rst1!InsuranceEndDate
'If IsNull(Rst1!InsuranceDemandType) = False Then cboInsuranceDemandType.text = DemandType(CByte(Rst1!InsuranceDemandType))
If IsNull(Rst1!InsuranceDemandType) = False Then cboInsuranceDemandType.ListIndex = CByte(Rst1!InsuranceDemandType) - 1

If IsNull(Rst1!InsuranceEachPeriod) = False Then txtInsuranceEachPeriod.text = Rst1!InsuranceEachPeriod
If IsNull(Rst1!InsuranceNextDueDate) = False Then txtInsuranceNextDueDate.text = Rst1!InsuranceNextDueDate
If IsNull(Rst1!InsurancePercentage) = False Then txtInsurancePercentage.text = Rst1!InsurancePercentage
If IsNull(Rst1!AnnualInsuranceCharge) = False Then txtAnnualInsuranceCharge.text = Rst1!AnnualInsuranceCharge
If IsNull(Rst1!TotalYearlyInsurance) = False Then txtTotalYearlyInsurance.text = Rst1!TotalYearlyInsurance
''

If IsNull(Rst1!Text1) = False Then Text1.text = Rst1!Text1
If IsNull(Rst1!Text2) = False Then Text2.text = Rst1!Text2
If IsNull(Rst1!Text3) = False Then Text3.text = Rst1!Text3
If IsNull(Rst1!SuppCaption1) = False Then lblSupplementary1.Caption = Rst1!SuppCaption1
If IsNull(Rst1!SuppCaption2) = False Then lblSupplementary2.Caption = Rst1!SuppCaption2
If IsNull(Rst1!SuppCaption3) = False Then lblSupplementary3.Caption = Rst1!SuppCaption3


Rst1.Close
Conn1.Close
OldUnit = cboUnit.text
PopulateBreaches
PopulateAssignments
End Sub

Public Sub PopulateBreaches()
   
'On Error Resume Next
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
         'txtUnitAddress1.text = strSQLQuery_
              

SetBreachGrid
populateGrid adoConn, sSQLQuery_, gridBreach
adoConn.Close
Set adoConn = Nothing
   
End Sub

Public Sub PopulateAssignments()
'On Error Resume Next
'Set the RDO Connections to the dataset
Dim sSQLQuery_ As String
 
Dim adoConn As New ADODB.Connection
adoConn.Open "DSN=" & Adsn & ";UID=;PWD="


sSQLQuery_ = "SELECT LeaseAssignments.AssignmentID, " & _
      "LeaseAssignments.AssignDate, " & _
      "LeaseAssignments.Tenant " & _
      "FROM LeaseAssignments " & _
      "WHERE LeaseAssignments.LeaseID = '" & txtLeaseID.text & "' "

SetAssignmentGrid
populateGrid adoConn, sSQLQuery_, gridAssignment

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
                    
         'txtUnitAddress1.text = strSQLQuery_
         
   Set rstBreach = conBreach.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

   gridBreach.Clear
   gridBreach.Rows = 2
   gridBreach.Cols = 7
    
   
   gridBreach.ColWidth(0) = 1
   gridBreach.ColWidth(1) = cboBreachType.Width + cmdSetBreachType.Width
   gridBreach.ColWidth(2) = txtCommenceDate.Width
   gridBreach.ColWidth(3) = txtInitiatedBy.Width
   gridBreach.ColWidth(4) = chkResolved.Width
   gridBreach.ColWidth(5) = txtDateReceived.Width
   gridBreach.ColWidth(6) = txtReceivedBy.Width
   
      
   Dim oColumn As rdoColumn
   Dim iColumn As Integer
   iColumn = 0
   
   gridBreach.Cols = rstBreach.rdoColumns.Count
   For Each oColumn In rstBreach.rdoColumns
        gridBreach.TextMatrix(0, iColumn) = oColumn.Name
        iColumn = iColumn + 1
   Next oColumn
   
   'SetMaintenanceHistoryControl
      
   rstBreach.Close
   conBreach.Close
   Set rstBreach = Nothing
   Set conBreach = Nothing
   
   'cmdSelected.Enabled = True

End Sub

Public Sub SetAssignmentGrid()
   
   Dim conAssignment As New RDO.rdoConnection
   Dim rstAssignment As rdoResultset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
   conAssignment.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conAssignment.CursorDriver = rdUseIfNeeded
   conAssignment.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT LeaseAssignments.AssignmentID, " & _
      "LeaseAssignments.AssignDate, " & _
      "LeaseAssignments.Tenant " & _
      "FROM LeaseAssignments " & _
      "WHERE LeaseAssignments.LeaseID = '" & txtLeaseID.text & "' "
                    
         'txtUnitAddress1.text = strSQLQuery_
         
   Set rstAssignment = conAssignment.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

   gridAssignment.Clear
   gridAssignment.Rows = 2
   gridAssignment.Cols = 3
    
   
   gridAssignment.ColWidth(0) = 1
   gridAssignment.ColWidth(1) = txtAssignDate.Width
   gridAssignment.ColWidth(2) = txtTenant.Width
   
      
   Dim oColumn As rdoColumn
   Dim iColumn As Integer
   iColumn = 0
   
   gridAssignment.Cols = rstAssignment.rdoColumns.Count
   For Each oColumn In rstAssignment.rdoColumns
        gridAssignment.TextMatrix(0, iColumn) = oColumn.Name
        iColumn = iColumn + 1
   Next oColumn
   
   'SetMaintenanceHistoryControl
      
   rstAssignment.Close
   conAssignment.Close
   Set rstAssignment = Nothing
   Set conAssignment = Nothing

   'cmdSelected.Enabled = True
End Sub

Public Sub GetUnits()
    Dim i As Integer

    'Get all the unoccupied units and put in cbounit
    cboUnit.Clear

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT UnitNumber FROM Units WHERE Occupied = 'N' ORDER BY UnitNumber"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    If Rst1.EOF = False Then
        While Rst1.EOF = False
            cboUnit.AddItem Rst1!UnitNumber
            Rst1.MoveNext
        Wend
    End If

    Rst1.Close

    'Get unit of current tenant and put in cbounit.text
    If TenantCode <> "" Then
        SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
        If IsNull(Rst1!CurrentRental) Then OldUnit = "" Else OldUnit = Rst1!CurrentRental
    
        Rst1.Close
        Conn1.Close
        If OldUnit = "" Then
            cboUnit.text = ""
        Else
            cboUnit.AddItem OldUnit, 0
            cboUnit.text = cboUnit.List(0)
        End If
    Else
        Conn1.Close
    End If
    
End Sub

Public Sub AddNew()
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

   cboTenant1.Visible = False
   cboTenant2.Visible = True
   cboTenant2.Enabled = True

   Call EmptyBoxes
   Call GetTenantsWithoutLease
   Call EnableBoxes

   Call FillcboType(cboBRDemandType, szaDemandtype)
   Call FillcboType(cboSCDemandType, szaDemandtype)
   Call FillcboType(cboIntDemandType, szaDemandtype)
   Call FillcboType(cboInsuranceDemandType, szaDemandtype)

   cboUnit.Enabled = True
   txtLeaseID.Enabled = True
   txtLeaseID.text = ""

   cmdAddNew.Visible = False

   cmdDelete.Visible = False

   cmdEdit.Visible = False

   cmdCancelNew.Visible = True
   cmdCancelNew.TabIndex = 26
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub mnuAdd_Click()

Call AddNew

End Sub

Public Sub Edit()

'LoadDept
End Sub

Private Sub mnuDelete_Click()

Call Delete

End Sub

Private Sub mnuDemands_Click()

Load frmDemands
Unload Me
frmDemands.Show

End Sub

Private Sub mnuEdit_Click()

Call Edit

End Sub

Public Sub Delete()

Dim Response

If cboTenant1.text = "" Then
    MsgBox "You must select a lease to delete", vbOKOnly + vbCritical, "No Lease selected"
    Exit Sub
Else
    Response = MsgBox("Are you sure you want to delete the lease for tenant: " & TenantName, vbYesNo + vbQuestion, "Delete Lease")
    If Response = vbYes Then
        Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
        Conn1.CursorDriver = rdUseOdbc
        Conn1.EstablishConnection rdDriverNoPrompt
        
        SQLStr1 = "SELECT * FROM LeaseDetails WHERE LeaseID = '" & txtLeaseID.text & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        
        Rst1.Delete
        Rst1.Close
        Conn1.Close
        
        Call EmptyBoxes
        Call GetTenantsWithLease
    End If
End If

End Sub

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

'Load frmMain
Unload Me
'frmMain.Show

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

Private Sub TabStrip1_Click()
   Frame1(mintCurFrame).Visible = False
End Sub

Private Sub gridSubUnit_Click()

End Sub

Private Sub gridAssignment_Click()
AssignmentButtonMode GridRowOnSelection
End Sub

Private Sub gridAssignment_RowColChange()
populateControl frmLease, gridAssignment
End Sub

Private Sub gridBreach_Click()
BreachButtonMode GridRowOnSelection
End Sub

Private Sub gridBreach_RowColChange()
populateControl frmLease, gridBreach
End Sub

Private Sub lblSupplementary1_Click()
txtSuppCaption1.Visible = True
txtSuppCaption1.Left = lblSupplementary1.Left
'txtSuppCaption1.Top = lblSupplementary1.Top
'txtSuppCaption1.Width = lblSupplementary1.Width
txtSuppCaption1.text = lblSupplementary1.Caption
txtSuppCaption1.SetFocus

End Sub

Private Sub lblSupplementary2_Click()
txtSuppCaption2.Visible = True
txtSuppCaption2.Left = lblSupplementary2.Left
'txtSuppCaption2.Top = lblSupplementary2.Top
'txtSuppCaption2.Width = lblSupplementary2.Width
txtSuppCaption2.text = lblSupplementary2.Caption
txtSuppCaption2.SetFocus

End Sub

Private Sub lblSupplementary3_Click()
txtSuppCaption3.Visible = True
txtSuppCaption3.Left = lblSupplementary3.Left
'txtSuppCaption3.Top = lblSupplementary3.Top
'txtSuppCaption3.Width = lblSupplementary3.Width
txtSuppCaption3.text = lblSupplementary3.Caption
'lblSupplementary3.Visible = False

txtSuppCaption3.SetFocus

End Sub

Private Sub optAnnualInsuranceCharge_Click()
 txtInsurancePercentage.Locked = True
 txtAnnualInsuranceCharge.Locked = False
 txtAnnualInsuranceCharge.SetFocus
 txtInsurancePercentage.text = "0.00"
End Sub

Private Sub optFixedTotal_Click()
   txtSCPercentage.Locked = True
   txtPPSqFoot.Locked = True
   txtAnnualService.Locked = False
   txtAnnualService.SetFocus
End Sub

Private Sub optGlobalData_Click()
   MousePointer = vbHourglass

   Dim rdoConn As New RDO.rdoConnection
   Dim rstRst As rdoResultset
   Dim szSQL As String

   txtGlobalAmount.text = Format(GetPPSF(gszCurUnitNum) * GetUnitTA(gszCurUnitNum), "0.00")
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
'
   If cbo1.text = "Yes" Then Call CalculateSC

   Set rstRst = Nothing
   Set rdoConn = Nothing
   
   MousePointer = vbDefault
End Sub

Private Sub Option4_Click()
 txtSCPercentage.Locked = False
txtPPSqFoot.Locked = True
txtAnnualService.Locked = True
'   txtSCPercentage.Enabled
txtSCPercentage.SetFocus
End Sub

Private Sub optInsurancePercentage_Click()
 txtInsurancePercentage.Locked = False
 txtAnnualInsuranceCharge.Locked = True
 txtInsurancePercentage.SetFocus
 txtAnnualInsuranceCharge.text = "0.00"
End Sub

Private Sub optPercentage_Click()
   txtSCPercentage.Locked = False
   txtPPSqFoot.Locked = True
   txtAnnualService.Locked = True
'   txtSCPercentage.Enabled
   txtSCPercentage.SetFocus
End Sub

Private Sub optSqFoot_Click()
   txtSCPercentage.Locked = True
   txtPPSqFoot.Locked = False
   txtAnnualService.Locked = True
   txtPPSqFoot.SetFocus
End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text6_Change()

End Sub

Private Sub txt10_LostFocus()

    If txt10.text <> "" Then
        If NumberCheck2(txt10.text) = False Then
            txt10.text = ""
        Else
            txt10.text = Round(CDbl(txt10.text), 2)
        End If
    End If

End Sub


Private Sub txt10b_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txt10b

End Sub

Private Sub txt14_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txt14
End Sub

Private Sub txt14_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txt14, KeyAscii
End Sub

Private Sub txt3_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txt3
End Sub

Private Sub txt3_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txt3, KeyAscii
End Sub

Private Sub txt4_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txt4
End Sub

Private Sub txt4_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txt4, KeyAscii
End Sub

Private Sub txt5_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txt5
End Sub

Private Sub txt5_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txt5, KeyAscii
End Sub

Private Sub txt9_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txt9
End Sub

Private Sub txtAnnualInsuranceCharge_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtAnnualInsuranceCharge_LostFocus()
  If txtAnnualInsuranceCharge.Locked Then Exit Sub
   Dim Area As String, Total As Double
   
   MousePointer = vbHourglass
'
   txtAnnualInsuranceCharge.text = Format(IIf(txtAnnualInsuranceCharge.text = "", 0, txtAnnualInsuranceCharge.text), "0.00")
'
   Total = CDbl(txtAnnualInsuranceCharge.text)
   txtTotalYearlyInsurance.text = Format(Total, "0.00")
'
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & (cboInsuranceFrequency.ListIndex + 1) & ";"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   txtInsuranceEachPeriod.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")
   
   Rst1.Close
   Conn1.Close
'
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
   If cbo1.text = "Yes" Then Call CalculateSC

   MousePointer = vbDefault
End Sub


Private Sub txtAssignDate_LostFocus()
'Added By Asif. 13/01/2006
TextBoxFormatDate txtAssignDate
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


Private Sub txtDateReceived_Change()
   'Added by Samrat. 16.01.2006
   TextBoxChangeDate txtDateReceived
End Sub

Private Sub txtDateReceived_KeyPress(KeyAscii As Integer)
   'Added by Samrat. 16.01.2006
   TextBoxKeyPrsDate txtDateReceived, KeyAscii
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
   
   Dim Total As Double, TotalServiceCharge As String
   
   txtInsurancePercentage.text = Format(IIf(txtInsurancePercentage.text = "", 0, txtInsurancePercentage.text), "0.0000")
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
'Added By Asif. 13/01/2006
TextBoxFormatDate txtInsuranceStartDate
End Sub

Private Sub txtLeaseID_GotFocus()
   Dim szaTenant() As String, szaUnit() As String

   If cboUnit.text = "" Then
      MsgBox "Please select unit.", vbCritical + vbOKOnly, "Unit"
      Exit Sub
   End If

   If txtLeaseID.text <> "" Then Exit Sub
   If cboTenant2.text <> "" Then szaTenant = Split(cboTenant2.text, " / ")
   If cboTenant1.text <> "" Then szaTenant = Split(cboTenant1.text, " / ")
   szaUnit = Split(cboUnit.text, " - ")

   txtLeaseID.text = OnlyAlpahNumericString(szaTenant(0)) & OnlyAlpahNumericString(szaUnit(0))

   txtLeaseID.SelStart = 0
   txtLeaseID.SelLength = Len(txtLeaseID.text)
End Sub

Private Function OnlyAlpahNumericString(szString As String) As String
   Dim i As Integer, X As Integer
   
   For i = 1 To Len(szString)
      X = Asc(Mid(szString, i, 1))
      If (X > 47 And X < 58) Or (X > 64 And X < 91) Or (X > 96 And X < 123) Then
         OnlyAlpahNumericString = OnlyAlpahNumericString & Mid(szString, i, 1)
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
'   If KeyAscii = 13 Or KeyAscii = 10 Then txtNet__LostFocus (tabPurExp.Tab)
'
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtPPSqFoot_LostFocus()
   If txtPPSqFoot.Locked Then Exit Sub
'   MsgBox optSqFoot.Value
   Dim Area As String, Total As Double
   
   'MousePointer = vbHourglass
'
   txtPPSqFoot.text = Format(IIf(txtPPSqFoot.text = "", 0, txtPPSqFoot.text), "0.0000")
'
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   
   SQLStr1 = "SELECT TotalArea FROM Units WHERE UnitNumber = '" & Left(cboUnit.text, 8) & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   If Not IsNull(Rst1!TotalArea) Then
      Area = Rst1!TotalArea
   Else
      Area = 0
   End If
   
   Total = Area * CDbl(txtPPSqFoot.text)
   txtAmount.text = Format(Total, "0.00")
   
   Rst1.Close
'   Conn1.Close

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
   If cbo1.text = "Yes" Then Call CalculateSC

   MousePointer = vbDefault
End Sub

Private Sub txt11_LostFocus()

If txt11.text <> "" Then
    If NumberCheck(txt11.text) = False Then
        txt11.text = ""
    Else
        If txt11.text <> "" And txt12.text <> "" And txt12a.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txt12_LostFocus()

If txt12.text <> "" Then
    If NumberCheck2(txt12.text) = False Then
        txt12.text = ""
    Else
        txt12.text = Round(CDbl(txt12.text), 2)
        If txt11.text <> "" And txt12.text <> "" And txt12a.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txt12a_LostFocus()

If txt12a.text <> "" Then
    If NumberCheck2(txt12a.text) = False Then
        txt12a.text = ""
    Else
        If txt11.text <> "" And txt12.text <> "" And txt12a.text <> "" Then CalculateInterest
    End If
End If

End Sub

Private Sub txt14_LostFocus()

'If txt14.text <> "" Then If CheckDate(txt14.text) = False Then txt14.text = ""
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txt14.text <> "" Then TextBoxFormatDate txt14

End Sub

Private Sub txt15_LostFocus()

If txt15.text <> "" Then If CheckDate(txt15.text) = False Then txt15.text = ""

End Sub

Private Sub txt16_LostFocus()

If txt16.text <> "" Then If CheckDate(txt16.text) = False Then txt16.text = ""

End Sub

Private Sub txt17_LostFocus()

If txt17.text <> "" Then
    If NumberCheck2(txt17.text) = False Then
        txt17.text = ""
    Else
        txt17.text = Round(CDbl(txt17.text), 2)
    End If
End If

End Sub

Private Sub txt18_LostFocus()

If txt18.text <> "" Then If CheckDate(txt18.text) = False Then txt18.text = ""

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

Private Sub txt3_LostFocus()
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txt3.text <> "" Then TextBoxFormatDate txt3
End Sub

Private Sub txt4_LostFocus()
' Added by Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txt4.text <> "" Then TextBoxFormatDate txt4
End Sub

Private Sub txt5_LostFocus()
'Added By Asif. 13/01/2006
' Modified by Samrat 08/02/2006
If txt5.text <> "" Then TextBoxFormatDate txt5
End Sub

Private Sub txt6_LostFocus()

If txt6.text <> "" Then
    If NumberCheck2(txt6.text) = False Then
        txt6.text = ""
        Exit Sub
    Else
        txt6.text = Round(CDbl(txt6.text), 2)
    End If
    If txt5.text <> "" Then Call CalculateBR
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
xy = CInt(Right(txt5.text, 2))
xm = CInt(Mid(txt5.text, 4, 2))
xd = CInt(Left(txt5.text, 2))
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

'Have to ensure that if the frequency is changed i.e. what calls CalculateBR then the base rate is changed. Txt5 or other will be based on txt7. txt7 has to change
'KOut next 2

'txt7.Enabled = True
'txt8.Enabled = True


'In conditional below remmed all set equal to txt5, 0, 2, 4, 10. Undone now

'Added And txt7.Text = "", to the conditional below not as of 22:11 20/11/2002

If txt6.text <> "" Then 'If there is a Base Rate for Year Figure
    Select Case cboFreqBR.text
        Case cboFreqBR.List(0): ' Weekly in advance
            
            txt8.text = Round((CDbl(txt6.text) / 52), 2)
            
            'Line below, why is it there, remmed 10:42, clearly unremmed
            
            txt7.text = txt5.text
                  
        Case cboFreqBR.List(1): ' Weekly in arrears
            
            txt8.text = Round((CDbl(txt6.text) / 52), 2)
            txt7.text = DateAdd("d", 7, txt5.text)
        Case cboFreqBR.List(2): ' Fortnightly in advance
        
            txt8.text = Round((CDbl(txt6.text) / 26), 2)
            txt7.text = txt5.text
        Case cboFreqBR.List(3): ' Fortnightly in arrears
            txt8.text = Round((CDbl(txt6.text) / 26), 2)
            txt7.text = DateAdd("d", 7, txt5.text)
        Case cboFreqBR.List(4): ' Monthly in advance
            txt8.text = Round((CDbl(txt6.text) / 12), 2)
            txt7.text = txt5.text
        Case cboFreqBR.List(5): ' Monthly in arrears
        
            txt8.text = Round((CDbl(txt6.text) / 12), 2)
            txt7.text = DateAdd("m", 1, txt5.text)
            
        Case cboFreqBR.List(6): ' Quarterly in advance
            txt8.text = Round((CDbl(txt6.text) / 4), 2)
            Select Case xm
                Case Is < qm2:
                    Select Case xm
                        Case Is < qm1:
                            txt7.text = quarterly1
                        Case qm1:
                            Select Case xd
                                Case Is < qd1:
                                    txt7.text = quarterly1
                                Case qd1:
                                    txt7.text = quarterly1
                                Case Is > qd1:
                                    txt7.text = quarterly2
                            End Select
                        Case Is > qm1:
                            txt7.text = quarterly2
                    End Select
                Case qm2:
                    Select Case xd
                        Case Is < qd2:
                            txt7.text = quarterly2
                        Case qd2:
                            txt7.text = quarterly2
                        Case Is > qd2:
                            txt7.text = quarterly3
                    End Select
                Case Is > qm2:
                    Select Case xm
                        Case Is < qm3:
                            txt7.text = quarterly3
                        Case qm3:
                            Select Case xd
                                Case Is < qd3:
                                    txt7.text = quarterly3
                                Case qd3:
                                    txt7.text = quarterly3
                                Case Is > qd3:
                                    txt7.text = quarterly4
                            End Select
                        Case Is > qm3:
                            Select Case xm
                                Case Is < qm4:
                                    txt7.text = quarterly4
                                Case qm4:
                                    Select Case xd
                                        Case Is < qd4:
                                            txt7.text = quarterly4
                                        Case qd4:
                                            txt7.text = quarterly4
                                        Case Is > qd4:
                                            txt7.text = DateAdd("yyyy", 1, quarterly1)
                                    End Select
                                Case Is > qm4:
                                   txt7.text = DateAdd("yyyy", 1, quarterly1)
                            End Select
                    End Select
            End Select
        Case cboFreqBR.List(7): ' Quarterly in arrears
        MsgBox cboFreqBR.List(7)
            txt8.text = Round((CDbl(txt6.text) / 4), 2)
            Select Case xm
                Case Is < qm2:
                    Select Case xm
                        Case Is < qm1:
                            txt7.text = quarterly2
                        Case qm1:
                            Select Case xd
                                Case Is < qd1:
                                    txt7.text = quarterly2
                                Case qd1:
                                    txt7.text = quarterly2
                                Case Is > qd1:
                                    txt7.text = quarterly3
                            End Select
                        Case Is > qm1:
                            txt7.text = quarterly3
                    End Select
                Case qm2:
                    Select Case xd:
                        Case Is < qd2:
                            txt7.text = quarterly3
                        Case qd2:
                            txt7.text = quarterly3
                        Case Is > qd2:
                            txt7.text = quarterly4
                    End Select
                Case Is > qm2:
                        Select Case xm
                            Case Is < qm3:
                                txt7.text = quarterly4
                            Case qm3:
                                Select Case xd
                                    Case Is < qd3:
                                        txt7.text = quarterly4
                                    Case qd3:
                                        txt7.text = quarterly4
                                    Case Is > qd3:
                                        'quarterly1 (next year)
                                        txt7.text = DateAdd("yyyy", 1, quarterly1)
                                End Select
                            Case Is > qm3:
                                Select Case xm
                                    Case Is < qm4:
                                        txt7.text = DateAdd("yyyy", 1, quarterly1)
                                    Case qm4:
                                        Select Case xd
                                            Case Is < qd4:
                                                txt7.text = DateAdd("yyyy", 1, quarterly1)
                                            Case qd4:
                                                txt7.text = DateAdd("yyyy", 1, quarterly1)
                                            Case Is > qd4:
                                                txt7.text = DateAdd("yyyy", 1, quarterly2)
                                        End Select
                                    Case Is > qm4:
                                        txt7.text = DateAdd("yyyy", 1, quarterly2)
                                End Select
                        End Select
            End Select
        Case cboFreqBR.List(8): ' Half yearly in advance
            txt8.text = Round((CDbl(txt6.text) / 2), 2)
            Select Case xm
                Case Is < hm1:
                    txt7.text = halfyearly1
                Case hm1:
                    Select Case xd
                        Case Is < hd1:
                            txt7.text = halfyearly1
                        Case hd1:
                            txt7.text = halfyearly1
                        Case Is > hd1:
                            txt7.text = halfyearly2
                    End Select
                Case Is > hm1:
                    Select Case xm:
                        Case Is < hm2:
                            txt7.text = halfyearly2
                        Case hm2:
                            Select Case xd
                                Case Is < hd2:
                                    txt7.text = halfyearly2
                                Case hd2:
                                    txt7.text = halfyearly2
                                Case Is > hd2:
                                    txt7.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm2:
                            txt7.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
            End Select
        Case cboFreqBR.List(9): ' Half yearly in arrears
            txt8.text = Round((CDbl(txt6.text) / 2), 2)
            Select Case xm
                Case Is < hm2:
                    Select Case xm
                        Case Is < hm1:
                            txt7.text = halfyearly2
                        Case hm1:
                            Select Case xd
                                Case Is < hd1:
                                    txt7.text = halfyearly2
                                Case hd1:
                                    txt7.text = halfyearly2
                                Case Is > hd1:
                                    txt7.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm1:
                            txt7.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
                Case hm2:
                    Select Case xd
                        Case Is < hd2:
                            txt7.text = DateAdd("yyyy", 1, halfyearly1)
                        Case hd2:
                            txt7.text = DateAdd("yyyy", 1, halfyearly1)
                        Case Is > hd2:
                            txt7.text = DateAdd("yyyy", 1, halfyearly2)
                    End Select
                Case Is > hm2:
                    txt7.text = DateAdd("yyyy", 1, halfyearly2)
            End Select
        Case cboFreqBR.List(10): ' Yearly in advance
            txt8.text = Round(CDbl(txt6.text), 2)
            'remmed below at 13:19
            txt7.text = txt5.text

'            Select Case xm
 '               Case Is < ym:
  '                  txt7.Text = yearly
   '             Case ym:
    '                Select Case xd
     '                   Case Is < yd:
      '                      txt7.Text = yearly
       '                 Case yd:
        '                    txt7.Text = yearly
         '               Case Is > yd:
          '                  txt7.Text = DateAdd("yyyy", 1, yearly)
           '         End Select
            '    Case Is > ym:
             '       txt7.Text = DateAdd("yyyy", 1, yearly)
            'End Select
        Case cboFreqBR.List(11): ' Yearly in arrears
            txt8.text = Round(CDbl(txt6.text), 2)
            txt7.text = DateAdd("yyyy", 1, txt5.text)
        '    Select Case xm
         '       Case Is < ym:
          '          txt7.Text = DateAdd("yyyy", 1, yearly)
           '     Case ym:
            '        Select Case xd
             '           Case Is < yd:
              '              txt7.Text = DateAdd("yyyy", 1, yearly)
               '         Case yd:
                '            txt7.Text = DateAdd("yyyy", 1, yearly)
                 '       Case Is > yd:
                  '          txt7.Text = DateAdd("yyyy", 2, yearly)
                   ' End Select
                'Case Is > ym:
                 '   txt7.Text = DateAdd("yyyy", 2, yearly)
            
            'End Select
    
    End Select





'kc remmed next 7 lines 11:27 9/10/2002 and moved it above the End If, as in SC
'why do we want to do this?? Also think they should be part of the primary conditional
'Undone as of 22:11 20/11/2002

If txt7.text <> "" And txt5.text <> "" Then
    b = DateDiff("yyyy", txt7.text, txt5.text)
    'txt15.Text = b
    
    If b <> 0 Then
        txt7.text = DateAdd("yyyy", b, txt7.text)
    End If
End If



End If



txt7.Enabled = False
txt8.Enabled = False

End Sub

Public Sub CalculateSC()

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

xy = CInt(Right(txt9.text, 2))
xm = CInt(Mid(txt9.text, 4, 2))
xd = CInt(Left(txt9.text, 2))
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

txt10b.Enabled = True

If cboUnit.text = "" Then
    MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
    Exit Sub
End If

Select Case cboFreqSC.text
    Case cboFreqSC.List(0): 'Weekly in advance
        txt10b.text = txt9.text
    Case cboFreqSC.List(1): 'Weekly in arrears
        txt10b.text = DateAdd("d", 7, txt9.text)
    Case cboFreqSC.List(2): 'Fortnightly in advance
        txt10b.text = txt9.text
    Case cboFreqSC.List(3): 'Fortnightly in arrears
        txt10b.text = DateAdd("d", 14, txt9.text)
    Case cboFreqSC.List(4): 'Monthly in advance
        txt10b.text = txt9.text
    Case cboFreqSC.List(5): 'Monthly in arrears
        txt10b.text = DateAdd("m", 1, txt9.text)
    Case cboFreqSC.List(6): 'Quarterly in advance
        Select Case xm
            Case Is < qm2:
                Select Case xm
                    Case Is < qm1:
                        txt10b.text = quarterly1
                    Case qm1:
                        Select Case xd
                            Case Is < qd1:
                                txt10b.text = quarterly1
                            Case qd1:
                                txt10b.text = quarterly1
                            Case Is > qd1:
                                txt10b.text = quarterly2
                        End Select
                    Case Is > qm1:
                        txt10b.text = quarterly2
                End Select
            Case qm2:
                Select Case xd
                    Case Is < qd2:
                        txt10b.text = quarterly2
                    Case qd2:
                        txt10b.text = quarterly2
                    Case Is > qd2:
                        txt10b.text = quarterly3
                End Select
            Case Is > qm2:
                Select Case xm
                    Case Is < qm3:
                        txt10b.text = quarterly3
                    Case qm3:
                        Select Case xd
                            Case Is < qd3:
                                txt10b.text = quarterly3
                            Case qd3:
                                txt10b.text = quarterly3
                            Case Is > qd3:
                                txt10b.text = quarterly4
                        End Select
                    Case Is > qm3:
                        Select Case xm
                            Case Is < qm4:
                                txt10b.text = quarterly4 'qm4
                            Case qm4:
                                Select Case xd
                                    Case Is < qd4:
                                        txt10b.text = quarterly4
                                    Case qd4:
                                        txt10b.text = quarterly4
                                    Case Is > qd4:
                                        txt10b.text = DateAdd("yyyy", 1, quarterly1)
                                End Select
                            Case Is > qm4:
                               txt10b.text = DateAdd("yyyy", 1, quarterly1)
                        End Select
                End Select
            End Select
    Case cboFreqSC.List(7): 'Quarterly in arrears
        Select Case xm
            Case Is < qm2:
                Select Case xm
                    Case Is < qm1:
                        txt10b.text = quarterly2
                    Case qm1:
                        Select Case xd
                            Case Is < qd1:
                                txt10b.text = quarterly2
                            Case qd1:
                                txt10b.text = quarterly2
                            Case Is > qd1:
                                txt10b.text = quarterly3
                        End Select
                    Case Is > qm1:
                        txt10b.text = quarterly3
                End Select
            Case qm2:
                Select Case xd:
                    Case Is < qd2:
                        txt10b.text = quarterly3
                    Case qd2:
                        txt10b.text = quarterly3
                    Case Is > qd2:
                        txt10b.text = quarterly4
                End Select
            Case Is > qm2:
                    Select Case xm
                        Case Is < qm3:
                            txt10b.text = quarterly4
                        Case qm3:
                            Select Case xd
                                Case Is < qd3:
                                    txt10b.text = quarterly4
                                Case qd3:
                                    txt10b.text = quarterly4
                                Case Is > qd3:
                                    'quarterly1 (next year)
                                    txt10b.text = DateAdd("yyyy", 1, quarterly1)
                            End Select
                        Case Is > qm3:
                            Select Case xm
                                Case Is < qm4:
                                    txt10b.text = DateAdd("yyyy", 1, quarterly1)
                                Case qm4:
                                    Select Case xd
                                        Case Is < qd4:
                                            txt10b.text = DateAdd("yyyy", 1, quarterly1)
                                        Case qd4:
                                            txt10b.text = DateAdd("yyyy", 1, quarterly1)
                                        Case Is > qd4:
                                            txt10b.text = DateAdd("yyyy", 1, quarterly2)
                                    End Select
                                Case Is > qm4:
                                    txt10b.text = DateAdd("yyyy", 1, quarterly2)
                            End Select
                    End Select
        End Select
    Case cboFreqSC.List(8): 'Half yearly in advance
        Select Case xm
                Case Is < hm1:
                    txt10b.text = halfyearly1
                Case hm1:
                    Select Case xd
                        Case Is < hd1:
                            txt10b.text = halfyearly1
                        Case hd1:
                            txt10b.text = halfyearly1
                        Case Is > hd1:
                            txt10b.text = halfyearly2
                    End Select
                Case Is > hm1:
                    Select Case xm:
                        Case Is < hm2:
                            txt10b.text = halfyearly2
                        Case hm2:
                            Select Case xd
                                Case Is < hd2:
                                    txt10b.text = halfyearly2
                                Case hd2:
                                    txt10b.text = halfyearly2
                                Case Is > hd2:
                                    txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm2:
                            txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
        End Select
    Case cboFreqSC.List(9): 'Half yearly in arrears
        Select Case xm
                Case Is < hm2:
                    Select Case xm
                        Case Is < hm1:
                            txt10b.text = halfyearly2
                        Case hm1:
                            Select Case xd
                                Case Is < hd1:
                                    txt10b.text = halfyearly2
                                Case hd1:
                                    txt10b.text = halfyearly2
                                Case Is > hd1:
                                    txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                            End Select
                        Case Is > hm1:
                            txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                    End Select
                Case hm2:
                    Select Case xd
                        Case Is < hd2:
                            txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                        Case hd2:
                            txt10b.text = DateAdd("yyyy", 1, halfyearly1)
                        Case Is > hd2:
                            txt10b.text = DateAdd("yyyy", 1, halfyearly2)
                    End Select
                Case Is > hm2:
                    txt10b.text = DateAdd("yyyy", 1, halfyearly2)
            End Select
    Case cboFreqSC.List(10): ' yearly in advance
        txt10b.text = txt9.text
    Case cboFreqSC.List(11): ' yearly in arrears
        txt10b.text = DateAdd("yyyy", 1, txt9.text)
End Select

If txt10b.text <> "" And txt9.text <> "" Then
    b = DateDiff("yyyy", txt10b.text, txt9.text)
    If b <> 0 Then
        txt10b.text = DateAdd("yyyy", b, txt10b.text)
    End If
End If

txt10b.Enabled = False
End Sub

Public Sub CalculateInsuranceCharge()

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
xm = CInt(Mid(txtInsuranceStartDate.text, 4, 2))
xd = CInt(Left(txtInsuranceStartDate.text, 2))
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

If cboUnit.text = "" Then
    MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
    Exit Sub
End If

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

a = CDbl(txt12a.text)
r = (BaseRate + CDbl(txt12.text)) / 100
d = CInt(txt11.text)

txt13.text = Round(a * r * d / 365, 2)

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
   
   Dim Total As Double, TotalServiceCharge As String
   
   On Error GoTo ErrorHander
   
   MousePointer = vbHourglass
   
   txtSCPercentage.text = Format(IIf(txtSCPercentage.text = "", 0, txtSCPercentage.text), "0.0000")
   
   ' The following code is to calculate SC payable according to percentage
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   
   SQLStr1 = "SELECT GlobalData.TotalSC FROM GlobalData,Units where Units.PropertyID = GlobalData.PropertyID AND Units.UnitNumber = '" & Left(cboUnit, 8) & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   TotalServiceCharge = Rst1!TotalSC
'
   Rst1.Close
   
   Total = CDbl(TotalServiceCharge) * (CDbl(txtSCPercentage.text) / 100)
   txtAmount.text = Format(Total, "0.00")
   
   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

'Debug.Print cboFreqSC.ListCount

   txtFinalAmout.text = Format((Total / CInt(Rst1!PARTOFYEAR)), "0.00")
   
   Rst1.Close
   Conn1.Close
'
   If cbo1.text = "Yes" Then Call CalculateSC

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
txtSuppCaption2.Visible = False
lblSupplementary2.Caption = txtSuppCaption2.text
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

    rstBreach!LeaseID = txtLeaseID.text
    rstBreach!BreachType = cboBreachType.BoundText
    rstBreach!CommenceDate = txtCommenceDate.text
    rstBreach!InitiatedBy = txtInitiatedBy.text
    If chkResolved.Value = 1 Then
        rstBreach!Resolved = True
    Else
        rstBreach!Resolved = False
    End If
    rstBreach!DateReceived = txtDateReceived.text
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
'    rstBreach.Close
'    conBreach.Close
'    Set rstBreach = Nothing
'    Set conBreach = Nothing
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

    rstAssignment!LeaseID = txtLeaseID.text
    rstAssignment!AssignDate = txtAssignDate.text
    rstAssignment!Tenant = txtTenant.text
    rstAssignment.Update

'    Next iRowIndex
'

    rstAssignment.Close
    conAssignment.Close
    Set rstAssignment = Nothing
    Set conAssignment = Nothing
    SaveAssignment = True
    PopulateAssignments
    Exit Function
    
Exception:
    
    MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
'    rstAssignment.Close
'    conAssignment.Close
'    Set rstAssignment = Nothing
'    Set conAssignment = Nothing
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
            txtCommenceDate.Enabled = False
            txtInitiatedBy.Enabled = False
            chkResolved.Enabled = False
            txtDateReceived.Enabled = False
            txtReceivedBy.Enabled = False
        
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
            txtCommenceDate.Enabled = True
            txtCommenceDate.text = ""
            txtInitiatedBy.Enabled = True
            txtInitiatedBy.text = ""
            chkResolved.Enabled = True
            txtDateReceived.Enabled = True
            txtDateReceived.text = ""
            txtReceivedBy.Enabled = True
            txtReceivedBy.text = ""
                    
        Case ComponentMode.EditMode
            cmdBreachNew.Enabled = False
            cmdBreachEdit.Enabled = False
            cmdBreachSave.Enabled = True
            cmdBreachCancel.Enabled = True
            
            gridBreach.Enabled = False
        
            cboBreachType.Enabled = True
            cmdSetBreachType.Enabled = True
            txtCommenceDate.Enabled = True
            txtInitiatedBy.Enabled = True
            chkResolved.Enabled = True
            txtDateReceived.Enabled = True
            txtReceivedBy.Enabled = True
            
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
        
            txtAssignDate.Enabled = False
            txtTenant.Enabled = False
        
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
        
            txtAssignDate.Enabled = True
            txtAssignDate.text = ""
            txtTenant.Enabled = True
            txtTenant = ""
                
        Case ComponentMode.EditMode
            cmdAssignmentNew.Enabled = False
            cmdAssignmentEdit.Enabled = False
            cmdAssignmentSave.Enabled = True
            cmdAssignmentCancel.Enabled = True
            
            gridAssignment.Enabled = False
        
            txtAssignDate.Enabled = True
            txtTenant.Enabled = True
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

Private Sub FillcboType(conCombo As Control, szaDemandtype() As String)
   Dim SQLStr1 As String, i As Integer

   conCombo.Clear
   
   i = 0
   While szaDemandtype(i) <> ""
       conCombo.AddItem szaDemandtype(i)
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
