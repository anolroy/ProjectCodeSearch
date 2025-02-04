VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   8430
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picNominal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3240
      Left            =   8280
      ScaleHeight     =   3210
      ScaleWidth      =   3465
      TabIndex        =   99
      Top             =   6750
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdCloseNominal 
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
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNominal 
         Height          =   2505
         Left            =   45
         TabIndex        =   101
         Top             =   675
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4419
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         WordWrap        =   -1  'True
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   107
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   106
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   120
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Nominal Code"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   195
         Left            =   1620
         TabIndex        =   104
         Top             =   135
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Nominal Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchNC 
         Height          =   255
         Left            =   45
         TabIndex        =   103
         Top             =   375
         Width           =   945
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1667;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchNName 
         Height          =   255
         Left            =   1035
         TabIndex        =   102
         Top             =   375
         Width           =   2370
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4180;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   1
         Left            =   45
         Top             =   75
         Width           =   3105
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   1200
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   90
      Top             =   6735
      Visible         =   0   'False
      Width           =   6285
      Begin VB.CommandButton cmdPicCLose 
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
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3525
         Left            =   45
         TabIndex        =   92
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6218
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         WordWrap        =   -1  'True
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   98
         Top             =   375
         Width           =   4590
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8096;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   97
         Top             =   375
         Width           =   1530
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2699;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   96
         Top             =   135
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   94
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   93
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5850
      End
   End
   Begin TabDlg.SSTab tabSettings 
      Height          =   6660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   11748
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   5
      TabsPerRow      =   8
      TabHeight       =   732
      ForeColor       =   4194368
      MouseIcon       =   "frmOptions.frx":29C12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Report Files"
      TabPicture(0)   =   "frmOptions.frx":29C2E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Deposit Accounts"
      TabPicture(1)   =   "frmOptions.frx":29C4A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label6(3)"
      Tab(1).Control(2)=   "Label6(2)"
      Tab(1).Control(3)=   "Label6(0)"
      Tab(1).Control(4)=   "Label6(1)"
      Tab(1).Control(5)=   "Label50(1)"
      Tab(1).Control(6)=   "Label50(2)"
      Tab(1).Control(7)=   "Label50(0)"
      Tab(1).Control(8)=   "cmbClient"
      Tab(1).Control(9)=   "cmbBank"
      Tab(1).Control(10)=   "cmbDNC"
      Tab(1).Control(11)=   "flxDeposit"
      Tab(1).Control(12)=   "Frame6"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Lease Notification"
      TabPicture(2)   =   "frmOptions.frx":29C66
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Email"
      TabPicture(3)   =   "frmOptions.frx":29C82
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Frame8"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "VAT"
      TabPicture(4)   =   "frmOptions.frx":29C9E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11(4)"
      Tab(4).Control(1)=   "Shape2"
      Tab(4).Control(2)=   "Label11(2)"
      Tab(4).Control(3)=   "Label11(3)"
      Tab(4).Control(4)=   "Label11(1)"
      Tab(4).Control(5)=   "Label11(0)"
      Tab(4).Control(6)=   "Label11(5)"
      Tab(4).Control(7)=   "Shape3"
      Tab(4).Control(8)=   "flxVat"
      Tab(4).Control(9)=   "chkInUse"
      Tab(4).Control(10)=   "cmdSaveVAT"
      Tab(4).Control(11)=   "txtVatDesp"
      Tab(4).Control(12)=   "txtVatRate"
      Tab(4).Control(13)=   "cmdExit"
      Tab(4).Control(14)=   "cmdCancelVAT"
      Tab(4).Control(15)=   "cmdUpdate"
      Tab(4).Control(16)=   "chkOnReturn"
      Tab(4).ControlCount=   17
      TabCaption(5)   =   "Control Accounts"
      TabPicture(5)   =   "frmOptions.frx":29CBA
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label1(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label50(3)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(2)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label1(3)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label1(4)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label1(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "txtClientList"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "flxTransactionTypes"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "cmdCASave"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdCAClose"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "cmdCACancel"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "cmdDelete"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "cmdCANew"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "cmdClientList"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Sub Types"
      TabPicture(6)   =   "frmOptions.frx":29CD6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(6)"
      Tab(6).Control(1)=   "Label1(7)"
      Tab(6).Control(2)=   "flxSubTypes"
      Tab(6).Control(3)=   "cmdSubTypeClose"
      Tab(6).Control(4)=   "cmdSubTypeSave"
      Tab(6).Control(5)=   "cmdSubTypeCancel"
      Tab(6).Control(6)=   "picSubType"
      Tab(6).ControlCount=   7
      TabCaption(7)   =   "Fund "
      TabPicture(7)   =   "frmOptions.frx":29CF2
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame9"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame9 
         Height          =   5505
         Left            =   -74910
         TabIndex        =   126
         Top             =   495
         Width           =   12840
         Begin VB.CommandButton cmdSaveFund 
            Caption         =   "Save"
            Height          =   465
            Left            =   4680
            TabIndex        =   128
            Top             =   1665
            Width           =   1185
         End
         Begin VB.CheckBox chkFundAssignment 
            Caption         =   "Fund Assignment to Client and Properties"
            Height          =   555
            Left            =   4320
            TabIndex        =   127
            Top             =   585
            Width           =   3300
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1230
         Left            =   -73065
         TabIndex        =   122
         Top             =   4320
         Width           =   8700
         Begin VB.CommandButton cmdSaveSettings 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3390
            TabIndex        =   125
            Top             =   405
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditSettings 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   2070
            TabIndex        =   124
            Top             =   405
            Width           =   1095
         End
         Begin VB.CommandButton cmdCloseSettings 
            Caption         =   "C&lose"
            Height          =   375
            Left            =   4710
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Close"
            Top             =   405
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   -72795
         TabIndex        =   118
         Top             =   3600
         Width           =   8610
         Begin VB.CommandButton cmdCloseOpt 
            Caption         =   "C&lose"
            Height          =   375
            Left            =   4905
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Close"
            Top             =   315
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditOptions 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   2250
            TabIndex        =   120
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSaveOptions 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3555
            TabIndex        =   119
            Top             =   315
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1230
         Left            =   -74325
         TabIndex        =   112
         Top             =   4680
         Width           =   12255
         Begin VB.CommandButton cmdNew 
            Caption         =   "&Add New"
            Height          =   345
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Add New"
            Top             =   405
            Width           =   1200
         End
         Begin VB.CommandButton cmdLDCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   345
            Left            =   5550
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Cancel"
            Top             =   405
            Width           =   1200
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4035
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Save"
            Top             =   405
            Width           =   1200
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2505
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Edit"
            Top             =   405
            Width           =   1200
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "C&lose"
            Height          =   345
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Close"
            Top             =   405
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkOnReturn 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -66990
         TabIndex        =   111
         Top             =   1090
         Width           =   375
      End
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
         Left            =   5040
         TabIndex        =   88
         Top             =   615
         Width           =   300
      End
      Begin VB.PictureBox picSubType 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -64230
         ScaleHeight     =   495
         ScaleWidth      =   1095
         TabIndex        =   85
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
         Begin VB.TextBox txtSubType 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdSubTypeCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -69375
         TabIndex        =   81
         Top             =   5145
         Width           =   1485
      End
      Begin VB.CommandButton cmdSubTypeSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   -71115
         TabIndex        =   80
         Top             =   5145
         Width           =   1485
      End
      Begin VB.CommandButton cmdSubTypeClose 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   -65895
         TabIndex        =   82
         Top             =   5145
         Width           =   1485
      End
      Begin VB.CommandButton cmdCANew 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2550
         TabIndex        =   69
         Top             =   5910
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7770
         TabIndex        =   72
         Top             =   5910
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdCACancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6030
         TabIndex        =   71
         Top             =   5910
         Width           =   1485
      End
      Begin VB.CommandButton cmdCAClose 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   9510
         TabIndex        =   73
         Top             =   5910
         Width           =   1485
      End
      Begin VB.CommandButton cmdCASave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4290
         TabIndex        =   70
         Top             =   5910
         Width           =   1485
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -72675
         TabIndex        =   56
         Top             =   5085
         Width           =   915
      End
      Begin VB.CommandButton cmdCancelVAT 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69075
         TabIndex        =   58
         Top             =   5085
         Width           =   915
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   -67275
         TabIndex        =   59
         Top             =   5085
         Width           =   915
      End
      Begin VB.TextBox txtVatRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   53
         Top             =   1080
         Width           =   1065
      End
      Begin VB.TextBox txtVatDesp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71715
         TabIndex        =   54
         Top             =   1080
         Width           =   3420
      End
      Begin VB.CommandButton cmdSaveVAT 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70875
         TabIndex        =   57
         Top             =   5085
         Width           =   915
      End
      Begin VB.CheckBox chkInUse 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -68115
         TabIndex        =   55
         Top             =   1080
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Caption         =   "Service Charge send by post:"
         Height          =   855
         Left            =   -73065
         TabIndex        =   51
         Top             =   3420
         Width           =   8685
         Begin VB.OptionButton optSC 
            Caption         =   "Service Charge Expenditure Statement"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton optSC 
            Caption         =   "Service Charge Budget Statement"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   40
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sending Email Configuration:"
         Enabled         =   0   'False
         Height          =   2580
         Left            =   -73065
         TabIndex        =   46
         Top             =   765
         Width           =   8730
         Begin VB.CheckBox chkTLS 
            Height          =   195
            Left            =   1710
            TabIndex        =   109
            Top             =   1665
            Width           =   1545
         End
         Begin VB.CheckBox chkSSL 
            Height          =   195
            Left            =   1710
            TabIndex        =   38
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox txtFromEmail 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   33
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtPws 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   6225
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   37
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtUName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   36
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7425
            MaxLength       =   5
            TabIndex        =   35
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtSMTP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   34
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label14 
            Caption         =   "Use TLS:"
            Height          =   255
            Left            =   135
            TabIndex        =   108
            Top             =   1710
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Use SSL:"
            Height          =   255
            Left            =   135
            TabIndex        =   87
            Top             =   1395
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "From Email Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Password:"
            Height          =   255
            Left            =   5505
            TabIndex        =   50
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "User Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Port:              (default: 25)"
            Height          =   255
            Left            =   5505
            TabIndex        =   48
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "SMTP Server Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DFDFDF&
         Caption         =   "Property wise Statement File Path:"
         Height          =   3990
         Index           =   1
         Left            =   -74520
         TabIndex        =   43
         Top             =   1980
         Width           =   12240
         Begin VB.CommandButton cmdStatementPathCancel 
            Caption         =   "&Close"
            Height          =   375
            Left            =   9480
            TabIndex        =   8
            Top             =   3435
            Width           =   1095
         End
         Begin VB.CommandButton cmdStatementPathSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   8295
            TabIndex        =   7
            Top             =   3435
            Width           =   1095
         End
         Begin VB.CommandButton cmdStatementPathEdit 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   7125
            TabIndex        =   6
            Top             =   3435
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxStFlPth 
            Height          =   2880
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   5080
            _Version        =   393216
            ForeColor       =   0
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   12632256
            BackColorSel    =   8454143
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
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
            _Band(0).Cols   =   3
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.Label Label6 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   735
            ForeColor       =   0
            VariousPropertyBits=   8388627
            Caption         =   "Property"
            Size            =   "1296;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label6 
            Height          =   255
            Index           =   5
            Left            =   3120
            TabIndex        =   44
            Top             =   240
            Width           =   1335
            ForeColor       =   0
            VariousPropertyBits=   8388627
            Caption         =   "Statement Path"
            Size            =   "2355;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFDFDF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00AFDFDF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   120
            Top             =   240
            Width           =   11910
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Atomatic ID Generation"
         Height          =   1575
         Left            =   -72795
         TabIndex        =   42
         Top             =   1980
         Visible         =   0   'False
         Width           =   8595
         Begin VB.CheckBox CheckManaging 
            Caption         =   "Managing agent"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   32
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox CheckProperty 
            Caption         =   "Property"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   30
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox CheckSupplier 
            Caption         =   "Supplier"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   28
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox CheckLessee 
            Caption         =   "Lessee"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox CheckClient 
            Caption         =   "Client / Landlord"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox CheckUnit 
            Caption         =   "Unit"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lease Alerts"
         Height          =   855
         Left            =   -72795
         TabIndex        =   24
         Top             =   945
         Width           =   8595
         Begin VB.CommandButton cmdExpRepGen 
            Caption         =   "Generate Report of Leases >>"
            Height          =   375
            Left            =   5160
            TabIndex        =   86
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtLeaseEndDays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkAlarm 
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Set Alarm                        days prior to the Lease Termination."
            Height          =   255
            Left            =   600
            TabIndex        =   41
            Top             =   400
            Width           =   4575
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFDFDF&
         Caption         =   "Report File Path:"
         Height          =   1155
         Index           =   0
         Left            =   -74595
         TabIndex        =   9
         Top             =   585
         Width           =   12285
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cance&l"
            Height          =   375
            Left            =   10755
            TabIndex        =   3
            Top             =   660
            Width           =   1095
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   9615
            TabIndex        =   2
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "File Path:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin MSForms.TextBox txtFilePath 
            Height          =   315
            Left            =   9585
            TabIndex        =   5
            Top             =   180
            Width           =   1815
            VariousPropertyBits=   679495711
            Size            =   "3201;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmdPathEdit 
            Height          =   315
            Left            =   11430
            TabIndex        =   1
            ToolTipText     =   "Edit the Path"
            Top             =   180
            Width           =   375
            Size            =   "661;556"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox txtAppPath 
            Height          =   555
            Left            =   795
            TabIndex        =   4
            Top             =   240
            Width           =   4575
            VariousPropertyBits=   -1467987949
            Size            =   "8070;979"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDeposit 
         Height          =   2685
         Left            =   -74310
         TabIndex        =   13
         Top             =   1935
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   4736
         _Version        =   393216
         ForeColor       =   0
         Cols            =   7
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
         _Band(0).Cols   =   7
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxVat 
         Height          =   3195
         Left            =   -74640
         TabIndex        =   52
         Top             =   1485
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5636
         _Version        =   393216
         ForeColor       =   0
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   8454143
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   5
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTransactionTypes 
         Height          =   4335
         Left            =   225
         TabIndex        =   65
         Top             =   1260
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSubTypes 
         Height          =   3855
         Left            =   -74460
         TabIndex        =   79
         Top             =   1170
         Width           =   12180
         _ExtentX        =   21484
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape Shape3 
         Height          =   330
         Left            =   -74640
         Top             =   675
         Width           =   12255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On VAT Return"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   -67080
         TabIndex        =   110
         Top             =   720
         Width           =   1050
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   2565
         TabIndex        =   89
         Top             =   615
         Width           =   2475
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4366;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Sub Type Description"
         Height          =   195
         Index           =   7
         Left            =   -70230
         TabIndex        =   83
         Top             =   720
         Width           =   4875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Sub Type Name"
         Height          =   195
         Index           =   6
         Left            =   -74415
         TabIndex        =   84
         Top             =   720
         Width           =   3960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Allow Posting"
         Height          =   195
         Index           =   5
         Left            =   11010
         TabIndex        =   76
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Nominal Account Name"
         Height          =   195
         Index           =   4
         Left            =   8415
         TabIndex        =   66
         Top             =   1035
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Nominal Code"
         Height          =   195
         Index           =   3
         Left            =   7020
         TabIndex        =   67
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   77
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   3
         Left            =   1965
         TabIndex        =   74
         Top             =   615
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Control Account"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   68
         Top             =   1020
         Width           =   2955
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "  ID"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   -74460
         TabIndex        =   64
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vat Code"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   1
         Left            =   -73695
         TabIndex        =   63
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   3
         Left            =   -71670
         TabIndex        =   62
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vat Rate"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   -72795
         TabIndex        =   61
         Top             =   720
         Width           =   585
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   975
         Left            =   -74640
         Top             =   4830
         Width           =   12300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InUse"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   -68100
         TabIndex        =   60
         Top             =   720
         Width           =   405
      End
      Begin MSForms.ComboBox cmbDNC 
         Height          =   285
         Left            =   -68880
         TabIndex        =   23
         Top             =   1200
         Width           =   1860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "987;5000"
      End
      Begin MSForms.ComboBox cmbBank 
         Height          =   285
         Left            =   -70800
         TabIndex        =   22
         Top             =   1200
         Width           =   1905
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3351;503"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "987;5000"
      End
      Begin MSForms.ComboBox cmbClient 
         Height          =   285
         Left            =   -74325
         TabIndex        =   21
         Top             =   1200
         Width           =   3510
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6191;503"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1234;5000"
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   0
         Left            =   -74325
         TabIndex        =   20
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Nominal:"
         Height          =   195
         Index           =   2
         Left            =   -68880
         TabIndex        =   19
         Top             =   1005
         Width           =   1170
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Bank:"
         Height          =   195
         Index           =   1
         Left            =   -70800
         TabIndex        =   18
         Top             =   1005
         Width           =   915
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   1
         Left            =   -73875
         TabIndex        =   17
         Top             =   1650
         Width           =   615
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Client"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   0
         Left            =   -74340
         TabIndex        =   16
         Top             =   1650
         Width           =   255
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "ID"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   2
         Left            =   -70800
         TabIndex        =   15
         Top             =   1650
         Width           =   375
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Bank"
         Size            =   "661;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   3
         Left            =   -68880
         TabIndex        =   14
         Top             =   1650
         Width           =   615
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Nominal"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         Caption         =   "Default Client Deposit Account Settings"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74265
         TabIndex        =   12
         Top             =   540
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VAT_NEW_ENTRY_   As Boolean
Dim VAT_MODIFIED_    As Boolean

Private iNewEditCC            As Byte
Private bEditSubType          As Boolean
Private bConfigureFlxDeposit  As Boolean
Private tAddNew               As Byte
Private DAYS_TO_REMIND        As String
Private IS_REMIND             As Boolean
Private NC()                  As String
Dim sTextBox As String
'Private szSQL_CC              As String
Private Sub cmdClientList_Click()
    picClient.Left = 480
    picClient.Top = 480
    sTextBox = "1"
    LoadflxClient
    
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                 rRow = 1
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
                    flxClient.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
          
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub LoadflxNominal()
     ConfigflxNominal
     Dim rstRec As New ADODB.Recordset
     Dim adoConn As New ADODB.Connection
     Dim szSQL As String
     Dim rRow As Integer
     adoConn.Open getConnectionString
     szSQL = "SELECT Code, Name, ClientID " & _
                "FROM   NominalLedger " & _
                "WHERE  Posting AND ClientID= '" & txtClientList.text & "' ORDER BY Code;"
     rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
     If rstRec.RecordCount = 0 Then
        flxNominal.Rows = 2
     Else
        flxNominal.Rows = rstRec.RecordCount + 1
     End If
    rRow = 2
      flxNominal.AddItem ""
      flxNominal.RowHeight(1) = 280
    While Not rstRec.EOF
     
       flxNominal.row = 1
       flxNominal.RowSel = 1
       flxNominal.ColSel = 1
       flxNominal.TextMatrix(rRow, 0) = ""
       flxNominal.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
       flxNominal.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
       flxNominal.RowHeight(rRow) = 280
       rstRec.MoveNext
       rRow = rRow + 1
    Wend
          
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdCloseNominal_Click()
        picNominal.Visible = False
        ControlHanlding DefaultMode
'        Dim adoconn As New ADODB.Connection
'        adoconn.Open getConnectionString
'        ConfigFlxTransctionTypes
'        LoadFlxTransactionTypes adoconn 'this function loads control accounts
'        adoconn.Close
'        Set adoconn = Nothing
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    cmdClientList.SetFocus
End Sub

Private Sub cmdSaveFund_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update shoppingCentre Set isFundAssign='" & IIf(chkFundAssignment.Value = 1, -1, 0) & "'"
    adoConn.Close
    MsgBox "Fund configuration has been saved", vbInformation, "Saved"
End Sub

Private Sub flxClient_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If sTextBox = "1" Then
        txtClientList.text = flxClient.TextMatrix(flxClient.row, 1)
        txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 2) ' here we are saving name in the tag
        configflxTransactionTypes
        LoadFlxTransactionTypes adoConn 'this function loads control accounts
    End If
    picClient.Visible = False
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub flxNominal_Click()
    If flxNominal.row = 0 Then Exit Sub
    PlaceTheCodeIngrid
    FocusControl cmdCASave
End Sub

Private Sub flxNominal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxNominal_Click
    End If
End Sub

Private Sub flxVat_Click()
    If flxVat.TextMatrix(1, 0) = "" Then Exit Sub

   txtVatRate.text = flxVat.TextMatrix(flxVat.row, 2)
   txtVatDesp.text = flxVat.TextMatrix(flxVat.row, 3)
   chkInUse.Value = IIf(flxVat.TextMatrix(flxVat.row, 4) = "YES", 1, 0)
   chkOnReturn.Value = IIf(flxVat.TextMatrix(flxVat.row, 5) = "YES", 1, 0)
End Sub

Private Sub flxVat_DblClick()
    cmdUpdate_Click
End Sub

Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
              flxClient.row = i
        End If
   Next i
End Sub


Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxClient.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
          
         
          'If sTextBox = "1" Then
           cmdClientList.SetFocus
'           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           'End If
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            flxClient.SetFocus
        End If
    End If
End Sub
'Private Sub cboCC_Click()
'   Dim iRow As Integer
'
''szHeader$ = "|<CAName|<Type|<NCode|<NName|ClientID|Fixed|<Posting|TypeID|saved|OriNC|OriPosting"
''Saved? S->Saved, C->New set of CA for a client, N->New CA, A->Amend, D->Decide
'   With flxTransactionTypes
'      For iRow = 1 To .Rows - 1
'         If .RowHeight(iRow) > 0 Then
'            If .TextMatrix(iRow, .col) = cboCC.Value Then
'               picCC.Visible = False
'               .Enabled = True
'                MsgBox "The code has been used already", vbInformation, "Warning"
'               Exit Sub
'            End If
'         End If
'      Next iRow
'
'      .TextMatrix(.row, .col) = cboCC.Value 'put the control code
'      .TextMatrix(.row, .col + 1) = cboCC.Column(1) 'put the nominal name
'      picCC.Visible = False
'      .Enabled = True
'      iNewEditCC = 1
'      cboCC.ListIndex = -1
'
'      If .TextMatrix(.row, 9) = "S" And _
'         .TextMatrix(.row, 3) <> .TextMatrix(.row, 10) Then
'         .TextMatrix(.row, 9) = "A" 'A->Amend
'         flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "NO"
'      End If
'      If .TextMatrix(.row, 9) = "C" And _
'         .TextMatrix(.row, 3) <> "" Then .TextMatrix(.row, 9) = "D"                    'D->Decide
'
'   End With
'End Sub
Private Sub PlaceTheCodeIngrid()
Dim iRow As Integer

'szHeader$ = "|<CAName|<Type|<NCode|<NName|ClientID|Fixed|<Posting|TypeID|saved|OriNC|OriPosting"
'Saved? S->Saved, C->New set of CA for a client, N->New CA, A->Amend, D->Decide
   With flxTransactionTypes
      For iRow = 1 To .Rows - 1
         If .RowHeight(iRow) > 0 Then
            If .TextMatrix(iRow, .col) = flxNominal.TextMatrix(flxNominal.row, 1) And flxNominal.TextMatrix(flxNominal.row, 1) <> "" Then
               picNominal.Visible = False
               .Enabled = True
                MsgBox "The code has been used already", vbInformation, "Warning"
               Exit Sub
            End If
         End If
      Next iRow
   
      .TextMatrix(.row, .col) = flxNominal.TextMatrix(flxNominal.row, 1)  'put the control code
      .TextMatrix(.row, .col + 1) = flxNominal.TextMatrix(flxNominal.row, 2)  'put the nominal name
      If .row > 6 Then
        If .row = 10 Then
            .TextMatrix(.row, 7) = "NO" 'Set allow Posting TO YES if .row >7
         ElseIf .row = 9 Then
            .TextMatrix(.row, 7) = "YES" 'Set allow Posting TO YES if .row >7
         ElseIf .row = 8 Then
            .TextMatrix(.row, 7) = "YES" 'Set allow Posting TO YES if .row >7
         ElseIf .row = 7 Then
            .TextMatrix(.row, 7) = "NO" 'Set allow Posting TO YES if .row >7
         ElseIf .row = 6 Then
            .TextMatrix(.row, 7) = "YES" 'Set allow Posting TO YES if .row >7
        
         End If
      End If
       picNominal.Visible = False
      .Enabled = True
      iNewEditCC = 1
     ' cboCC.ListIndex = -1
      'We use 9 for Balance sheet or profit and loss from now on
      If .TextMatrix(.row, 3) <> .TextMatrix(.row, 10) Then
         .TextMatrix(.row, 11) = "A" 'A->Amend
        ' flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "NO"
      End If
      'If .TextMatrix(.row, 9) = "C" And .TextMatrix(.row, 3) <> "" Then .TextMatrix(.row, 9) = "D"                     'D->Decide
      
   End With
    
End Sub
Private Sub cboCC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then ControlHanlding DefaultMode
End Sub

'Private Sub cboPosting_Click()
'   flxTransactionTypes.TextMatrix(flxTransactionTypes.row, flxTransactionTypes.col) = cboPosting.text
'   If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 9) = "S" Then
'        flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 9) = "A"  'A->Amend
'    End If
'
'   picPosting.Visible = False
'   flxTransactionTypes.Enabled = True
'   iNewEditCC = 1
'   cboPosting.ListIndex = -1
'End Sub

'Private Sub cboPosting_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If KeyAscii = 27 Then picPosting.Visible = False
'End Sub

'Private Sub cmbCAClient_Click()
'   Dim iRow    As Integer
'   Dim bFound  As Boolean
'
'   For iRow = 1 To flxTransactionTypes.Rows - 1
'      flxTransactionTypes.RowHeight(iRow) = 0
'   Next iRow
'
'   bFound = False
'   For iRow = 1 To flxTransactionTypes.Rows - 1
'      If flxTransactionTypes.TextMatrix(iRow, 5) = cmbCAClient.Value Then
'         flxTransactionTypes.RowHeight(iRow) = 285
'         bFound = True
'      End If
'   Next iRow
'   If Not bFound Then
'      For iRow = 1 To flxTransactionTypes.Rows - 1
'         flxTransactionTypes.RowHeight(iRow) = 285
'      Next iRow
'   End If
'
'   If UBound(NC, 2) > 1 Then 'need to understand this part
'      Dim Data()       As String
'      Dim i            As Integer
'      ReDim Data(2, 0) As String
'
'      i = 0
'      For iRow = 0 To UBound(NC, 2) - 1
'         If NC(2, iRow) = cmbCAClient.Value Then
'            Data(0, i) = NC(0, iRow)               'Nominal Code
'            Data(1, i) = NC(1, iRow)               'Nominal Code Name
'            Data(2, i) = NC(2, iRow)               'Client ID
'            i = i + 1
'            ReDim Preserve Data(2, i) As String
'         End If
'      Next iRow
'
'      cboCC.Column() = Data()
'   End If
'End Sub

Private Sub cmdCACancel_Click()
   If MsgBox("Do you like to discard the changes?", vbQuestion + vbYesNo, "Control Code") = vbNo Then Exit Sub

   Dim i As Integer

    ControlHanlding DefaultMode
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    configflxTransactionTypes
    LoadFlxTransactionTypes adoConn 'this function loads control accounts
    adoConn.Close
    Set adoConn = Nothing
    
''   Dim iRow    As Integer
'''Saved? S->Saved, C->New set of CA for a client, N->New CA, A->Amend, D->Decide
''
''   iRow = 1
''   Do While iRow < flxTransactionTypes.Rows
''      If flxTransactionTypes.TextMatrix(iRow, 9) = "N" Then
''         flxTransactionTypes.RemoveItem iRow
''         iRow = iRow - 1
''      End If
''
''      If flxTransactionTypes.TextMatrix(iRow, 9) = "A" Then
''         If flxTransactionTypes.TextMatrix(iRow, 3) <> flxTransactionTypes.TextMatrix(iRow, 10) Then
''            For i = 0 To cboCC.ListCount - 1
''               If cboCC.Column(0, i) = flxTransactionTypes.TextMatrix(iRow, 10) Then
''                  flxTransactionTypes.TextMatrix(iRow, 3) = flxTransactionTypes.TextMatrix(iRow, 10)
''                  flxTransactionTypes.TextMatrix(iRow, 4) = cboCC.Column(1, i)
''               End If
''            Next i
''         End If
''
''         flxTransactionTypes.TextMatrix(iRow, 7) = flxTransactionTypes.TextMatrix(iRow, 11)
''         flxTransactionTypes.TextMatrix(iRow, 9) = "S"
''      End If
''
''      If flxTransactionTypes.TextMatrix(iRow, 9) = "D" Then
''         flxTransactionTypes.TextMatrix(iRow, 3) = ""
''         flxTransactionTypes.TextMatrix(iRow, 4) = ""
''         flxTransactionTypes.TextMatrix(iRow, 7) = ""
''         flxTransactionTypes.TextMatrix(iRow, 9) = "C"
''         flxTransactionTypes.TextMatrix(iRow, 10) = ""
''         flxTransactionTypes.TextMatrix(iRow, 11) = ""
''      End If
''      iRow = iRow + 1
''   Loop
End Sub

Private Sub cmdCAClose_Click()
   If iNewEditCC <> 0 Then
      If MsgBox("Control Accounts have been modified." & Chr(13) & _
                "Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Control Code") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Public Sub ControlHanlding(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         If frmMMain.IsRibbonVersion Then
'            picCC.Visible = False
            flxTransactionTypes.Enabled = True
            iNewEditCC = 0
            cmdCANew.Enabled = True
            cmdCASave.Enabled = False
            cmdCACancel.Enabled = False
            cmdDelete.Enabled = True
            cmdCAClose.Enabled = True
         Else
            cmdCANew.Visible = False
            cmdCASave.Visible = False
            cmdCACancel.Visible = False
            cmdDelete.Visible = False
         End If

      Case ComponentMode.NewEntryMode
         iNewEditCC = 1
         cmdCANew.Enabled = False
         cmdCASave.Enabled = True
         cmdCACancel.Enabled = True
         cmdDelete.Enabled = False
         cmdCAClose.Enabled = False

      Case ComponentMode.EditMode
         cmdCANew.Enabled = False
         cmdCASave.Enabled = True
         cmdCACancel.Enabled = True
         cmdDelete.Enabled = False
         cmdCAClose.Enabled = False

         iNewEditCC = 2

      Case ComponentMode.SavedMode
'         picCC.Visible = False

         iNewEditCC = 0
   End Select
End Sub
'
'Private Sub cmdCAClose_KeyPress(KeyAscii As Integer)
'   Dim X As MSForms.ReturnInteger
'   X = 27
'   If KeyAscii = 27 And picCC.Visible Then cboCC_KeyPress (X)
'End Sub

Private Sub cmdCAClose_LostFocus()
'   tabSettings.SetFocus
End Sub

Private Sub cmdCAClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCAClose_Click
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCancelVAT_Click()
    txtVatDesp.text = ""
    txtVatRate.text = ""
    chkInUse.Value = 0
    flxVat.Enabled = True
    VAT_MODIFIED_ = False
    txtVatRate.Locked = False
    txtVatDesp.Locked = False
    chkInUse.Enabled = True
    chkOnReturn.Enabled = True

    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    LoadflxVat adoConn
    cmdUpdate.Enabled = True
    cmdSave.Enabled = False
    cmdCancelVAT.Enabled = False
    adoConn.Close
    Set adoConn = Nothing
    FocusControl cmdClose
End Sub

Private Sub cmdCANew_Click()
'bellow 2 lines has been added by anol 14 Mar 2015
'issue 561
   cmdCANew.Enabled = False
   
   Exit Sub
   If txtClientList.text = "" Then Exit Sub

   Load frmCtrlAcc
   frmCtrlAcc.AddNew = True
   ControlHanlding NewEntryMode
   frmCtrlAcc.lblClient.Caption = txtClientList.text & "/" & txtClientList.Tag
   frmCtrlAcc.Show
   Me.Enabled = False
End Sub

Private Sub cmdCASave_Click() 'We are saving control accounts here
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim iRow As Integer
   Dim d    As Integer
   
   'If iNewEditCC = 0 Then Exit Sub
 'We dont need these validation ' anol 20200701
''   Dim szForm As String
''
''   szForm = WhoIsOpen(Me.Name)
''   If szForm <> "NONE" Then
''      MsgBox "Please close """ & szForm & """ and retry to save control account", vbInformation + vbOKOnly, "Control Account"
''      Exit Sub
''   End If
''   'test all the control account are filled
''   For iRow = 1 To flxTransactionTypes.Rows - 2
''      If Trim(flxTransactionTypes.TextMatrix(iRow, 3)) = "" Then
''          MsgBox "Please fill all control account", vbInformation + vbOKOnly, "Control Account"
''          Exit Sub
''      End If
''   Next
   'If picCC.Visible Then cboCC_Click

   'On Error GoTo ErrorHandler

  'Here  d variable is writing CADisOrder field in the nominal ledger table which is mainly maintaing the serial number for each control account for 1 client

  'DNA means  D->Decide N->New CA A->Amend
  'CAType Means it will store a single letter for S P I O R  to mark as control account P for purchase control account O for output control account
  'CAFixed is the marker if it is a control account or not if it is control account then it is marking it as true else false
  'CAPosting we are not using this field
  'CAFixed if you set any control account then CAFixed is true this is the boolean for if this account is a control account or not
  'Exit Sub
   adoConn.Open getConnectionString
        d = 0
             szSQL = "UPDATE NominalLedger " & _
                            "SET    CAName = ' ', " & _
                            "CAFixed = false , " & _
                            "CADisOrder = " & d & ", " & _
                            "Posting = true, " & _
                            "CAType = '' " & _
                            "WHERE ClientID = '" & txtClientList.text & "';" 'updating for everything into zero state
            adoConn.Execute szSQL
   
        d = 1
   For iRow = 1 To flxTransactionTypes.Rows - 1
   
'         If (flxTransactionTypes.TextMatrix(iRow, 9) = "D" Or _
'               flxTransactionTypes.TextMatrix(iRow, 9) = "N" Or _
'               flxTransactionTypes.TextMatrix(iRow, 9) = "A") And _
'                     flxTransactionTypes.TextMatrix(iRow, 3) <> "" Then
       ' If iRow = 8 Then
               ' I   Need to keep original account number form Nominal account table because when you empty the code for existing setup control account
               ' in where clause you cannot search with an empty code to update then you need the original value from database
               'If flxTransactionTypes.TextMatrix(iRow, 11) = "A" Then 'this means you have changed the code, now it can be either empty or a new code
                            'it it is a new code you also need to deal with the old code and make it empty
'                        If flxTransactionTypes.TextMatrix(iRow, 3) = "" Then 'if you are making control empty then this part
'                                 d = 0
'                                 'We are not using CAPosting field
'                                 szSQL = "UPDATE NominalLedger " & _
'                                            "SET    CAName = '', " & _
'                                            "CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", False, True) & ", " & _
'                                            "CADisOrder = " & d & ", " & _
'                                            "Posting = true, " & _
'                                            "CAType = '', " & _
'                                            "Type = '" & flxTransactionTypes.TextMatrix(iRow, 9) & "' " & _
'                                            "WHERE  Code = '" & flxTransactionTypes.TextMatrix(iRow, 10) & "' AND " & _
'                                            "ClientID = '" & txtClientList.text & "';" 'updating for new nominal code
'                                         adoConn.Execute szSQL
'
'                         Else        'if you are changing  control account  then this part
                                 'Update for new account
                                 szSQL = "UPDATE NominalLedger " & _
                                        "SET    CAName = '" & flxTransactionTypes.TextMatrix(iRow, 1) & "', " & _
                                        "CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", False, True) & ", " & _
                                        "CADisOrder = " & d & ", " & _
                                        "Posting = true, " & _
                                        "Type = '" & flxTransactionTypes.TextMatrix(iRow, 9) & "', " & _
                                        "CAType = '" & flxTransactionTypes.TextMatrix(iRow, 8) & "', " & _
                                        "CAPosting = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", True, False) & " " & _
                                        "WHERE  Code = '" & flxTransactionTypes.TextMatrix(iRow, 3) & "' AND " & _
                                        "ClientID = '" & txtClientList.text & "';" 'updating for new nominal code
                                        'szSQL = "UPDATE NominalLedger SET    CAName = 'Managing Agents control Account (B/S)', CAFixed = True, CADisOrder = 1, Posting = true, Type = 'YES', CAType = 'MF', CAPosting = False WHERE  Code = '5040' AND ClientID = 'CARSONPR';"
                                         adoConn.Execute szSQL
''                                    'Clear for old account
''                                     szSQL = "UPDATE NominalLedger " & _
''                                        "SET    CAName = '" & flxTransactionTypes.TextMatrix(iRow, 1) & "', " & _
''                                        "CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", False, True) & ", " & _
''                                        "CADisOrder =0, " & _
''                                        "CAType = '' " & _
''                                        "WHERE  Code = '" & flxTransactionTypes.TextMatrix(iRow, 10) & "' AND " & _
''                                        "ClientID = '" & txtClientList.text & "';" 'updating for old nominal code
''
''                                        'szSQL = "UPDATE NominalLedger SET    CAName = 'Client/Landlord Control Account (B/S)', CAFixed = False, CADisOrder =0, CAType = '', CAPosting = True WHERE  Code = '2109' AND ClientID = 'CARSONPR';"
''                                         adoConn.Execute szSQL
'                        End If
          '     End If
                d = d + 1
'                Else
'                         szSQL = "UPDATE NominalLedger " & _
'                                        "SET    CAName = '" & flxTransactionTypes.TextMatrix(iRow, 1) & "', " & _
'                                        "CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", False, True) & ", " & _
'                                        "CADisOrder = " & d & ", " & _
'                                        "Posting = true, " & _
'                                        "Type = '" & flxTransactionTypes.TextMatrix(iRow, 9) & "', " & _
'                                        "CAType = '" & flxTransactionTypes.TextMatrix(iRow, 8) & "', " & _
'                                        "CAPosting = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "YES", True, False) & " " & _
'                                        "WHERE  Code = '" & flxTransactionTypes.TextMatrix(iRow, 3) & "' AND " & _
'                                        "ClientID = '" & txtClientList.text & "';" 'updating for new nominal code
'                                        'szSQL = "UPDATE NominalLedger SET    CAName = 'Managing Agents control Account (B/S)', CAFixed = True, CADisOrder = 1, Posting = true, Type = 'YES', CAType = 'MF', CAPosting = False WHERE  Code = '5040' AND ClientID = 'CARSONPR';"
'                                         adoConn.Execute szSQL
'                                          d = d + 1
'                End If
            'Updating the base NL one which has been changed
           'by anol 17 Nov 2015
'              If (flxTransactionTypes.TextMatrix(iRow, 9) = "A" And flxTransactionTypes.TextMatrix(iRow, 3) <> flxTransactionTypes.TextMatrix(iRow, 10)) Then
'                    szSQL = "UPDATE NominalLedger " & _
'                    "SET   CAName = NULL, " & _
'                           " CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 6) = "YES", True, False) & ", " & _
'                           "CAType = '', " & _
'                           "CADisOrder = 0, Posting = 1, " & _
'                           "CAPosting = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "NA", True, False) & " " & _
'                    "WHERE  Code = '" & flxTransactionTypes.TextMatrix(iRow, 10) & "' AND " & _
'                           "ClientID = '" & txtClientList.text & "';" 'updating for old nominal code
'                           adoconn.Execute szSQL
'             End If
'         End If
'         If flxTransactionTypes.TextMatrix(iRow, 3) = "" Then 'when you empty a grid line .
'                     szSQL = "UPDATE NominalLedger " & _
'                    "SET   CAName = NULL, " & _
'                           " CAFixed = " & IIf(flxTransactionTypes.TextMatrix(iRow, 6) = "YES", True, False) & ", " & _
'                           "CAType = '', " & _
'                           "CADisOrder = 0, Posting = 1, " & _
'                           "CAPosting = " & IIf(flxTransactionTypes.TextMatrix(iRow, 7) = "NA", True, False) & " " & _
'                    "WHERE  CAName = '" & flxTransactionTypes.TextMatrix(iRow, 2) & "' AND " & _
'                           "ClientID = '" & txtClientList.text & "';" 'updating for old nominal code
'                           adoconn.Execute szSQL
'         End If
   Next iRow
   LoadFlxTransactionTypes adoConn 'this function loads control accounts
   adoConn.Close
   Set adoConn = Nothing
    
   ControlHanlding DefaultMode
   ShowMsgInTaskBar "Control Account codes have been updated", "Y", "P"
   cmdCASave.Enabled = False
   Exit Sub
ErrorHandler:
   If iNewEditCC = 2 Then MsgBox "System could not edit Control Code." & Chr(13) & Err.Number & " " & Err.description, vbCritical + vbOKOnly, "Error"
   ControlHanlding DefaultMode
End Sub

Private Sub cmdClose_Click()
   If tAddNew > 0 Then
      If MsgBox("Do you want to save before close the form?", vbQuestion + vbYesNo, "Default Client Deposit A/C") = vbNo Then
         Unload Me
      Else
         Exit Sub
      End If
   End If
   Unload Me
End Sub

Private Sub cmdClose_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdClose_Click
End Sub

Private Sub cmdCloseOpt_Click()
    Unload Me
End Sub

Private Sub cmdCloseOpt_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdCloseOpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCloseOpt_Click
End Sub

Private Sub cmdCloseSettings_Click()
   Unload Me
End Sub

Private Sub cmdCloseSettings_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdCloseSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCloseSettings_Click
End Sub

Private Sub cmdDelete_Click()
'bellow 2 lines has been added by anol 14 Mar 2015
'issue 561
    cmdDelete.Enabled = False
    Exit Sub
    'End of miodification
   If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 6) = "YES" Then
      ShowMsgInTaskBar "This is main control account and will not be deleted.", "Y", "N"
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If NoTransFound(adoConn) Then
      If MsgBox("Do you wish to delete the control code?", vbQuestion + vbYesNo, "Control Code") = vbYes Then
'szHeader$ = "|<CAName|<Type|<NCode|<NName|ClientID|Fixed|<Posting|TypeID|saved|OriNC|OriPosting"
'                 1       2     3      4       5      6        7      8     9     10      11
'Saved? S->Saved, C->New set of CA for a client, N->New CA, A->Amend, D->Decide
         If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 9) = "S" Or _
               flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 9) = "A" Then
            adoConn.Execute "UPDATE NominalLedger " & _
                            "SET    CAName = '', CAFixed = False, " & _
                                   "CADisOrder = 0, CAType = '', CAPosting = False " & _
                            "WHERE  Code = '" & flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 3) & "' AND " & _
                                   "ClientID = '" & txtClientList.text & "';"
         End If
         flxTransactionTypes.RemoveItem flxTransactionTypes.row
         ShowMsgInTaskBar "The control account has been deleted", "Y", "P"
      End If
   Else
      ShowMsgInTaskBar "There are transactions have been posted on this control account", "Y", "N"
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function NoTransFound(adoConn As ADODB.Connection) As Boolean
   NoTransFound = True
   If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 3) = "" Then Exit Function

   Dim adoRST As New ADODB.Recordset

   adoRST.Open "SELECT * FROM NLPosting WHERE NOMINAL_CODE = '" & flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 3) & "';", adoConn, adOpenStatic, adLockReadOnly

   NoTransFound = IIf(adoRST.EOF Or adoRST.BOF, True, False)

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Sub cmdEdit_Click()
   ButtonHanlding EditMode
  
   tAddNew = 2
End Sub

Private Sub cmdEditOptions_Click()
   txtLeaseEndDays.Locked = False
   txtLeaseEndDays.SetFocus
   SelTxtInCtrl txtLeaseEndDays
   chkAlarm.Enabled = True

   cmdSaveOptions.Enabled = True
   cmdEditOptions.Enabled = False
   CheckUnit.Enabled = True
   CheckClient.Enabled = True
   CheckLessee.Enabled = True
   CheckSupplier.Enabled = True
   CheckProperty.Enabled = True
   CheckManaging.Enabled = True
End Sub

Private Sub cmdEditSettings_Click()
   Frame4.Enabled = True
   txtFromEmail.SetFocus
   cmdSaveSettings.Enabled = True
   cmdEditSettings.Enabled = False
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdExit_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdExit_Click
End Sub

Private Sub cmdExpRepGen_Click()
   If txtLeaseEndDays.text = "" Then Exit Sub

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpiredLeases.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   ' Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue Val(txtLeaseEndDays.text)

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdLDCancel_Click()
   If tAddNew = 1 Then
      If MsgBox("Would you like to cancel adding new record?", vbQuestion + vbYesNo, "Default Client Deposit A/C") = vbYes Then
         tAddNew = 0
         ButtonHanlding DefaultMode
         Exit Sub
      End If
   Else
      If MsgBox("Would you like to cancel modifying the record?", vbQuestion + vbYesNo, "Default Client Deposit A/C") = vbYes Then
         tAddNew = 0
         flxDeposit.row = 0
         ButtonHanlding DefaultMode
      End If
   End If
End Sub

Private Sub cmdNew_Click()
   ButtonHanlding NewEntryMode
   tAddNew = 1
End Sub

Private Sub ButtonHanlding(ByVal mode As ComponentMode)
   Select Case mode

   Case ComponentMode.DefaultMode
      cmdNew.Enabled = True
      cmdEdit.Enabled = False
      cmdSave.Enabled = False
      cmdLDCancel.Enabled = False
      cmdClose.Enabled = True

      cmbClient.Locked = True
      cmbBank.Locked = True
      cmbDNC.Locked = True
      cmbClient.text = ""
      cmbBank.text = ""
      cmbDNC.text = ""

   Case ComponentMode.NewEntryMode
      cmdNew.Enabled = False
      cmdEdit.Enabled = False
      cmdSave.Enabled = True
      cmdLDCancel.Enabled = True
      cmdClose.Enabled = True

      cmbClient.Locked = False
      cmbBank.Locked = False
      cmbDNC.Locked = False
      cmbClient.text = ""
      cmbBank.text = ""
      cmbDNC.text = ""

   Case ComponentMode.EditMode
      cmdNew.Enabled = False
      cmdEdit.Enabled = False
      cmdSave.Enabled = True
      cmdLDCancel.Enabled = True
      cmdClose.Enabled = True

      cmbClient.Locked = True
      cmbBank.Locked = False
      cmbDNC.Locked = False

   Case ComponentMode.GridRowOnSelection
      cmdNew.Enabled = True
      cmdEdit.Enabled = True
      cmdSave.Enabled = False
      cmdLDCancel.Enabled = False
      cmdClose.Enabled = True

      cmbClient.Locked = True
      cmbBank.Locked = True
      cmbDNC.Locked = True
   End Select
End Sub

Private Sub cmdOK_Click()
   If Not cmdPathEdit.Enabled Then
      Dim sSQLQuery_ As String
      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString

      sSQLQuery_ = "UPDATE SECONDARYCODE " & _
                   "SET SECONDARYCODE.VALUE = '" & txtFilePath.text & "' " & _
                   "WHERE PrimaryCode = 'FPATH' AND " & _
                     "CODE = 'FPATH2'"
'Debug.Print sSQLQuery_
      adoConn.Execute sSQLQuery_
   End If

   szReportPath = txtFilePath.text

   Unload Me
End Sub

Private Sub cmdPathEdit_Click()
   txtFilePath.Locked = False
   cmdPathEdit.Enabled = False
End Sub

Private Sub cmdSave_Click()
   Dim i As Integer

   If cmbClient.Value = "" Then
      Exit Sub
   End If
   
   For i = 1 To flxDeposit.Rows - 1
      If flxDeposit.TextMatrix(i, 1) = cmbClient.Value Then Exit For
   Next i

   If i < flxDeposit.Rows And tAddNew = 1 Then
      ShowMsgInTaskBar "You can only have one deposit default for each client."
      Exit Sub
   End If

   If cmbBank.text = "" Then
      ShowMsgInTaskBar "Please select a bank code from the drop down list."
      cmbBank.SetFocus
      Exit Sub
   End If

   If cmbDNC.text = "" Then
      ShowMsgInTaskBar "Please select a nominal code from the drop down list.", , "N"
      cmbDNC.SetFocus
      Exit Sub
   End If

'  Saving record  *******************************************************************************
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "UPDATE Client " & _
           "SET spare1 = '" & cmbBank.Value & "', " & _
               "spare2 = '" & cmbDNC.Value & "' " & _
           "WHERE ClientID = '" & cmbClient.Value & "';"

   adoConn.Execute szSQL

   ButtonHanlding DefaultMode
   tAddNew = 0
   LoadClientData adoConn

   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "The data has been saved successfully."
End Sub

Private Sub cmdSaveOptions_Click()
   If txtLeaseEndDays.text = "" Then
      ShowMsgInTaskBar "Please enter the value for alarm days.", , "D"
      txtLeaseEndDays.SetFocus
      Exit Sub
   End If

   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL, sSql As String
   
   adoConn.Open getConnectionString
   
   On Error GoTo ErrHandler

'  Remember choice
   
   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If
   
   adoRST.Close
   Set adoRST = Nothing

   If UBound(szaChoice) < 0 Then ReDim szaChoice(5) As String

   If CheckUnit.Value Then
      szaChoice(0) = "U"
   Else
      szaChoice(0) = ""
   End If
   If CheckClient.Value Then
      szaChoice(1) = "CL"
   Else
      szaChoice(1) = ""
   End If
   If CheckLessee.Value Then
      szaChoice(2) = "L"
   Else
      szaChoice(2) = ""
   End If
   If CheckSupplier.Value Then
      szaChoice(3) = "S"
   Else
      szaChoice(3) = ""
   End If
   If CheckProperty.Value Then
      szaChoice(4) = "P"
   Else
      szaChoice(4) = ""
   End If
   If CheckManaging.Value Then
      szaChoice(5) = "MA"
   Else
      szaChoice(5) = ""
   End If
   
   szChoice = Join(szaChoice, "#")

   szSQL = "UPDATE SecondaryCode SET SecondaryCode.VALUE = '" & szChoice & "' WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoConn.Execute szSQL
  
   If (Not IS_REMIND = chkAlarm.Value) Then
        IS_REMIND = chkAlarm.Value
        Dim s As String
        s = IIf(chkAlarm.Value, "Y", "N")
        sSql = "UPDATE SECONDARYCODE " & _
                   "SET SECONDARYCODE.VALUE = '" & s & "' " & _
                   "WHERE PrimaryCode = 'ATLD' AND Code = 'TA';"
        adoConn.Execute sSql
   End If
                        
   adoConn.Execute "UPDATE SECONDARYCODE " & _
                   "SET SECONDARYCODE.VALUE = '" & txtLeaseEndDays.text & "' " & _
                   "WHERE PrimaryCode = 'ATLD' AND Code = 'TL';"
   
   
   adoConn.Close
   Set adoConn = Nothing
   
   If ((Not IS_REMIND = chkAlarm.Value) Or (Not DAYS_TO_REMIND = txtLeaseEndDays.text)) Then
       Dim conUnit_ As New ADODB.Connection
       Dim rstUnit_ As New ADODB.Recordset
       Dim sSQLQuery_ As String, sTerminateDate As String, dTermDate As Date, sMsg As String
       Dim iDays As Double

       iDays = CInt(txtLeaseEndDays.text) * -1
       sMsg = "Lease Expires in " & txtLeaseEndDays.text & " Days."
       
        conUnit_.Open getConnectionString
    
        sSQLQuery_ = "SELECT Alarm, Reminder_ID, TerminateDate, LeaseID " & _
                     "FROM LEASEDETAILS WHERE not isNull(TerminateDate)"
    
        rstUnit_.Open sSQLQuery_, conUnit_, adOpenDynamic, adLockOptimistic
        
        While Not rstUnit_.EOF
            If (chkAlarm.Value) Then
                dTermDate = rstUnit_!TerminateDate
                dTermDate = DateAdd("d", iDays, dTermDate)
                rstUnit_!Alarm = "Y"
                rstUnit_!Reminder_ID = NewReminder(CStr(dTermDate), "0830", sMsg, "LeaseDetails", rstUnit_!LeaseID)
            Else
                rstUnit_!Alarm = "N"
                ClearReminder rstUnit_!Reminder_ID
            End If
            rstUnit_.Update
            rstUnit_.MoveNext
        Wend
        
       rstUnit_.Close
       Set rstUnit_ = Nothing
       conUnit_.Close
       Set conUnit_ = Nothing
   End If
   
   ShowMsgInTaskBar "Data has been updated successfully."
   txtLeaseEndDays.Locked = True
   cmdSaveOptions.Enabled = False
   cmdEditOptions.Enabled = True
   chkAlarm.Enabled = False
   CheckUnit.Enabled = False
   CheckClient.Enabled = False
   CheckLessee.Enabled = False
   CheckSupplier.Enabled = False
   CheckProperty.Enabled = False
   CheckManaging.Enabled = False
   
   Exit Sub

ErrHandler:
   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
End Sub

Private Sub cmdSaveSettings_Click()
   Frame4.Enabled = False
   cmdSaveSettings.Enabled = False
   cmdEditSettings.Enabled = True

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM ShoppingCentre;"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   ' issue 323 Emails are not being sent out from STVM ..this was due to space in smtp fixed by anol 20170228
   adoRST.Fields.Item("Email1").Value = Trim(txtFromEmail.text)
   adoRST.Fields.Item("SMTP").Value = Trim(txtSMTP.text)
   adoRST.Fields.Item("UName").Value = Trim(txtUName.text)
   adoRST.Fields.Item("Pws").Value = Trim(txtPws.text)
   adoRST.Fields.Item("Port").Value = Trim(txtPort.text)
   adoRST.Fields.Item("SSL").Value = chkSSL.Value 'added by anol 20161125
   adoRST.Fields.Item("TLS").Value = chkTLS.Value 'added by anol 20191130
   adoRST.Update

   Set adoRST = Nothing
   Set adoConn = Nothing

   szFromEmail = txtFromEmail.text
   szSMTPserver = txtSMTP.text
   szUName = txtUName.text
   szPws = txtPws.text
   szPort = txtPort.text
   szSSL = chkSSL.Value
   szTLS = chkTLS.Value
'
'   MsgBox "Heighly recommendation: Restart the Prestige.", vbInformation + vbOKOnly, "System Message"
End Sub

Private Sub cmdSaveVAT_Click()
   Dim sSQLQuery_ As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   If Trim(txtVatRate.text) = "" Then
        MsgBox "Please enter VAT Rate"
        Exit Sub
   End If
   adoConn.Open getConnectionString

   sSQLQuery_ = "SELECT VAT_ID, VAT_RATE, DESCRIPTIONS, IN_USE,OnReturn " & _
                "FROM tlbVatCode where VAT_ID =" & flxVat.TextMatrix(flxVat.row, 0) & ";"

   adoRST.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
   If Not adoRST.EOF Then
        adoRST.Fields.Item("VAT_RATE").Value = txtVatRate.text
        adoRST.Fields.Item("DESCRIPTIONS").Value = txtVatDesp.text
        If chkInUse.Value = 1 Then
             adoRST.Fields.Item("IN_USE").Value = True
             adoRST.Update
        End If
        If chkInUse.Value = 0 Then
             adoRST.Fields.Item("IN_USE").Value = False
             adoRST.Update
        End If
        If chkOnReturn = 1 Then
             adoRST.Fields.Item("OnReturn").Value = True
             adoRST.Update
        End If
        If chkOnReturn = 0 Then
             adoRST.Fields.Item("OnReturn").Value = False
             adoRST.Update
        End If
   End If
   



'   While Not adoRst.EOF
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 2) <> adoRst.Fields.Item("VAT_RATE").Value Then
'         adoRst.Fields.Item("VAT_RATE").Value = flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 2)
'         adoRst.Update
'      End If
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 3) <> adoRst.Fields.Item("DESCRIPTIONS").Value Then
'         adoRst.Fields.Item("DESCRIPTIONS").Value = flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 3)
'         adoRst.Update
'      End If
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 4) = "YES" And Not adoRst.Fields.Item("IN_USE").Value Then
'         adoRst.Fields.Item("IN_USE").Value = True
'         adoRst.Update
'      End If
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 4) = "NO" And adoRst.Fields.Item("IN_USE").Value Then
'         adoRst.Fields.Item("IN_USE").Value = False
'         adoRst.Update
'      End If
'      'on return column
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 5) = "YES" Then
'         adoRst.Fields.Item("OnReturn").Value = True
'         adoRst.Update
'      End If
'      If flxVat.TextMatrix(adoRst.Fields.Item("VAT_ID").Value + 1, 5) = "NO" Then
'         adoRst.Fields.Item("OnReturn").Value = False
'         adoRst.Update
'      End If
'
'      adoRst.MoveNext
'   Wend

   adoRST.Close
   Set adoRST = Nothing

   LoadflxVat adoConn
   adoConn.Close
   Set adoConn = Nothing
   VAT_MODIFIED_ = False

   MsgBox "VAT has been updated successfully.", vbInformation, "Data saved"
   cmdUpdate.Enabled = True
   cmdSaveVAT.Enabled = False
   flxVat.Enabled = True
   cmdCancelVAT.Enabled = False
End Sub

Private Sub cmdStatementPathCancel_Click()
   Unload Me
End Sub

Private Sub cmdStatementPathCancel_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdStatementPathEdit_Click()
   If flxStFlPth.row < 1 Then
      ShowMsgInTaskBar "Please select a property from the list.", , "N"
      Exit Sub
   End If

   flxStFlPth.TextMatrix(flxStFlPth.row, 2) = SelectAFile("rpt")
End Sub

Private Sub cmdSubTypeClose_Click()
   If bEditSubType Then
      If MsgBox("Sub Types have been modified." & Chr(13) & _
                "Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Sub Types") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Private Sub cmdSubTypeClose_LostFocus()
   tabSettings.SetFocus
End Sub

Private Sub cmdSubTypeClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdSubTypeClose_Click
End Sub

Private Sub cmdSubTypeSave_Click()
   Dim adoConn    As New ADODB.Connection
   Dim iRow       As Integer
   Dim szSQL      As String

   adoConn.Open getConnectionString

   For iRow = 1 To flxSubTypes.Rows - 1
      If flxSubTypes.TextMatrix(iRow, 3) = "M" Then

         szSQL = "UPDATE NLSubTypes " & _
                 "SET    STDescription = '" & flxSubTypes.TextMatrix(iRow, 2) & "' " & _
                 "WHERE  STCode = '" & flxSubTypes.TextMatrix(iRow, 0) & "';"
         adoConn.Execute szSQL
         flxSubTypes.TextMatrix(iRow, 3) = ""
      End If
   Next iRow
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdUpdate_Click()
    If flxVat.TextMatrix(flxVat.row, 2) <> txtVatRate.text Then VAT_MODIFIED_ = True
    If flxVat.TextMatrix(flxVat.row, 3) = txtVatDesp.text Then VAT_MODIFIED_ = True
    If flxVat.TextMatrix(flxVat.row, 4) = "YES" And chkInUse.Value = 0 Then VAT_MODIFIED_ = True
    If flxVat.TextMatrix(flxVat.row, 4) = "NO" And chkInUse.Value = 1 Then VAT_MODIFIED_ = True
    
    If flxVat.TextMatrix(flxVat.row, 5) = "YES" And chkOnReturn.Value = 0 Then VAT_MODIFIED_ = True
    If flxVat.TextMatrix(flxVat.row, 5) = "NO" And chkOnReturn.Value = 1 Then VAT_MODIFIED_ = True
    
    flxVat.TextMatrix(flxVat.row, 2) = txtVatRate.text
    flxVat.TextMatrix(flxVat.row, 3) = txtVatDesp.text
    flxVat.TextMatrix(flxVat.row, 4) = IIf(chkInUse.Value = 1, "YES", "NO")
    flxVat.TextMatrix(flxVat.row, 5) = IIf(chkOnReturn.Value = 1, "YES", "NO")
   
    cmdSaveVAT.Enabled = True
    cmdUpdate.Enabled = False
    cmdCancelVAT.Enabled = True
    flxVat.Enabled = False
    txtVatRate.Locked = False
    txtVatDesp.Locked = False
    chkInUse.Enabled = True
    chkOnReturn.Enabled = True
    FocusControl txtVatRate
    txtVatRate.SelStart = 0
    txtVatRate.SelLength = Len(txtVatRate.text)
End Sub

Private Sub flxDeposit_RowColChange()
   If flxDeposit.TextMatrix(1, 0) = "" Then Exit Sub

   cmbClient.Value = flxDeposit.TextMatrix(flxDeposit.row, 1)
   cmbBank.Value = flxDeposit.TextMatrix(flxDeposit.row, 3)
   cmbDNC.Value = flxDeposit.TextMatrix(flxDeposit.row, 5)

   ButtonHanlding GridRowOnSelection
End Sub

Private Sub flxStFlPth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxStFlPth.ToolTipText = flxStFlPth.TextMatrix(flxStFlPth.MouseRow, flxStFlPth.MouseCol)
End Sub

Private Sub flxSubTypes_DblClick()
   If flxSubTypes.col = 2 Then
      txtSubType.Top = 0
      txtSubType.Left = 0
      txtSubType.Width = flxSubTypes.ColWidth(2)
      txtSubType.Height = 285
      txtSubType.text = flxSubTypes.TextMatrix(flxSubTypes.row, 2)

      picSubType.Top = flxSubTypes.CellTop + flxSubTypes.Top
      picSubType.Left = flxSubTypes.CellLeft + flxSubTypes.Left
      picSubType.Visible = True
      flxSubTypes.Enabled = False
      txtSubType.SetFocus
      picSubType.Width = txtSubType.Width
      picSubType.Height = txtSubType.Height
   End If
End Sub

Private Sub flxTransactionTypes_DblClick()
   If (flxTransactionTypes.col = 1 Or flxTransactionTypes.col = 2) Then
      If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 6) = "NO" Then
         Load frmCtrlAcc
         frmCtrlAcc.AddNew = False
         ControlHanlding EditMode
         frmCtrlAcc.lblClient.Caption = txtClientList.text & "/" & txtClientList.Tag
         frmCtrlAcc.txtCtrlName.text = flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 1)
         frmCtrlAcc.cboType.Value = flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 8)

         frmCtrlAcc.Show
         Me.Enabled = False
      Else
         'ShowMsgInTaskBar "This is main control account, cannot be amended", "Y", "N"
      End If
   End If
   If flxTransactionTypes.col = 3 Or flxTransactionTypes.col = 4 Then
      If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 2) = "" Then Exit Sub
      ControlHanlding EditMode
      txtSearchNC.text = ""
      txtSearchNName.text = ""
      
      Call LoadflxNominal
      flxTransactionTypes.col = 3
      'strnage thing happen hare . if you set manually 1500/any other fixed top this pic control disappears
      'So instead I increased the form height
      picNominal.Top = flxTransactionTypes.CellTop + flxTransactionTypes.Top
      picNominal.Left = flxTransactionTypes.CellLeft + flxTransactionTypes.Left
      If flxNominal.Rows > 2 Then
            picNominal.Visible = True
            flxTransactionTypes.Enabled = False
            'txtSearchNC.SetFocus
      Else
            MsgBox "Please create chart of account for this client first", vbInformation, "Warning!"
      End If
      FocusControl txtSearchNC
   End If
   If flxTransactionTypes.col = 7 And (flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "NO" Or flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "YES") Then
      If flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "NO" Then
'            ControlHanlding EditMode
'            flxTransactionTypes.Enabled = False
            'flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "YES"
            'flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 11) = "A"
      Else
    '       ShowMsgInTaskBar "This is fixed control account code", "Y", "N"
            'flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 7) = "NO"
             'flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 11) = "A"
      End If
      cmdCASave.Enabled = True
   End If
End Sub

Private Sub flxTransactionTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
End Sub

Private Sub flxTransactionTypes_RowColChange()
   HighLightRowFlxGrid flxTransactionTypes, flxTransactionTypes.row
End Sub

Private Sub Form_Load()
  ' frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 13140 '8985
   Me.Height = 7200 '5895
   Me.BackColor = MODULEBACKCOLOR
   
    tabSettings.BackColor = Me.BackColor
    txtVatRate.Locked = True
    txtVatDesp.Locked = True
    chkInUse.Enabled = False
    chkOnReturn.Enabled = False
    
   tabSettings.Tab = 0
   tAddNew = 0
   flxDeposit.row = 0
   bEditSubType = False

   txtAppPath.text = App.Path
   txtFilePath.text = szReportPath
   txtFilePath.Locked = True

   bConfigureFlxDeposit = False
   ButtonHanlding DefaultMode

   Dim sSQLQuery_ As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   adoConn.Open getConnectionString

   sSQLQuery_ = "SELECT VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PrimaryCode = 'ATLD' AND Code = 'TL';"
   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
   txtLeaseEndDays.text = adoRST.Fields.Item("VALUE").Value
   txtLeaseEndDays.Locked = True
   DAYS_TO_REMIND = txtLeaseEndDays.text
   adoRST.Close

   sSQLQuery_ = "SELECT VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PrimaryCode = 'ATLD' AND Code = 'TA';"
   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
   chkAlarm.Value = IIf(adoRST.Fields.Item("VALUE").Value = "Y", 1, 0)
   chkAlarm.Enabled = False
   IS_REMIND = chkAlarm.Value
   adoRST.Close

   sSQLQuery_ = "SELECT * " & _
                "FROM ShoppingCentre;"
   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
   txtFromEmail.text = IIf(IsNull(adoRST.Fields.Item("Email1").Value), "", adoRST.Fields.Item("Email1").Value)
   txtSMTP.text = IIf(IsNull(adoRST.Fields.Item("SMTP").Value), "", adoRST.Fields.Item("SMTP").Value)
   txtUName.text = IIf(IsNull(adoRST.Fields.Item("UName").Value), "", adoRST.Fields.Item("UName").Value)
   txtPws.text = IIf(IsNull(adoRST.Fields.Item("Pws").Value), "", adoRST.Fields.Item("Pws").Value)
   txtPort.text = IIf(IsNull(adoRST.Fields.Item("Port").Value), "", adoRST.Fields.Item("Port").Value)
   chkSSL.Value = IIf((IIf(IsNull(adoRST.Fields.Item("SSL").Value), False, adoRST.Fields.Item("SSL").Value)), 1, 0) 'added by anol 20161125
   chkTLS.Value = IIf((IIf(IsNull(adoRST.Fields.Item("TLS").Value), False, adoRST.Fields.Item("TLS").Value)), 1, 0) 'added by anol 20191130
   adoRST.Close

'  #############################
'  Load all combos
'  #############################
   LoadComboes adoConn

'  #############################
'  All Properties statement path
'  #############################
   StatementPath adoConn

   LoadflxVat adoConn

'  #############################
'  Control Accounts loading
'  #############################
   configflxTransactionTypes

   If Not frmMMain.IsRibbonVersion Then GoTo Closing_FormLoad
   loadfirstclient adoConn 'this function loads control accounts
   LoadFlxTransactionTypes adoConn 'this function loads control accounts
'   populateGridDefinedHeader adoConn, szSQL_CC, flxTransactionTypes, 0

   sSQLQuery_ = "SELECT Code, Name, ClientID " & _
                "FROM   NominalLedger " & _
                "WHERE  Posting AND ClientID <> 'NONE' " & _
                "ORDER BY Code;"

'   populateCombo adoConn, sSQLQuery_, cboCC, "40 pt;100 pt;0 pt"
'   cboCC.ColumnWidths = "40pt;"
   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockOptimistic

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
   Else
      Dim TotalRow As Long, TotalCol As Long

      TotalRow = adoRST.RecordCount
      TotalCol = adoRST.Fields.Count
      ReDim NC(TotalCol - 1, TotalRow) As String

      Dim i As Integer, j As Integer

      For i = 0 To adoRST.RecordCount - 1
         For j = 0 To adoRST.Fields.Count - 1
            NC(j, i) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
         Next j
         adoRST.MoveNext
      Next i

      adoRST.Close
      Set adoRST = Nothing
   End If
'
'   cboPosting.AddItem "YES"
'   cboPosting.AddItem "NO"

'  #############################
'  Nominal Ledger Sub Types
'  #############################
   ConfigureFlxSubTypes
   LoadFlxSubTypes adoConn

Closing_FormLoad:
   adoConn.Close
   Set adoConn = Nothing
'   cboCC.Width = flxTransactionTypes.ColWidth(3) + flxTransactionTypes.ColWidth(4)
'   picCC.Width = cboCC.Width
'   picCC.Height = cboCC.Height
'   picCC.BackColor = &H8000000F
'   cboCC.Top = 0
'   cboCC.Left = 0
'   cboPosting.Width = flxTransactionTypes.ColWidth(7)
'   picPosting.Width = cboPosting.Width
'   picPosting.Height = cboPosting.Height
'   picPosting.BackColor = &H8000000F
'   cboPosting.Top = 0
'   cboPosting.Left = 0
   'If cmbCAClient.ListCount > 0 Then cmbCAClient.ListIndex = 0
   Call WheelHook(Me.hWnd)
End Sub
Private Sub loadfirstclient(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes
   txtClientList.text = adoRST("CLIENTID").Value
   txtClientList.Tag = adoRST("CLIENTNAME").Value
   adoRST.Close

NoRes:
   Set adoRST = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   Set adoRST = Nothing
End Sub
Private Sub LoadFlxTransactionTypes(adoConn As ADODB.Connection) 'This function loads control accounts
'szHeader$ = "|<CAName|<Type|<NCode|<NName|<ClientID|<Fixed|<Posting|<TypeID"
'                                              0         0              0
   Dim adoRST     As New ADODB.Recordset
   Dim rRow       As Integer
   Dim K          As Integer
   Dim szSQL      As String

        If txtClientList.text = "" Then 'this is control account client
                    Exit Sub
        End If
        Call configflxTransactionTypes
        rRow = 1
'        szSQL = "SELECT T.CAName, S.Value, T.Code AS NCode, " & _
'                 "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
'                 "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
'            "FROM NominalLedger AS T, SecondaryCode AS S " & _
'            "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
'                  "T.ClientID = '" & txtClientList.text & "' " & _
'            "ORDER BY T.CADisOrder;"
'
'        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        
'        If adoRst.RecordCount > 0 Then
'                While Not adoRst.EOF
'                   flxTransactionTypes.TextMatrix(rRow, 1) = IIf(IsNull(adoRst.Fields.Item("CAName").Value) = True, "", adoRst.Fields.Item("CAName").Value)
'                   flxTransactionTypes.TextMatrix(rRow, 2) = adoRst.Fields.Item("Value").Value
'                   flxTransactionTypes.TextMatrix(rRow, 3) = adoRst.Fields.Item("NCode").Value
'                   flxTransactionTypes.TextMatrix(rRow, 4) = adoRst.Fields.Item("NName").Value
'                   flxTransactionTypes.TextMatrix(rRow, 5) = adoRst.Fields.Item("ClientID").Value
'                   flxTransactionTypes.TextMatrix(rRow, 6) = adoRst.Fields.Item("Fixed").Value
'                   flxTransactionTypes.TextMatrix(rRow, 7) = adoRst.Fields.Item("P").Value 'CAPosting
'                   flxTransactionTypes.TextMatrix(rRow, 8) = adoRst.Fields.Item("Type").Value
'                   flxTransactionTypes.TextMatrix(rRow, 9) = "S"
'                   flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
'                   flxTransactionTypes.TextMatrix(rRow, 11) = flxTransactionTypes.TextMatrix(rRow, 7)
'                   adoRst.MoveNext
'                   rRow = rRow + 1
'                   flxTransactionTypes.AddItem ""
'                Wend
'        Else
'now we need to update CAFIXED/caposting equivallent to allow posting
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Sales Ledger Control"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Sales"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T, SecondaryCode AS S " & _
                                "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
                                      "T.ClientID = '" & txtClientList.text & "' AND CAName='Sales Ledger Control' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                        flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                        flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
                                        'flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
                                         'flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
        
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"   'Fixed
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "S" 'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        rRow = rRow + 1
                        flxTransactionTypes.AddItem ""
            
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Purchase Ledger Control"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Purchase"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T, SecondaryCode AS S " & _
                                "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Purchase Ledger Control'" & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                            flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                            flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                             flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "P" 'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        rRow = rRow + 1
                        flxTransactionTypes.AddItem ""
            
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Input VAT"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Input VAT"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T, SecondaryCode AS S " & _
                                "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Input VAT' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                          flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES" 'we are not using 6 col
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "I" 'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        rRow = rRow + 1
                        flxTransactionTypes.AddItem ""
            
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Output VAT"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Output VAT"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T, SecondaryCode AS S " & _
                                "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
                                      "T.ClientID = '" & txtClientList.text & "' AND CAName='Output VAT' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                        flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                        flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                          flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "O"
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        rRow = rRow + 1
                        flxTransactionTypes.AddItem ""
            
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Retained Earnings"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Retained Earnings"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T, SecondaryCode AS S " & _
                                "WHERE T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND " & _
                                      "T.ClientID = '" & txtClientList.text & "' AND CAName='Retained Earnings' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                          flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "R"  'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Rent & Other Amounts Payable (P&L)"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Rent & Other Amounts Payable (P&L)"
                        szSQL = "SELECT T.CAName,T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T " & _
                                "WHERE " & _
                                      "T.ClientID = '" & txtClientList.text & "' AND CAName='Rent & Other Amounts Payable (P&L)' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                           flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "RO"
                        flxTransactionTypes.TextMatrix(rRow, 9) = "2"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Client/Landlord Control Account (B/S)"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Client/Landlord Control Account (B/S)"
                        szSQL = "SELECT T.CAName ,T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T " & _
                                "WHERE " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Client/Landlord Control Account (B/S)'  " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                          flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "RP"  'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Accruals Control Account (B/S)"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Accruals Control Account (B/S)"
                        szSQL = "SELECT T.CAName, T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T " & _
                                "WHERE  " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Accruals Control Account (B/S)'  " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
'                                           flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRst.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
'                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "AC" 'This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Management Fee Payable (P&L)"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Management Fee Payable (P&L)"
                        szSQL = "SELECT T.CAName,T.CAFixed, T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T " & _
                                "WHERE " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Management Fee Payable (P&L)'" & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                         flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                         flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
                                           flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRST.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "MP" ''This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "2"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        flxTransactionTypes.TextMatrix(rRow, 1) = "Managing Agents control Account (B/S)"
                        flxTransactionTypes.TextMatrix(rRow, 2) = "Managing Agents control Account (B/S)"
                        szSQL = "SELECT T.CAName, T.CAFixed,T.Code AS NCode, " & _
                                     "T.Name AS NName, T.ClientID, T.CAFixed AS Fixed, " & _
                                     "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder " & _
                                "FROM NominalLedger AS T " & _
                                "WHERE  " & _
                                      "T.ClientID = '" & txtClientList.text & "'  AND CAName='Managing Agents control Account (B/S)' " & _
                                "ORDER BY T.CADisOrder;"
                            
                                  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                  If Not adoRST.EOF Then
                                            flxTransactionTypes.TextMatrix(rRow, 3) = adoRST.Fields.Item("NCode").Value
                                            flxTransactionTypes.TextMatrix(rRow, 4) = adoRST.Fields.Item("NName").Value
                                             flxTransactionTypes.TextMatrix(rRow, 7) = IIf(adoRST.Fields.Item("CAFixed").Value = False, "YES", "NO")
                                  Else
                                         flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                                  End If
                                  adoRST.Close
                                  Set adoRST = Nothing
                        flxTransactionTypes.TextMatrix(rRow, 5) = txtClientList.text
                        flxTransactionTypes.TextMatrix(rRow, 6) = "YES"
                        flxTransactionTypes.TextMatrix(rRow, 7) = "NO"
                        flxTransactionTypes.TextMatrix(rRow, 8) = "MF" ''This is very short name for the control Names
                        flxTransactionTypes.TextMatrix(rRow, 9) = "1"
                        flxTransactionTypes.TextMatrix(rRow, 10) = flxTransactionTypes.TextMatrix(rRow, 3)
                        flxTransactionTypes.AddItem ""
                        rRow = rRow + 1
                        
                        
'         End If
'         adoRst.Close
   
   
   Set adoRST = Nothing
End Sub

Private Sub ConfigureFlxSubTypes()
   With flxSubTypes
      .Cols = 4
      .Rows = 2
      .ColWidth(0) = 0
      .RowHeight(0) = 0
      .ColWidth(1) = Label1(7).Left - Label1(6).Left
      .ColWidth(2) = .Width + .Left - Label1(7).Left - 300
      .ColWidth(3) = 0                                         'Modified Flag
   End With
End Sub

Private Function LoadFlxSubTypes(adoConn As ADODB.Connection)
   Dim rRow As Integer, szSQL As String
   Dim rstRec As New ADODB.Recordset

   szSQL = "SELECT * " & _
           "FROM NLSubTypes;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSubTypes.Clear
      flxSubTypes.Rows = 2

      rRow = 1
      While Not rstRec.EOF
         flxSubTypes.RowHeight(rRow) = 285
         flxSubTypes.TextMatrix(rRow, 0) = rstRec!STCode
         flxSubTypes.TextMatrix(rRow, 1) = rstRec!STName
         flxSubTypes.TextMatrix(rRow, 2) = rstRec!STDescription
         flxSubTypes.TextMatrix(rRow, 3) = ""
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSubTypes.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close

   Set rstRec = Nothing
End Function

Private Sub configflxTransactionTypes()
   Dim szHeader As String, iCol As Integer

   With flxTransactionTypes
      .Clear
      .Cols = 12
      .Rows = 2
      .RowHeight(0) = 0
      szHeader$ = "|<CAName|<Type|<NCode|<NName|ClientID|Fixed|<Posting|TypeID|saved|OriNC|OriPosting"
'                   CAName S.Value, T.NCode, T.NName, T.ClientID, T.Fixed, IIF(T.Posting, 'YES', 'NO'), T.Type
      .FormatString = szHeader$

      .ColWidth(0) = 0                                     '
      .ColWidth(1) = Label1(2).Left - Label1(1).Left       'Control Acc Name
      .ColWidth(2) = Label1(3).Left - Label1(2).Left       'Type
      .ColWidth(3) = Label1(4).Left - Label1(3).Left       'Nominal code
      .ColWidth(4) = Label1(5).Left - Label1(4).Left
      .ColWidth(5) = 0                                     'Client
      .ColWidth(6) = 0                                     'fixed?
      .ColWidth(7) = .Width + .Left - Label1(5).Left - 300 'Allow Posting
      .ColWidth(8) = 0                                     'Type ID
      .ColWidth(9) = 0                                     'Saved? S->Saved, C->New set of CA for a client, N->New CA, A->Amend, D->Decide
      .ColWidth(10) = 0                                    'Original NC
      .ColWidth(11) = 0                                    'Original AllowPosting
   End With
End Sub
Private Sub ConfigflxNominal()
   Dim szHeader As String, iCol As Integer

   With flxNominal
      .Clear
      .Cols = 3
      .Rows = 2
      .RowHeight(0) = 0
      szHeader$ = "|<Nominal Code|<Nominal Name"
'                   CAName S.Value, T.NCode, T.NName, T.ClientID, T.Fixed, IIF(T.Posting, 'YES', 'NO'), T.Type
      .FormatString = szHeader$

      .ColWidth(0) = 80                                     '
      .ColWidth(1) = 900    'Control Acc Name
      .ColWidth(2) = 3000       'Type
   End With
End Sub
Private Sub cmdStatementPathSave_Click()
   If MsgBox("Do you wish to save?", vbQuestion + vbYesNo, "Statement File") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT PropertyID, StPath " & _
           "FROM GlobalData;"

   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxStFlPth.Rows - 1
      adoRST.Find ("PropertyID = '" & flxStFlPth.TextMatrix(iRow, 0) & "'"), , , 1

      adoRST.Fields.Item("StPath").Value = flxStFlPth.TextMatrix(iRow, 2)
      adoRST.Update
   Next iRow

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   ShowMsgInTaskBar "Sucessfully updated"
   Unload Me
End Sub

Private Sub StatementPath(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer
   
   flxStFlPth.Cols = 3
   flxStFlPth.Rows = 2
   flxStFlPth.RowHeight(0) = 0

   flxStFlPth.ColWidth(0) = 0
   flxStFlPth.ColWidth(1) = Label6(5).Left - Label6(4).Left - flxStFlPth.Left
   flxStFlPth.ColWidth(2) = flxStFlPth.Width - Label6(5).Left - 200

   szSQL = "SELECT Property.PropertyName, Property.PropertyID, StPath " & _
           "FROM GlobalData, Property " & _
           "WHERE GlobalData.PropertyID = Property.PropertyID;"
   
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   iRow = 1
   While Not adoRST.EOF
      flxStFlPth.TextMatrix(iRow, 0) = adoRST.Fields.Item("PropertyID").Value
      flxStFlPth.TextMatrix(iRow, 1) = adoRST.Fields.Item("PropertyName").Value
      flxStFlPth.TextMatrix(iRow, 2) = IIf(IsNull(adoRST.Fields.Item("StPath").Value), "", adoRST.Fields.Item("StPath").Value)
      adoRST.MoveNext
      If Not adoRST.EOF Then flxStFlPth.AddItem ""
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing

   flxStFlPth.row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim X

   If VAT_MODIFIED_ Then
      X = MsgBox("You made some changes in the VAT, do you wish to save?", vbQuestion + vbYesNoCancel, "Data Saving")
      If X = vbCancel Then Cancel = 1
      If X = vbYes Then cmdSaveVAT_Click
   End If

   'If Cancel = 0 Then frmMMain.fraCmdButton.Enabled = True
   
'   Call WheelUnHook(Me.hwnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub tabSettings_Click(PreviousTab As Integer)
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim szChoice As String, szaChoice() As String
   Dim rsShopingCentre As New ADODB.Recordset
   adoConn.Open getConnectionString

   Select Case tabSettings.Tab
   Case 1:
        If Not bConfigureFlxDeposit Then
           ConfigureFlxDeposit
           LoadClientData adoConn
        End If
   Case 5: 'This is control account Balance
        ControlHanlding DefaultMode
        cmdClientList.SetFocus
   Case 7: 'this is fund configuration
        rsShopingCentre.Open "Select isFundAssign from shoppingCentre", adoConn, adOpenStatic, adLockReadOnly
        If Not rsShopingCentre.EOF Then
                chkFundAssignment.Value = IIf(rsShopingCentre("isFundAssign").Value = True, 1, 0)
        End If
        rsShopingCentre.Close
        Set rsShopingCentre = Nothing
   End Select
   
   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If

   If UBound(szaChoice) > 0 Then
      If szaChoice(0) <> "" Then
         If InStr(szaChoice(0), "U") > 0 Then CheckUnit.Value = Checked           'U - Unit
      Else
         CheckUnit.Value = Unchecked
      End If
      If szaChoice(1) <> "" Then
         If InStr(szaChoice(1), "CL") > 0 Then CheckClient.Value = Checked         'CL - Client / Landlord
      Else
         CheckClient.Value = Unchecked
      End If
      If szaChoice(2) <> "" Then
         If InStr(szaChoice(2), "L") > 0 Then CheckLessee.Value = Checked         'L - Lessee
      Else
         CheckLessee.Value = Unchecked
      End If
      If szaChoice(3) <> "" Then
         If InStr(szaChoice(3), "S") > 0 Then CheckSupplier.Value = Checked       'S - Supplier
      Else
         CheckSupplier.Value = Unchecked
      End If
      If szaChoice(4) <> "" Then
         If InStr(szaChoice(4), "P") > 0 Then CheckProperty.Value = Checked       'P - Property
      Else
         CheckProperty.Value = Unchecked
      End If
      If szaChoice(5) <> "" Then
         If InStr(szaChoice(5), "MA") > 0 Then CheckManaging.Value = Checked      'MA - Managing Agent
      Else
         CheckManaging.Value = Unchecked
      End If
   End If
   
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadClientData(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iCol As Integer

   szSQL = "SELECT Client.ClientID, Client.ClientName, Client.spare1, " & _
               "Client.spare2, NLB.Name as BCName, NLN.Name as NCName " & _
           "FROM Client, " & _
               "(SELECT Name, Code FROM NominalLedger, Client WHERE Code = Client.spare1 and Client.spare1 <> '') AS NLB, " & _
               "(SELECT Name, Code FROM NominalLedger, Client WHERE Code = Client.spare2 and Client.spare2 <> '') AS NLN " & _
           "WHERE Client.spare1 <> '' AND Client.spare2 <> '' AND " & _
               "Client.spare1 = NLB.Code AND Client.spare2 = NLN.Code;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With adoRST
      iCol = 1
      flxDeposit.Clear
      flxDeposit.Rows = 2
      While Not .EOF
         flxDeposit.TextMatrix(iCol, 0) = iCol
         flxDeposit.TextMatrix(iCol, 1) = .Fields.Item("ClientID").Value
         flxDeposit.TextMatrix(iCol, 2) = .Fields.Item("ClientName").Value
         flxDeposit.TextMatrix(iCol, 3) = .Fields.Item("spare1").Value
         flxDeposit.TextMatrix(iCol, 4) = .Fields.Item("BCName").Value
         flxDeposit.TextMatrix(iCol, 5) = .Fields.Item("spare2").Value
         flxDeposit.TextMatrix(iCol, 6) = .Fields.Item("NCName").Value
         .MoveNext
         If Not .EOF Then flxDeposit.AddItem ""
         iCol = iCol + 1
      Wend
      .Close
   End With
   Set adoRST = Nothing
   flxDeposit.row = 0
End Sub

Private Sub LoadComboes(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim Data() As String, i As Integer

   szSQL = "SELECT NL.Name, NL.Code " & _
           "FROM NominalLedger as NL;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim Data(1, adoRST.RecordCount - 1) As String
   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST.Fields.Item("Code").Value
      Data(1, i) = adoRST.Fields.Item("Name").Value
      i = i + 1
      adoRST.MoveNext
   Wend
   cmbBank.Clear
   cmbBank.Column() = Data()
   cmbDNC.Clear
   cmbDNC.Column() = Data()

   adoRST.Close

   If Not frmMMain.IsRibbonVersion Then GoTo Closing_LoadComboes

   szSQL = "SELECT C.ClientID, C.ClientName " & _
           "FROM Client AS C, NominalLedger AS N " & _
           "WHERE C.ClientID = N.ClientID " & _
           "GROUP BY C.ClientID, C.ClientName;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.RecordCount > 0 Then
      ReDim Data(1, adoRST.RecordCount - 1) As String
      i = 0
      While Not adoRST.EOF
         Data(0, i) = adoRST.Fields.Item("ClientID").Value
         Data(1, i) = adoRST.Fields.Item("ClientName").Value
         i = i + 1
         adoRST.MoveNext
      Wend
      cmbClient.Clear
      cmbClient.Column() = Data()
'      cmbCAClient.Clear
'      cmbCAClient.Column() = Data()
   End If
   adoRST.Close
   
   szSQL = ""

Closing_LoadComboes:
   Set adoRST = Nothing
End Sub

Private Sub ConfigureFlxDeposit()
   flxDeposit.Clear
   flxDeposit.Cols = 7
   flxDeposit.Rows = 2
   flxDeposit.RowHeight(0) = 0

   flxDeposit.ColWidth(0) = Label6(1).Left - Label6(0).Left    ' id
   flxDeposit.ColWidth(1) = 0                                  ' Client Id
   flxDeposit.ColWidth(2) = Label6(2).Left - Label6(1).Left
   flxDeposit.ColWidth(3) = 0                                  ' Bank ID
   flxDeposit.ColWidth(4) = Label6(3).Left - Label6(2).Left
   flxDeposit.ColWidth(5) = 0                                  ' Nominal Code
   flxDeposit.ColWidth(6) = flxDeposit.Width + flxDeposit.Left - Label6(3).Left - 280

   bConfigureFlxDeposit = True
End Sub

Private Sub tabSettings_LostFocus()
   Select Case tabSettings.Tab
      Case 1:
         FocusControl cmbClient
      Case 2:
         If chkAlarm.Enabled Then
            FocusControl chkAlarm
         Else
            FocusControl cmdEditOptions
         End If
      Case 3:
         If txtFromEmail.Enabled Then
            FocusControl txtFromEmail
         Else
            FocusControl cmdEditSettings
         End If
      Case 4:
         FocusControl cmdUpdate
      Case 5:
'         cmbCAClient.SetFocus
      Case 6:
         FocusControl cmdSubTypeSave
   End Select
End Sub

Private Sub tabSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tabSettings.MousePointer = vbDefault
   Me.MousePointer = vbArrow
End Sub

'  System loads all Bank Records into the grid of the form
Public Function LoadflxVat(adoConn As ADODB.Connection)
   Dim rRow As Integer, szSQL As String
   Dim rstRec As New ADODB.Recordset
   flxVat.RowHeight(0) = 0
   flxVat.Cols = 6
   flxVat.ColWidth(0) = Label11(1).Left - Label11(0).Left
   flxVat.ColAlignment(0) = vbLeftJustify
   flxVat.ColWidth(1) = Label11(2).Left - Label11(1).Left
   flxVat.ColAlignment(1) = vbLeftJustify
   flxVat.ColWidth(2) = Label11(3).Left - Label11(2).Left
   flxVat.ColAlignment(2) = vbRightJustify
   flxVat.ColWidth(3) = Label11(4).Left - Label11(3).Left + 200
   flxVat.ColWidth(4) = 1100 ''Label11(4).Left - Label11(3).Left 'flxVat.Left + flxVat.Width - Label11(4).Left - 340
   flxVat.ColWidth(5) = 1100 'Label11(4).Left - Label11(3).Left
   szSQL = "SELECT * " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxVat.Clear
      flxVat.Rows = 2

      rRow = 1
      While Not rstRec.EOF
         flxVat.TextMatrix(rRow, 0) = rstRec!VAT_ID
         flxVat.TextMatrix(rRow, 1) = rstRec!VAT_CODE
         flxVat.TextMatrix(rRow, 2) = rstRec!VAT_RATE
         flxVat.TextMatrix(rRow, 3) = rstRec!DESCRIPTIONS
         flxVat.TextMatrix(rRow, 4) = IIf(rstRec!IN_USE, "YES", "NO")
         flxVat.TextMatrix(rRow, 5) = IIf(rstRec!OnReturn = 0, "NO", "YES")
         rstRec.MoveNext
         If Not rstRec.EOF Then flxVat.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close

   Set rstRec = Nothing
End Function

Private Sub txtFromEmail_GotFocus()
   SelTxtInCtrl txtFromEmail
End Sub

Private Sub txtFromEmail_LostFocus()
   Dim szErrMsg As String

   If Trim(txtFromEmail.text) <> "" Then
      If Not ValidateEmail(txtFromEmail.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "From Email Address"
         SelTxtInCtrl txtFromEmail
         txtFromEmail.SetFocus
      End If
   End If
End Sub

Private Sub txtPort_GotFocus()
   SelTxtInCtrl txtPort
End Sub

Private Sub txtPws_GotFocus()
   SelTxtInCtrl txtPws
End Sub

Private Sub txtSearchNC_Change()
      'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchNC.text) > 0 Then
        txtSearchNName.text = ""
   End If

   For i = flxNominal.Rows - 1 To 1 Step -1
        flxNominal.RowHeight(i) = 240
        If InStr(1, UCase(flxNominal.TextMatrix(i, 1)), UCase(txtSearchNC.text), vbTextCompare) = 0 Then
              flxNominal.RowHeight(i) = 0
        End If
        If flxNominal.RowHeight(i) = 240 Then
              flxNominal.row = i
        End If
   Next i
End Sub

Private Sub txtSearchNC_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxNominal.SetFocus
    End If
    If KeyCode = 13 Then
           flxNominal.SetFocus
    End If
End Sub

Private Sub txtSearchNC_KeyPress(KeyAscii As MSForms.ReturnInteger)
      If KeyAscii = 27 Then
            picNominal.Visible = False
    End If
End Sub

Private Sub txtSearchNName_Change()
     'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchNName.text) > 0 Then
        txtSearchNC.text = ""
   End If

   For i = flxNominal.Rows - 1 To 1 Step -1
        flxNominal.RowHeight(i) = 240
        If InStr(1, UCase(flxNominal.TextMatrix(i, 2)), UCase(txtSearchNName.text), vbTextCompare) = 0 Then
            flxNominal.RowHeight(i) = 0
        End If
        If flxNominal.RowHeight(i) = 240 Then
            flxNominal.row = i
        End If
   Next i
End Sub

Private Sub txtSearchNName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxNominal.SetFocus
    End If
    If KeyCode = 13 Then
           flxNominal.SetFocus
    End If
End Sub

Private Sub txtSMTP_GotFocus()
   SelTxtInCtrl txtSMTP
End Sub

Private Sub txtSubType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If flxSubTypes.TextMatrix(flxSubTypes.row, 2) <> txtSubType.text Then
         flxSubTypes.TextMatrix(flxSubTypes.row, 2) = txtSubType.text
         flxSubTypes.TextMatrix(flxSubTypes.row, 3) = "M"
      End If
      flxSubTypes.Enabled = True
      picSubType.Visible = False
      txtSubType.text = ""
      bEditSubType = True
   End If
   If KeyAscii = 27 Then
      flxSubTypes.Enabled = True
      picSubType.Visible = False
      txtSubType.text = ""
   End If
End Sub

Private Sub txtSubType_LostFocus()
   txtSubType_KeyPress 13
End Sub

Private Sub cmdSubTypeCancel_Click()
   txtSubType_KeyPress 27
End Sub

Private Sub txtUName_GotFocus()
   SelTxtInCtrl txtUName
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
         ' PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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



Private Sub txtVatDesp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtVatDesp.text) = 0 Then
                FocusControl chkInUse
        Else
                FocusControl cmdSaveVAT
        End If
    End If
End Sub



Private Sub txtVatRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtVatDesp
    End If
    
End Sub
