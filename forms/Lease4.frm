VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLease4 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lease Information"
   ClientHeight    =   13950
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   19950
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lease4.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13950
   ScaleWidth      =   19950
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   13185
      TabIndex        =   315
      Top             =   8505
      Visible         =   0   'False
      Width           =   6630
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
         Index           =   0
         Left            =   6330
         Style           =   1  'Graphical
         TabIndex        =   316
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   4335
         Left            =   90
         TabIndex        =   323
         Top             =   675
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   7646
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
      Begin VB.Shape Shape2 
         Height          =   5010
         Left            =   45
         Top             =   45
         Width           =   6540
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   322
         Top             =   180
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "2170;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   315
         TabIndex        =   320
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2302;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   321
         Top             =   390
         Width           =   3555
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6271;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   319
         Top             =   195
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
         Index           =   2
         Left            =   5175
         TabIndex        =   318
         Top             =   180
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   5265
         TabIndex        =   317
         Top             =   390
         Width           =   1170
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2064;450"
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
         Index           =   15
         Left            =   90
         Top             =   135
         Width           =   6345
      End
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   2
      Left            =   5670
      ScaleHeight     =   405
      ScaleWidth      =   3150
      TabIndex        =   203
      Top             =   1530
      Visible         =   0   'False
      Width           =   3150
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   204
         Top             =   105
         Width           =   2700
      End
   End
   Begin VB.PictureBox fraList 
      BackColor       =   &H80000004&
      Height          =   3870
      Index           =   0
      Left            =   13320
      ScaleHeight     =   3810
      ScaleWidth      =   5295
      TabIndex        =   182
      Top             =   7605
      Visible         =   0   'False
      Width           =   5355
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
         TabIndex        =   198
         Top             =   20
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3255
         Index           =   0
         Left            =   15
         TabIndex        =   195
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
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1365
         TabIndex        =   188
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
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   187
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
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   186
         Top             =   15
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
         Left            =   750
         TabIndex        =   185
         Top             =   15
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   184
         Top             =   0
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   183
         Top             =   1680
         Width           =   1095
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
   Begin VB.Frame Frame1 
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
      Height          =   990
      Index           =   8
      Left            =   75
      TabIndex        =   119
      Top             =   7005
      Width           =   18300
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy &Lease"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2925
         TabIndex        =   146
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4230
         TabIndex        =   128
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8190
         TabIndex        =   127
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1605
         TabIndex        =   126
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "&Terminate"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10815
         TabIndex        =   125
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel New"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5550
         TabIndex        =   124
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9495
         TabIndex        =   123
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Lease"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6885
         TabIndex        =   122
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13440
         TabIndex        =   121
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12120
         TabIndex        =   120
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Lease"
      Height          =   1725
      Index           =   16
      Left            =   80
      TabIndex        =   108
      Top             =   0
      Width           =   18360
      Begin VB.CommandButton cmdUsage 
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
         Index           =   0
         Left            =   14805
         TabIndex        =   311
         Top             =   1170
         Width           =   300
      End
      Begin VB.CommandButton cmdUnitNumber 
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
         Index           =   0
         Left            =   5130
         TabIndex        =   2
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton cmdTenants 
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
         Left            =   5490
         TabIndex        =   1
         Top             =   270
         Width           =   300
      End
      Begin VB.CommandButton cmdLease 
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
         Left            =   5130
         TabIndex        =   0
         Top             =   270
         Width           =   300
      End
      Begin VB.TextBox txtSageAccountNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   310
         Top             =   300
         Width           =   1410
      End
      Begin VB.CheckBox chkMultipleLH 
         Caption         =   "Show Multiple Leaseholders"
         Height          =   315
         Left            =   7185
         TabIndex        =   179
         Top             =   300
         Width           =   2475
      End
      Begin VB.CheckBox chkExpLease 
         Caption         =   "Show Expired Leases only"
         Height          =   315
         Left            =   7185
         TabIndex        =   111
         Top             =   720
         Width           =   2310
      End
      Begin VB.TextBox txtUnitName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1140
         Width           =   3795
      End
      Begin VB.TextBox txtTenant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox txtUnitNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   202
         Top             =   720
         Width           =   3795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lease Status :"
         Height          =   195
         Index           =   13
         Left            =   7155
         TabIndex        =   206
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   12
         Left            =   8190
         TabIndex        =   205
         Top             =   1155
         Width           =   660
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usage:"
         Height          =   195
         Index           =   4
         Left            =   10785
         TabIndex        =   148
         Top             =   1185
         Width           =   465
      End
      Begin MSForms.ComboBox cboUsage 
         Height          =   315
         Left            =   11715
         TabIndex        =   147
         Top             =   1170
         Width           =   3075
         VariousPropertyBits=   746604569
         MaxLength       =   50
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5424;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Number:"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   118
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaseholder:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   117
         Top             =   300
         Width           =   900
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   11715
         TabIndex        =   116
         Top             =   750
         Width           =   3075
         VariousPropertyBits=   746604575
         BackColor       =   15134203
         BorderStyle     =   1
         Size            =   "5424;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClient 
         Height          =   285
         Left            =   11715
         TabIndex        =   115
         Top             =   345
         Width           =   3075
         VariousPropertyBits=   746604575
         BackColor       =   15134203
         BorderStyle     =   1
         Size            =   "5424;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   3
         Left            =   10785
         TabIndex        =   114
         Top             =   750
         Width           =   645
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   2
         Left            =   10770
         TabIndex        =   113
         Top             =   345
         Width           =   465
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name:"
         Height          =   195
         Left            =   255
         TabIndex        =   112
         Top             =   1140
         Width           =   765
      End
   End
   Begin VB.PictureBox fraList 
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
      Height          =   4695
      Index           =   1
      Left            =   3300
      ScaleHeight     =   4665
      ScaleWidth      =   8880
      TabIndex        =   102
      Top             =   8790
      Visible         =   0   'False
      Width           =   8910
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
         Left            =   8190
         TabIndex        =   190
         Top             =   360
         Width           =   300
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
         Left            =   8190
         TabIndex        =   189
         Top             =   30
         Width           =   300
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7455
         TabIndex        =   196
         Top             =   1035
         Width           =   1200
      End
      Begin VB.TextBox txtSearchTenant 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   45
         TabIndex        =   191
         Top             =   1035
         Width           =   1275
      End
      Begin VB.TextBox txtSearchName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   192
         Top             =   1035
         Width           =   2565
      End
      Begin VB.TextBox txtSearchUnitName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5685
         TabIndex        =   194
         Top             =   1035
         Width           =   1740
      End
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
         Index           =   1
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
         Height          =   3255
         Left            =   45
         TabIndex        =   200
         Top             =   1380
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5741
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
      Begin MSForms.TextBox txtSearchCompany 
         Height          =   285
         Left            =   3960
         TabIndex        =   193
         Top             =   1035
         Width           =   1710
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3016;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Number"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   229
         Top             =   765
         Width           =   900
      End
      Begin MSForms.TextBox txtPropertyList 
         Height          =   285
         Left            =   720
         TabIndex        =   199
         Tag             =   "ALL"
         Top             =   360
         Width           =   7470
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "13176;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   720
         TabIndex        =   197
         Tag             =   "ALL"
         Top             =   30
         Width           =   7470
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "13176;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   7515
         TabIndex        =   181
         Top             =   765
         Width           =   840
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Index           =   2
         Left            =   5730
         TabIndex        =   105
         Top             =   765
         Width           =   855
         ForeColor       =   16711680
         VariousPropertyBits=   8388627
         Caption         =   "Unit Name"
         Size            =   "1508;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   1305
         TabIndex        =   104
         Top             =   765
         Width           =   1320
         ForeColor       =   16711680
         VariousPropertyBits=   8388627
         Caption         =   "Name"
         Size            =   "2328;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   103
         Top             =   765
         Width           =   735
         ForeColor       =   16711680
         VariousPropertyBits=   8388627
         Caption         =   "Lessee ID"
         Size            =   "1296;450"
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
         Height          =   240
         Index           =   0
         Left            =   90
         Top             =   765
         Width           =   8550
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   107
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   106
         Top             =   70
         Width           =   465
      End
   End
   Begin TabDlg.SSTab tabLease 
      Height          =   5160
      Left            =   75
      TabIndex        =   3
      Top             =   1830
      Width           =   18345
      _ExtentX        =   32359
      _ExtentY        =   9102
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   8
      TabsPerRow      =   11
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lease &Details"
      TabPicture(0)   =   "Lease4.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(10)"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Rent Charges"
      TabPicture(1)   =   "Lease4.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Label36"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Rent Re&view"
      TabPicture(2)   =   "Lease4.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Brea&ks"
      TabPicture(3)   =   "Lease4.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(0)"
      Tab(3).Control(1)=   "Frame1(9)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Service &Charges"
      TabPicture(4)   =   "Lease4.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(2)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Interest Charges"
      TabPicture(5)   =   "Lease4.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label49"
      Tab(5).Control(1)=   "Frame3"
      Tab(5).Control(2)=   "Frame1(11)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "&Breaches"
      TabPicture(6)   =   "Lease4.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(12)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Assignment"
      TabPicture(7)   =   "Lease4.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1(13)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "I&nsurance"
      TabPicture(8)   =   "Lease4.frx":09AA
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "Frame1(14)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "&Licences"
      TabPicture(9)   =   "Lease4.frx":09C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame1(7)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "&Memo"
      TabPicture(10)  =   "Lease4.frx":09E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1(5)"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4740
         Index           =   7
         Left            =   -74910
         TabIndex        =   292
         Top             =   360
         Width           =   18060
         Begin VB.CommandButton Command3 
            Caption         =   "&Save Licence"
            Height          =   375
            Left            =   12555
            TabIndex        =   309
            Top             =   2205
            Width           =   1665
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Cancel Licence"
            Height          =   375
            Left            =   12555
            TabIndex        =   308
            Top             =   3000
            Width           =   1665
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2040
            ScrollBars      =   2  'Vertical
            TabIndex        =   301
            Top             =   555
            Width           =   1515
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Edit Licence"
            Height          =   375
            Left            =   12510
            TabIndex        =   300
            Top             =   1620
            Width           =   1665
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&New Licence"
            Height          =   375
            Left            =   12510
            TabIndex        =   299
            Top             =   825
            Width           =   1665
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   405
            ScrollBars      =   2  'Vertical
            TabIndex        =   298
            Top             =   555
            Width           =   1635
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9945
            ScrollBars      =   2  'Vertical
            TabIndex        =   297
            Top             =   555
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5715
            ScrollBars      =   2  'Vertical
            TabIndex        =   296
            Top             =   555
            Width           =   4215
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   295
            Top             =   555
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11490
            Style           =   1  'Graphical
            TabIndex        =   294
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3840
            ScrollBars      =   2  'Vertical
            TabIndex        =   293
            Top             =   555
            Width           =   1875
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLicences 
            Height          =   2835
            Left            =   405
            TabIndex        =   302
            Top             =   855
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   5001
            _Version        =   393216
            Cols            =   5
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
         Begin VB.Label Label17 
            Caption         =   "Date"
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   307
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Type"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   306
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Licensor"
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   305
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Status"
            Height          =   255
            Index           =   4
            Left            =   9945
            TabIndex        =   304
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Description"
            Height          =   255
            Index           =   3
            Left            =   5715
            TabIndex        =   303
            Top             =   315
            Width           =   2115
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4740
         Index           =   13
         Left            =   -74910
         TabIndex        =   280
         Top             =   360
         Width           =   18150
         Begin VB.CommandButton cmdAssignmentNew 
            Caption         =   "&New Assignment"
            Height          =   375
            Left            =   13005
            TabIndex        =   291
            Top             =   900
            Width           =   1665
         End
         Begin VB.CommandButton cmdAssignmentCancel 
            Caption         =   "&Cancel Assignment"
            Height          =   375
            Left            =   13005
            TabIndex        =   290
            Top             =   3300
            Width           =   1665
         End
         Begin VB.CommandButton cmdAssignmentSave 
            Caption         =   "&Save Assignment"
            Height          =   375
            Left            =   13005
            TabIndex        =   289
            Top             =   2505
            Width           =   1665
         End
         Begin VB.CommandButton cmdAssignmentEdit 
            Caption         =   "&Edit Assignment"
            Height          =   375
            Left            =   13005
            TabIndex        =   288
            Top             =   1695
            Width           =   1665
         End
         Begin VB.TextBox txtAssignment_Date 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2205
            ScrollBars      =   2  'Vertical
            TabIndex        =   285
            Top             =   555
            Width           =   1635
         End
         Begin VB.TextBox txtAssignee 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3870
            ScrollBars      =   2  'Vertical
            TabIndex        =   284
            Top             =   555
            Width           =   2715
         End
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6630
            ScrollBars      =   2  'Vertical
            TabIndex        =   283
            Top             =   555
            Width           =   4155
         End
         Begin VB.CommandButton Command8 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12150
            Style           =   1  'Graphical
            TabIndex        =   282
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   10830
            ScrollBars      =   2  'Vertical
            TabIndex        =   281
            Top             =   555
            Width           =   1275
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAssignment 
            Height          =   2835
            Left            =   2205
            TabIndex        =   287
            Top             =   900
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   5001
            _Version        =   393216
            Cols            =   5
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
         Begin VB.Label Label17 
            Caption         =   "Description"
            Height          =   255
            Index           =   8
            Left            =   6750
            TabIndex        =   314
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Assignee"
            Height          =   255
            Index           =   7
            Left            =   3870
            TabIndex        =   313
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Label54"
            Height          =   255
            Index           =   6
            Left            =   2205
            TabIndex        =   312
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label Label17 
            Caption         =   "Status"
            Height          =   255
            Index           =   5
            Left            =   10830
            TabIndex        =   286
            Top             =   315
            Width           =   1515
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4650
         Index           =   11
         Left            =   -74910
         TabIndex        =   257
         Top             =   405
         Width           =   18195
         Begin VB.TextBox txtInterestDescription 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4710
            MaxLength       =   100
            TabIndex        =   278
            Top             =   4095
            Width           =   6555
         End
         Begin VB.OptionButton optManIntCal 
            Caption         =   "Manually calculated Interest"
            Height          =   255
            Left            =   3690
            TabIndex        =   277
            Top             =   2565
            Width           =   2295
         End
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   1080
            Index           =   4
            Left            =   3645
            TabIndex        =   270
            Top             =   2880
            Width           =   7695
            Begin VB.TextBox txtAmtCrgIntOn 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3060
               TabIndex        =   273
               Top             =   240
               Width           =   1080
            End
            Begin VB.TextBox txtInt2bChrg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3060
               Locked          =   -1  'True
               TabIndex        =   272
               Top             =   640
               Width           =   1080
            End
            Begin VB.TextBox txtNoIntDays 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6000
               TabIndex        =   271
               Top             =   240
               Width           =   960
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Number of Interest days                                       days."
               Height          =   195
               Index           =   1
               Left            =   4245
               TabIndex        =   276
               Top             =   240
               Width           =   3195
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Interest to be charged:                                         £   "
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   275
               Top             =   645
               Width           =   3000
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Specified Amount to charge Interest on:  £"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   274
               Top             =   240
               Width           =   2925
            End
         End
         Begin VB.TextBox txtIntPayableAfterDays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   269
            Top             =   2025
            Width           =   960
         End
         Begin VB.OptionButton optAutoIntCal 
            Caption         =   "System calculates Interest charged on total O/S balance.    Interest Payable after                                       days."
            Height          =   375
            Left            =   3690
            TabIndex        =   268
            Top             =   2025
            Width           =   7455
         End
         Begin VB.Frame Frame1 
            Caption         =   "Additional Interest Rate:"
            Height          =   660
            Index           =   3
            Left            =   3690
            TabIndex        =   264
            Top             =   1260
            Width           =   7695
            Begin VB.TextBox txtAdditionalIntRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   265
               Top             =   240
               Width           =   960
            End
            Begin MSForms.OptionButton optLSR 
               Height          =   375
               Left            =   2975
               TabIndex        =   267
               Top             =   180
               Width           =   4335
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "7646;661"
               Value           =   "0"
               Caption         =   "Lease Specific Additional Interest Rate                                       %"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.OptionButton optGIR 
               Height          =   375
               Left            =   765
               TabIndex        =   266
               Top             =   180
               Width           =   1815
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               DisplayStyle    =   5
               Size            =   "3201;661"
               Value           =   "1"
               Caption         =   "Global Interest Rate"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.ComboBox cboIntCrgable 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5250
            TabIndex        =   258
            Text            =   "No"
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Index           =   2
            Left            =   3690
            TabIndex        =   279
            Top             =   4095
            Width           =   870
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund:"
            Height          =   195
            Index           =   0
            Left            =   3690
            TabIndex        =   263
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type:"
            Height          =   195
            Index           =   6
            Left            =   7290
            TabIndex        =   262
            Top             =   720
            Width           =   990
         End
         Begin MSForms.ComboBox cboIntChargeDept 
            Height          =   315
            Left            =   4230
            TabIndex        =   261
            Top             =   720
            Width           =   2835
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5001;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;35277"
         End
         Begin MSForms.ComboBox cboIntDemandType 
            Height          =   315
            Left            =   8430
            TabIndex        =   260
            Top             =   720
            Width           =   2835
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5001;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705;70555"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Interest Chargeable:"
            Height          =   195
            Left            =   3690
            TabIndex        =   259
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4650
         Index           =   9
         Left            =   -74955
         TabIndex        =   250
         Top             =   405
         Width           =   18195
         Begin VB.ComboBox cboBreakClause 
            Height          =   315
            Left            =   6660
            TabIndex        =   255
            Text            =   "No"
            Top             =   585
            Width           =   780
         End
         Begin VB.TextBox txtBreakDate 
            Height          =   315
            Left            =   6615
            MaxLength       =   10
            TabIndex        =   252
            Top             =   1125
            Width           =   1960
         End
         Begin VB.ComboBox cboBreak 
            Height          =   315
            Left            =   6615
            TabIndex        =   251
            Top             =   1665
            Width           =   2000
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Break Clause:"
            Height          =   195
            Left            =   5580
            TabIndex        =   256
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Break Type:"
            Height          =   195
            Left            =   5580
            TabIndex        =   254
            Top             =   1710
            Width           =   795
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Break Date:"
            Height          =   195
            Left            =   5580
            TabIndex        =   253
            Top             =   1125
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Index           =   10
         Left            =   -74955
         TabIndex        =   230
         Top             =   360
         Width           =   18150
         Begin VB.TextBox txtLeaseStDt 
            Height          =   285
            Left            =   10590
            MaxLength       =   10
            TabIndex        =   239
            Top             =   1155
            Width           =   1995
         End
         Begin VB.TextBox txtYearEnd 
            Height          =   285
            Left            =   10590
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   238
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtLeaseEndDate 
            Height          =   285
            Left            =   10590
            MaxLength       =   10
            TabIndex        =   237
            Top             =   1845
            Width           =   1995
         End
         Begin VB.CheckBox chkSubLease 
            Caption         =   "Yes"
            Height          =   315
            Left            =   5790
            TabIndex        =   236
            Top             =   1155
            Width           =   735
         End
         Begin VB.CommandButton cmdLeaseType 
            Caption         =   ". . ."
            Height          =   300
            Left            =   7830
            TabIndex        =   235
            Top             =   1845
            Width           =   405
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "Lease4.frx":09FE
            Left            =   5790
            List            =   "Lease4.frx":0A00
            TabIndex        =   234
            Top             =   1845
            Width           =   1995
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5790
            TabIndex        =   233
            Top             =   2925
            Width           =   975
         End
         Begin VB.CommandButton cmdUnitNumber 
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
            Height          =   285
            Index           =   1
            Left            =   7710
            Style           =   1  'Graphical
            TabIndex        =   232
            Top             =   450
            Width           =   255
         End
         Begin VB.TextBox txtUnitName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   450
            Width           =   1905
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Start Date:"
            Height          =   195
            Index           =   6
            Left            =   8670
            TabIndex        =   249
            Top             =   1155
            Width           =   1185
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Year End:"
            Height          =   195
            Index           =   5
            Left            =   8670
            TabIndex        =   248
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease End Date:"
            Height          =   195
            Index           =   7
            Left            =   8670
            TabIndex        =   247
            Top             =   1845
            Width           =   1110
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Head Lease:"
            Height          =   195
            Index           =   9
            Left            =   4110
            TabIndex        =   246
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Lease:"
            Height          =   195
            Index           =   10
            Left            =   4110
            TabIndex        =   245
            Top             =   1155
            Width           =   735
         End
         Begin MSForms.CheckBox chkOLED 
            Height          =   345
            Left            =   8670
            TabIndex        =   244
            Top             =   2355
            Width           =   3945
            VariousPropertyBits=   1015031835
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6950;609"
            Value           =   "0"
            Caption         =   "Override lease end date"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkHoldingOver 
            Height          =   345
            Left            =   4095
            TabIndex        =   243
            Top             =   2355
            Width           =   1875
            VariousPropertyBits=   1015031835
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3307;609"
            Value           =   "0"
            Caption         =   "Holding Over"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Type"
            Height          =   195
            Left            =   4110
            TabIndex        =   242
            Top             =   1845
            Width           =   780
         End
         Begin MSForms.CheckBox chkGPrataDmd 
            Height          =   255
            Left            =   8670
            TabIndex        =   241
            Top             =   2925
            Width           =   3945
            VariousPropertyBits=   1015031835
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6950;450"
            Value           =   "0"
            Caption         =   "Generate Pro-rata Demand Automatically"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apportionment (%):"
            Height          =   195
            Index           =   8
            Left            =   4110
            TabIndex        =   240
            Top             =   2925
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Index           =   12
         Left            =   -74940
         TabIndex        =   207
         Top             =   390
         Width           =   18240
         Begin VB.TextBox txtMemo2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   11295
            TabIndex        =   219
            Top             =   765
            Width           =   1590
         End
         Begin VB.CommandButton cmdDeleteBreaches 
            Caption         =   "&Delete Breaches"
            Height          =   375
            Left            =   13050
            TabIndex        =   227
            Top             =   3285
            Width           =   1575
         End
         Begin VB.CommandButton cmdSetBreachType 
            Caption         =   "..."
            Height          =   285
            Left            =   3600
            TabIndex        =   211
            Top             =   765
            Width           =   255
         End
         Begin VB.TextBox txtReceivedBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8760
            TabIndex        =   218
            Top             =   765
            Width           =   1635
         End
         Begin VB.TextBox txtCommenceDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3840
            ScrollBars      =   2  'Vertical
            TabIndex        =   213
            Top             =   765
            Width           =   1605
         End
         Begin VB.TextBox txtInitiatedBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5490
            TabIndex        =   215
            Top             =   765
            Width           =   1935
         End
         Begin VB.CheckBox chkResolved 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10440
            TabIndex        =   222
            Top             =   765
            Width           =   665
         End
         Begin VB.TextBox txtDateReceived 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7455
            TabIndex        =   217
            Top             =   765
            Width           =   1275
         End
         Begin VB.CommandButton cmdBreachNew 
            Caption         =   "&New Breaches"
            Height          =   375
            Left            =   13065
            TabIndex        =   223
            Top             =   1095
            Width           =   1575
         End
         Begin VB.CommandButton cmdBreachCancel 
            Caption         =   "&Cancel Breaches"
            Height          =   375
            Left            =   13065
            TabIndex        =   226
            Top             =   2700
            Width           =   1575
         End
         Begin VB.CommandButton cmdBreachSave 
            Caption         =   "&Save Breaches"
            Height          =   375
            Left            =   13050
            TabIndex        =   225
            Top             =   2115
            Width           =   1575
         End
         Begin VB.CommandButton cmdBreachEdit 
            Caption         =   "&Edit Breaches"
            Height          =   375
            Left            =   13065
            TabIndex        =   224
            Top             =   1620
            Width           =   1575
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBreach 
            Height          =   2790
            Left            =   1320
            TabIndex        =   208
            Top             =   1125
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   4921
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label44 
            Caption         =   "Received By"
            Height          =   255
            Index           =   4
            Left            =   8760
            TabIndex        =   228
            Top             =   540
            Width           =   1515
         End
         Begin VB.Label Label44 
            Caption         =   "Breach Type"
            Height          =   210
            Index           =   0
            Left            =   1320
            TabIndex        =   221
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Commence Date"
            Height          =   300
            Index           =   1
            Left            =   3825
            TabIndex        =   220
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Initiated By"
            Height          =   255
            Index           =   2
            Left            =   5475
            TabIndex        =   216
            Top             =   540
            Width           =   1515
         End
         Begin VB.Label Label44 
            Caption         =   "Resolved"
            Height          =   195
            Index           =   5
            Left            =   10440
            TabIndex        =   214
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label44 
            Caption         =   "Date Received"
            Height          =   255
            Index           =   3
            Left            =   7455
            TabIndex        =   212
            Top             =   540
            Width           =   1275
         End
         Begin VB.Label Label44 
            Caption         =   "Memo"
            Height          =   255
            Index           =   6
            Left            =   11295
            TabIndex        =   210
            Top             =   540
            Width           =   555
         End
         Begin MSForms.ComboBox cboBreachType 
            Height          =   285
            Left            =   1320
            TabIndex        =   209
            Top             =   765
            Width           =   2295
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4048;503"
            TextColumn      =   2
            ColumnCount     =   2
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
      End
      Begin VB.Frame Frame1 
         Height          =   4740
         Index           =   6
         Left            =   -74880
         TabIndex        =   149
         Top             =   360
         Width           =   18150
         Begin VB.TextBox txtRRDemandType 
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
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   152
            Top             =   360
            Width           =   2745
         End
         Begin VB.CommandButton cmdRentReviewDemandType 
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
            Left            =   4140
            TabIndex        =   153
            Top             =   360
            Width           =   300
         End
         Begin VB.CommandButton cmdNewRentAnalysis 
            Caption         =   "&New Rent Review"
            Height          =   375
            Left            =   15105
            TabIndex        =   150
            Top             =   660
            Width           =   1700
         End
         Begin VB.CommandButton cmdCancelRentAnalysis 
            Caption         =   "&Cancel Rent Review"
            Height          =   375
            Left            =   15105
            TabIndex        =   163
            Top             =   3240
            Width           =   1700
         End
         Begin VB.CommandButton cmdSaveRentAnalysis 
            Caption         =   "&Save Rent Review"
            Height          =   375
            Left            =   15105
            TabIndex        =   160
            Top             =   1950
            Width           =   1700
         End
         Begin VB.CommandButton cmdEditRentAnalysis 
            Caption         =   "&Edit Rent Review"
            Height          =   375
            Left            =   15105
            TabIndex        =   162
            Top             =   1305
            Width           =   1700
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
            Height          =   285
            Left            =   11340
            Locked          =   -1  'True
            TabIndex        =   157
            ToolTipText     =   "Rent Increase/Decrease Amount"
            Top             =   360
            Width           =   1080
         End
         Begin VB.TextBox txtRentIncreaseDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10110
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   156
            ToolTipText     =   "Rent Increase/Decrease Date"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtRentReviewDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4500
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   154
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox txtSerial 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   151
            Top             =   360
            Width           =   900
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   12510
            TabIndex        =   158
            Top             =   390
            Width           =   255
         End
         Begin VB.CommandButton cmdDelRentAnalysis 
            Caption         =   "&Delete Rent Review"
            Enabled         =   0   'False
            Height          =   375
            Left            =   15105
            TabIndex        =   161
            Top             =   2595
            Width           =   1700
         End
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5745
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   155
            Top             =   360
            Width           =   4305
         End
         Begin VB.CheckBox chkRRStatus 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   13005
            TabIndex        =   159
            Top             =   390
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentAnalysis 
            Height          =   2955
            Left            =   360
            TabIndex        =   164
            Top             =   660
            Width           =   13560
            _ExtentX        =   23918
            _ExtentY        =   5212
            _Version        =   393216
            Cols            =   7
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
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent I/D Amt"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   11340
            TabIndex        =   173
            ToolTipText     =   "Rent Increase/Decrease Amount"
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent I/D Dt"
            Height          =   195
            Index           =   4
            Left            =   10110
            TabIndex        =   172
            ToolTipText     =   "Rent Increase/Decrease Date"
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Dt"
            Height          =   195
            Index           =   2
            Left            =   4530
            TabIndex        =   171
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Review No."
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   170
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Alarm"
            Height          =   195
            Index           =   6
            Left            =   12510
            TabIndex        =   169
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Demand Type"
            Height          =   195
            Index           =   1
            Left            =   1335
            TabIndex        =   168
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Height          =   195
            Index           =   3
            Left            =   5775
            TabIndex        =   167
            Top             =   120
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Completed"
            Height          =   195
            Index           =   7
            Left            =   12960
            TabIndex        =   166
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   $"Lease4.frx":0A02
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   165
            Top             =   3660
            Width           =   12375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Notes"
         Height          =   4680
         Index           =   5
         Left            =   -74955
         TabIndex        =   86
         Top             =   360
         Width           =   18180
         Begin VB.Frame Frame1 
            Caption         =   "Attactment Files:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   900
            Index           =   15
            Left            =   180
            TabIndex        =   174
            Top             =   3690
            Width           =   17850
            Begin VB.CommandButton cmdOpenFile 
               Caption         =   "&Open File"
               Height          =   435
               Left            =   8520
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   240
               Width           =   1350
            End
            Begin VB.CommandButton cmdClinetAddAtch 
               Caption         =   "&Add New"
               Height          =   435
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   240
               Width           =   1350
            End
            Begin VB.CommandButton cmdDeleteFile 
               Caption         =   "&Delete File"
               Height          =   435
               Left            =   10080
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   240
               Width           =   1350
            End
            Begin MSForms.ComboBox cmbFiles 
               Height          =   285
               Left            =   120
               TabIndex        =   178
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
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1763;4233"
            End
         End
         Begin VB.TextBox txtMemo 
            Height          =   3435
            Left            =   195
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   240
            Width           =   17850
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   1
         Left            =   -74920
         TabIndex        =   57
         Top             =   320
         Width           =   18135
         Begin VB.TextBox txtFreqBR 
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
            Height          =   315
            Left            =   4320
            TabIndex        =   73
            Top             =   405
            Width           =   1800
         End
         Begin VB.CommandButton cmdFreqBR 
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
            Left            =   6165
            TabIndex        =   61
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtComparenextDueDate1 
            Height          =   285
            Left            =   15480
            TabIndex        =   324
            Top             =   4140
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdBRDemandType 
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
            Left            =   2880
            TabIndex        =   59
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtBRDemandType 
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
            Height          =   315
            Left            =   135
            TabIndex        =   58
            Top             =   405
            Width           =   2700
         End
         Begin VB.CommandButton cmdRCFund 
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
            Left            =   9990
            TabIndex        =   64
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtRCFund 
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
            Height          =   315
            Left            =   8805
            TabIndex        =   74
            Top             =   405
            Width           =   1170
         End
         Begin VB.TextBox txtRCFundCode 
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
            Height          =   315
            Left            =   7658
            TabIndex        =   63
            Top             =   405
            Width           =   1125
         End
         Begin VB.TextBox txtStopRC 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   16530
            TabIndex        =   67
            Top             =   405
            Width           =   1125
         End
         Begin VB.CommandButton cmdDelRentCrg 
            Caption         =   "&Delete Rent"
            Height          =   375
            Left            =   12405
            TabIndex        =   72
            Top             =   4065
            Width           =   1335
         End
         Begin VB.CommandButton cmdEditRentCrg 
            Caption         =   "&Edit Rent"
            Height          =   375
            Left            =   7695
            TabIndex        =   69
            Top             =   4065
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveRentCrg 
            Caption         =   "&Save Rent"
            Height          =   375
            Left            =   9270
            TabIndex        =   70
            Top             =   4065
            Width           =   1335
         End
         Begin VB.CommandButton cmdNewRentCrg 
            Caption         =   "&New Rent"
            Height          =   375
            Left            =   6120
            TabIndex        =   68
            Top             =   4065
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelRentCrg 
            Caption         =   "&Cancel Rent"
            Height          =   375
            Left            =   10830
            TabIndex        =   71
            Top             =   4065
            Width           =   1335
         End
         Begin VB.TextBox txtBRChargingFigure 
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
            Height          =   315
            Left            =   12240
            MaxLength       =   12
            TabIndex        =   66
            Top             =   405
            Width           =   1260
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
            Height          =   315
            Left            =   15135
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   405
            Width           =   1365
         End
         Begin VB.TextBox txtNextDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F9EE&
            Height          =   315
            Left            =   6495
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   405
            Width           =   1080
         End
         Begin VB.TextBox txtTotalRentYear 
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
            Height          =   315
            Left            =   13545
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   405
            Width           =   1545
         End
         Begin VB.TextBox txtRentStartDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3255
            TabIndex        =   60
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox txtRentChargesIDEdit 
            Height          =   285
            Left            =   14475
            TabIndex        =   77
            Top             =   4125
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentCharges 
            Height          =   3195
            Left            =   120
            TabIndex        =   78
            Top             =   810
            Width           =   17580
            _ExtentX        =   31009
            _ExtentY        =   5636
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Stop Date"
            Height          =   195
            Index           =   9
            Left            =   16155
            TabIndex        =   129
            Top             =   165
            Width           =   705
         End
         Begin MSForms.ComboBox cboBRChargingMth 
            Height          =   315
            Left            =   10380
            TabIndex        =   65
            Top             =   405
            Width           =   1815
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3201;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   3
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "0;3527;35277"
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charging Methods"
            Height          =   195
            Index           =   5
            Left            =   10380
            TabIndex        =   101
            Top             =   165
            Width           =   1305
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   195
            Index           =   6
            Left            =   12240
            TabIndex        =   100
            Top             =   165
            Width           =   555
         End
         Begin VB.Label lblDefaultDescption 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " description"
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   97
            Top             =   4095
            Visible         =   0   'False
            Width           =   840
         End
         Begin MSForms.TextBox txtRentDesc 
            Height          =   330
            Left            =   2520
            TabIndex        =   91
            Top             =   4095
            Visible         =   0   'False
            Width           =   3135
            VariousPropertyBits=   746604571
            MaxLength       =   50
            BorderStyle     =   1
            Size            =   "5530;582"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkRentDes 
            Height          =   255
            Left            =   795
            TabIndex        =   90
            Top             =   4035
            Visible         =   0   'False
            Width           =   1695
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2990;450"
            Value           =   "1"
            Caption         =   "Default description"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Each Period"
            Height          =   195
            Index           =   8
            Left            =   15135
            TabIndex        =   85
            Top             =   165
            Width           =   825
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Next Due Dt"
            Height          =   195
            Index           =   3
            Left            =   6480
            TabIndex        =   84
            Top             =   180
            Width           =   900
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Total/Year"
            Height          =   195
            Index           =   7
            Left            =   13575
            TabIndex        =   83
            Top             =   165
            Width           =   735
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Frequency"
            Height          =   195
            Index           =   2
            Left            =   4380
            TabIndex        =   82
            Top             =   165
            Width           =   750
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
            Height          =   195
            Index           =   1
            Left            =   3300
            TabIndex        =   81
            Top             =   165
            Width           =   720
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Demand Type"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   960
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            Caption         =   "Fund"
            Height          =   195
            Index           =   4
            Left            =   7695
            TabIndex        =   79
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4800
         Index           =   2
         Left            =   -74920
         TabIndex        =   31
         Top             =   310
         Width           =   18180
         Begin VB.CommandButton cmdFreqSC 
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
            Left            =   5760
            TabIndex        =   12
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtFreqSC 
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
            Height          =   315
            Left            =   4500
            TabIndex        =   11
            Top             =   405
            Width           =   1260
         End
         Begin VB.TextBox txtSCDemandType 
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
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Top             =   405
            Width           =   3060
         End
         Begin VB.CommandButton cmdSCDemandType 
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
            Left            =   3150
            TabIndex        =   9
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtSCFundCode 
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
            Height          =   315
            Left            =   7119
            TabIndex        =   14
            Top             =   405
            Width           =   1125
         End
         Begin VB.TextBox txtSCFundName 
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
            Height          =   315
            Left            =   8265
            TabIndex        =   15
            Top             =   405
            Width           =   1170
         End
         Begin VB.CommandButton cmdSCFund 
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
            Left            =   9450
            TabIndex        =   16
            Top             =   405
            Width           =   300
         End
         Begin VB.TextBox txtCapAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   17055
            TabIndex        =   23
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtChargingFigure 
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
            Height          =   315
            Left            =   12465
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   19
            Top             =   405
            Width           =   1155
         End
         Begin VB.TextBox txtStopSC 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   15915
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   405
            Width           =   1110
         End
         Begin VB.TextBox txtSCDueEachPeriod 
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
            Height          =   315
            Left            =   14760
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   405
            Width           =   1110
         End
         Begin VB.TextBox txtSCTotalAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E6EDFB&
            Height          =   315
            Left            =   13710
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   405
            Width           =   975
         End
         Begin VB.TextBox txtPayableFrom 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   3525
            MaxLength       =   10
            TabIndex        =   10
            Top             =   405
            Width           =   930
         End
         Begin VB.TextBox txtSCNextDueDt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   13
            Top             =   405
            Width           =   930
         End
         Begin VB.CommandButton cmdSCDelete 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Delete Charge"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13380
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4140
            Width           =   1575
         End
         Begin VB.TextBox txtSCCharge 
            Height          =   285
            Left            =   6705
            TabIndex        =   88
            Top             =   4140
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdSCCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Cancel Charge"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11910
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4140
            Width           =   1335
         End
         Begin VB.CommandButton cmdSCNew 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&New Charge"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4140
            Width           =   1335
         End
         Begin VB.CommandButton cmdSCSave 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save Charge"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   4140
            Width           =   1335
         End
         Begin VB.CommandButton cmdSCEdit 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Edit Charge"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8970
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   4140
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSC 
            Height          =   3240
            Left            =   75
            TabIndex        =   93
            Top             =   810
            Width           =   17985
            _ExtentX        =   31724
            _ExtentY        =   5715
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cap Amount"
            Height          =   195
            Index           =   11
            Left            =   17055
            TabIndex        =   180
            Top             =   180
            Width           =   855
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stop Date"
            Height          =   195
            Index           =   10
            Left            =   15870
            TabIndex        =   140
            Top             =   165
            Width           =   705
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total/Year"
            Height          =   195
            Index           =   8
            Left            =   13710
            TabIndex        =   134
            Top             =   165
            Width           =   735
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Demand Type"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   139
            Top             =   165
            Width           =   1095
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule"
            Height          =   195
            Index           =   5
            Left            =   9765
            TabIndex        =   138
            Top             =   165
            Width           =   660
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            Height          =   195
            Index           =   1
            Left            =   3525
            TabIndex        =   137
            Top             =   165
            Width           =   720
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency"
            Height          =   195
            Index           =   2
            Left            =   4500
            TabIndex        =   136
            Top             =   165
            Width           =   750
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Due Dt"
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   135
            Top             =   165
            Width           =   900
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Each Period"
            Height          =   195
            Index           =   9
            Left            =   14760
            TabIndex        =   133
            Top             =   165
            Width           =   825
         End
         Begin MSForms.ComboBox cboSCChargingMth 
            Height          =   315
            Left            =   10995
            TabIndex        =   18
            Top             =   405
            Width           =   1395
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2461;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "529;35277"
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charging Methods"
            Height          =   195
            Index           =   6
            Left            =   11040
            TabIndex        =   132
            Top             =   165
            Width           =   1305
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   195
            Index           =   7
            Left            =   12510
            TabIndex        =   131
            Top             =   165
            Width           =   555
         End
         Begin VB.Label lblSC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property/Fund"
            Height          =   195
            Index           =   4
            Left            =   7140
            TabIndex        =   130
            Top             =   165
            Width           =   1035
         End
         Begin MSForms.ComboBox cboSchedule 
            Height          =   315
            Left            =   9765
            TabIndex        =   17
            Top             =   405
            Width           =   1200
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2117;556"
            TextColumn      =   2
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "529;35277"
         End
         Begin VB.Label lblDefaultDescption 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " description"
            Height          =   195
            Index           =   4
            Left            =   2595
            TabIndex        =   98
            Top             =   4200
            Visible         =   0   'False
            Width           =   840
         End
         Begin MSForms.TextBox txtSCDesc 
            Height          =   330
            Left            =   3465
            TabIndex        =   92
            Top             =   4140
            Visible         =   0   'False
            Width           =   3135
            VariousPropertyBits=   746604571
            MaxLength       =   50
            BorderStyle     =   1
            Size            =   "5530;582"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkSCDes 
            Height          =   255
            Left            =   1740
            TabIndex        =   94
            Top             =   4140
            Visible         =   0   'False
            Width           =   1815
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3201;450"
            Value           =   "1"
            Caption         =   "Default description"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2415
         Index           =   0
         Left            =   -71640
         TabIndex        =   56
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Frame Frame3 
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
         Height          =   3015
         Left            =   -72240
         TabIndex        =   29
         Top             =   1080
         Width           =   7785
      End
      Begin VB.Frame Frame2 
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
         Height          =   3015
         Left            =   -73327
         TabIndex        =   6
         Top             =   735
         Width           =   10035
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   14
         Left            =   60
         TabIndex        =   33
         Top             =   360
         Width           =   18180
         Begin VB.TextBox txtFreqIC 
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
            Height          =   315
            Left            =   4815
            TabIndex        =   35
            Top             =   450
            Width           =   1350
         End
         Begin VB.CommandButton cmdFreqIC 
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
            Left            =   6165
            TabIndex        =   36
            Top             =   450
            Width           =   300
         End
         Begin VB.CommandButton cmdInsDemandType 
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
            Left            =   3285
            TabIndex        =   32
            Top             =   450
            Width           =   300
         End
         Begin VB.TextBox txtInsDemandType 
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
            Height          =   315
            Left            =   225
            TabIndex        =   30
            Top             =   450
            Width           =   3060
         End
         Begin VB.CommandButton cmdICFundCode 
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
            Left            =   9975
            TabIndex        =   40
            Top             =   450
            Width           =   300
         End
         Begin VB.TextBox txtICFundName 
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
            Height          =   315
            Left            =   8790
            TabIndex        =   39
            Top             =   450
            Width           =   1170
         End
         Begin VB.TextBox txtICFundCode 
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
            Height          =   315
            Left            =   7650
            TabIndex        =   38
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtStopIC 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   16200
            TabIndex        =   45
            Top             =   450
            Width           =   1095
         End
         Begin VB.TextBox txtInsPercentage 
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
            Height          =   315
            Left            =   12285
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   42
            Top             =   450
            Width           =   1155
         End
         Begin VB.TextBox txtTotalYearlyIns 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E6EDFB&
            Height          =   315
            Left            =   13500
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   450
            Width           =   1200
         End
         Begin VB.TextBox txtInsEachPeriod 
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
            Height          =   315
            Left            =   14745
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   450
            Width           =   1395
         End
         Begin VB.CommandButton cmdInsDelete 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Delete Ins."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12465
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   3585
            Width           =   1320
         End
         Begin VB.CommandButton cmdIncEdit 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Edit Ins."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8100
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   3585
            Width           =   1320
         End
         Begin VB.CommandButton cmdIncSave 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save Ins."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9555
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   3585
            Width           =   1320
         End
         Begin VB.CommandButton cmdIncNew 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&New Ins."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6645
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   3585
            Width           =   1320
         End
         Begin VB.CommandButton cmdIncCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Cancel Ins."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11010
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   3585
            Width           =   1320
         End
         Begin VB.TextBox txtInsStartDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3690
            TabIndex        =   34
            Top             =   450
            Width           =   1095
         End
         Begin VB.TextBox txtInsNextDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F9EE&
            Height          =   315
            Left            =   6495
            TabIndex        =   37
            Top             =   450
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxIns 
            Height          =   2715
            Left            =   210
            TabIndex        =   89
            Top             =   810
            Width           =   17085
            _ExtentX        =   30136
            _ExtentY        =   4789
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblInc 
            AutoSize        =   -1  'True
            Caption         =   "Stop Date:"
            Height          =   195
            Index           =   9
            Left            =   16200
            TabIndex        =   145
            Top             =   210
            Width           =   735
         End
         Begin VB.Label lblInc 
            AutoSize        =   -1  'True
            Caption         =   "Each Period"
            Height          =   195
            Index           =   8
            Left            =   14760
            TabIndex        =   143
            Top             =   225
            Width           =   825
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total/Year"
            Height          =   195
            Index           =   7
            Left            =   13455
            TabIndex        =   144
            Top             =   225
            Width           =   735
         End
         Begin VB.Label lblInc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charging Methods"
            Height          =   195
            Index           =   5
            Left            =   10320
            TabIndex        =   142
            Top             =   210
            Width           =   1305
         End
         Begin MSForms.ComboBox cboIncCharMth 
            Height          =   315
            Left            =   10365
            TabIndex        =   41
            Top             =   450
            Width           =   1875
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3307;556"
            TextColumn      =   2
            ColumnCount     =   6
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "529;35277"
         End
         Begin VB.Label lblInc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   195
            Index           =   6
            Left            =   12285
            TabIndex        =   141
            Top             =   210
            Width           =   555
         End
         Begin VB.Label lblDefaultDescption 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " description"
            Height          =   195
            Index           =   8
            Left            =   1755
            TabIndex        =   99
            Top             =   3675
            Visible         =   0   'False
            Width           =   840
         End
         Begin MSForms.TextBox txtInsDesc 
            Height          =   330
            Left            =   2685
            TabIndex        =   96
            Top             =   3630
            Visible         =   0   'False
            Width           =   3135
            VariousPropertyBits=   746604571
            MaxLength       =   50
            BorderStyle     =   1
            Size            =   "5530;582"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkInsDes 
            Height          =   375
            Left            =   885
            TabIndex        =   95
            Top             =   3570
            Visible         =   0   'False
            Width           =   1695
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2990;661"
            Value           =   "1"
            Caption         =   "Default description"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fund"
            Height          =   195
            Index           =   4
            Left            =   7680
            TabIndex        =   55
            Top             =   210
            Width           =   360
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Demand Type"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   54
            Top             =   210
            Width           =   960
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
            Height          =   195
            Index           =   1
            Left            =   3705
            TabIndex        =   53
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frequency"
            Height          =   195
            Index           =   2
            Left            =   4845
            TabIndex        =   52
            Top             =   210
            Width           =   750
         End
         Begin VB.Label lblInc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Next Due"
            Height          =   195
            Index           =   3
            Left            =   6525
            TabIndex        =   51
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rent Payable:"
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
         Left            =   -65160
         TabIndex        =   7
         Top             =   3960
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label49 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72900
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Label Label41 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4950
      TabIndex        =   4
      Top             =   2880
      Width           =   1515
   End
End
Attribute VB_Name = "frmLease4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form ID: frmLease4
'Form Name: Lease Details

Option Explicit

Dim BREACH_NEW_ENTRY_ As Boolean
Dim ASSIGNMENT_NEW_ENTRY_ As Boolean
Dim RENTCHARGES_EDIT As Integer
Dim SERVICECHARGES_EDIT As Integer
Dim INSURANCECHARGES_EDIT As Integer
Dim gridBreach_EDIT  As Integer
Dim Breach_EDIT As Integer
Dim RENT_REVIEW_ADDNEW_MODE As Boolean ' true means adding new record and false means this is in edit mode rentreview
Dim PROPERTY_ID As String
Dim COPY_LEASE As Boolean

Dim Rst1 As New ADODB.Recordset
Dim Conn2 As New ADODB.Connection
Dim Rst2 As New ADODB.Recordset
Dim szSQL As String
Dim SQLStr2 As String

Public FormLoad As Boolean
Const RRID = 8                'Rent Review ID
Dim szaProperty() As String
Dim szaTenantBalance()  As String
Dim szSel As String
Dim strSessionClientID As String
Dim strSessionPropertyID As String
Dim szUndoLeaseEndDate As String
Dim szUndoLeaseStatus As String
Dim bReviewLocked As Boolean
Dim szLeaseStatus As Boolean
Public var
Dim sText As String
Dim strAssignmentID As String
Dim strBreachID As String
'Dim txtComparenextDueDate1 As TextBox
Dim strLeaseId As String
Dim szClientID As String

Private Sub cboBRChargingMth_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtBRChargingFigure.SetFocus
    End If
End Sub

Private Sub cboBRDemandType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtRentStartDate.SetFocus
    End If
End Sub

Private Sub cboBreachType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        txtCommenceDate.SetFocus
    End If
End Sub



Private Sub cboFreqSC_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtSCNextDueDt.SetFocus
    End If
End Sub

Private Sub cboIncCharMth_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtInsPercentage.SetFocus
    End If
End Sub

Private Sub cboInsDemandType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtInsStartDate.SetFocus
    End If
End Sub

Private Sub cboInsDept_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cboIncCharMth.SetFocus
    End If
End Sub

'Private Sub cboRentChargeDept_Click()
'    'added by anol 23 07 2016
'    If cboRentChargeDept.ListIndex >= 0 Then
'        cboRentChargeDept.ToolTipText = cboRentChargeDept.Column(1)
'    End If
'
'End Sub



Private Sub cboRentChargeDept_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cboBRChargingMth.SetFocus
    End If
End Sub

Private Sub cboRRDemandType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtRentReviewDate.SetFocus
    End If
End Sub

Private Sub cboSCChargingMth_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtChargingFigure.SetFocus
    End If
End Sub

Private Sub cboSCDemandType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 9 Then
'        txtPayableFrom.SetFocus
'    End If
End Sub

Private Sub cboSCDemandType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtPayableFrom.SetFocus
    End If
End Sub

Private Sub cboSCDept_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cboSchedule.SetFocus
    End If
End Sub

Private Sub cboSchedule_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cboSCChargingMth.SetFocus
    End If
End Sub

Private Sub cboUsage_LostFocus()
    Dim adoConn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    If Trim(cboUsage.text) <> "" Then
        adoConn.Open getConnectionString
        szSQL = "SELECT value " & _
                "FROM SECONDARYCODE  " & _
                "where value='" & cboUsage.text & "' AND PRIMARYCODE = 'UUSE';"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If adoRST.EOF Then
            MsgBox "Please select a correct value from the list", vbInformation, "Please select a correct value."
            cboUsage.text = ""
            FocusControl cboUsage
        End If
        adoConn.Close
        Set adoConn = Nothing
    End If
End Sub

Private Sub chkIncluseExlessee_Click()
    cmdtenants_Click
End Sub

Private Sub cmdCancelBreaches_Click()

End Sub

Private Sub cmdBRDemandType_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "RCDemandType"
    Call LoadDemandTypes
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Sub LoadDemandTypes()
    Dim rRow As Integer
   Dim szSQL As String
    Dim szSQL1 As String

   Dim iSel As Integer
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
    Dim adoRst1 As New ADODB.Recordset
   Dim i As Integer, szaData() As String
   Dim iAllDemandType As Integer
   
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1265
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Demand ID"
   lblClientID(1).Caption = "Demand Type"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup(0).Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoConn.Open getConnectionString
   'szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   szSQL1 = "SELECT COUNT(*) AS C_I FROM DemandTypes"
   adoRst1.Open szSQL1, adoConn, adOpenStatic, adLockReadOnly
   If adoRst1!C_I = 0 Then
      
      Exit Sub
   End If

   adoRst1.Close

'  RENT           *****************************************************
   If szSel = "RCDemandType" Or szSel = "cmdRentReviewDemandType" Then
        szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (CategoryCode = 1 OR CategoryCode = 4) AND " & _
                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
   End If
   If szSel = "SCDemandType" Then
         szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (CategoryCode = 2 OR CategoryCode = 4 OR CategoryCode = 5) AND " & _
                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
   End If
   If szSel = "ICDemandType" Then
        szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (CategoryCode = 3 OR CategoryCode = 4) AND " & _
                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
   Else
        If szSQL = "" Then
                          szSQL = "SELECT DISTINCT ID, Type " & _
                        "FROM LRentCharges, DemandTypes " & _
                        "WHERE LRentCharges.BRDemandType = DemandTypes.ID AND " & _
                             "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
        End If
   
   End If
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iAllDemandType = rstRec.RecordCount
   ReDim szaData(1, iAllDemandType) As String
   
   
   If rstRec.EOF Then
        MsgBox "You need to setup demandtypes", vbInformation, "Warning"
        flxClientList.Clear
        flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Type").Value
                    'flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundID").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadFrequencyBR()
    Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
    Dim adoRst1 As New ADODB.Recordset
   Dim i As Integer, szaData() As String
   Dim iAllDemandType As Integer
   
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1265
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Frequencies Name"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup(0).Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoConn.Open getConnectionString
   'szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   szSQL = "SELECT * FROM Frequencies"
   
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iAllDemandType = rstRec.RecordCount
   ReDim szaData(1, iAllDemandType) As String
   
   
   If rstRec.EOF Then
        MsgBox "You need to setup Frequencies", vbInformation, "Warning"
        flxClientList.Clear
        flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Frequency").Value
                    'flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundID").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdClose1_Click()
     fraList(0).Visible = False
     fraList(1).Enabled = True
     tabLease.Enabled = True
End Sub

Private Sub cmdDeleteBreaches_Click()
    If cmdSaveNew.Visible Then Exit Sub

'   If cmdDeleteBreaches.Caption = "&Delete Breaches" Then
'        If MsgBox("Would you like to delete this Breach information?", vbQuestion + vbYesNo, "Breach information") = vbNo Then Exit Sub
'        gridBreach.TextMatrix(gridBreach.row, 7) = "DELETED"
'        gridBreach.RowHeight(gridBreach.row) = 0
'        'MsgBox "This Breach information has been marked for deletition. It will be permanently removed when you save this lease.", vbInformation + vbOKOnly, "Breach information"
'        cboBreachType.ListIndex = -1
'        txtCommenceDate.text = ""
'        txtInitiatedBy.text = ""
'        txtDateReceived.text = ""
'        txtReceivedBy.text = ""
'        txtMemo2.text = ""
'        chkResolved.Value = 0
'        Breach_EDIT = 0
'   Else
'        gridBreach.TextMatrix(gridBreach.row, 7) = ""
'        MsgBox "This Breach information has been retrieved.", vbInformation + vbOKOnly, "Breach information"
'   End If
   
    Dim adoConn1   As New ADODB.Connection
    If MsgBox("Would you like to delete this Breach information?", vbQuestion + vbYesNo, "Please Confirm to delete current Breach information") = vbYes Then
        adoConn1.Open getConnectionString
        adoConn1.Execute "Delete from LeaseBreaches where BreachID=" & flxSC.TextMatrix(flxSC.row, 6) & ""
        adoConn1.Close
        Set adoConn1 = Nothing
        Call loadFlxBreach
      
        cboBreachType.ListIndex = -1
        txtCommenceDate.text = ""
        txtInitiatedBy.text = ""
        txtDateReceived.text = ""
        txtReceivedBy.text = ""
        txtMemo2.text = ""
        chkResolved.Value = 0
        Breach_EDIT = 0
    End If

   BreachButtonMode DefaultMode
End Sub

Private Sub cmdFreqBR_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "FreqBR"
    Call LoadFrequencyBR
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdFreqIC_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "FreqIC"
    Call LoadFrequencyBR
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdFreqSC_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "FreqSC"
    Call LoadFrequencyBR
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdICFundCode_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "ICFund"
    Call LoadFunds
    Frame5.Top = 1930
    Frame5.Left = txtSCFundCode.Left - 1100
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdInsDemandType_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "ICDemandType"
    Call LoadDemandTypes
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdInsFreq_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "FreqIC"
    Call LoadFrequencyBR
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdPropertyList_Click()
   fraList(1).Enabled = False
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
   fraList(0).Left = fraList(1).Left + 500 'tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = fraList(1).Top + 200 'tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "Property"
End Sub
Private Sub LoadflxProperty()
   flxSupplier(0).Clear
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
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0


   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
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
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim iRows As Integer
   flxSupplier(0).Rows = 2
   iRows = 1
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = "ALL"
      flxSupplier(0).TextMatrix(iRows, 2) = "ALL"
      flxSupplier(0).AddItem ""
   iRows = 2
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("PropertyID").Value
      flxSupplier(0).TextMatrix(iRows, 2) = adoRST.Fields.Item("PropertyName").Value
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend
 
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
  
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadflxUnits(Optional ByVal Filter As String)

   flxSupplier(0).Clear
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 70
   flxSupplier(0).ColWidth(1) = 1500
   flxSupplier(0).ColWidth(2) = 3300
   flxSupplier(0).ColWidth(3) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify


   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(1)

   lblSearch0(0).Caption = "Unit Number"
   lblSearch1(0).Caption = "Unit Description"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0


   'On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString


   If szSel = "HeadLease" Then
        szSQL = "SELECT LeaseID,L.UnitNumber,U.UnitName  FROM LeaseDetails L ,Units U where " & _
                "U.UNitnumber=L.unitnumber AND L.Status =true AND L.UnitNumber is not null Order by L.Unitnumber"
   ElseIf szSel = "Unit" Then 'loading the units that are not occupied
        szSQL = "SELECT Units.UnitNumber, Units.UnitName FROM Units LEFT JOIN " & _
                " (SELECT LeaseDetails.UnitNumber FROM LeaseDetails WHERE LeaseDetails.Status = TRUE) as X " & _
                "ON X.UnitNumber=Units.UnitNumber where X.UnitNumber is null"
   End If
    
 
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Len(Filter) > 0 Then
        adoRST.Filter = Filter
   End If
   Dim iRows As Integer
   flxSupplier(0).Rows = 2

   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("UnitNumber").Value
      flxSupplier(0).TextMatrix(iRows, 2) = adoRST.Fields.Item("UnitName").Value
      If szSel = "HeadLease" Then
            flxSupplier(0).TextMatrix(iRows, 3) = adoRST.Fields.Item("LeaseID").Value
      End If
      'flxSupplier(0).RowHeight(iRows) = 280
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend
   adoRST.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
  
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdTerminationClose_Click()
    tabLease.Enabled = True
    Frame1(16).Enabled = True
End Sub

Private Sub cmdTerminationClose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         fraList(3).Visible = False
    End If
End Sub

Private Sub cmdRCFund_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "RCFund"
    Call LoadFunds
    Frame5.Top = 1930
    Frame5.Left = txtRCFundCode.Left - 1100
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdRentReviewDemandType_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
'    szSel = "RCDemandType"
    szSel = "cmdRentReviewDemandType"
    Call LoadDemandTypes
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdSCDemandType_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "SCDemandType"
    Call LoadDemandTypes
    Frame5.Top = 1930
    Frame5.Left = txtBRDemandType.Left + 500
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdSCFund_Click()
    tabLease.Enabled = False
    Frame1(16).Enabled = False
    Frame1(8).Enabled = False
    szSel = "SCFund"
    Call LoadFunds
    Frame5.Top = 1930
    Frame5.Left = txtSCFundCode.Left - 1100
    Frame5.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdUnitNumber_Click(Index As Integer)
   If Index = 0 Then
        If cmdSaveNew.Visible Then 'this shall load Unit grid
            szSel = "Unit"
            LoadflxUnits
         
            'tabTenant.Enabled = False 'it is already false by other lessee grid
            txtSearch1.Visible = True
            txtSearch2.Visible = True
         
            txtSearch1.text = ""
            txtSearch2.text = ""
         
            fraList(0).Left = 270 'tabTenant.Left + txtDNC(1).Left
            fraList(0).Top = 720 'tabTenant.Top + txtDNC(1).Top
            fraList(0).Visible = True
            fraList(0).ZOrder 0
            txtSearch1.SetFocus
            
        End If
   ElseIf Index = 1 Then ' this shall load HeadLease grid
        szSel = "HeadLease"
        tabLease.Enabled = False
        LoadflxUnits
         
       'tabTenant.Enabled = False 'it is already false by other lessee grid
       txtSearch1.Visible = True
       txtSearch2.Visible = True
    
       txtSearch1.text = ""
       txtSearch2.text = ""
    
       fraList(0).Left = 675 'tabTenant.Left + txtDNC(1).Left
       fraList(0).Top = 2610 'tabTenant.Top + txtDNC(1).Top
       fraList(0).Visible = True
       fraList(0).ZOrder 0
       txtSearch1.SetFocus
      
   End If
End Sub




Private Sub Command1_Click()
'    Dim adoconn As New ADODB.Connection
'    Dim rsDemandType As New ADODB.Recordset
'    Dim szSQL As String
'    Dim dictAllDemandType As New Dictionary
'    Dim iRow As Integer
'    Dim result
'
'    For iRow = 1 To 5
'        dictAllDemandType.Add iRow, iRow
'    Next iRow
'     For iRow = 1 To 6
'         result = dictAllDemandType(iRow)
'    Next iRow
'    Debug.Print IsEmpty(result)
End Sub

Private Sub flxClientList_Click()
    fraList(1).Enabled = True
    tabLease.Enabled = True
    Frame5.Visible = False
    Dim adoConn As New ADODB.Connection
    If szSel = "Client" Then
         txtClientList.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtClientList.text = flxClientList.TextMatrix(flxClientList.row, 2)
         If flxClientList.TextMatrix(flxClientList.row, 1) <> "ALL" Then
             strSessionClientID = flxClientList.TextMatrix(flxClientList.row, 1)
         Else
             strSessionClientID = ""
         End If
       
         txtPropertyList.Tag = "ALL"
         txtPropertyList.text = "ALL"
         cboClientList_Click
         FocusControl cmdPropertyList
    ElseIf szSel = "RCFund" Then
         txtRCFundCode.Tag = flxClientList.TextMatrix(flxClientList.row, 3)
         txtRCFundCode.text = flxClientList.TextMatrix(flxClientList.row, 1)
         txtRCFund.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl cboBRChargingMth
   ElseIf szSel = "SCFund" Then
         txtSCFundCode.Tag = flxClientList.TextMatrix(flxClientList.row, 3)
         txtSCFundCode.text = flxClientList.TextMatrix(flxClientList.row, 1)
         txtSCFundName.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl cboSchedule
   ElseIf szSel = "ICFund" Then
         txtICFundCode.Tag = flxClientList.TextMatrix(flxClientList.row, 3)
         txtICFundCode.text = flxClientList.TextMatrix(flxClientList.row, 1)
         txtICFundName.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl cboIncCharMth
    ElseIf szSel = "RCDemandType" Then
         txtBRDemandType.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtBRDemandType.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl txtRentStartDate
     ElseIf szSel = "cmdRentReviewDemandType" Then
     'here you need to put textbox name
         txtRRDemandType.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtRRDemandType.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl txtRentReviewDate
    ElseIf szSel = "SCDemandType" Then
         txtSCDemandType.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtSCDemandType.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl txtPayableFrom
    ElseIf szSel = "ICDemandType" Then
         txtInsDemandType.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtInsDemandType.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         FocusControl txtInsStartDate
         ' szSel = "FreqBR"
    ElseIf szSel = "FreqBR" Then
         txtFreqBR.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtFreqBR.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         Call cboFreqBR_LostFocus
         FocusControl txtNextDueDate
    ElseIf szSel = "FreqSC" Then
         txtFreqSC.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtFreqSC.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         Call cboFreqSC_LostFocus
         FocusControl txtSCNextDueDt
    ElseIf szSel = "FreqIC" Then
         txtFreqIC.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
         txtFreqIC.text = flxClientList.TextMatrix(flxClientList.row, 2)
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         Frame1(8).Enabled = True
         Call cboInsFreq_LostFocus
         FocusControl txtInsNextDueDate
    End If
   
End Sub

Private Sub flxClientList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClientList_Click
    End If
End Sub

Private Sub flxIns_Click()
    If flxIns.TextMatrix(1, 0) = "" Then Exit Sub
   txtICFundCode.Tag = flxIns.TextMatrix(flxIns.row, 17)
   txtICFundCode.text = flxIns.TextMatrix(flxIns.row, 5)
   txtICFundName.text = flxIns.TextMatrix(flxIns.row, 18)
   txtInsStartDate.text = flxIns.TextMatrix(flxIns.row, 2)
   txtFreqIC.Tag = flxIns.TextMatrix(flxIns.row, 11)
   txtFreqIC.text = flxIns.TextMatrix(flxIns.row, 3)
   txtInsDemandType.Tag = flxIns.TextMatrix(flxIns.row, 12)
   txtInsDemandType.text = flxIns.TextMatrix(flxIns.row, 19)
   txtInsNextDueDate.text = flxIns.TextMatrix(flxIns.row, 4)
   cboIncCharMth.Value = flxIns.TextMatrix(flxIns.row, 13)
   txtInsPercentage.text = flxIns.TextMatrix(flxIns.row, 7)
   txtTotalYearlyIns.text = flxIns.TextMatrix(flxIns.row, 8)
   txtInsEachPeriod.text = flxIns.TextMatrix(flxIns.row, 9)
   txtInsDesc.text = flxIns.TextMatrix(flxIns.row, 14)
   txtStopIC.text = flxIns.TextMatrix(flxIns.row, 16)

   ControlsModeInsuranceCharges GridRowOnSelection
End Sub

Private Sub flxLeaseList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxLeaseList_Click
    End If
End Sub

Private Sub flxRentAnalysis_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call AllDemandType(adoConn)
    adoConn.Close
    txtSerial.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 1)
     txtRRDemandType.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 2)
    txtRRDemandType.Tag = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 3)
    txtRentReviewDate.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 4)
    txtComments.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 5)
    txtRentIncreaseDate.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 6)
    txtRentIncreaseAmount.text = flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 7)
    chkAlarm.Value = IIf(flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 9) = "NO", 0, 1)
    chkRRStatus.Value = IIf(flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 11) = "NO", 0, 1)
End Sub

Private Sub flxRentCharges_Click()
   If flxRentCharges.TextMatrix(1, 0) = "" Then Exit Sub

'   On Error Resume Next

   txtBRDemandType.Tag = flxRentCharges.TextMatrix(flxRentCharges.row, 0)
   txtBRDemandType.text = flxRentCharges.TextMatrix(flxRentCharges.row, 18)
   txtRCFundCode.Tag = flxRentCharges.TextMatrix(flxRentCharges.row, 6)
   txtRCFundCode.text = flxRentCharges.TextMatrix(flxRentCharges.row, 7)
   txtRCFund.text = flxRentCharges.TextMatrix(flxRentCharges.row, 17)
   txtFreqBR.Tag = flxRentCharges.TextMatrix(flxRentCharges.row, 3)
   txtFreqBR.text = flxRentCharges.TextMatrix(flxRentCharges.row, 19)
   txtRentStartDate.text = flxRentCharges.TextMatrix(flxRentCharges.row, 2)
   cboBRChargingMth.Value = flxRentCharges.TextMatrix(flxRentCharges.row, 8)
   txtBRChargingFigure.text = flxRentCharges.TextMatrix(flxRentCharges.row, 10)
   txtTotalRentYear.text = flxRentCharges.TextMatrix(flxRentCharges.row, 11)
   txtNextDueDate.text = flxRentCharges.TextMatrix(flxRentCharges.row, 5)
   txtRentDueEachPeriod.text = flxRentCharges.TextMatrix(flxRentCharges.row, 12)
   txtRentChargesIDEdit.text = flxRentCharges.TextMatrix(flxRentCharges.row, 13)
   txtRentDesc.text = ""
   If flxRentCharges.TextMatrix(flxRentCharges.row, 14) <> "" Then
      txtRentDesc.text = flxRentCharges.TextMatrix(flxRentCharges.row, 14)
      txtSCDesc.Visible = True
      chkRentDes.Value = 0
   Else
      txtRentDesc.text = ""
      txtRentDesc.Visible = False
      chkRentDes.Value = 1
   End If
   txtStopRC.text = flxRentCharges.TextMatrix(flxRentCharges.row, 16)

   ControlsModeRentCharges GridRowOnSelection
End Sub

Private Sub flxRentCharges_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flxRentCharges.MouseCol = 7 Then
        flxRentCharges.ToolTipText = flxRentCharges.TextMatrix(flxRentCharges.MouseRow, 17)
    Else
         flxRentCharges.ToolTipText = ""
    End If
End Sub

Private Sub flxSC_Click()
   If flxSC.TextMatrix(1, 0) = "" Then Exit Sub
   txtSCDemandType.Tag = flxSC.TextMatrix(flxSC.row, 0)
   txtSCDemandType.text = flxSC.TextMatrix(flxSC.row, 20)
   txtPayableFrom.text = flxSC.TextMatrix(flxSC.row, 2)
   'cboFreqSC.ListIndex = flxSC.TextMatrix(flxSC.Row, 3) - 1
   txtFreqSC.Tag = flxSC.TextMatrix(flxSC.row, 17)
   txtFreqSC.text = flxSC.TextMatrix(flxSC.row, 3)
   txtSCNextDueDt.text = flxSC.TextMatrix(flxSC.row, 4)
   'cboSCDept.ListIndex = flxSC.TextMatrix(flxSC.Row, 5) - 1
   txtSCFundCode.Tag = flxSC.TextMatrix(flxSC.row, 5)
   txtSCFundCode.text = flxSC.TextMatrix(flxSC.row, 6)
   txtSCFundName.text = flxSC.TextMatrix(flxSC.row, 19)
   'If flxSC.TextMatrix(flxSC.Row, 7) <> "" Then cboSchedule.ListIndex = CInt(flxSC.TextMatrix(flxSC.Row, 7)) - 1
   If Val(flxSC.TextMatrix(flxSC.row, 7)) > 0 Then cboSchedule.Value = CInt(flxSC.TextMatrix(flxSC.row, 7))
   'cboSCChargingMth.ListIndex = flxSC.TextMatrix(flxSC.Row, 9) - 1
   cboSCChargingMth.Value = flxSC.TextMatrix(flxSC.row, 9)
   txtChargingFigure.text = flxSC.TextMatrix(flxSC.row, 11)
   txtSCTotalAmount.text = flxSC.TextMatrix(flxSC.row, 12)
   txtSCDueEachPeriod.text = flxSC.TextMatrix(flxSC.row, 13)
   txtSCCharge.text = flxSC.TextMatrix(flxSC.row, 14)

   txtSCDesc.text = ""
   If flxSC.TextMatrix(flxSC.row, 15) <> "" Then
      txtSCDesc.text = flxSC.TextMatrix(flxSC.row, 15)
      txtSCDesc.Visible = True
      chkSCDes.Value = 0
   Else
      txtSCDesc.text = ""
      txtSCDesc.Visible = False
      chkSCDes.Value = 1
   End If
   txtStopSC.text = flxSC.TextMatrix(flxSC.row, 18)

   ControlsModeServiceCharges GridRowOnSelection
   Exit Sub
End Sub

Private Sub flxSupplier_Click(Index As Integer)
   fraList(0).Visible = False
   tabLease.Enabled = True
   fraList(1).Enabled = True
   Dim adoConn As New ADODB.Connection
   If szSel = "Client" Then
        txtClientList.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtClientList.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
        If flxSupplier(0).TextMatrix(flxSupplier(0).row, 1) <> "ALL" Then
            strSessionClientID = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        Else
            strSessionClientID = ""
        End If
      
        txtPropertyList.Tag = "ALL"
        txtPropertyList.text = "ALL"
        cboClientList_Click
        FocusControl cmdPropertyList
   End If
    If szSel = "Property" Then
        txtPropertyList.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtPropertyList.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
        If flxSupplier(0).TextMatrix(flxSupplier(0).row, 1) <> "ALL" Then
            strSessionPropertyID = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        Else
            strSessionPropertyID = ""
        End If
            
        
        cboPropertyList_Click
        'adoconn.Open getConnectionString
        'LoadDept adoconn 'load the fund for all charges
        'adoconn.Close
        txtSearchTenant.SetFocus
   End If
   If szSel = "Unit" Then
         If txtTenant.text = "" Then
              MsgBox "Please select a tenant.", vbCritical + vbOKOnly, "Lease"
              txtUnitNumber.text = ""
              cmdtenants_Click
              
              Exit Sub
           End If
           txtUnitNumber.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
           If Trim(txtUnitNumber.text) = "" Then
                txtUnitName(0).text = ""
                Exit Sub ' No unit is selected from the grid
          End If
           txtUnitName(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
           Dim Conn1 As New ADODB.Connection
        
           Conn1.Open getConnectionString
        
           szSQL = "SELECT ClientName, PropertyName, Property.PropertyID,Units.UnitName " & _
                     "FROM Client, Property, Units " & _
                     "WHERE Client.ClientID = Property.ClientID And " & _
                         "Property.PropertyID = Units.PropertyID And " & _
                         "Units.UnitNumber = '" & txtUnitNumber.text & "';"
        
           Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
        'Rst1.Close
        'Exit Sub
           txtClient.text = Rst1!ClientName
           txtProperty.text = Rst1!PropertyName
           txtUnitName(0).text = Rst1!UnitName
           If Len(PROPERTY_ID) > 0 And COPY_LEASE = True Then
                 If PROPERTY_ID <> Rst1!propertyID Then
                    MsgBox " The unit selected does not match the client or the property of the lease being copied from." & vbCrLf & _
                        "There is an inconsistency in the demand types on this lease. Please assign the correct demand types to this lease before saving."
                 End If
           End If
           PROPERTY_ID = Rst1!propertyID
        
           AllDemandType Conn1
           'LoadDept Conn1 'load  fund for all charges
           Rst1.Close
           Conn1.Close
           Set Rst1 = Nothing
           Set Conn1 = Nothing
        '***************************************************************************
        'CREATE THE LEASE ID AUTOMATICALLY
        '***************************************************************************
           If cmdAddNew.Visible Then Exit Sub
        
           Dim szaTenant() As String ', szaUnit() As String
        
           If txtTenant.text <> "" Then
           'issue 540 Lease details form - Not working correctly
           'Fixed by anol 22 Feb 2015
              'szaTenant = Split(txtTenant.text, " / ")
        '   szaUnit = Split(cboUnit.text, " - ")
             ' txtLeaseID.text = OnlyNumericString(szaTenant(0)) & OnlyNumericString(txtUnitName(0).text & Format(Now, "YYMMDDHHMMSS"))
             strLeaseId = OnlyNumericString(txtTenant.Tag) & OnlyNumericString(txtUnitName(0).text & Format(Now, "YYMMDDHHMMSS"))
          End If
        
           
   End If
   If szSel = "HeadLease" Then
        txtUnitName(1).Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 3)
        txtUnitName(1).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   flxSupplier(0).Clear
   
End Sub

Private Sub flxSupplier_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplier_Click (Index)
    End If
End Sub

Private Sub Label16_Click(Index As Integer)
    If Index = 12 Then
        If Label16(12).FontUnderline = True Then
            If chkOLED.Value = 0 And Trim(txtLeaseEndDate.text) = "" Then
                MsgBox "Lease end date cannot be empty when lease override is false", vbInformation, "Warning"
                txtLeaseEndDate.SetFocus
                Exit Sub
            End If
            frmTerminationDate.SourceOfCalling = "Link"
            frmTerminationDate.LeaseEndDate = txtLeaseEndDate.text
            frmTerminationDate.LeaseOverRide = CBool(chkOLED.Value)
            'frmTerminationDate.txtTerminationDate.text = Label16(12).Tag
            If Label16(12).Tag = "" Then
                DisplayDateform Me, "", Date
            Else
                DisplayDateform Me, "", Label16(12).Tag
            End If
        End If
    End If
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRentIncreaseDate.SetFocus
    End If
End Sub

Private Sub txtICFundCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtICFundName
    End If
End Sub

Private Sub txtICFundName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdICFundCode
    End If
End Sub

Private Sub txtInitiatedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateReceived.SetFocus
    End If
End Sub

Private Sub txtInsEachPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtStopIC.SetFocus
    End If
End Sub

Private Sub txtInsNextDueDate_Change()
     If txtInsNextDueDate.text <> "" Then
        TextBoxChangeDate txtInsNextDueDate
   End If
End Sub

Private Sub txtInsNextDueDate_LostFocus()
     TextBoxFormatDate txtInsNextDueDate
     If IsDate(txtInsNextDueDate.text) And IsDate(txtInsStartDate.text) Then
        If DateDiff("d", txtInsStartDate.text, txtInsNextDueDate.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the insurance start date", vbInformation, "Warning"
            txtInsNextDueDate.text = ""
            FocusControl txtInsNextDueDate
        End If
    End If
End Sub



Private Sub txtMemo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdBreachSave.Enabled Then
        FocusControl cmdBreachSave
    End If
End Sub

Private Sub txtNextDueDate_Change()
    If txtNextDueDate.text <> "" Then
        TextBoxChangeDate txtNextDueDate
   End If
End Sub

Private Sub txtNextDueDate_LostFocus()
    TextBoxFormatDate txtNextDueDate
    If IsDate(txtNextDueDate.text) And IsDate(txtRentStartDate.text) Then
        If DateDiff("d", txtRentStartDate.text, txtNextDueDate.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the Rent start date", vbInformation, "Warning"
            txtNextDueDate.text = ""
            FocusControl txtNextDueDate
        End If
    End If
End Sub

Private Sub txtRCFund_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtRCFund.ToolTipText = txtRCFund.text
End Sub

Private Sub txtRCFundCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtRCFundCode.ToolTipText = txtRCFundCode.text
End Sub

Private Sub txtReceivedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdBreachSave
    End If
End Sub

Private Sub txtRentDueEachPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtStopRC.SetFocus
    End If
End Sub

Private Sub txtSCDueEachPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtStopSC.SetFocus
    End If
End Sub

Private Sub txtSCFundCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtSCFundCode.ToolTipText = txtSCFundCode.text
End Sub

Private Sub txtSCFundName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtSCFundName.ToolTipText = txtSCFundName.text
End Sub

Private Sub txtSCNextDueDt_Change()
     If txtSCNextDueDt.text <> "" Then
        TextBoxChangeDate txtSCNextDueDt
   End If
End Sub

Private Sub txtSCNextDueDt_LostFocus()
    TextBoxFormatDate txtSCNextDueDt
     If IsDate(txtSCNextDueDt.text) And IsDate(txtPayableFrom.text) Then
        If DateDiff("d", txtPayableFrom.text, txtSCNextDueDt.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the Service Charge start date", vbInformation, "Warning"
            txtSCNextDueDt.text = ""
            FocusControl txtSCNextDueDt
        End If
    End If
End Sub

Private Sub txtSCTotalAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSCDueEachPeriod.SetFocus
    End If
End Sub

Private Sub txtSearch1_Change()
'   Dim i As Integer
'
'   If Len(txtSearch1.text) > 0 Then
'      txtSearch2.text = ""
'   End If
'
'   For i = 1 To flxSupplier(0).Rows - 1
'      flxSupplier(0).RowHeight(i) = 240
'      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtSearch1.text))) <> UCase(txtSearch1.text) Then
'         flxSupplier(0).RowHeight(i) = 0
'      End If
'   Next i
  'Updated by anol 22 Dec 2015
   Dim i As Integer
   Dim j As Integer
   Dim Filter As String
   If szSel = "Unit" Or szSel = "HeadLease" Then
        If Len(txtSearch1.text) > 0 Then
              txtSearch2.text = ""
              Filter = " UnitNumber LIKE '" + UCase(txtSearch1.text) + "*'"
           End If
        
           If Len(txtSearch2.text) > 0 Then
              txtSearch1.text = ""
              Filter = " UnitName LIKE '" + UCase(txtSearch2.text) + "*'"
           End If
        Call LoadflxUnits(Filter)
        Exit Sub
   End If
   If szSel = "Client" Or szSel = "Property" Then
        j = 1
   Else
        j = 0
   End If
   If Len(txtSearch1.text) > 0 Then
        txtSearch2.text = ""
   End If
  
   For i = flxSupplier(0).Rows - 1 To 1 Step -1
        flxSupplier(0).RowHeight(i) = 240
        If InStr(1, UCase(flxSupplier(0).TextMatrix(i, j)), UCase(txtSearch1.text), vbTextCompare) = 0 Then
              flxSupplier(0).RowHeight(i) = 0
        End If
        If flxSupplier(0).RowHeight(i) = 240 Then
              flxSupplier(0).row = i
        End If
   Next i
   
End Sub

Private Sub txtSearch2_Change()
'   Dim i As Integer
'
'   If Len(txtSearch2.text) > 0 Then
'      txtSearch1.text = ""
'   End If
'
'   For i = 1 To flxSupplier(0).Rows - 1
'      flxSupplier(0).RowHeight(i) = 240
'      If UCase(Left(flxSupplier(0).TextMatrix(i, 1), Len(txtSearch2.text))) <> UCase(txtSearch2.text) Then
'         flxSupplier(0).RowHeight(i) = 0
'      End If
'   Next i
  'Updated by anol 10 Dec 2015
   Dim i As Integer
   Dim Filter As String
'   Dim j As Integer
'   If szSel = "Client" Or szSel = "Property" Then
'        j = 1
'   Else
'        j = 0
'   End If
   If szSel = "Unit" Or szSel = "HeadLease" Then
        If Len(txtSearch1.text) > 0 Then
              txtSearch2.text = ""
              Filter = " UnitNumber LIKE '%" + UCase(txtSearch1.text) + "*'"
           End If
        
           If Len(txtSearch2.text) > 0 Then
              txtSearch1.text = ""
              Filter = " UnitName LIKE '%" + UCase(txtSearch2.text) + "*'"
           End If
        Call LoadflxUnits(Filter)
        Exit Sub
   End If
   If Len(txtSearch2.text) > 0 Then
        txtSearch1.text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
        flxSupplier(0).RowHeight(i) = 240
        If InStr(1, UCase(flxSupplier(0).TextMatrix(i, 2)), UCase(txtSearch2.text), vbTextCompare) = 0 Then
            flxSupplier(0).RowHeight(i) = 0
        End If
        If flxSupplier(0).RowHeight(i) = 240 Then
            flxSupplier(0).row = i
        End If
   Next i
End Sub

Private Sub txtSearch2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxSupplier(0).SetFocus
    End If
End Sub
Private Sub cmdClientList_Click()
'    fraList(1).Enabled = False
'    LoadflxClient
'
'   'tabTenant.Enabled = False 'it is already false by other lessee grid
'   txtSearch1.Visible = True
'   txtSearch2.Visible = True
'
'   txtSearch1.text = ""
'   txtSearch2.text = ""
'
'   'fraList(0).Width = 5115
'   'Picture1.Width = 5815
'   'cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
'   'Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
'  ' flxSupplier(0).Width = 4695
'   fraList(0).Left = fraList(1).Left + 500 'tabTenant.Left + txtDNC(1).Left
'   fraList(0).Top = fraList(1).Top + 200 'tabTenant.Top + txtDNC(1).Top
'   fraList(0).Visible = True
'   fraList(0).ZOrder 0
'   txtSearch1.SetFocus
    szSel = "Client"
    fraList(1).Enabled = False
    'tabTenant.Enabled = False
    Call LoadClient
    Frame5.Top = fraList(1).Top + 200 'tabTenant.Top + txtDNC(1).Top
    Frame5.Left = fraList(1).Left + 500 'tabTenant.Left + txtDNC(1).Left
    Frame5.Visible = True
    fraList(1).Enabled = False
    tabLease.Enabled = False
    Frame5.ZOrder 0
End Sub
Private Sub LoadClient()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   'TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
'        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Client ID"
   lblClientID(1).Caption = "Client Name"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
'   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"
   
   rsFundMatrix.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   If rsFundMatrix("isfundAssign").Value = False Then
'        iSel = 0
'        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
'   Else
'        iSel = 1
'        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
'                szPropertySelection1 & "' and ClientID='" & szClientID & "' and isDeleted=false"
'   End If
'   rsFundMatrix.Close
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
'        If iSel = 0 Then
            ShowMsgInTaskBar "CLIENT has not been setup for this company.", , "N"
'         Else
'            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
'         End If
      flxClientList.Clear
      flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("CLIENTID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("CLIENTNAME").Value
                    flxClientList.TextMatrix(rRow, 3) = "" 'rstRec.Fields.Item("FundID").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub txtSearchClientID_Change()
       
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240

      If InStr(1, UCase(flxClientList.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        FocusControl txtSearchClientName
    End If
End Sub



Private Sub txtSearchClientName_Change()
       'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
        flxClientList.RowHeight(i) = 240
       
        If InStr(1, UCase(flxClientList.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
        End If
       
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
End Sub
Private Sub LoadflxClient()
   flxSupplier(0).Clear
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
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0


   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

    szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim iRows As Integer
   flxSupplier(0).Rows = 2
   iRows = 1
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = "ALL"
      flxSupplier(0).TextMatrix(iRows, 2) = "ALL"
      flxSupplier(0).AddItem ""
   iRows = 2
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("CLIENTID").Value
      flxSupplier(0).TextMatrix(iRows, 2) = adoRST.Fields.Item("CLIENTNAME").Value
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend
 
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
  
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub


Private Sub cboBRChargingMth_Click()
   If cboBRChargingMth.text = "Global" Then
      If txtRCFund.text = "" Then
         MsgBox "Please select the Fund before choosing the Global charging method.", vbCritical + vbOKOnly, "Fund"
         Exit Sub
      End If
      BRGlobal
      txtBRChargingFigure.text = txtTotalRentYear.text
      txtBRChargingFigure.Locked = True
   Else
      txtBRChargingFigure.text = ""
      txtTotalRentYear.text = "0.00"
      txtRentDueEachPeriod.text = "0.00"
      txtBRChargingFigure.Locked = False
   End If
End Sub

Private Sub cboBRChargingMth_GotFocus()
   If txtRentStartDate.Enabled Then Exit Sub

   If txtFreqBR.text = "" Then
      MsgBox "Please select the Frequency before choosing the Charging Method.", vbCritical + vbOKOnly, "Charging Method"
      FocusControl cmdFreqBR
      Exit Sub
   End If
End Sub

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

Private Sub cboBreakClause_Click()
   If cboBreakClause.text = "Yes" Then
      If txtLeaseEndDate.ForeColor = vbRed Then
         MsgBox "This lease has expired, then you cannot add new break.", vbCritical + vbOKOnly, "Lease Expired"
         Exit Sub
      End If

      Frame1(0).Enabled = True
      txtBreakDate.SetFocus
   Else
      Frame1(0).Enabled = False
   End If
End Sub

Private Sub cboBreakClause_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboClientList_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

'*************************************************************************************************
'  Listing all properties according to selected client
'   If txtClientList.Tag <> "ALL" Then
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'              "ORDER BY PropertyID;"
'   Else
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "ORDER BY PropertyID;"
'   End If
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboPropertyList.Column() = Data()
'   cboPropertyList.ListIndex = 0
'*************************************************************************************************
   ConfigFlxLeaseList
   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
      LoadFlxLeaseList "", "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
      LoadFlxLeaseList "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
      LoadFlxLeaseList "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
      LoadFlxLeaseList "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
'**************************************************************************************************

'NoRes:
'   adoRst.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
   If adoRST.State = 1 Then
        adoRST.Close
   End If
   If adoConn.State = 1 Then
        adoConn.Close
   End If
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cboFreqBR_Click()
   If txtRentStartDate.text <> "" And txtNextDueDate.text = "" Then NextDueDate txtFreqBR, txtRentStartDate, txtNextDueDate, txtBRDemandType.Tag
End Sub

Private Sub cboFreqBR_LostFocus()
    'issue 761 done by anol 20190429
    If txtFreqBR.text = "" Then Exit Sub
    Call chargingMonthValidateBRFreq
    If txtRentStartDate.text <> "" And Trim(txtNextDueDate.text) = "" Then
        'Do not ask any ques just put the calculated value
         NextDueDate txtFreqBR, txtRentStartDate, txtNextDueDate, txtBRDemandType.Tag
    ElseIf txtRentStartDate.text <> "" And Trim(txtNextDueDate.text) <> "" Then
         txtComparenextDueDate1 = txtNextDueDate.text
          NextDueDate txtFreqBR, txtRentStartDate, txtComparenextDueDate1, txtBRDemandType.Tag
          If txtComparenextDueDate1 <> txtNextDueDate.text Then
            If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
                  txtNextDueDate.text = txtComparenextDueDate1
            End If
         End If
        
    End If
    'If txtRentStartDate.text <> "" And Trim(txtNextDueDate.text) = "" Then NextDueDate cboFreqBR, txtRentStartDate, txtNextDueDate, txtBRDemandType.Tag
End Sub

Private Sub cboFreqSC_GotFocus()
   If cmdSCNew.Enabled Then Exit Sub

   If txtPayableFrom.text = "" Then
      MsgBox "Please enter a start date before selecting a frequency.", vbInformation + vbOKOnly, "Start date required"
      txtPayableFrom.SetFocus
   End If
End Sub

Private Sub cboFreqSC_LostFocus()
    
   If txtFreqSC.text = "" Then Exit Sub
    
   If txtUnitNumber.text = "" Then
       MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
       Exit Sub
   End If
   Call chargingMonthValidateSCFreq
'fixed by anol 15 Nov 2015
'issue 571

'then issue 561 reverse requirement 20180405
  ' If txtPayableFrom.text <> "" And txtSCNextDueDt.text = "" Then NextDueDate cboFreqSC, txtPayableFrom, txtSCNextDueDt, txtSCDemandType.tag
  
   'issue 761 done by anol 20190429
    If txtPayableFrom.text <> "" And Trim(txtSCNextDueDt.text) = "" Then
        'Do not ask any ques just put the calculated value
         NextDueDate txtFreqSC, txtPayableFrom, txtSCNextDueDt, txtSCDemandType.Tag
    ElseIf txtPayableFrom.text <> "" And Trim(txtSCNextDueDt.text) <> "" Then
         txtComparenextDueDate1 = txtSCNextDueDt.text
          NextDueDate txtFreqSC, txtPayableFrom, txtComparenextDueDate1, txtSCDemandType.Tag
          If txtComparenextDueDate1 <> txtSCNextDueDt.text Then
            If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
                  txtSCNextDueDt = txtComparenextDueDate1
            End If
         End If
        
    End If
   
  
End Sub



Private Sub cboInsDept_LostFocus()
   On Error GoTo ErrHanlder

   If cmdIncNew.Enabled Then Exit Sub

   If txtICFundName.text = "" Then
        MsgBox "Please select a valid fund from the list.", vbCritical + vbOKOnly, "Wrong Fund"
        txtICFundCode.text = ""
        txtICFundCode.Tag = ""
        txtICFundName.text = ""
        FocusControl cmdICFundCode
   End If
   Exit Sub

ErrHanlder:
        txtICFundCode.text = ""
        txtICFundCode.Tag = ""
        txtICFundName.text = ""
End Sub

Private Sub cboInsFreq_GotFocus()
   If cmdIncNew.Enabled Then Exit Sub

   If txtInsStartDate.text = "" Then
      MsgBox "You must enter the Insurance Start Date before enter frequency.", vbInformation + vbOKOnly, "Insurance start date missing"
      txtInsStartDate.SetFocus
   End If
End Sub

Private Sub cboInsFreq_LostFocus()
   If cmdIncNew.Enabled Then Exit Sub

   If txtFreqIC.text = "" Then Exit Sub
   Call chargingMonthValidateInsFreq
   'If txtInsNextDueDate.text = "" And txtInsStartDate.text <> "" Then NextDueDate cboInsFreq, txtInsStartDate, txtInsNextDueDate, txtInsDemandType.tag
   'txtComparenextDueDate1 shall hold calculated next due date value
    If txtInsNextDueDate.text = "" And Trim(txtInsStartDate.text) <> "" Then
        'Do not ask any ques just put the calculated value
         NextDueDate txtFreqIC, txtInsStartDate, txtInsNextDueDate, txtInsDemandType.Tag
    ElseIf txtInsNextDueDate.text <> "" And Trim(txtInsStartDate.text) <> "" Then
         txtComparenextDueDate1 = txtInsNextDueDate.text
          NextDueDate txtFreqIC, txtInsStartDate, txtComparenextDueDate1, txtInsDemandType.Tag
          If txtComparenextDueDate1 <> txtInsNextDueDate.text Then
            If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
                  txtInsNextDueDate.text = txtComparenextDueDate1
            End If
         End If
    End If
    
    
End Sub

Private Sub cboIntChargeDept_LostFocus()
   On Error GoTo ErrHanlder

   If IsNull(cboIntChargeDept.Value) Then
      MsgBox "Please select a valid fund from the drop down list.", vbCritical + vbOKOnly, "Wrong Fund"
      cboIntChargeDept.text = ""
      cboIntChargeDept.SetFocus
   End If
   Exit Sub

ErrHanlder:
   cboIntChargeDept.text = ""
End Sub

Private Sub cboIntCrgable_Click()
   Dim sIntRate As Single

   If cboIntCrgable.text = "Yes" Then
      Conn2.Open getConnectionString

      sIntRate = GlobalIntRate(PROPERTY_ID, Conn2)
      If sIntRate <= 0 Then
         MsgBox "Please set up global base interest rate to proceed.", vbCritical + vbOKOnly, "Global Base interest rate"
         Conn2.Close
         cboIntCrgable.text = "No"
         Exit Sub
      End If

      Conn2.Close

      Frame3.Enabled = True
      cboIntChargeDept.SetFocus
   Else
      Frame3.Enabled = False
   End If
End Sub

Private Sub cboIntCrgable_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cboPropertyList_Click()
   ConfigFlxLeaseList
   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
      LoadFlxLeaseList "", "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
      LoadFlxLeaseList "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
      LoadFlxLeaseList "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
      LoadFlxLeaseList "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' ", _
                         "LeaseDetails.SageAccountNumber ASC"
End Sub

Private Sub cboRentChargeDept_LostFocus()
   If txtRentStartDate.Enabled Then Exit Sub

   On Error GoTo ErrHanlder

   If (txtRCFund.text) = "" Then
      MsgBox "Please select a valid fund from the drop down list.", vbCritical + vbOKOnly, "Wrong Fund"
      txtRCFund.text = ""
      cmdRCFund.SetFocus
   End If
   Exit Sub

ErrHanlder:
    txtRCFund.text = ""
End Sub

Private Sub cboSCChargingMth_Click()
   If cboSCChargingMth.text = "Global" Then
      If txtSCFundName.text = "" Then
         MsgBox "Please select the Fund before choosing the Global charging method.", vbCritical + vbOKOnly, "Fund"
         Exit Sub
      End If
      SCGlobal
      txtChargingFigure.text = txtSCTotalAmount.text
      txtChargingFigure.Locked = True
   Else
      txtChargingFigure.text = ""
      txtSCTotalAmount.text = ""
      txtSCDueEachPeriod.text = ""
      txtChargingFigure.Locked = False
   End If
End Sub

Private Sub cboSCChargingMth_GotFocus()
   If cmdSCNew.Enabled Then Exit Sub
   
   If txtFreqSC.text = "" Then
      MsgBox "Please select the Frequency before choosing the Charging Method.", vbCritical + vbOKOnly, "Charging Method"
      FocusControl cmdFreqSC
      Exit Sub
   End If
End Sub

Private Sub cboSCDept_LostFocus()
   On Error GoTo ErrHanlder

   If cmdSCNew.Enabled Then Exit Sub

   If txtSCFundName.text = "" Then
      MsgBox "Please select a valid fund from the list.", vbCritical + vbOKOnly, "Wrong Fund"
      txtSCFundName.text = ""
      txtSCFundCode.text = ""
      txtSCFundCode.Tag = ""
      FocusControl cmdSCFund
   End If
   Exit Sub

ErrHanlder:
    txtSCFundName.text = ""
      txtSCFundCode.text = ""
      txtSCFundCode.Tag = ""
End Sub

Private Sub cboSchedule_Click()
   If cboSchedule.text = "MULTIPLE" Then
      frmMulSch.Show
      Me.Enabled = False
   End If
End Sub

'Private Sub cboTenant_Click()
'   txtTenant.text = cboTenant.text
'End Sub

'Private Sub cboTenant_GotFocus()
'   Const CB_SHOWDROPDOWN = &H14F
'   Dim Tmp
'   Tmp = CboShowDown(cboTenant.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
'End Sub

'Private Sub cboUnit_Click()
'   If txtTenant.text = "" Then
'      MsgBox "Please select a tenant.", vbCritical + vbOKOnly, "Lease"
'      cboUnit.text = ""
'      cmdtenants_Click
'      Exit Sub
'   End If
'
'   Dim Conn1 As New ADODB.Connection
'
'   Conn1.Open getConnectionString
'
'   szSQL = "SELECT ClientName, PropertyName, Property.PropertyID " & _
'             "FROM Client, Property, Units " & _
'             "WHERE Client.ClientID = Property.ClientID And " & _
'                 "Property.PropertyID = Units.PropertyID And " & _
'                 "Units.UnitNumber = '" & txtUnitNumber.text & "';"
'
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''Rst1.Close
''Exit Sub
'   txtClient.text = Rst1!ClientName
'   txtProperty.text = Rst1!PropertyName
'   txtUnitName(0).text = cboUnit.Column(1)
'   PROPERTY_ID = Rst1!propertyID
'
'   AllDemandType Conn1
'
'   Rst1.Close
'   Conn1.Close
'   Set Rst1 = Nothing
'   Set Conn1 = Nothing
''***************************************************************************
''CREATE THE LEASE ID AUTOMATICALLY
''***************************************************************************
'   If cmdAddNew.Visible Then Exit Sub
'
'   Dim szaTenant() As String ', szaUnit() As String
'
'   If txtTenant.text <> "" Then
'   'issue 540 Lease details form - Not working correctly
'   'Fixed by anol 22 Feb 2015
'      szaTenant = Split(txtTenant.text, " / ")
''   szaUnit = Split(cboUnit.text, " - ")
'      txtLeaseID.text = OnlyNumericString(szaTenant(0)) & OnlyNumericString(txtUnitNumber.text & Format(Now, "YYMMDDHHMMSS"))
'  End If
'
'   tabLease.Enabled = True
'End Sub

Private Sub chkHoldingOver_Change()
   If chkHoldingOver.Value Then
      chkOLED.Value = True
      chkOLED.Locked = True
   Else
      chkOLED.Locked = False
   End If
End Sub

Private Sub chkInsDes_Click()
   lblDefaultDescption(tabLease.Tab).Visible = Not chkInsDes.Value
   txtInsDesc.Visible = Not chkInsDes.Value
   If Not chkInsDes.Value Then
      txtInsDesc.SetFocus
      chkInsDes.Caption = "Default"
   Else
      chkInsDes.Caption = chkInsDes.Caption & " description"
'      chkInsDes.SetFocus
   End If
End Sub

Private Sub chkOLED_Change()
   If chkHoldingOver.Value Then Exit Sub
   txtLeaseEndDate.Locked = chkOLED.Value

   If chkOLED.Value = False And txtLeaseEndDate.text = "" And strLeaseId <> "" Then
      MsgBox "You must input the lease end date.", vbInformation + vbOKOnly, "Lease Override"
   End If
End Sub

Private Sub chkRentDes_Click()
   lblDefaultDescption(tabLease.Tab).Visible = Not chkRentDes.Value
   txtRentDesc.Visible = Not chkRentDes.Value
   If Not chkRentDes.Value Then
      txtRentDesc.SetFocus
      chkRentDes.Caption = "Default"
   Else
      chkRentDes.Caption = chkRentDes.Caption & " description"
   End If
End Sub

Private Sub chkSCDes_Click()
    On Error GoTo Err
   lblDefaultDescption(tabLease.Tab).Visible = Not chkSCDes.Value
   txtSCDesc.Visible = Not chkSCDes.Value
   If Not chkSCDes.Value Then
      txtSCDesc.SetFocus
      chkSCDes.Caption = "Default"
   Else
      chkSCDes.Caption = chkSCDes.Caption & " description"
      chkSCDes.Visible = True
   End If
   Exit Sub
Err:
End Sub

Private Sub chkSubLease_Click()
   If chkSubLease.Value = 1 Then
       cmdUnitNumber(1).Enabled = True
   Else
       cmdUnitNumber(1).Enabled = False
       txtUnitName(1).text = ""
   End If
End Sub

Private Function OpenedLeasePreviewForm() As Boolean
   Dim frm As Form

   OpenedLeasePreviewForm = False
   For Each frm In Forms
      If frm.Name = "frmLeaseViewSummary" Then
         OpenedLeasePreviewForm = True
         Exit For
      End If
   Next frm
End Function

Private Sub cmdAddNew_Click()
   If OpenedLeasePreviewForm Then
      MsgBox "Please close the Lease Preview form.", vbCritical + vbOKOnly, "Add new lease"
      Exit Sub
   End If
   
   If MsgBox("Do you want to add new lease?", vbQuestion + vbYesNo, "Lease - New") = vbNo Then Exit Sub
   chkExpLease.Enabled = False
   chkMultipleLH.Enabled = False
   cmdUnitNumber(0).Enabled = True
   
   tabLease.Enabled = False
   txtUnitNumber.text = ""
   Label16(12).Tag = "" 'clearing the lease status
   Label16(12).Caption = ""
   ConfigFlxRentCharges
   ConfigFlxSC
   ConfigFlxInsurance
   ControlsModeRentCharges DefaultMode
   ControlsModeServiceCharges DefaultMode
   ControlsModeInsuranceCharges DefaultMode

   ConfigFlxRentAnalysis
   ConfigAssignmentGrid

   Dim adoConn As New ADODB.Connection
   Dim adoRst1 As New ADODB.Recordset

   adoConn.Open getConnectionString

   adoRst1.Open "SELECT * FROM GlobalData", adoConn, adOpenStatic, adLockReadOnly

   If adoRst1.RecordCount = 0 Then
      MsgBox "You Need to Enter the Global Data before you can add a lease record.", vbOKOnly + vbInformation, "Global Data"
      adoRst1.Close

      Exit Sub
   Else
      adoRst1.Close
   End If

   Call EmptyBoxes
   Rem out by anol 20180511 below procedure
   'Call GetTenantsWithoutLease
   'Call GetUnitWithoutLease(adoConn)
   Call EnableBoxes
   Call AllDemandType(adoConn)

   FocusControl cmdTenants

   adoConn.Close
   Set adoConn = Nothing
   
   cmdAddNew.Visible = False
   cmdCopy.Visible = False
   cmdSaveNew.Visible = True
   cmdTerminate.Visible = False
   cmdDelete.Visible = False
   cmdEdit.Visible = False
   cmdCancelNew.Visible = True
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
   fraList(1).Visible = False
End Sub

'Private Sub GetUnitWithoutLease(adoConn As ADODB.Connection)
'   Dim i As Integer
'   'Dim szaData() As String
'
''   szSQL = "SELECT Units.UnitNumber, Units.UnitName " & _
''             "FROM Units " & _
''             "WHERE Units.UnitNumber NOT IN " & _
''               "(SELECT LeaseDetails.UnitNumber " & _
''                "FROM LeaseDetails " & _
''                "WHERE LeaseDetails.Status = TRUE)"
'    'Modified tardy SQL by anol 20180403
'  szSQL = "SELECT Units.UnitNumber, Units.UnitName FROM Units LEFT JOIN " & _
'        " (SELECT LeaseDetails.UnitNumber FROM LeaseDetails WHERE LeaseDetails.Status = TRUE) as X " & _
'        "ON X.UnitNumber=Units.UnitNumber where X.UnitNumber is null"
'
'   rst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
''   cboUnit.Clear
''
''   If Not Rst1.EOF Then
''      ReDim szaData(1, Rst1.RecordCount - 1) As String
''
''      If Not Rst1.EOF Then
''         While Not Rst1.EOF
''            szaData(0, i) = Rst1!UnitNumber
''            szaData(1, i) = Rst1!UnitNumber & " - " & Rst1!UnitName
''            i = i + 1
''            Rst1.MoveNext
''         Wend
''      End If
''      cboUnit.Column() = szaData()
''   End If
'
'   rst1.Close
'End Sub

Private Sub cmdAddNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MousePointer = vbArrow
End Sub

Private Sub cmdAssignmentCancel_Click()
   AssignmentButtonMode DefaultMode
End Sub

Private Sub cmdAssignmentEdit_Click()
ASSIGNMENT_NEW_ENTRY_ = False
AssignmentButtonMode EditMode

End Sub

Private Sub cmdAssignmentNew_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new Assignment.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If

    ASSIGNMENT_NEW_ENTRY_ = True
   'AssignmentButtonMode NewEntryMode
    cmdAssignmentNew.Enabled = False
    cmdAssignmentEdit.Enabled = False
    cmdAssignmentSave.Enabled = True
    cmdAssignmentCancel.Enabled = True

    gridAssignment.Enabled = False

    txtAssignment_Date.Locked = False
    
    
    
End Sub
Private Function ValidateSaveAssignment() As Boolean
    
End Function
Private Sub cmdAssignmentSave_Click()
If SaveAssignment Then
   ShowMsgInTaskBar "The assignment information saved successfully."
Else
   ShowMsgInTaskBar "Could not save assignment information", , "N"
End If
AssignmentButtonMode DefaultMode
End Sub

Private Sub cmdBreachCancel_Click()
   BreachButtonMode ExpensesMode
   cboBreachType.ListIndex = -1
   txtCommenceDate.text = ""
   txtInitiatedBy.text = ""
   txtDateReceived.text = ""
   txtReceivedBy.text = ""
   txtMemo2.text = ""
   chkResolved.Value = 0
   Breach_EDIT = 0
End Sub

Private Sub cmdBreachEdit_Click()
   BREACH_NEW_ENTRY_ = False
   BreachButtonMode EditMode
   Breach_EDIT = gridBreach.row
   FocusControl cboBreachType
End Sub

Private Sub cmdBreachNew_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new Breach.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If

   BREACH_NEW_ENTRY_ = True
   strBreachID = returnBreachID
   BreachButtonMode NewEntryMode
   cboBreachType.ListIndex = -1
   txtCommenceDate.text = ""
   txtInitiatedBy.text = ""
   txtDateReceived.text = ""
   txtReceivedBy.text = ""
   txtMemo2.text = ""
   chkResolved.Value = 0
   Breach_EDIT = 0
End Sub
Private Function returnBreachID() As Integer
   Dim rstSet As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   
   rstSet.Open "SELECT MAX(BreachID) AS TID FROM LeaseBreaches;", adoConn, adOpenStatic, adLockReadOnly
   returnBreachID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID + 1))
   rstSet.Close
   Set rstSet = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function
Private Sub cmdBreachSave_Click()
   If ValidateSaveBreaches = False Then
        Exit Sub
   End If
   If SaveBreaches Then
      ShowMsgInTaskBar "The breach information saved successfully."
      cboBreachType.text = ""
      txtCommenceDate.text = ""
      txtInitiatedBy.text = ""
      txtDateReceived.text = ""
      txtReceivedBy.text = ""
      txtMemo2.text = ""
      chkResolved.Value = 0
   Else
      ShowMsgInTaskBar "Could not save breach information", , "N"
   End If
   BreachButtonMode DefaultMode
End Sub

Private Sub cmdCancelEdit_Click()
    
   If cmdSaveRentCrg.Enabled Then
      MsgBox "Please save/cancel the Rent Charge first.", vbCritical + vbOKOnly, "Rent Charge"
      Exit Sub
   End If
   If cmdSCSave.Enabled Then
      MsgBox "Please save/cancel the Service Charge first.", vbCritical + vbOKOnly, "Service Charge"
      Exit Sub
   End If
   If cmdIncSave.Enabled Then
      MsgBox "Please save/cancel the Insurace Charge first.", vbCritical + vbOKOnly, "Insurance Charge"
      Exit Sub
   End If
   If cmdSaveRentAnalysis.Enabled Then
      MsgBox "Please save/cancel the Rent Review first.", vbCritical + vbOKOnly, "Rent Review Charge"
      Exit Sub
   End If
   chkMultipleLH.Enabled = True
   chkExpLease.Enabled = True
   Call EmptyBoxes
   Call SetAddNewMode
   Call DisableBoxes
   ConfigFlxRentCharges
   ConfigFlxSC
   ConfigFlxRentAnalysis
   ConfigGridBreach
   ControlsModeRentCharges DefaultMode
   ControlsModeServiceCharges DefaultMode
   tabLease.Tab = 0
   Label16(12).Caption = ""
   Label16(12).Tag = ""
End Sub


Private Sub cmdCancelNew_Click()
   If MsgBox("Do you want to cancel the adding new lease?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      chkMultipleLH.Enabled = True
      chkExpLease.Enabled = True
      COPY_LEASE = False
      Call EmptyBoxes
      Call SetAddNewMode
      Call DisableBoxes
      ConfigGridBreach
      ConfigFlxInsurance
      ConfigFlxLeaseList
      ConfigFlxRentAnalysis
      ConfigFlxRentCharges
      ConfigFlxSC
   End If
End Sub

Private Sub cmdCancelRentAnalysis_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
    bReviewLocked = False
   UnlockTextBoxes False
   RentReviewButtonMode DefaultMode
End Sub

Private Sub cmdCancelRentCrg_Click()
   If MsgBox("Are you sure you wish to cancel Rent Charge changes?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      ControlsModeRentCharges ExpensesMode
      RENTCHARGES_EDIT = 0
      chkRentDes.Visible = False
   End If
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub

   AddNewAttachmentInCombo cmbFiles, "Lease", strLeaseId

   ShowMsgInTaskBar "File has been saved successfully."
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdCopy_Click()
   If OpenedLeasePreviewForm Then
      MsgBox "Please close the Lease Preview form.", vbCritical + vbOKOnly, "Copy lease"
      Exit Sub
   End If
   
   frmMMain.MousePointer = vbArrow
   
   COPY_LEASE = True
   
   tabLease.Enabled = False
   cmdAddNew.Visible = False
   cmdTerminate.Visible = False
   cmdDelete.Visible = False
   cmdEdit.Visible = False
   cmdCancelNew.Visible = True
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
   cmdSaveNew.Visible = True
   cmdCopy.Visible = False
   fraList(1).Visible = False
   
   txtTenant.text = ""
   txtUnitNumber.text = ""
'   txtUnitName(0).Enabled = True
   cmdUnitNumber(0).Enabled = True
   txtUnitName(0).text = ""
   strLeaseId = ""
   txtClient.text = ""
   txtProperty.text = ""
   Label16(12).Tag = "" 'clearing the lease status
   Label16(12).Caption = ""
   txtLeaseStDt.text = ""
   txtLeaseEndDate.text = ""
   
   chkExpLease.Enabled = True
   txtClient.Enabled = False
   txtProperty.Enabled = False
   
   Dim adoConn As New ADODB.Connection
   Dim adoRst1 As New ADODB.Recordset

   adoConn.Open getConnectionString

   adoRst1.Open "SELECT * FROM GlobalData", adoConn, adOpenStatic, adLockReadOnly

   If adoRst1.RecordCount = 0 Then
      MsgBox "You Need to Enter the Global Data before you can add a lease record.", vbOKOnly + vbInformation, "Global Data"
      adoRst1.Close
      
      Exit Sub
   Else
      adoRst1.Close
''      'issue 594 1)
''      'user cannot go back to browsing state lease
''        Dim iRow As Integer
''        For iRow = 1 To flxRentCharges.Rows - 1
''            If flxRentCharges.TextMatrix(iRow, 13) <> "" Then
''            End If
''        Next iRow
''
''        For iRow = 1 To flxSC.Rows - 1
''            If flxSC.TextMatrix(iRow, 14) <> "" Then
''            End If
''
''        Next iRow
''
''        For iRow = 1 To flxIns.Rows - 1
''            If flxIns.TextMatrix(iRow, 0) <> "" Then
''            End If
''        Next iRow

   End If
    Rem out by anol 20180511
   'Call GetTenantsWithoutLease
'   Call GetUnitWithoutLease(adoConn)
   Call EnableBoxes
   
   adoConn.Close
   Set adoConn = Nothing
   cmdLease.Visible = False
   cmdLease.Enabled = False
   cmdTenants.Enabled = True
   cmdTenants.Visible = True
   'cboTenant.Enabled = True
   txtUnitNumber.Enabled = True
   
   fraList(1).Visible = False
End Sub

Private Sub cmdDelete_Click()
   If strLeaseId = "" Then
        MsgBox "Please Select a valid Tenant", vbInformation, "Warning!!"
        Exit Sub
   End If
   If MsgBox("Are you sure to delete the lease?", vbQuestion + vbYesNo, "Lease - Delete") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   adoConn.Execute "DELETE * FROM LeaseDetails WHERE LeaseID = '" & strLeaseId & "';"

   adoConn.Close
   Set adoConn = Nothing

   MsgBox "Lease has been deleted successfully", vbInformation + vbOKOnly, "Lease - Delete"

   Call EmptyBoxes
   Call SetAddNewMode
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub

   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), strLeaseId, "Lease"

   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdDelRentAnalysis_Click()
   If flxRentAnalysis.row = 0 Then Exit Sub
   If flxRentAnalysis.TextMatrix(flxRentAnalysis.row, RRID) = "" Then Exit Sub
   If MsgBox("Are you sure, you want to delete the record?", vbQuestion + vbYesNo, "Delete Record") = vbNo Then Exit Sub
   bReviewLocked = False
   DeleteRentReview
   flxRentAnalysis.Enabled = False
   
   If flxRentAnalysis.Rows = 2 Then flxRentAnalysis.Rows = 3
   
   If flxRentAnalysis.row < flxRentAnalysis.Rows - 1 Then
      Dim i As Integer
      
      For i = flxRentAnalysis.row + 1 To flxRentAnalysis.Rows - 1
            If Val(flxRentAnalysis.TextMatrix(i, 1)) > 0 Then 'conditaion line added by anol issue 533
                flxRentAnalysis.TextMatrix(i, 1) = flxRentAnalysis.TextMatrix(i, 1) - 1
            End If
      Next i
   End If
   
   flxRentAnalysis.RemoveItem flxRentAnalysis.row
   
   RentReviewButtonMode DefaultMode
   UnlockTextBoxes False
End Sub

Private Sub DeleteRentReview()
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset
   Dim szSQL   As String
   Dim i       As Integer

   adoConn.Open getConnectionString
   szSQL = "SELECT LeaseID FROM RentAnalysis WHERE ID = " & flxRentAnalysis.TextMatrix(flxRentAnalysis.row, RRID) & ";"
   adoRST.Open szSQL, adoConn
   szSQL = adoRST.Fields.Item(0).Value
   adoRST.Close

   adoConn.Execute "DELETE * FROM RentAnalysis WHERE ID = " & flxRentAnalysis.TextMatrix(flxRentAnalysis.row, RRID) & ";"
   
   szSQL = "SELECT LeaseID, COUNT(LeaseID) AS C, MAX(SerialNumber) AS M " & _
           "FROM RentAnalysis " & _
           "WHERE LeaseID = '" & szSQL & "' " & _
           "GROUP BY LeaseID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then 'this caluse added by anol 20171112 issue 533
     If adoRST.Fields.Item("C").Value < Val(adoRST.Fields.Item("M").Value) Then
      szSQL = "SELECT * FROM RentAnalysis WHERE LeaseID = '" & adoRST.Fields.Item("LeaseID").Value & "';"
      adoRST.Close
      
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
      
      i = 1
      While Not adoRST.EOF
         adoRST.Fields.Item("SerialNumber").Value = i
         adoRST.Update
         i = i + 1
         
         adoRST.MoveNext
      Wend
     End If
 End If
     
   adoRST.Close
   Set adoRST = Nothing
   
   adoConn.Close
   Set adoConn = Nothing

   If flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 9) = "YES" Then _
      ClearReminder flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 10)

   MsgBox "Record has been deleted successfully.", vbOKOnly + vbInformation, "Delete Rent review"
End Sub

Private Sub cmdDelRentCrg_Click()
   If cmdSaveNew.Visible Then Exit Sub
   'Primary key for LRentCharges table is RentCharges which is loading at the column number 13
   'Now what I need to do, direct delete from database without retreiving and load again
   
'   If cmdDelRentCrg.Caption = "&Delete Rent" Then
'      If MsgBox("Would you like to delete this Rent Charge?", vbQuestion + vbYesNo, "Rent Charge") = vbNo Then Exit Sub
'
'      flxRentCharges.TextMatrix(flxRentCharges.row, 15) = "DELETED"
'      flxRentCharges.RowHeight(flxRentCharges.row) = 0
'      'MsgBox "This rent charge has been marked for deletition. It will be permanently removed when you save this lease.", vbInformation + vbOKOnly, "Rent Charge"
'   Else
'      flxRentCharges.TextMatrix(flxRentCharges.row, 15) = ""
'      MsgBox "This rent charge has been retrieved.", vbInformation + vbOKOnly, "Rent Charge"
'   End If
    Dim adoConn1   As New ADODB.Connection
    If MsgBox("Would you like to delete this Rent Charge?", vbQuestion + vbYesNo, "Please Confirm to delete current Rent Charge") = vbYes Then
        adoConn1.Open getConnectionString
        adoConn1.Execute "Delete from LRentCharges where RentCharges='" & flxRentCharges.TextMatrix(flxRentCharges.row, 13) & "'"
        LoadFlxRentCharges adoConn1
        adoConn1.Close
        Set adoConn1 = Nothing
    End If
    
    
    
   ControlsModeRentCharges DefaultMode
End Sub

Private Sub cmdEdit_Click()
   If OpenedLeasePreviewForm Then
      MsgBox "Please close the Lease Preview form.", vbCritical + vbOKOnly, "Edit lease"
      Exit Sub
   End If
   
   If strLeaseId = "" Then
       MsgBox "You must select a Lease to edit.", vbOKOnly + vbCritical, "No Lease Selected"
       FocusControl cmdLease
       Exit Sub
   End If

'   If MsgBox("Do you want to edit the lease?", vbQuestion + vbYesNo, "Lease - Edit") = vbNo Then Exit Sub

   Dim szaTenant() As String
   Dim Conn1 As New ADODB.Connection
   Dim szText As String, iCboIndex As Integer
   ControlsModeRentCharges ExpensesMode
   ControlsModeServiceCharges ExpensesMode
   ControlsModeInsuranceCharges ExpensesMode
   BreachButtonMode ExpensesMode
   Call EnableBoxes
   Conn1.Open getConnectionString
'   LoadDept Conn1 'load  fund for all charges
   Conn1.Close
   cboIntCrgable.Enabled = False
   '******************
   'EnableBoxes method unlock the cboUnit combo,
   'but user will not able to change the unit number once the lease created
   
   cmdTenants.Enabled = False
   '********************
'   cboTenant.Enabled = False

   cmdAddNew.Visible = False
   cmdTerminate.Visible = False
   cmdDelete.Visible = False
   cmdEdit.Visible = False
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = True
   cmdSaveEdit.TabIndex = 25
   cmdCancelEdit.Visible = True
   cmdCancelEdit.TabIndex = 26

   If txtLeaseEndDate.ForeColor = vbRed Then
      SelTxtInCtrl txtLeaseEndDate
      txtLeaseEndDate.SetFocus
   End If
   
   'Added by anol 06 Nov 2014
   'issue 471 Note 718
   cmdSCDemandType.Enabled = False
   txtPayableFrom.Locked = True
   cmdFreqSC.Enabled = False
   txtSCNextDueDt.Locked = True
    '04 Feb 2015 Anol
   txtCapAmount.Locked = True
   cmdSCFund.Enabled = False
   cboSchedule.Locked = True
   cboSCChargingMth.Locked = True
   txtChargingFigure.Locked = True
   txtSCTotalAmount.Locked = True
   txtSCDueEachPeriod.Locked = True
   txtStopSC.Locked = True
   'added by anol 21 June 2016
   cmdInsDemandType.Enabled = False
   txtInsPercentage.Locked = True
   txtInsNextDueDate.Locked = True
   'End of addition
   chkOLED.Locked = False
   txtLeaseEndDate.Locked = False
   cmdUnitNumber(1).Enabled = True
   cmdRentReviewDemandType.Enabled = False
   'added by anol 21 July 2016
   CreateLink
   
End Sub
Private Sub CreateLink()
    If Left(UCase(Trim(Label16(12).Caption)), 10) = UCase("Terminated") Then
        Label16(12).FontUnderline = True
    Else
        Label16(12).FontUnderline = False
    End If
End Sub
Private Sub cmdEditRentAnalysis_Click()
   RENT_REVIEW_ADDNEW_MODE = False
   flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 0) = "X"
   UnlockTextBoxes True
   bReviewLocked = False
   If flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 12) = "YES" Then
        bReviewLocked = True
        'When review run in true then locked editing except comment
        txtRentReviewDate.Locked = True
        txtRentIncreaseDate.Locked = True
        txtRentIncreaseAmount.Locked = True
        txtSerial.Locked = True
        'cboRRDemandType.Locked = True
        cmdRentReviewDemandType.Enabled = False
        txtComments.Locked = False
        chkAlarm.Enabled = False
        chkRRStatus.Enabled = False
        'ShowMsgInTaskBar "Rent increase has been applied, You cannot edit the line"
   End If
   cmdEditRentAnalysis.Enabled = False
   flxRentAnalysis.Enabled = False
   RentReviewButtonMode EditMode
End Sub

Private Sub UnlockTextBoxes(bState As Boolean)
   txtRentReviewDate.Locked = Not bState
   txtRentIncreaseDate.Locked = Not bState
   txtRentIncreaseAmount.Locked = Not bState
   txtSerial.Locked = Not bState
   'cboRRDemandType.Locked = Not bState
   cmdRentReviewDemandType.Enabled = bState
   txtComments.Locked = Not bState
   chkAlarm.Enabled = bState
   chkRRStatus.Enabled = bState

   If Not bState Then
      txtRentReviewDate.text = ""
      txtRentIncreaseDate.text = ""
      txtRentIncreaseAmount.text = ""
      txtSerial.text = ""
      txtComments.text = ""
      'cboRRDemandType.text = ""
      

txtRRDemandType.Tag = ""
txtRRDemandType.text = ""
      chkAlarm.Value = 0
      chkRRStatus.Value = 0
   End If
End Sub

Private Sub cmdEditRentCrg_Click()
   ControlsModeRentCharges EditMode
   RENTCHARGES_EDIT = flxRentCharges.row
   chkRentDes.Visible = True
   FocusControl cmdBRDemandType
End Sub

'Private Sub cmdGridUnitLookup_Click1()
'   fraList(1).Visible = False
'
''Resolved by BOSL
''Issue No: 0000445.
''To avoid triggering the search boxes events more than once.
''Modified By: Asif. 26 Jul 2014
'
''   txtSearchTenant.text = ""
''   txtSearchName.text = ""
'
'End Sub



Private Sub cmdGridUnitLookup_Click(Index As Integer)
    If Index = 1 Then
        fraList(1).Visible = False
        fraList(1).Enabled = True
        'tabTenant.Enabled = True
        tabLease.Enabled = True
        Frame1(16).Enabled = True
    ElseIf Index = 0 Then
        Frame5.Visible = False
        Frame5.Enabled = True
        'tabTenant.Enabled = True
        tabLease.Enabled = True
        fraList(1).Enabled = True
        tabLease.Enabled = True
        If szSel = "RCFund" Then
            tabLease.Enabled = True
            Frame1(16).Enabled = True
            Frame1(8).Enabled = True
        End If
    End If
    
    
End Sub

Private Sub cmdIncCancel_Click()
   If MsgBox("Are you sure you wish to cancel Insurance Charge changes?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      ControlsModeInsuranceCharges ExpensesMode
      INSURANCECHARGES_EDIT = 0
   End If
End Sub

Private Sub cmdIncEdit_Click()
   If Not cmdIncNew.Enabled Then
      ControlsModeInsuranceCharges EditMode
      INSURANCECHARGES_EDIT = flxIns.row
      'added by anol 21 June 2016
        cmdInsDemandType.Enabled = True
        txtInsPercentage.Locked = False
        txtInsNextDueDate.Locked = False
        'End of addition
   End If
End Sub

Private Sub cmdIncNew_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new charge.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If
   Dim Conn As New ADODB.Connection
   Conn.Open getConnectionString
   AllDemandType Conn
   'LoadDept Conn
   Conn.Close
   Set Conn = Nothing
   ControlsModeInsuranceCharges NewEntryMode
   INSURANCECHARGES_EDIT = 0
   FocusControl cmdInsDemandType
End Sub
Private Function chargingMonthValidate() As Boolean 'insurance
        On Error GoTo XX
        Dim X
        X = cboIncCharMth.Column(1)
        chargingMonthValidate = True
        Exit Function
XX:
        MsgBox "Please select a valid charging method", vbInformation, "Warning"
        FocusControl cboIncCharMth
End Function
Private Function chargingMonthValidateSC() As Boolean
        On Error GoTo XX
        Dim X
        X = cboSCChargingMth.Column(1)
        chargingMonthValidateSC = True
        Exit Function 'chaging month correctly
XX:
        MsgBox "Please select a valid charging method", vbInformation, "Warning"
        FocusControl cboSCChargingMth
End Function
Private Function chargingMonthValidateBR() As Boolean
        On Error GoTo XX
        Dim X
        X = cboBRChargingMth.Column(1)
        chargingMonthValidateBR = True
        Exit Function 'chaging month correctly'
XX:
        MsgBox "Please select a valid charging method", vbInformation, "Warning"
        FocusControl cboBRChargingMth
End Function
Private Function chargingMonthValidateBRFreq() As Boolean
        On Error GoTo XX
        Dim X
        X = txtFreqBR.Tag
        If IsNull(txtFreqBR.Tag) Then
            MsgBox "Please select a valid Frequency", vbInformation, "Warning"
            txtFreqBR.text = ""
            FocusControl cmdFreqBR
            Exit Function
        End If
        chargingMonthValidateBRFreq = True
        Exit Function 'chaging month correctly'txtFreqBR.tag
XX:
        MsgBox "Please select a valid Frequency", vbInformation, "Warning"
        FocusControl cmdFreqBR
End Function
Private Function chargingMonthValidateSCFreq() As Boolean
        On Error GoTo XX
        Dim X
        X = txtFreqSC.Tag
        If txtFreqSC.Tag = "" Then
            MsgBox "Please select a valid Frequency", vbInformation, "Warning"
            txtFreqSC.text = ""
            FocusControl cmdFreqSC
            Exit Function
        End If
        chargingMonthValidateSCFreq = True
        Exit Function 'chaging month correctly'txtFreqBR.tag
XX:
        MsgBox "Please select a valid Frequency", vbInformation, "Warning"
        FocusControl cmdFreqSC
End Function
Private Function chargingMonthValidateInsFreq() As Boolean
        On Error GoTo XX
        Dim X
        X = txtFreqIC.Tag
        If IsNull(txtFreqIC.Tag) Then
            MsgBox "Please select a valid Frequency", vbInformation, "Warning"
            txtFreqIC.text = ""
            FocusControl cmdFreqIC
            Exit Function
        End If
        chargingMonthValidateInsFreq = True
        Exit Function 'chaging month correctly'txtFreqBR.tag
XX:
        MsgBox "Please select a valid Frequency", vbInformation, "Warning"
        FocusControl cmdFreqIC
End Function
Private Sub cmdIncSave_Click()
'   If MsgBox("Do you want to save now?", vbQuestion + vbYesNo, "Save") = vbYes Then
   If txtICFundName.text = "" Then
      MsgBox "You must select a Fund of insurance charge.", vbOKOnly + vbCritical, "SC - Department"
      tabLease.Tab = 8
      FocusControl cmdICFundCode
      Exit Sub
   End If
   
   If txtInsStartDate.text = "" Then
      MsgBox "You must enter a insurance Charge Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 8
      txtInsStartDate.SetFocus
      Exit Sub
   End If
    If Trim(txtInsPercentage.text) = "" Then
        txtInsPercentage.text = "0.00"
        Exit Sub
   End If
   
  
   If IsNull(txtFreqIC.Tag) Then
        MsgBox "Please select a valid frequency", vbInformation, "Warning"
        FocusControl cmdFreqIC
        Exit Sub
   End If
    If Trim(txtInsNextDueDate.text) = "" Then
      MsgBox "You must enter a insurance Next Due Date!", vbOKOnly + vbCritical, "Date Required"
      FocusControl txtInsNextDueDate
      Exit Sub
   End If
   If IsDate(txtInsNextDueDate.text) And IsDate(txtInsStartDate.text) Then
        If DateDiff("d", txtInsStartDate.text, txtInsNextDueDate.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the insurance start date", vbInformation, "Warning"
            txtInsNextDueDate.text = ""
            FocusControl txtInsNextDueDate
            Exit Sub
        End If
    End If
   If txtFreqIC.text = "" Then
      MsgBox "You must select a insurance Charge Frequency!", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 8
      FocusControl cmdFreqIC
      Exit Sub
   End If
   If chargingMonthValidate = False Then
        Exit Sub
   End If
   If chargingMonthValidateInsFreq = False Then
        Exit Sub
   End If
   
   If txtInsDemandType.text = "" Then
      MsgBox "You must choose Demand type from the dropdown menu.", vbCritical + vbOKOnly, "Data Required"
      tabLease.Tab = 8
      cmdInsDemandType.SetFocus
      Exit Sub
   End If
   If cboIncCharMth.text = "" Then
      MsgBox "You must choose paying method from the dropdown menu.", vbCritical + vbOKOnly, "Data Required"
      tabLease.Tab = 8
      cboIncCharMth.SetFocus
   End If
   If txtInsPercentage.text = "" Then
      MsgBox "You must enter a value for insurance Charge charging amount.!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 8
      txtInsPercentage.SetFocus
      Exit Sub
   End If
   If Not chkInsDes.Value And Trim(txtInsDesc.text) = "" Then
      MsgBox "You must enter description of the insurance.", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 8
      txtInsDesc.SetFocus
      Exit Sub
   End If

''*** Saving data in the grid.
'       Add a row at the bottom of the grid
   If INSURANCECHARGES_EDIT = 0 Then
      If flxIns.TextMatrix(flxIns.Rows - 1, 0) <> "" Then flxIns.AddItem ""

      flxIns.TextMatrix(flxIns.Rows - 1, 0) = UniqueID()
      flxIns.TextMatrix(flxIns.Rows - 1, 1) = txtInsDemandType.text
      flxIns.TextMatrix(flxIns.Rows - 1, 17) = CInt(txtICFundCode.Tag)
      flxIns.TextMatrix(flxIns.Rows - 1, 2) = txtInsStartDate.text
      flxIns.TextMatrix(flxIns.Rows - 1, 3) = txtFreqIC.text                            'Frequncy
      flxIns.TextMatrix(flxIns.Rows - 1, 11) = CInt(txtFreqIC.Tag)
      flxIns.TextMatrix(flxIns.Rows - 1, 4) = txtInsNextDueDate.text
      flxIns.TextMatrix(flxIns.Rows - 1, 5) = txtICFundCode.text
      flxIns.TextMatrix(flxIns.Rows - 1, 12) = txtInsDemandType.Tag
      flxIns.TextMatrix(flxIns.Rows - 1, 6) = cboIncCharMth.Column(1)
      flxIns.TextMatrix(flxIns.Rows - 1, 13) = cboIncCharMth.Column(0)
      flxIns.TextMatrix(flxIns.Rows - 1, 7) = txtInsPercentage.text
      flxIns.TextMatrix(flxIns.Rows - 1, 8) = txtTotalYearlyIns.text
      flxIns.TextMatrix(flxIns.Rows - 1, 9) = txtInsEachPeriod.text
      flxIns.TextMatrix(flxIns.Rows - 1, 14) = Trim(txtInsDesc.text)
      flxIns.TextMatrix(flxIns.Rows - 1, 16) = Trim(txtStopIC.text)
      flxIns.TextMatrix(flxIns.Rows - 1, 18) = txtICFundName.text
   Else
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 1) = txtInsDemandType.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 17) = CInt(txtICFundCode.Tag)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 2) = txtInsStartDate.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 3) = txtFreqIC.text                  'Frequncy
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 11) = CInt(txtFreqIC.Tag)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 4) = txtInsNextDueDate.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 5) = txtICFundCode.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 12) = txtInsDemandType.Tag
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 6) = cboIncCharMth.Column(1)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 13) = cboIncCharMth.Column(0)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 7) = txtInsPercentage.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 8) = txtTotalYearlyIns.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 9) = txtInsEachPeriod.text
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 14) = Trim(txtInsDesc.text)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 16) = Trim(txtStopIC.text)
      flxIns.TextMatrix(INSURANCECHARGES_EDIT, 18) = txtICFundName.text
   End If

   ControlsModeInsuranceCharges DefaultMode
   INSURANCECHARGES_EDIT = 0
   ShowMsgInTaskBar "The insurance charge grid has been updated."
'   End If
End Sub

Private Sub cmdInsDelete_Click()
   If cmdSaveNew.Visible Then Exit Sub
    'Primary key for LInsuranceCharges  table is InsCharges   which is loading at the column number 0
   'Now what I need to do, direct delete from database without retreiving and load again
   
'   If cmdInsDelete.Caption = "&Delete Ins." Then
'      If MsgBox("Would you like to delete this Insurance?", vbQuestion + vbYesNo, "Insurance Charge") = vbNo Then Exit Sub
'
'      flxIns.TextMatrix(flxIns.row, 15) = "DELETED"
'      flxIns.RowHeight(flxIns.row) = 0
'      'MsgBox "This insurance charge has been marked for deletition. It will be permanently removed when you save this lease.", vbInformation + vbOKOnly, "Insurance Charge"
'   Else
'      flxIns.TextMatrix(flxIns.row, 15) = ""
'      MsgBox "This insurance charge has been retrieved.", vbInformation + vbOKOnly, "Insurance Charge"
'   End If
    Dim adoConn1   As New ADODB.Connection
    If MsgBox("Would you like to delete this Insurance?", vbQuestion + vbYesNo, "Please Confirm to delete current Insurance Charge") = vbYes Then
        adoConn1.Open getConnectionString
        adoConn1.Execute "Delete from LInsuranceCharges where InsCharges='" & flxIns.TextMatrix(flxIns.row, 0) & "'"
        LoadFlxIns adoConn1
        adoConn1.Close
        Set adoConn1 = Nothing
    End If
    
   ControlsModeInsuranceCharges DefaultMode
End Sub

Private Sub cmdLease_Click()
  ' Call PrepareList
  Dim Conn As New ADODB.Connection
  
'  If frmMMain.Leasee4_LesseList_isUptoDate = False Then
'        Conn.Open getConnectionString
'        TenantAccountBalance Conn
'        Conn.Close
'        frmMMain.Leasee4_LesseList_isUptoDate = True
'  End If
  cmdPropertyList.Enabled = True
  cmdClientList.Enabled = True
   If strSessionPropertyID = "" Then
        txtPropertyList.Tag = "ALL"
        txtPropertyList.text = "ALL"
   Else
        txtPropertyList.Tag = strSessionPropertyID
        txtPropertyList.text = strSessionPropertyID
   End If
   'issue  402 the client/property selection should remain active until the user changes it or closes the lease details form'by anol 20170608
   If strSessionClientID = "" Then
        txtClientList.Tag = "ALL"
        txtClientList.text = "ALL"
   Else
        txtClientList.Tag = strSessionClientID
        txtClientList.text = strSessionClientID
   End If
   'added by anol 03 08 2016
   txtSearchTenant.text = ""
   txtSearchName.text = ""
   txtSearchUnitName.text = ""
   Text8.text = ""
   'below procedure loading lessee list directly from LeaseDetails table
   cboClientList_Click
   tabLease.Enabled = False
   Frame1(16).Enabled = False
   fraList(1).Top = Frame1(16).Top + txtTenant.Top + txtTenant.Height + 5
   fraList(1).Left = Frame1(16).Left + txtTenant.Left + 5
   fraList(1).Visible = True
   fraList(1).ZOrder 0
   FocusControl txtSearchTenant
End Sub

Private Sub PrepareList()
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
'   adoConn.Open getConnectionString
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientList.Column() = Data()
'   cboClientList.ListIndex = 0
'   adoRst.Close
''*************************************** PROPERTY COMBO ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "ORDER BY PropertyID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboPropertyList.Column() = Data()
'   cboPropertyList.ListIndex = 0
'
'NoRes:
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'
'    'Resolved by BOSL
'    'Issue No: 0000445. Commented the following two lines to avoid calling same functions twice.
'    'Modified By: Asif. 02 Aug 2014
'
''   ConfigFlxLeaseList
''   LoadFlxLeaseList "", "LeaseDetails.SageAccountNumber ASC"
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'
'   ConfigFlxLeaseList
'   LoadFlxLeaseList "", "LeaseDetails.SageAccountNumber ASC"
End Sub

Private Sub ConfigFlxLeaseList()
   Dim szHeader As String

   flxLeaseList.Clear
   flxLeaseList.Cols = 12
   flxLeaseList.RowHeight(0) = 0
   szHeader$ = "|<LeaseID|<Tenant ID|<Tenant Name|<Unit Name"
   flxLeaseList.FormatString = szHeader$
   flxLeaseList.ColWidth(0) = 80 'Label5(0).Left - flxLeaseList.Left   '240        Solid column
   flxLeaseList.ColWidth(1) = 0          'Lease ID
   flxLeaseList.ColWidth(2) = Label5(1).Left - Label5(0).Left - 60 '1400       'Client ID
   flxLeaseList.ColWidth(3) = 2600 'Label5(2).Left - Label5(1).Left + 30 '+ 200         'Client Name
   flxLeaseList.ColWidth(4) = 1800 'flxLeaseList.Left + flxLeaseList.Width - Label5(2).Left - 300 'Unit number
   flxLeaseList.ColWidth(5) = 1600           'Unit Name
   flxLeaseList.ColAlignment(5) = vbLeftJustify
   flxLeaseList.ColWidth(6) = 0          'Client Name
   flxLeaseList.ColWidth(7) = 0          'Property NAME
   flxLeaseList.ColWidth(8) = 0          'Property ID
   flxLeaseList.ColWidth(9) = 0          'Usage
   flxLeaseList.ColWidth(10) = 1200           'Balance
   flxLeaseList.ColWidth(11) = 0           'Extra
   flxLeaseList.Rows = 2
End Sub


'Resolved by BOSL
'Issue No: 0000445.
'The function generates the expression of matching string pattern by using SQL LIKE operation and
'uses the in-built Filter function of the ADODB recordset to filter the records that match with the
'expression and finally bind the filtered records to the grid.
'Modified By: Asif. 26 Jul 2014
Private Function FilterTenantsList(Filter As String) As String
   Debug.Print 2
   Dim iRow As Integer
   iRow = 1
   Dim szWhere As String
   Dim szOrderby As String
   Dim tempstr As String
   
   szOrderby = "LeaseDetails.SageAccountNumber ASC"
   
   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = ""
      
   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
      
   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' "
      
   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
                         
''   Dim Filter As String
''   'Wild card search has been implemented by anol
''   'issue 0000445: Searching issues found through out Prestige
''   'Date 22 Feb 2015
''   If Len(txtSearchTenant.text) > 0 Then
''      txtSearchName.text = ""
''      txtSearchUnitName.text = ""
''      tempstr = Replace(UCase(txtSearchTenant.text), "'", "''")
''      Filter = " SageAccountNumber LIKE '%" + tempstr + "*'"
''
''   End If
''
''   If Len(txtSearchName.text) > 0 Then
''      txtSearchTenant.text = ""
''      txtSearchUnitName.text = ""
''      tempstr = Replace(UCase(txtSearchName.text), "'", "''")
''      Filter = " CompanyName LIKE '%" + tempstr + "*'"
''   End If
''
''   If Len(txtSearchUnitName.text) > 0 Then
''      txtSearchTenant.text = ""
''      txtSearchName.text = ""
''      tempstr = Replace(UCase(txtSearchUnitName.text), "'", "''")
''      Filter = " UnitName LIKE '%" + tempstr + "*'"
''   End If
   
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
'    If chkExpLease.Value = 1 Then 'This line was added by anol 26 Jun as it was not showing newly added tenant
       
    If cmdSaveNew.Visible Then 'entering new rows in leasedetail table

        szSQL = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName,'' as UnitNumber, '' as UnitName " & _
                "FROM Tenants LEFT JOIN LeaseDetails ON Tenants.SageAccountNumber=LeaseDetails.SageAccountNumber " & _
                "WHERE LeaseDetails.SageAccountNumber is null  AND " & _
                "(Tenants.Comments IS NULL OR Tenants.Comments = '')  ORDER BY Tenants.SageAccountNumber"
    Else 'user is surfing lease detail and user can also edit lease
            
         szSQL = "SELECT LeaseID, LeaseDetails.SageAccountNumber as SageAccountNumber, " & _
               "Tenants.CompanyName as CompanyName, UnitName, LeaseDetails.UnitNumber, LeaseDetails.Usage, " & _
               "ClientName, PropertyName, Property.PropertyID " & _
               "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
               "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
               "LeaseDetails.Status = " & IIf(chkExpLease.Value = 0, "True", "False") & " And " & _
               "Units.PropertyId = Property.PropertyID And " & _
               "Property.ClientID = Client.ClientID AND " & _
               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
               "" & szWhere & " " & _
                "ORDER BY " & szOrderby & ""
      'added by anol 2023-07-06
        szSQL = "SELECT SQL1.LeaseID, SQL1.SageAccountNumber, SQL1.CompanyName, SQL1.UnitName, SQL1.UnitNumber, SQL1.Usage, SQL1.ClientName, SQL1.PropertyName, SQL1.PropertyID, round(SQL2.Amt,2) as amt" & _
            " FROM" & _
            " (" & _
            szSQL & _
            " ) AS SQL1" & _
            " LEFT JOIN" & _
            " (" & _
            "     SELECT SageAccountNumber, SUM(Switch(type=1, Amount, type=23, Amount, type=2, -Amount, type=3, -Amount, type=4, -Amount)) AS Amt" & _
            "     FROM tlbReceipt" & _
            "     GROUP BY SageAccountNumber" & _
            " ) AS SQL2" & _
            " ON SQL1.SageAccountNumber = SQL2.SageAccountNumber;"
    End If

   
            
   Dim adoRST As New ADODB.Recordset
   
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        adoRST.Filter = Filter
   End If
   
'   MsgBox adoRst.RecordCount
      
   flxLeaseList.Clear
   ConfigFlxLeaseList
   If adoRST.RecordCount = 0 Then
         flxLeaseList.Rows = 2
   Else
        flxLeaseList.Rows = adoRST.RecordCount + 1
   End If
   
'   Dim i As Integer, j As Integer
   
'   For i = 0 To adoRst.RecordCount - 1
'      For j = 0 To adoRst.Fields.count - 1
'         flxLeaseList.TextMatrix(i + 1, j) = IIf(IsNull(adoRst.Fields(j)), "", adoRst.Fields(j))
'      Next j
'      adoRst.MoveNext
'   Next i
   If cmdSaveNew.Visible Then
    'when we are creating new lease
        While Not adoRST.EOF
           flxLeaseList.TextMatrix(iRow, 1) = "" 'adoRst!SageAccountNumber & " / " & adoRst!CompanyName
           flxLeaseList.TextMatrix(iRow, 2) = adoRST!SageAccountNumber
           flxLeaseList.TextMatrix(iRow, 3) = adoRST!CompanyName
           flxLeaseList.TextMatrix(iRow, 4) = adoRST!UnitNumber '"" 'Unit Number
           flxLeaseList.TextMatrix(iRow, 5) = adoRST!UnitName '"" 'Unit Name
           
           
           flxLeaseList.TextMatrix(iRow, 6) = ""
           flxLeaseList.TextMatrix(iRow, 7) = ""
           flxLeaseList.TextMatrix(iRow, 8) = ""
           flxLeaseList.TextMatrix(iRow, 9) = ""
           adoRST.MoveNext
           iRow = iRow + 1
           If iRow = 11 Then
                  fraList(1).Visible = True
                  fraList(2).Visible = True
                  Label16(11).Visible = True
                  fraList(1).ZOrder 0
                  fraList(2).ZOrder 0
                  fraList(2).Refresh
                  Label16(11).Refresh
                  flxLeaseList.Refresh
            End If
        Wend
   
   Else
          'on existing lease
          While Not adoRST.EOF
              flxLeaseList.TextMatrix(iRow, 1) = adoRST!LeaseID
              flxLeaseList.TextMatrix(iRow, 2) = adoRST!SageAccountNumber
              flxLeaseList.TextMatrix(iRow, 3) = adoRST!CompanyName
              flxLeaseList.TextMatrix(iRow, 4) = adoRST!UnitNumber
              flxLeaseList.TextMatrix(iRow, 5) = adoRST!UnitName
              flxLeaseList.TextMatrix(iRow, 6) = adoRST!ClientName
              flxLeaseList.TextMatrix(iRow, 7) = adoRST!PropertyName
              flxLeaseList.TextMatrix(iRow, 8) = adoRST!propertyID
              flxLeaseList.TextMatrix(iRow, 9) = IIf(IsNull(adoRST!Usage), "", adoRST!Usage)
              flxLeaseList.TextMatrix(iRow, 10) = Format(IIf(IsNull(adoRST!amt), 0, adoRST!amt), "0.00")
              adoRST.MoveNext
        '      If Not adoRst.EOF Then flxLeaseList.AddItem ""
              iRow = iRow + 1
              If iRow = 11 Then
                      fraList(1).Visible = True
                      fraList(2).Visible = True
                      Label16(11).Visible = True
                      fraList(1).ZOrder 0
                      fraList(2).ZOrder 0
                      fraList(2).Refresh
                      Label16(11).Refresh
                      flxLeaseList.Refresh
              End If
           Wend
   End If
   If flxLeaseList.Rows > 1 Then
       flxLeaseList.row = 1
   End If
'   MsgBox iRow
'   SetControlStyle flxLeaseList
   Label16(11).Visible = False
   fraList(2).Visible = False
   
'   fraList(1).Visible = False
'   fraList(2).Visible = False
   adoRST.Close
   Set adoRST = Nothing
   
   'UpdateBalance

   adoConn.Close
   Set adoConn = Nothing

End Function
Private Sub UpdateBalance()
   Dim i As Integer, j As Integer
   
   For i = 1 To flxLeaseList.Rows - 1
      For j = 0 To UBound(szaTenantBalance, 2) - 1
         If flxLeaseList.TextMatrix(i, 2) = szaTenantBalance(0, j) Then
            'flxLeaseList.TextMatrix(i, 10) = Format(szaTenantBalance(1, j), "0.00")
            flxLeaseList.TextMatrix(i, 10) = RoundingNumber2(szaTenantBalance(1, j))
            Exit For
         End If
      Next j
      If j = UBound(szaTenantBalance, 2) Then flxLeaseList.TextMatrix(i, 10) = "0.00"
   Next i
End Sub
'Private Sub TenantAccountBalance(adoConn As ADODB.Connection)
'   Dim szSQL As String, i As Integer, iIndex As Integer
'   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset
'
'   szSQL = "SELECT COUNT(SageAccountNumber), 2 " & _
'           "From " & _
'            "(" & _
'             "SELECT tlbReceipt.SageAccountNumber  " & _
'             "From tlbReceipt " & _
'             "GROUP BY tlbReceipt.SageAccountNumber" & _
'            ");"
'   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRptDr.EOF Then
'      adoRptDr.Close
'      Set adoRptDr = Nothing
'      Exit Sub
'   End If
'
'   ReDim szaTenantBalance(1, adoRptDr.Fields.Item(0).Value) As String
'   adoRptDr.Close
'
'   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
'           "FROM tlbReceipt AS Rpt " & _
'           "WHERE Type = 1 OR Type = 23 " & _
'           "GROUP BY SageAccountNumber " & _
'           "ORDER BY SageAccountNumber;"
''Debug.Print szSQL
'   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   iIndex = 0
'   While Not adoRptDr.EOF
'      szaTenantBalance(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
''If adoRptDr.Fields.Item("SageAccountNumber").Value = "Payden01" Then
''MsgBox ""
''End If
'      szaTenantBalance(1, iIndex) = adoRptDr.Fields.Item("Dr").Value
'      iIndex = iIndex + 1
'      adoRptDr.MoveNext
'   Wend
'
'   adoRptDr.Close
'
'   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
'           "FROM tlbReceipt AS Rpt " & _
'           "WHERE Type <> 1 AND Type <> 23 " & _
'           "GROUP BY SageAccountNumber;"
''Debug.Print szSQL
'   adoRptCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRptCr.EOF
'      For i = 0 To iIndex - 1
'         If szaTenantBalance(0, i) = adoRptCr.Fields.Item("SageAccountNumber").Value Then
'            Exit For
'         End If
'      Next i
''If adoRptCr.Fields.Item("SageAccountNumber").Value = "Payden01" Then
''MsgBox ""
''End If
'      If i < iIndex Then
'         szaTenantBalance(1, i) = Val(szaTenantBalance(1, i)) - adoRptCr.Fields.Item("Cr").Value
'      Else
'         iIndex = iIndex + 1
'         szaTenantBalance(0, iIndex) = adoRptCr.Fields.Item("Cr").Value
'      End If
'      adoRptCr.MoveNext
'   Wend
'
'   adoRptCr.Close
'
'   Set adoRptDr = Nothing
'   Set adoRptCr = Nothing
'End Sub
Private Sub TenantAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber), 2 " & _
           "From " & _
            "(" & _
             "SELECT tlbReceipt.SageAccountNumber  " & _
             "From tlbReceipt " & _
             "GROUP BY tlbReceipt.SageAccountNumber" & _
            ");"
   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'will it collect distict or not?ans : this is disttinct because you are using group by
   If adoRptDr.EOF Then
      adoRptDr.Close
      Set adoRptDr = Nothing
      Exit Sub
   End If

   ReDim szaTenantBalance(1, adoRptDr.Fields.Item(0).Value) As String
   adoRptDr.Close

     szSQL = "SELECT tlbReceipt.SageAccountNumber  " & _
             "From tlbReceipt " & _
             "GROUP BY tlbReceipt.SageAccountNumber order by SageAccountNumber ;"
'Debug.Print szSQL
   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoRptDr.EOF
      szaTenantBalance(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
        'If adoRptDr.Fields.Item("SageAccountNumber").Value = "Payden01" Then
        'MsgBox ""
        'End If
      'szaTenantBalance(1, iIndex) = RoundingNumber(adoRptDr.Fields.Item("Dr").Value, 2)
      szaTenantBalance(1, iIndex) = 0
      iIndex = iIndex + 1
      adoRptDr.MoveNext
   Wend

   adoRptDr.Close

   szSQL = "SELECT SageAccountNumber, Type, SUM(Amount) AS Amt " & _
           "FROM tlbReceipt " & _
           "GROUP BY SageAccountNumber,Type;"
'Debug.Print szSQL
   adoRptCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRptCr.EOF
      For i = 0 To iIndex - 1
         If szaTenantBalance(0, i) = adoRptCr.Fields("SageAccountNumber").Value Then
            If adoRptCr.Fields.Item("Type").Value = 1 Or adoRptCr.Fields.Item("Type").Value = 23 Then
                 szaTenantBalance(1, i) = Val(szaTenantBalance(1, i)) + adoRptCr.Fields.Item("Amt").Value
            End If
            If adoRptCr.Fields.Item("Type").Value = 2 Or adoRptCr.Fields.Item("Type").Value = 3 Or adoRptCr.Fields.Item("Type").Value = 4 Then
                 szaTenantBalance(1, i) = Val(szaTenantBalance(1, i)) - adoRptCr.Fields.Item("Amt").Value
            End If
         End If
         
       Next
       adoRptCr.MoveNext
     
   Wend

   adoRptCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

Private Sub LoadFlxLeaseList(szWhere As String, szOrderby As String)
   Dim conLease As New ADODB.Connection
   Dim adoLease As New ADODB.Recordset
   Dim szSQL As String

   'On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conLease.Open getConnectionString
'AND IsNull(TerminateDate) has been added by anol 13 May 2015
'The bug was that it was showing termindated lease
'Removed later
   szSQL = "SELECT LeaseID, LeaseDetails.SageAccountNumber, " & _
               "Tenants.CompanyName, UnitName, LeaseDetails.UnitNumber, LeaseDetails.Usage, " & _
               "ClientName, PropertyName, Property.PropertyID " & _
           "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
               "LeaseDetails.Status = " & IIf(chkExpLease.Value = 0, "True", "False") & " And " & _
               "Units.PropertyId = Property.PropertyID And " & _
               "Property.ClientID = Client.ClientID AND " & _
               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
               "" & szWhere & " " & _
           "ORDER BY " & szOrderby & ""
'added by anol 2023-07-06
        szSQL = "SELECT SQL1.LeaseID, SQL1.SageAccountNumber, SQL1.CompanyName, SQL1.UnitName, SQL1.UnitNumber, SQL1.Usage, SQL1.ClientName, SQL1.PropertyName, SQL1.PropertyID, round(SQL2.Amt,2) as amt" & _
            " FROM" & _
            " (" & _
            szSQL & _
            " ) AS SQL1" & _
            " LEFT JOIN" & _
            " (" & _
            "     SELECT SageAccountNumber, SUM(Switch(type=1, Amount, type=23, Amount, type=2, -Amount, type=3, -Amount, type=4, -Amount)) AS Amt" & _
            "     FROM tlbReceipt" & _
            "     GROUP BY SageAccountNumber" & _
            " ) AS SQL2" & _
            " ON SQL1.SageAccountNumber = SQL2.SageAccountNumber;"


   adoLease.Open szSQL, conLease, adOpenStatic, adLockReadOnly

   If adoLease.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1
   
    'Resolved by BOSL
    'Issue No: 0000445.
    'Modified By: Asif. 02 Aug 2014

    flxLeaseList.Rows = adoLease.RecordCount + 1
'    flxLeaseList.RowHeight(1) = 0
''   flxLeaseList.Rows = 2
'    iRow = 2
   While Not adoLease.EOF
      flxLeaseList.TextMatrix(iRow, 1) = adoLease!LeaseID
      flxLeaseList.TextMatrix(iRow, 2) = adoLease!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 3) = adoLease!CompanyName
      flxLeaseList.TextMatrix(iRow, 4) = adoLease!UnitNumber
      flxLeaseList.TextMatrix(iRow, 5) = adoLease!UnitName
      flxLeaseList.TextMatrix(iRow, 6) = adoLease!ClientName
      flxLeaseList.TextMatrix(iRow, 7) = adoLease!PropertyName
      flxLeaseList.TextMatrix(iRow, 8) = adoLease!propertyID
      flxLeaseList.TextMatrix(iRow, 9) = IIf(IsNull(adoLease!Usage), "", adoLease!Usage)
      flxLeaseList.TextMatrix(iRow, 10) = Format(IIf(IsNull(adoLease!amt), 0, adoLease!amt), "0.00")
      adoLease.MoveNext
'      If Not adoLease.EOF Then flxLeaseList.AddItem ""
      iRow = iRow + 1
      If iRow = 11 Then
            fraList(1).Visible = True
            fraList(2).Visible = True
            Label16(11).Visible = True
            fraList(1).ZOrder 0
            fraList(2).ZOrder 0
            fraList(2).Refresh
            Label16(11).Refresh
            fraList(1).Refresh
            flxLeaseList.Refresh
      End If
      
   Wend
   Label16(11).Visible = False
   fraList(2).Visible = False
'   If frmMMain.Leasee4_LesseList_isUptoDate = False Then
'        TenantAccountBalance conLease
'        frmMMain.Leasee4_LesseList_isUptoDate = True
'   End If
'   UpdateBalance
NoRes:
   adoLease.Close
   conLease.Close
   Set adoLease = Nothing
   Set conLease = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
   
   conLease.Close
   Set adoLease = Nothing
   Set conLease = Nothing
End Sub

Private Sub cmdLease_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraList(1).Visible = True Then
      fraList(1).Visible = False
   End If
End Sub

Private Sub cmdLeaseType_Click()
   frmSecondaryCode.PRIMARY_CODE_SHOW = "LTYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   LoadType "LTYP", cboType
End Sub

Private Sub cmdNewRentAnalysis_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new review.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If

   RENT_REVIEW_ADDNEW_MODE = True
   bReviewLocked = False
   UnlockTextBoxes True
   cmdNewRentAnalysis.Enabled = False
   flxRentAnalysis.Enabled = False
   RentReviewButtonMode NewEntryMode
   SetRentReviewSlNo txtSerial
   'cboRRDemandType.SetFocus
   FocusControl cmdRentReviewDemandType
End Sub

Private Sub SetRentReviewSlNo(txtControl As TextBox)
   Dim adoConn As New ADODB.Connection
   Dim adoRST As ADODB.Recordset
   Dim szSQL As String

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ";"
   Set adoRST = New ADODB.Recordset

   szSQL = "SELECT ID " & _
           "FROM RENTANALYSIS " & _
           "WHERE LeaseID = '" & strLeaseId & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   txtControl.text = IIf(adoRST.EOF, 0, adoRST.RecordCount) + 1
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdNewRentCrg_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new charge.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If
   Dim Conn As New ADODB.Connection
   Conn.Open getConnectionString
   AllDemandType Conn
'   LoadDept Conn
   Conn.Close
   Set Conn = Nothing
   ControlsModeRentCharges NewEntryMode
   RENTCHARGES_EDIT = 0
   chkRentDes.Visible = True
   FocusControl cmdBRDemandType
End Sub

'Private Sub cmdNDD_OK_Click()
'  If (tabLease.Tab = 1) Then
'      txtNextDueDate.text = txtNDD.text
'   ElseIf (tabLease.Tab = 4) Then
'      txtSCNextDueDt.text = txtNDD.text
'   ElseIf (tabLease.Tab = 8) Then
'      txtInsNextDueDate.text = txtNDD.text
'   End If
'   Frame4(0).Visible = False
'   tabLease.Enabled = True
'   txtNDD.text = ""
'End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdSaveRentAnalysis_Click()
   Dim lID As Long, i As Integer, bFound As Boolean
   Dim Conn1 As New ADODB.Connection
   bReviewLocked = False
   If cmdEditRentAnalysis.Enabled = True And cmdNewRentAnalysis.Enabled = True Then Exit Sub

   If txtSerial.text = "" Then
      MsgBox "Please enter serial number.", vbCritical + vbOKOnly, "Rent Analysis"
      txtSerial.SetFocus
      Exit Sub
   End If
   If txtRRDemandType.text = "" Or txtRRDemandType.Tag = "" Then
      MsgBox "Please enter the demand type.", vbCritical + vbOKOnly, "Rent Analysis"
      FocusControl cmdRentReviewDemandType
      Exit Sub
   End If
   If txtRentIncreaseDate.text = "" And txtRentIncreaseAmount.text = "" Then
      If txtRentReviewDate.text = "" Then
         MsgBox "Please enter rent review date.", vbCritical + vbOKOnly, "Rent Analysis"
         txtRentReviewDate.SetFocus
         Exit Sub
      End If
   Else
      If txtRentIncreaseDate.text = "" Then
         MsgBox "Please enter the Rent Increase/Decrease Date.", vbCritical + vbOKOnly, "Rent Analysis"
         txtRentIncreaseDate.SetFocus
         Exit Sub
      End If
'      If Val(txtRentIncreaseAmount.text) = 0 Then
'         MsgBox "Please enter the Rent Increase/Decrease Amount.", vbCritical + vbOKOnly, "Rent Analysis"
'         txtRentIncreaseAmount.SetFocus
'         Exit Sub
'      End If
   End If
'   szFlxHeader$ = "|<Serial|<DemandType|<RRDemandType|<RentReviewDate|<Comments|<RentIncreaseDate" & _
'                  "|>RentIncreaseAmount|ID|<Alarm|<AlarmID|<RRStatus"

   If RENT_REVIEW_ADDNEW_MODE Then                 'Check for double entry
      bFound = False
      For i = 1 To flxRentAnalysis.Rows - 1
         If txtRRDemandType.Tag = flxRentAnalysis.TextMatrix(i, 3) Then
            If flxRentAnalysis.TextMatrix(i, 4) <> "" Then
               If txtRentReviewDate.text = flxRentAnalysis.TextMatrix(i, 4) Then
                  bFound = True
                  Exit For
               End If
            Else
               If txtRentIncreaseDate.text = flxRentAnalysis.TextMatrix(i, 6) And _
                  txtRentIncreaseAmount.text = flxRentAnalysis.TextMatrix(i, 7) Then
                  bFound = True
                  Exit For
               End If
            End If
         End If
      Next i
      If bFound Then
         MsgBox "There is a same review exits. Duplicate entry will not be saved.", vbCritical + vbOKOnly, "Rent Analysis"
         Exit Sub
      End If
'---------------------------------------------------------------------------------------------
      bFound = False
      For i = 1 To flxRentAnalysis.Rows - 1
         If txtRentIncreaseDate.text <> "" Then
            If txtRentIncreaseDate.text = flxRentAnalysis.TextMatrix(i, 6) Then
               bFound = True
               Exit For
            End If
         End If
      Next i
      If bFound Then
         MsgBox "There is a same review exits. Duplicate entry will not be saved.", vbCritical + vbOKOnly, "Rent Analysis"
         Exit Sub
      End If
   End If

   Conn1.Open getConnectionString

   szSQL = "SELECT * FROM RentAnalysis"

   If Not RENT_REVIEW_ADDNEW_MODE Then                                           'EDIT MODE
      lID = FindXID
      szSQL = szSQL + " WHERE ID = " & lID & ";"

      Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

      flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 0) = "X"
   Else                                                                          'ADD NEW MODE
      szSQL = szSQL + ";"

      Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew

      Conn1.Execute "UPDATE LeaseDetails SET RentIncreaseAmount = 1 WHERE LeaseID = '" & strLeaseId & "';"
   End If

   Rst1!LeaseID = strLeaseId
   Rst1!SerialNumber = txtSerial.text
   Rst1!RRDemandType = CInt(txtRRDemandType.Tag)
   If txtRentReviewDate.text <> "" Then _
      Rst1!RentReviewDate = CDate(Format(txtRentReviewDate.text, "dd mmmm yyyy"))
   Rst1!Comments = txtComments.text
   If txtRentIncreaseDate.text <> "" Then
      Rst1!RentIncreaseDate = CDate(Format(txtRentIncreaseDate.text, "dd mmmm yyyy"))
   Else
      Rst1!RentIncreaseDate = Null
   End If
   Rst1!RentIncreaseAmount = IIf(txtRentIncreaseAmount.text = "", 0, CCur(Val(txtRentIncreaseAmount.text)))

   If chkAlarm.Value Then
      Rst1!Reminder_ID = NewReminder(Format(CDate(Rst1!RentReviewDate), "YYYYMMDD"), "010000", _
                         "Rent Review for " & txtProperty.text & " " & txtTenant.text, _
                         "Rent Review", txtSerial.text)
   Else
      If flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 9) = "YES" Then
         ClearReminder flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 10)
         Rst1!Reminder_ID = ""
      End If
   End If

   Rst1!RRStatus = IIf(chkRRStatus.Value, "Y", "N")
   'On add mode this status update is fine but in edit mode you should leave this field as it is
   'issue 537
   If RENT_REVIEW_ADDNEW_MODE Then
       Rst1!Status = "N"
   End If

   Rst1.Update

   Rst1.Close
   Conn1.Close

   Set Rst1 = Nothing
   Set Conn1 = Nothing

   cmdEditRentAnalysis.Enabled = True
   cmdNewRentAnalysis.Enabled = True
   flxRentAnalysis.Enabled = True
   UnlockTextBoxes False
   ConfigFlxRentAnalysis
   LoadFlxRentAnalysis
   RentReviewButtonMode DefaultMode

   ShowMsgInTaskBar "Data has been updated.", "Y", "P"
End Sub

Private Function FindXID() As Long
   Dim iRow As Integer
   
   For iRow = 1 To flxRentAnalysis.Rows - 1
      If flxRentAnalysis.TextMatrix(iRow, 0) = "X" Then
         FindXID = CLng(flxRentAnalysis.TextMatrix(iRow, RRID))
         Exit Function
      End If
   Next iRow
   FindXID = -1
End Function
Private Function ValidationOnDemandType() As Boolean
    'written by anol 2023.06.19
'We found that this problem is caused when the user creates a new lease by copying an existing lease that belongs to a different
'client or property from that of the unit of the new lease that is being created.
'To prevent this we need to check that unit of the new lease being created from the copy belongs to the same client and property.
'A message needs to be created on lost focus of selecting the unit to assign to the newly copied lease. This message should say
'“There is an inconsistency in the demand types on this lease. Please assign the correct demand types to this lease before saving.”
'“The unit selected does not match the client or the property of the lease being copied from.”

    Dim adoConn As New ADODB.Connection
    Dim rsDemandType As New ADODB.Recordset
    Dim szSQL As String
    Dim dictAllDemandType As New Dictionary
    Dim iRow As Integer
    Dim result
    adoConn.Open getConnectionString
    rsDemandType.Open "Select ID,Type from demandtypes D,Units U where D.PropertyID=U.PropertyID AND U.UnitNumber='" & txtUnitNumber.text & "' order by ID ", adoConn, adOpenStatic, adLockReadOnly
    While Not rsDemandType.EOF
        Debug.Print rsDemandType!Id
        'Note: you cannot add number as item or key. SO I converted them as  cstr =string
        dictAllDemandType.Add CStr(rsDemandType!Id), CStr(rsDemandType!Id) ' key is not important value is important. we are using values to be checked for exitst
        rsDemandType.MoveNext
    Wend
   ' adoconn.Close
   
  ' dictAllDemandType.removeAll
    rsDemandType.Close
     adoConn.Close
    For iRow = 1 To flxRentCharges.Rows - 1
        If flxRentCharges.TextMatrix(iRow, 0) <> "" Then
            result = dictAllDemandType.Exists(flxRentCharges.TextMatrix(iRow, 0)) 'exits checks if the value exits in the dictionary
            If Not result Then
                Exit Function
            End If
        End If
    Next iRow
    For iRow = 1 To flxSC.Rows - 1
        If flxSC.TextMatrix(iRow, 0) <> "" Then
            result = dictAllDemandType.Exists(flxSC.TextMatrix(iRow, 0))
            If Not result Then
                Exit Function
            End If
        End If
    Next iRow
    For iRow = 1 To flxIns.Rows - 1
        If flxIns.TextMatrix(iRow, 12) <> "" Then
            result = dictAllDemandType.Exists(flxIns.TextMatrix(iRow, 12))
            If Not result Then
                Exit Function
            End If
        End If
    Next iRow
        ValidationOnDemandType = True
End Function

Private Sub cmdSaveEdit_Click()
'   If MsgBox("Do you want to update the lease?", vbQuestion + vbYesNo, "Lease - Update") = vbNo Then Exit Sub
   Dim adoConn  As New ADODB.Connection
   
   If ValidationSaveLease = False Then
        Exit Sub
   End If
   If ValidationOnDemandType = False Then
        MsgBox " " & vbCrLf & _
                        "There is an inconsistency in the demand types on this lease. Please assign the correct demand types to this lease before saving."
        Exit Sub
   End If
   COPY_LEASE = False
   If SaveUpdateLease(False) Then ' if false then system will edit records
   'You will not clear leaseID unless User changes from grid because it will used for reloading the whole form information
                'ShowMsgInTaskBar "The lease record has been updated."
                MsgBox "The lease record has been updated.", vbInformation, "Lease information"
                chkExpLease.Enabled = True
                chkMultipleLH.Enabled = True
                
                'I am commenting those three subproc becuase I dont want to clear the data
                      ConfigFlxRentCharges
                      ConfigFlxSC
                      ConfigFlxInsurance
                ControlsModeRentCharges DefaultMode
                ControlsModeServiceCharges DefaultMode
                ControlsModeInsuranceCharges DefaultMode
                UnlockTextBoxes False
                ''   'added by anol 22 Feb 2015
                ''      If Conn2.State = 0 Then
                ''      Conn2.Open getConnectionString
                ''      End If
                ''      Call FillCbos(Conn2)
                ''      If Conn2.State = 1 Then
                ''            Conn2.Close
                ''       End If
                ''       'End of modification
                ConfigFlxRentAnalysis
                ConfigAssignmentGrid
                adoConn.Open getConnectionString
                GetRecord adoConn 'Could not load the lease information, lease detail ID incorrect!! this message is coming after save edit which is wrong
                If adoConn.State = 1 Then
                    adoConn.Close
                    Set adoConn = Nothing
                End If
                'adoConn.Open getConnectionString
                '      LoadingForm adoConn
        If Trim(strLeaseId) = "" Then
            'GetRecord adoConn ' this function loads form information based on txtLeaseID.text
            'AllDemandType adoConn
            'you need to clear everything here
            EmptyBoxes
        Else
            'Label16(12).Caption = ""
        End If
        Call DisableBoxes
'        adoConn.Close
'        Set adoConn = Nothing
   End If
End Sub

Private Sub cmdSaveNew_Click()
    If ValidationSaveLease = False Then
       Exit Sub
    End If
   
    If MsgBox("Do you want to save the lease?", vbQuestion + vbYesNo, "Lease - Save") = vbNo Then Exit Sub
   
    cmdTenants.Visible = False
    cmdLease.Visible = True
    cmdLease.Enabled = True
   
    If SaveUpdateLease(True) Then 'if the value is true then system will add NEW records
'         ConfigFlxRentCharges
'         ConfigureFlxSC
'         ConfigFlxInsurance
      ControlsModeRentCharges DefaultMode
      ControlsModeServiceCharges DefaultMode
      ControlsModeInsuranceCharges DefaultMode
      DisableBoxes
'         ConfigFlxRentAnalysis
'         SetAssignmentGrid
   End If

   If COPY_LEASE Then
         MsgBox "The new (Copied) lease record has been saved", vbOKOnly + vbInformation, "Saved"
         COPY_LEASE = False
   Else
         ShowMsgInTaskBar "The new lease record has been saved successfully."
   End If
      
   
   chkMultipleLH.Enabled = True
   chkExpLease.Enabled = True
End Sub
Private Function ValidationSaveLease() As Boolean
    If Trim(txtTenant.text) = "" Then
        MsgBox "Please Select a valid Tenant", vbInformation, "Warning!!"
        FocusControl cmdTenants
        Exit Function
   End If
   If Trim(txtUnitNumber.text) = "" Then
        MsgBox "Please Select a valid UnitNumber", vbInformation, "Warning!!"
        FocusControl cmdUnitNumber(0)
        Exit Function
   End If
   
'    If txtTenant.text = "" Then
'      MsgBox "You must select a lessee.", vbOKOnly + vbCritical, "Lessee"
'      cmdTenants.SetFocus
'      ValidationSaveLease = False
'      Exit Function
'   End If

   If txtLeaseStDt.text = "" Then
      MsgBox "You must enter a Lease Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 0
      FocusControl txtLeaseStDt
      ValidationSaveLease = False
      Exit Function
   End If

   If txtLeaseEndDate.text = "" And chkOLED.Value = False Then
      MsgBox "You must enter a Lease End Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 0
      txtLeaseEndDate.SetFocus
      ValidationSaveLease = False
      Exit Function
   End If

   If cmdSaveRentCrg.Enabled Then
      MsgBox "Please save/cancel your Rent Charge entries first.", vbCritical + vbOKOnly, "Rent Charge"
      tabLease.Tab = 1
      cmdSaveRentCrg.SetFocus
      Exit Function
   End If

   If cmdSCSave.Enabled Then
      MsgBox "Please save/cancel your Service Charge entries first.", vbCritical + vbOKOnly, "Service Charge"
      tabLease.Tab = 4
      cmdSCSave.SetFocus
      Exit Function
   End If

   If cmdIncSave.Enabled Then
      MsgBox "Please save/cancel your Insurance Charge entries first.", vbCritical + vbOKOnly, "Insurance Charge"
      tabLease.Tab = 8
      cmdIncSave.SetFocus
      Exit Function
   End If

   If cmdSaveRentAnalysis.Enabled Then
      MsgBox "Please save/cancel your Rent review first.", vbCritical + vbOKOnly, "Rent Review"
      tabLease.Tab = 2
      cmdSaveRentAnalysis.SetFocus
      Exit Function
   End If

   If cboIntCrgable.text = "Yes" Then
      If cboIntChargeDept.text = "" Then
         MsgBox "You must select a department for the interest charge!", vbOKOnly + vbCritical, "Interest Charge - Department"
         tabLease.Tab = 5
         cboIntChargeDept.SetFocus
         ValidationSaveLease = False
         Exit Function
      End If
      If cboIntDemandType.text = "" Then
         MsgBox "You must select a demand type for the interest charge!", vbOKOnly + vbCritical, "Interest Charge - Demand Type"
         tabLease.Tab = 5
         cboIntDemandType.SetFocus
         ValidationSaveLease = False
         Exit Function
      End If
      If optLSR.Value And (txtAdditionalIntRate.text = "" Or Val(txtAdditionalIntRate.text) <= 0) Then
         MsgBox "You must input lease specific additional interest rate.", vbOKOnly + vbCritical, "Interest Charge - Interest Rate"
         tabLease.Tab = 5
         txtAdditionalIntRate.SetFocus
         ValidationSaveLease = False
         Exit Function
      End If
      If optAutoIntCal.Value And (txtIntPayableAfterDays.text = "" Or Val(txtIntPayableAfterDays.text) < 1) Then
         MsgBox "You must input a positive integer number of days after interest will be calculated.", vbOKOnly + vbCritical, "Interest Charge - Interest Days"
         tabLease.Tab = 5
         txtIntPayableAfterDays.SetFocus
         ValidationSaveLease = False
         Exit Function
      End If
      If optManIntCal.Value Then
         If (txtAmtCrgIntOn.text = "" Or Val(txtAmtCrgIntOn.text) = 0) Then
            MsgBox "Please input specified amount to charge interest on.", vbCritical + vbOKOnly, "Interest Charge - Manual Calculation"
            tabLease.Tab = 5
            txtAmtCrgIntOn.SetFocus
            ValidationSaveLease = False
            Exit Function
         End If
         If txtNoIntDays.text = "" Or Val(txtNoIntDays.text) < 1 Then
            MsgBox "You must enter number of days interest will charge after!", vbOKOnly + vbCritical, "Interest Charge - Manual Calculation"
            tabLease.Tab = 5
            txtNoIntDays.SetFocus
            ValidationSaveLease = False
            Exit Function
         End If
      End If
   End If

   If txtUnitNumber.text = "" Then
       MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
       ValidationSaveLease = False
       Exit Function
   End If
   ValidationSaveLease = True
End Function
Private Function ReturnLeaseStatus(bSaveUpdate As Boolean) As Boolean
    If bSaveUpdate = False Then  'update lease mode
        If szLeaseStatus = False And chkOLED.Value = True Then
            If MsgBox("This lease is set to override the lease end date. Do you wish to reactivate this lease?", vbYesNo, "Please confirm!") = vbYes Then
                ReturnLeaseStatus = True
                'if it is yes, shall I delete termination date?
            Else
                ReturnLeaseStatus = False
            End If
        ElseIf szLeaseStatus = False And chkOLED.Value = False Then
             If DateDiff("d", Now, txtLeaseEndDate.text) < 0 And txtLeaseEndDate.ForeColor = vbBlack Then
                    ReturnLeaseStatus = False
             Else
                If MsgBox("This lease has not yet expired. Do you wish to reactivate this lease?", vbYesNo, "Please confirm!") = vbYes Then
                    ReturnLeaseStatus = True
                Else
                    ReturnLeaseStatus = False
                End If
             End If
        Else
            ReturnLeaseStatus = True
        End If
    Else 'It will set leasestatus true on add new mode
            ReturnLeaseStatus = True
    End If
    
End Function
'**** if the value of bSaveUpdate is true then system will add NEW records
'**** otherwise system will update existing record.
Private Function SaveUpdateLease(bSaveUpdate As Boolean) As Boolean
   Dim i As Integer, iKount As Integer, bUpdated As Boolean
   Dim szSplitLeaseID As String
   Dim szMsg As String

   
'   Dim szaUnit() As String
   Dim Conn1 As New ADODB.Connection

'   szaUnit = Split(cboUnit.text, " - ")

   'save the details to a new record
   Conn1.Open getConnectionString
   Conn1.BeginTrans
   GetGlobalDataForProperty txtUnitNumber.text, Conn1

   szSQL = "SELECT * FROM LeaseDetails " & _
             "WHERE LeaseID = '" & strLeaseId & "';"
   If Rst1.State = 1 Then
        Rst1.Close
   End If
   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
'Debug.Print szSQL
   Dim szaTemp() As String

   '****************************************************
   ' Adding data in the LeaseDetails Table
   '****************************************************
   Dim szaTenant() As String
'bSaveUpdate = False
   If bSaveUpdate Then 'if the value of bSaveUpdate is true then system will add NEW records
      If Not Rst1.EOF Or Not Rst1.BOF Then
          'MsgBox "Cannot save the lease. The lease reference already exist.", vbInformation, "Save Lease"
          'you need not clear txtLeaseId.text = "" 20190428 by anol
          strLeaseId = ""
          'txtLeaseID.SetFocus
          SaveUpdateLease = False
          Rst1.Close
          Set Rst1 = Nothing
          Exit Function
      End If

      Rst1.AddNew
      Rst1!LeaseID = strLeaseId
      'rst1!IncrementalID = Right(txtLeaseID.text, 12)
'      szaTenant = Split(txtTenant.text, " / ")
      ReDim szaTenant(2) As String
      szaTenant(0) = txtTenant.Tag 'sageaccount ID
      szaTenant(1) = txtTenant.text  'Tenant Name
      Rst1!CreatedDate = Now
      Rst1!CreatedBy = User
   Else
      ReDim szaTenant(2) As String
      'you  should save the previously clicked Lessee Other wise there will be a problem when user goto search window again
      szaTenant(0) = txtTenant.Tag 'flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sageaccount ID
      szaTenant(1) = txtTenant.text 'flxLeaseList.TextMatrix(flxLeaseList.row, 3) 'Tenant Name
   End If

   Rst1!SageAccountNumber = szaTenant(0)
   Rst1!CompanyName = szaTenant(1)
   'added by anol 04 Feb 2015
   Rst1!CapAmount = Val(txtCapAmount.text)
   'End of modification
   
  If chkSubLease.Value Then
   'Modified by anol 13 Feb 2015
   'issue 537
      If txtUnitName(1).text <> "" Then
         Rst1!HeadLease = txtUnitName(1).Tag
       Else
         Rst1!HeadLease = ""
      End If
  Else
      Rst1!HeadLease = ""
  End If
   Rst1!UnitNumber = txtUnitNumber.text
   Rst1!TYPEOFSTORE = cboType.text
   Rst1!Usage = IIf(cboUsage.text = "", "", cboUsage.text)

   If txtLeaseStDt.text <> "" Then Rst1!StartDate = CDate(Format(txtLeaseStDt.text, "dd mmmm yyyy"))
   If txtYearEnd.text <> "" Then Rst1!YearEnd = CDate(Format(txtYearEnd.text, "dd mmmm yyyy"))

   If Not chkOLED.Value Then
      Rst1!EndDate = CDate(Format(txtLeaseEndDate.text, "dd mmmm yyyy"))
      Rst1!OLED = CBool(False)
   Else
      If txtLeaseEndDate.text <> "" Then _
         Rst1!EndDate = CDate(Format(txtLeaseEndDate.text, "dd mmmm yyyy"))
      Rst1!OLED = CBool(True)
   End If
   Rst1!HoldingOver = chkHoldingOver.Value
   
   If chkGPrataDmd.Value = 0 Or Not chkGPrataDmd.Value Then
      Rst1!GPrataDmd = False
   Else
      Rst1!GPrataDmd = True
   End If

'   ****************************************************
'    Update Rent Payable if there any rent
'   ****************************************************
   For i = 1 To flxRentCharges.Rows - 1
      If flxRentCharges.TextMatrix(i, 15) = "" Then iKount = iKount + 1
   Next i

   If flxRentCharges.TextMatrix(1, 0) <> "" And iKount > 0 Then
      Rst1!BRPayable = "Y"
   Else
      Rst1!BRPayable = "N"
   End If

   i = 0
   iKount = 0
'   ****************************************************
'    Adding data from Service Charge tab
'   ****************************************************
   If flxSC.TextMatrix(1, 0) <> "" Then
      Rst1!SCPayable = "Y"
   Else
      Rst1!SCPayable = "N"
   End If

'   ****************************************************
'    Adding data from Interest Charge tab
'   ****************************************************
   If cboIntCrgable.text = "Yes" Then
      Rst1!InterestChargeable = "Y"
      Rst1!IntChargeDept = cboIntChargeDept.Column(0)
      Rst1!IntDemandType = CInt(IIf(cboIntDemandType.text <> "", cboIntDemandType.Column(0), 0))
      If optGIR.Value Then
         Rst1!AdditionalInterest = 0
      Else
         Rst1!AdditionalInterest = Val(txtAdditionalIntRate.text)
      End If

      If optAutoIntCal.Value Then      'Automatic
         Rst1!ServiceChargeDept = "AUTO"              'ServiceChargeDept-> this field is using for INTEREST
         Rst1!DaysAfterInterestPayable = CInt(txtIntPayableAfterDays.text)
         Rst1!InterestChargedOn = CCur(0)
         Rst1!SCTOLimit = 0                           'SCTOLimit-> this field is using for INTEREST
         Rst1!InterestAmount = 0
      Else                             'Manual
         Rst1!ServiceChargeDept = "MANU"
         Rst1!InterestChargedOn = CCur(txtAmtCrgIntOn.text)
         Rst1!SCTOLimit = CInt(txtNoIntDays.text)     'SCTOLimit-> this field is using for INTEREST
         Rst1!InterestAmount = txtInt2bChrg.text
      End If

      Rst1!Text1 = txtInterestDescription.text
   Else
      Rst1!InterestChargeable = "N"
      Rst1!InterestChargedOn = 0
      Rst1!AdditionalInterest = 0
   End If

'   ****************************************************
'    Adding data from Break Clause tab
'   ****************************************************
   If cboBreakClause.text = "Yes" Then
      Rst1!BreakClause = "Y"
   
      If txtBreakDate.text <> "" Then Rst1!BreakDate = CDate(Format(txtBreakDate.text, "dd mmmm yyyy"))
      If cboBreak.text <> "" Then Rst1!BreakType = cboBreak.text
   Else
      Rst1!BreakClause = "N"
   End If
   
'   ****************************************************
'    Adding data from Rent Review tab
'   ****************************************************
'   If txtRentReviewDt.text <> "" Then Rst1!RentReviewDate = CDate(Format(txtRentReviewDt.text, "dd mmmm yyyy"))
'   If txtRentIncDt.text <> "" Then Rst1!RentIncreaseDate = CDate(Format(txtRentIncDt.text, "dd mmmm yyyy"))
'   If txtRentIncAmt.text <> "" Then Rst1!RentIncreaseAmount = CDbl(txtRentIncAmt.text)
   
'   ****************************************************
'    Adding data from Supplementary tab
'   ****************************************************
'   If txtSupp1.text <> "" Then Rst1!SuppText1 = txtSupp1.text
'   If txtSupp2.text <> "" Then Rst1!SuppText2 = txtSupp2.text
'   If txtSupp3.text <> "" Then Rst1!SuppText3 = txtSupp3.text
'   If lblSupplementary1.Caption <> "" Then Rst1!SuppCaption1 = lblSupplementary1.Caption
'   If lblSupplementary2.Caption <> "" Then Rst1!SuppCaption2 = lblSupplementary2.Caption
'   If lblSupplementary3.Caption <> "" Then Rst1!SuppCaption3 = lblSupplementary3.Caption

'   If txtDtFlgDate.text <> "" Then Rst1!DateFlagDate = CDate(Format(txtDtFlgDate.text, "dd mmmm yyyy"))
'   If txtDtFlgDesc.text <> "" Then Rst1!DateFlagDescription = txtDtFlgDesc.text
'
'   If txtDtFlgDt2.text <> "" Then Rst1!DateFlagDt2 = Format(txtDtFlgDt2.text, "dd mmmm yyyy")
'   If txtDtFlgDesc2.text <> "" Then Rst1!DateFlagDescription2 = txtDtFlgDesc2.text
'
'   If txtDtFlgDt3.text <> "" Then Rst1!DateFlagDt3 = Format(txtDtFlgDt3.text, "dd mmmm yyyy")
'   If txtDtFlgDesc3.text <> "" Then Rst1!DateFlagDescription3 = txtDtFlgDesc3.text

   If txtMemo.text <> "" Then Rst1!Notes = txtMemo.text
'   ****************************************************
'    Adding data from Insurance Charge tab
'   ****************************************************
   If flxIns.TextMatrix(1, 0) <> "" Then
      Rst1!InsurancePayable = "Y"
   Else
      Rst1!InsurancePayable = "N"
   End If

'   ********************************************************
'    Mark the Lease as live. Lease will be expired only by Lease terminating command.
   '    User can type in past date. @*#*#*@
'   ********************************************************
   bUpdated = False
   
'   If var = "" if you use terminaton date form this condition shall switch based on that form use
            If txtLeaseEndDate <> "" And Label16(12).Tag <> "" Then
                     If DateDiff("d", txtLeaseEndDate.text, Label16(12).Tag) > 0 And chkOLED.Value = 0 Then
                             MsgBox "Lease termination date cannot be greater than lease end date."
                             Conn1.RollbackTrans
                             Exit Function
                     End If
            End If
       If Not chkOLED.Value Then 'this will not check OLED
        
          If DateDiff("d", Now, txtLeaseEndDate.text) < 0 And txtLeaseEndDate.ForeColor = vbBlack Then
              Rst1!Status = False
             'fixed by anol terminate date was not writing 2016 MAR 03
                If Not IsDate(var) Then
                    If IsDate(Label16(12).Tag) = True Then
                        Rst1!TerminateDate = Format(Label16(12).Tag, "dd mmmm yyyy")  'fixed by anol was giving an error when this tag was empty so added date validation 2020-07-08
                    End If
                Else
                     Rst1!TerminateDate = Format(var, "dd mmmm yyyy")
                End If
                szSQL = "UPDATE Units " & _
                       "SET Occupied = 'N' " & _
                       "WHERE UnitNumber = '" & txtUnitNumber.text & "';"
             Conn1.Execute szSQL
             bUpdated = True
          Else
             Rst1!Status = ReturnLeaseStatus(bSaveUpdate) 'True
             If Rst1!Status = True Then
                Rst1!TerminateDate = Null
             Else
                If Not IsDate(var) Then
                    Rst1!TerminateDate = Format(Label16(12).Tag, "dd mmmm yyyy")
                Else
                    If IsDate(var) Then
                        Rst1!TerminateDate = Format(var, "dd mmmm yyyy")
                    Else
                        Rst1!TerminateDate = Null
                    End If
'                     Rst1!TerminateDate = Format(var, "dd mmmm yyyy")
                End If
                szSQL = "UPDATE Units " & _
                       "SET Occupied = 'N' " & _
                       "WHERE UnitNumber = '" & txtUnitNumber.text & "';"
                Conn1.Execute szSQL
             End If
'             var = Null
          End If
       Else 'if OLED is on then status shall be true
          Rst1!Status = ReturnLeaseStatus(bSaveUpdate) 'True
          If Rst1!Status = True Then
                Rst1!TerminateDate = Null
          Else
                Rst1!TerminateDate = Format(Date, "dd mmmm yyyy")
                szSQL = "UPDATE Units " & _
                       "SET Occupied = 'N' " & _
                       "WHERE UnitNumber = '" & txtUnitNumber.text & "';"
                Conn1.Execute szSQL
          End If
       End If
'    Else 'var has been set from termination date form
'            If txtLeaseEndDate <> "" Then
'                     If DateDiff("d", txtLeaseEndDate.text, var) > 0 And chkOLED.Value = 0 Then
'                             MsgBox "Lease termination date cannot be greater than lease end date."
'                             Conn1.RollbackTrans
'                             Exit Function
'                     End If
'            End If
'
'            Rst1!Status = False
'            Rst1!TerminateDate = Format(Label16(12).Tag, "dd mmmm yyyy")
'            szSQL = "UPDATE Units " & _
'                      "SET Occupied = 'N' " & _
'                      "WHERE UnitNumber = '" & txtUnitNumber.text & "';"
'            Conn1.Execute szSQL
'            bUpdated = True
'    End If
   
    Label16(12).Tag = ""
    
    If CBool(Rst1!Status) = True Then
        szLeaseStatus = True
    Else
        szLeaseStatus = False
    End If
    If CBool(Rst1!Status) = True Then
        szMsg = " Current"
    ElseIf CBool(Rst1!Status) = False And IsNull(Rst1!TerminateDate) = False Then
        szMsg = " Terminated (" & Rst1!TerminateDate & ")"
        Label16(12).Tag = Rst1!TerminateDate
    ElseIf CBool(Rst1!Status) = False And IsNull(Rst1!TerminateDate) = True Then 'if status is false and there is no termination date that means lease has expired
        szMsg = " Expired (" & Rst1!EndDate & ")"
    Else
        szMsg = ""
    End If
    szUndoLeaseStatus = Label16(12).Caption
    Label16(12).Caption = szMsg

'  ********************************************************
   Rst1.Update
   Rst1.Close
   ' here you need to put a check of double active lessee on same unit and put the commit trans and rollback trans
   If double_active_lessee(Conn1) Then
        Conn1.RollbackTrans
        'MsgBox "Two lessees cannot be active against the same unit at the same time", vbInformation, "Warning"
       
        Conn1.Close
        txtLeaseEndDate.text = szUndoLeaseEndDate
        Label16(12).Caption = szUndoLeaseStatus
        Call DisableBoxes
        cmdAddNew.Visible = True
        cmdUnitNumber(0).Enabled = False
        cmdAddNew.TabIndex = 25
        cmdTerminate.Visible = True
        cmdDelete.Visible = True
        cmdTerminate.TabIndex = 26
        cmdEdit.Visible = True
        cmdSaveNew.Visible = False
        cmdCancelNew.Visible = False
        cmdSaveEdit.Visible = False
        cmdCancelEdit.Visible = False
        Exit Function
   Else
        Conn1.CommitTrans
        Conn1.Close
   End If
   
   If COPY_LEASE = True Then
        COPY_LEASE = False
              'issue 594 1)
      'user cannot go back to browsing state lease
        Dim iRow As Integer
        For iRow = 1 To flxRentCharges.Rows - 1
            If flxRentCharges.TextMatrix(iRow, 13) <> "" Then
                flxRentCharges.TextMatrix(iRow, 13) = ""
            End If
        Next iRow

        For iRow = 1 To flxSC.Rows - 1
            If flxSC.TextMatrix(iRow, 14) <> "" Then
                 flxSC.TextMatrix(iRow, 14) = ""
            End If

        Next iRow

        For iRow = 1 To flxIns.Rows - 1
            If flxIns.TextMatrix(iRow, 0) <> "" Then
                flxIns.TextMatrix(iRow, 0) = ""
            End If
        Next iRow
   End If
   Conn1.Open getConnectionString
'   ****************************************************
'    Adding data from Rent Charge tab
'   ****************************************************
   If flxRentCharges.TextMatrix(1, 0) <> "" Then
''      Delete all Rent Charges of the Current Lease if exist
'      SQLStr2 = "DELETE * " & _
'                "FROM LRentCharges " & _
'                "WHERE LeaseID = '" & txtLeaseID.text & "' And " & _
'                  "(ISNULL(spare3) OR spare3 = '');"
'      Conn1.Execute SQLStr2

'      Add Rent Charges in the LRentCharges table
     ' Dim iRow As Integer

      For iRow = 1 To flxRentCharges.Rows - 1
         SQLStr2 = "SELECT * " & _
                   "FROM LRentCharges " & _
                   "WHERE RentCharges = '" & flxRentCharges.TextMatrix(iRow, 13) & "';"
         Rst2.Open SQLStr2, Conn1, adOpenDynamic, adLockOptimistic
         'Issue 823 Date 23-01-2020 fund code was saving empty without giving any warning.I have added validation
         If flxRentCharges.TextMatrix(iRow, 6) = "" Then
            MsgBox "Please enter a fund code for rent charges before saving this lease!", vbInformation, "Fund code is missing"
            tabLease.Tab = 1
            Rst2.Close
            Exit Function
         End If
         If flxRentCharges.TextMatrix(iRow, 3) = "" Then
            MsgBox "Please select a frequency before you save rent charges!", vbInformation, "Frequency is missing"
            tabLease.Tab = 1
            Rst2.Close
            Exit Function
         End If
         If Rst2.EOF Then
            Rst2.AddNew
            Rst2!RentCharges = UniqueID()
         End If
         szSplitLeaseID = Rst2!RentCharges

         Rst2!LeaseID = strLeaseId
         Rst2!RentChargeDept = flxRentCharges.TextMatrix(iRow, 6)
         Rst2!BRfrequency = flxRentCharges.TextMatrix(iRow, 3)
         Rst2!BRStartDate = CDate(Format(flxRentCharges.TextMatrix(iRow, 2), "dd mmmm yyyy"))
         Rst2!BRNextDueDate = CDate(Format(flxRentCharges.TextMatrix(iRow, 5), "dd mmmm yyyy"))
         Rst2!BRTotal = flxRentCharges.TextMatrix(iRow, 11)
         Rst2!BRAmount = flxRentCharges.TextMatrix(iRow, 12)
'  UpdateDT_DSR --> Updating the DT in the demand split table.

         If Rst2!BRDemandType <> flxRentCharges.TextMatrix(iRow, 0) Then
            UpdateDT_DSR Conn1, Rst2!BRDemandType, flxRentCharges.TextMatrix(iRow, 0)
         End If
         Rst2!BRDemandType = flxRentCharges.TextMatrix(iRow, 0)
         Rst2!RentDesc = flxRentCharges.TextMatrix(iRow, 14)
         Rst2!spare1 = flxRentCharges.TextMatrix(iRow, 8)
         Rst2!spare2 = flxRentCharges.TextMatrix(iRow, 10)
         Rst2!spare3 = IIf(flxRentCharges.TextMatrix(iRow, 15) = "", Null, flxRentCharges.TextMatrix(iRow, 15)) 'writing the delete flag
         If flxRentCharges.TextMatrix(iRow, 16) <> "" Then
            Rst2!StopRC = Format(flxRentCharges.TextMatrix(iRow, 16), "dd mmmm yyyy")
         Else
            Rst2!StopRC = ""
         End If

         Rst2.Update
         Rst2.Close
      
         SetFDD_Lease "RC", szSplitLeaseID, _
                      CDate(Format(flxRentCharges.TextMatrix(iRow, 5), "dd mmmm yyyy")), _
                      flxRentCharges.TextMatrix(iRow, 3), _
                      flxRentCharges.TextMatrix(iRow, 0), _
                      PROPERTY_ID, Conn1
      Next iRow
   End If

'   ****************************************************
'    Adding data from Service Charge tab
'   ****************************************************
   If flxSC.TextMatrix(1, 0) <> "" Then
'      'Delete all Rent Charges of the Current Lease if exist
'      SQLStr2 = "DELETE * " & _
'                "FROM LServiceCharges " & _
'                "WHERE LeaseID = '" & txtLeaseID.text & "'"
'      Conn1.Execute SQLStr2

      For iRow = 1 To flxSC.Rows - 1
         'Add Rent Charges in the LRentCharges table
         SQLStr2 = "SELECT * " & _
                   "FROM LServiceCharges " & _
                   "WHERE ServiceCharge = '" & flxSC.TextMatrix(iRow, 14) & "';"
         Rst2.Open SQLStr2, Conn1, adOpenDynamic, adLockOptimistic

         If Rst2.EOF Then
            Rst2.AddNew
            Rst2!ServiceCharge = UniqueID()
         End If
         szSplitLeaseID = Rst2!ServiceCharge

         Rst2!LeaseID = strLeaseId
         If Rst2!SCDemandType <> flxSC.TextMatrix(iRow, 0) Then
            UpdateDT_DSR Conn1, Rst2!SCDemandType, flxSC.TextMatrix(iRow, 0)
         End If
         Rst2!SCDemandType = flxSC.TextMatrix(iRow, 0)
         Rst2!SCPayableFrom = CDate(Format(flxSC.TextMatrix(iRow, 2), "dd mmmm yyyy"))
         Rst2!SCFrequency = flxSC.TextMatrix(iRow, 17)
         Rst2!SCNextDueDate = CDate(Format(flxSC.TextMatrix(iRow, 4), "dd mmmm yyyy"))
         Rst2!ServiceChargeDept = flxSC.TextMatrix(iRow, 5)
         If flxSC.TextMatrix(iRow, 7) <> "" Then Rst2!ScheduleID = flxSC.TextMatrix(iRow, 7)
         Rst2!ChargingMethod = CInt(flxSC.TextMatrix(iRow, 9))
         Rst2!CMFigure = CDbl(Val(flxSC.TextMatrix(iRow, 11)))
         Rst2!SCTotal = Val(flxSC.TextMatrix(iRow, 12))
         Rst2!SCAmount = Val(flxSC.TextMatrix(iRow, 13))

         Rst2!SCDesc = flxSC.TextMatrix(iRow, 15)
         Rst2!spare3 = IIf(flxSC.TextMatrix(iRow, 16) = "", Null, flxSC.TextMatrix(iRow, 16)) 'Delete flag
         If flxSC.TextMatrix(iRow, 18) <> "" Then
            Rst2!StopSC = Format(flxSC.TextMatrix(iRow, 18), "dd mmmm yyyy")
         Else
            Rst2!StopSC = ""
         End If
         Rst2.Update
         Rst2.Close

         SetFDD_Lease "SC", szSplitLeaseID, _
                      CDate(Format(flxSC.TextMatrix(iRow, 4), "dd mmmm yyyy")), _
                      flxSC.TextMatrix(iRow, 17), _
                      flxSC.TextMatrix(iRow, 0), _
                      PROPERTY_ID, Conn1
      Next iRow
   End If
'   ****************************************************
'    Adding data from Insurance Charge tab
'   ****************************************************
   If flxIns.TextMatrix(1, 0) <> "" Then
'      'Delete all Rent Charges of the Current Lease if exist
'      SQLStr2 = "DELETE * " & _
'                "FROM LInsuranceCharges " & _
'                "WHERE LeaseID = '" & txtLeaseID.text & "'"
'      Conn1.Execute SQLStr2

      For iRow = 1 To flxIns.Rows - 1
         'Add Rent Charges in the LRentCharges table
         SQLStr2 = "SELECT * " & _
                   "FROM LInsuranceCharges " & _
                   "WHERE InsCharges = '" & flxIns.TextMatrix(iRow, 0) & "';"
         Rst2.Open SQLStr2, Conn1, adOpenDynamic, adLockOptimistic
         
         If Rst2.EOF Then
            Rst2.AddNew
            Rst2!InsCharges = UniqueID 'flxIns.TextMatrix(iRow, 0)
         End If
         szSplitLeaseID = Rst2!InsCharges

         Rst2!LeaseID = strLeaseId
         Rst2!InsuranceDept = flxIns.TextMatrix(iRow, 17) 'which is in the grid col:
         Rst2!InsuranceStartDate = CDate(Format(flxIns.TextMatrix(iRow, 2), "dd mmmm yyyy"))
         Rst2!InsuranceFrequency = flxIns.TextMatrix(iRow, 11) 'which is in the grid col:3

'         If Rst2!InsuranceDemandType <> flxIns.TextMatrix(iRow, 12) And FIX_MODE__DT Then
         If Rst2!InsuranceDemandType <> flxIns.TextMatrix(iRow, 12) Then
            UpdateDT_DSR Conn1, Rst2!InsuranceDemandType, flxIns.TextMatrix(iRow, 12)
         End If
         
         Rst2!InsuranceDemandType = flxIns.TextMatrix(iRow, 12) 'which is in the grid col:
         Rst2!InsuranceNextDueDate = CDate(Format(flxIns.TextMatrix(iRow, 4), "dd mmmm yyyy"))
         Rst2!ChargingType = CInt(flxIns.TextMatrix(iRow, 13))
         Rst2!ChargingFigure = CDbl(flxIns.TextMatrix(iRow, 7))
         Rst2!TotalYearlyInsurance = CDbl(flxIns.TextMatrix(iRow, 8))
         Rst2!InsuranceEachPeriod = CDbl(flxIns.TextMatrix(iRow, 9))
         Rst2!InsDesc = flxIns.TextMatrix(iRow, 14)
         Rst2!spare3 = IIf(flxIns.TextMatrix(iRow, 15) = "", Null, flxIns.TextMatrix(iRow, 15))
         If flxIns.TextMatrix(iRow, 16) <> "" Then
            Rst2!StopIC = Format(flxIns.TextMatrix(iRow, 16), "dd mmmm yyyy")
         Else
            Rst2!StopIC = ""
         End If

         Rst2.Update
         Rst2.Close

         SetFDD_Lease "IC", szSplitLeaseID, _
                      CDate(Format(flxIns.TextMatrix(iRow, 4), "dd mmmm yyyy")), _
                      flxIns.TextMatrix(iRow, 11), _
                      flxIns.TextMatrix(iRow, 12), _
                      PROPERTY_ID, Conn1
      Next iRow
   End If
'*********************************************************************************
''Breach information saving
   If gridBreach.TextMatrix(1, 0) <> "" Then
      For iRow = 1 To gridBreach.Rows - 1
         SQLStr2 = "SELECT * " & _
                   "FROM LeaseBreaches " & _
                   "WHERE BreachID = " & gridBreach.TextMatrix(iRow, 6) & ";"
                   
         Rst2.Open SQLStr2, Conn1, adOpenDynamic, adLockOptimistic
         If Rst2.EOF Then
            Rst2.AddNew
         End If
         Rst2!LeaseID = strLeaseId
         Rst2!BreachType = gridBreach.TextMatrix(iRow, 8)
         Rst2!CommenceDate = gridBreach.TextMatrix(iRow, 1)
         Rst2!InitiatedBy = gridBreach.TextMatrix(iRow, 2)
         Rst2!Resolved = IIf(gridBreach.TextMatrix(iRow, 5) = "No", 0, 1)
         Rst2!DateReceived = gridBreach.TextMatrix(iRow, 3)
         Rst2!ReceivedBy = gridBreach.TextMatrix(iRow, 4)
         Rst2!deleteFlag = gridBreach.TextMatrix(iRow, 7)
         Rst2!LeaseMemo = gridBreach.TextMatrix(iRow, 9)
         Rst2.Update
         Rst2.Close
      Next iRow
     
   End If

'*********************************************************************************
   If bSaveUpdate And Not bUpdated Then
      szSQL = "UPDATE Units " & _
              "SET    OCCUPIED   = 'Y' " & _
              "WHERE  UNITNUMBER = '" & txtUnitNumber.text & "';"
      Conn1.Execute szSQL
   End If

   UpdateStopAutoDmd Conn1
''   'here you need to put a check of double active lessee on same unit and put the commit trans and rollback trans
''   If double_active_lessee(Conn1) Then
''        Conn1.RollbackTrans
''        MsgBox "Two lessee cannot be active at the same time", vbInformation, "Warning"
''   Else
''        Conn1.CommitTrans
''   End If
   Conn1.Close
   Set Conn1 = Nothing

   'Call DisableBoxes

   cmdAddNew.Visible = True
   cmdUnitNumber(0).Enabled = False
   cmdAddNew.TabIndex = 25
   cmdTerminate.Visible = True
   cmdDelete.Visible = True
   cmdTerminate.TabIndex = 26
   cmdEdit.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False

   'Call EmptyBoxes ' this function is clearing all txtboxes

   SaveUpdateLease = True
End Function
Private Function double_active_lessee(Conn1) As Boolean 'FALSE MEANS DATA IS FINE AND TRUE MEANS THERE IS INCONSISTENCY
    Dim rsLessee As New ADODB.Recordset
    rsLessee.Open "SELECT  count(LeaseDetails.SageaccountNumber) as cnt,LeaseDetails.SageaccountNumber " & _
                  "FROM LeaseDetails where status=true group by LeaseDetails.SageaccountNumber having  count(LeaseDetails.SageaccountNumber) >1 ;", Conn1, adOpenKeyset, adLockOptimistic
    If Not rsLessee.EOF Then
        double_active_lessee = True
    End If
    
    If double_active_lessee = True Then
            MsgBox "It is not possible to reactivate this lease as there is already an active lease against this lessee(" & rsLessee("SageaccountNumber").Value & ") . " & vbCrLf & _
            "You must cease or terminate the current active lease before you can reactivate any expired lease for this lessee.", vbInformation, "Warning"
            rsLessee.Close
            Set rsLessee = Nothing
            Exit Function
    End If
    rsLessee.Close
    Set rsLessee = Nothing
    
    'Below check has been added by anol 2019-08-13 issue 798 Cannot Log into Everon
    rsLessee.Open "SELECT  count(LeaseDetails.UNITNumber) as cnt,LeaseDetails.UNITNumber FROM LeaseDetails where status=true group by LeaseDetails.UNITNumber having " & _
                  " count(LeaseDetails.UNITNumber) >1 ;", Conn1, adOpenKeyset, adLockOptimistic
    If Not rsLessee.EOF Then
            double_active_lessee = True
    End If
    If double_active_lessee = True Then
        MsgBox "Two lessees cannot be active against the same unit(" & rsLessee("UNITNumber").Value & ") at the same time", vbInformation, "Warning"
        rsLessee.Close
        Set rsLessee = Nothing
        Exit Function
    End If
    rsLessee.Close
    Set rsLessee = Nothing
End Function
Private Sub UpdateDT_DSR(adoConn As ADODB.Connection, byInsDT_old As Integer, byInsDT_new As Integer)
   adoConn.Execute "UPDATE DemandSplitRecords AS DSR, DemandRecords AS D " & _
                   "SET DSR.TypeOfDemand = " & byInsDT_new & " " & _
                   "WHERE DSR.DemandID = D.DemandID AND " & _
                     "D.SageAccountNumber = '" & flxLeaseList.TextMatrix(flxLeaseList.row, 2) & "' AND " & _
                     "DSR.TypeOfDemand = " & byInsDT_old & ";"
End Sub

Private Sub cmdSaveRentCrg_Click()
'   If MsgBox("Do you want to save now?", vbQuestion + vbYesNo, "Save") = vbYes Then
   If Not IsDate(txtNextDueDate.text) Then
            MsgBox "Next Due Date is empty.", vbCritical + vbExclamation, "Next Due Date Required"
            Exit Sub
   End If
    If Trim(txtBRChargingFigure.text) = "" Then
        txtBRChargingFigure.text = "0.00"
        
   End If
  
   
   If txtFreqBR.text = "" Then
      MsgBox "You must select a Rent Frequency!", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 1
      FocusControl cmdFreqBR
      Exit Sub
   End If
   If txtRentStartDate.text = "" Then
      MsgBox "You must enter a Rent Start Date!", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 1
      txtRentStartDate.SetFocus
      Exit Sub
   End If
   If IsDate(txtNextDueDate.text) And IsDate(txtRentStartDate.text) Then
        If DateDiff("d", txtRentStartDate.text, txtNextDueDate.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the Rent start date", vbInformation, "Warning"
            txtNextDueDate.text = ""
            FocusControl txtNextDueDate
            Exit Sub
        End If
    End If
    'Or Trim(cboRentChargeDept.text) = "" validation has been added by anol 23-01-2020 issue 833
   If txtRCFund.text = "" Then
        MsgBox "Please select a valid fund!", vbOKOnly + vbCritical, "fund"
        cmdRCFund.SetFocus
        Exit Sub
   End If
   If txtBRDemandType.text = "" Then
      MsgBox "You must choose Demand type from the dropdown menu.", vbCritical + vbOKOnly, "Data Required"
      tabLease.Tab = 1
      FocusControl cmdBRDemandType
      Exit Sub
   End If
   If txtTotalRentYear.text = "" Then
      MsgBox "You must enter a Total rent for the year.", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 1
      txtTotalRentYear.SetFocus
      Exit Sub
   End If
   If Not chkRentDes.Value And Trim(txtRentDesc.text) = "" Then
      MsgBox "You must enter description of the rent charges.", vbOKOnly + vbCritical, "Date Required"
      tabLease.Tab = 1
      txtRentDesc.SetFocus
      Exit Sub
   End If
   If chargingMonthValidateBR = False Then
        Exit Sub
   End If
   If chargingMonthValidateBRFreq = False Then
        Exit Sub
   End If
   
'      ' Add a row at the bottom of the grid
   If RENTCHARGES_EDIT = 0 Then
      If flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 0) <> "" Then flxRentCharges.AddItem ""
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 0) = CInt(txtBRDemandType.Tag)
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 1) = txtBRDemandType.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 2) = Format(txtRentStartDate.text, "dd/mm/yyyy")
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 3) = CInt(txtFreqBR.Tag)
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 4) = txtFreqBR.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 5) = Format(txtNextDueDate.text, "dd/mm/yyyy")
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 6) = txtRCFundCode.Tag
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 7) = txtRCFundCode.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 8) = cboBRChargingMth.Column(0)
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 9) = cboBRChargingMth.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 10) = txtBRChargingFigure.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 11) = Format(txtTotalRentYear.text, "0.00")
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 12) = txtRentDueEachPeriod.text
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 13) = UniqueID()
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 14) = Trim(txtRentDesc.text)
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 16) = Trim(txtStopRC.text)
      flxRentCharges.TextMatrix(flxRentCharges.Rows - 1, 17) = txtRCFund.text
   Else
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 0) = CInt(txtBRDemandType.Tag)
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 1) = txtBRDemandType.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 2) = Format(txtRentStartDate.text, "dd/mm/yyyy")
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 3) = CInt(txtFreqBR.Tag)
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 4) = txtFreqBR.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 5) = Format(txtNextDueDate.text, "dd/mm/yyyy")
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 6) = txtRCFundCode.Tag
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 7) = txtRCFundCode.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 8) = cboBRChargingMth.Column(0)
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 9) = cboBRChargingMth.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 10) = txtBRChargingFigure.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 11) = Format(txtTotalRentYear.text, "0.00")
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 12) = txtRentDueEachPeriod.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 13) = txtRentChargesIDEdit.text
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 14) = Trim(txtRentDesc.text)
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 16) = Trim(txtStopRC.text)
      flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 17) = txtRCFund.text
   End If

   ControlsModeRentCharges DefaultMode
   RENTCHARGES_EDIT = 0
   chkRentDes.Visible = False
   ShowMsgInTaskBar "The rent charge grid has been updated."
   'End If
End Sub

Private Sub cmdSCCancel_Click()
   If MsgBox("Are you sure you wish to cancel Service Charge changes?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      ControlsModeServiceCharges ExpensesMode
      'added by anol 04 Feb 2015
      txtCapAmount.Locked = True
      txtCapAmount.Enabled = False
      'Added by anol 25 FEB 2015
      cmdSCDemandType.Enabled = False
      txtPayableFrom.Locked = True
      cmdFreqSC.Enabled = False
      txtSCNextDueDt.Locked = True
     cmdSCFund.Enabled = False
      cboSchedule.Locked = True
      cboSCChargingMth.Locked = True
      txtChargingFigure.Locked = True
      txtSCTotalAmount.Locked = True
      txtSCDueEachPeriod.Locked = True
      txtStopSC.Locked = True
   'End of Modification
      SERVICECHARGES_EDIT = 0
   End If
End Sub

Private Sub cmdSCDelete_Click()
   If cmdSaveNew.Visible Then Exit Sub
   'Primary key for LServiceCharges table is ServiceCharge which is loading at the column number 14
   'Now what I need to do, direct delete from database without retreiving and load again
   'Modified by anol 2020-06-04
'   If cmdSCDelete.Caption = "&Delete Charge" Then
'      If MsgBox("Would you like to delete this Service Charge?", vbQuestion + vbYesNo, "Service Charge") = vbNo Then Exit Sub
'
'      flxSC.TextMatrix(flxSC.row, 16) = "DELETED"
'      flxSC.RowHeight(flxSC.row) = 0
'      'MsgBox "This Service charge has been marked for deletition. It will be permanently removed when you save this lease.", vbInformation + vbOKOnly, "Service Charge"
'   Else
'      flxSC.TextMatrix(flxSC.row, 16) = ""
'      MsgBox "This Service charge has been retrieved.", vbInformation + vbOKOnly, "Service Charge"
'   End If
    
    Dim adoConn1   As New ADODB.Connection
    If MsgBox("Would you like to delete this Service Charge", vbQuestion + vbYesNo, "Please Confirm to delete current Service Charge") = vbYes Then
        adoConn1.Open getConnectionString
        adoConn1.Execute "Delete from LServiceCharges where ServiceCharge='" & flxSC.TextMatrix(flxSC.row, 14) & "'"
        LoadFlxSC adoConn1
        adoConn1.Close
        Set adoConn1 = Nothing
    End If
    
   ControlsModeServiceCharges DefaultMode
End Sub

Private Sub cmdSCEdit_Click()
    ControlsModeServiceCharges EditMode
   SERVICECHARGES_EDIT = flxSC.row
    'Added by anol 04 FEB 2015
   txtCapAmount.Locked = False
   txtCapAmount.Enabled = True
    'Added by anol 06 Nov 2014
   'issue 471 Note 718
  cmdSCDemandType.Enabled = True
   txtPayableFrom.Locked = False
   cmdFreqSC.Enabled = True
   txtSCNextDueDt.Locked = False
   cmdSCFund.Enabled = True
   cboSchedule.Locked = False
   cboSCChargingMth.Locked = False
   txtChargingFigure.Locked = False
   txtSCTotalAmount.Locked = False
   txtSCDueEachPeriod.Locked = False
   txtStopSC.Locked = False
   'End of Modification
   FocusControl cmdSCDemandType
End Sub

Private Sub cmdSCNew_Click()
   If txtLeaseEndDate.ForeColor = vbRed Then
      MsgBox "This lease has expired, then you cannot add new charge.", vbCritical + vbOKOnly, "Lease Expired"
      Exit Sub
   End If
   Dim Conn As New ADODB.Connection
   Conn.Open getConnectionString
   AllDemandType Conn
'   LoadDept Conn
   Conn.Close
   Set Conn = Nothing
   ControlsModeServiceCharges NewEntryMode
   'added by anol 04 Feb 2015
   txtCapAmount.Locked = False
   txtCapAmount.Enabled = True
   'Added by anol 22 Jan 2015
   cmdSCDemandType.Enabled = True
   txtPayableFrom.Locked = False
   cmdFreqSC.Enabled = True
   txtSCNextDueDt.Locked = False
   cmdSCFund.Enabled = True
   cboSchedule.Locked = False
   cboSCChargingMth.Locked = False
   txtChargingFigure.Locked = False
   txtSCTotalAmount.Locked = False
   txtSCDueEachPeriod.Locked = False
   txtStopSC.Locked = False
   'End of Modification
   SERVICECHARGES_EDIT = 0
   chkSCDes.Visible = True
   FocusControl cmdSCDemandType
End Sub

Private Sub cmdSCSave_Click()
'   If MsgBox("Do you want to save now?", vbQuestion + vbYesNo, "Save") = vbYes Then
   If txtSCDemandType.text = "" Then
      MsgBox "You must select a Demand Type", vbOKOnly + vbCritical, "Demand Type Required"
      tabLease.Tab = 4
      FocusControl cmdSCDemandType
      Exit Sub
   End If
   
   If txtPayableFrom.text = "" Then
      MsgBox "You must enter a Service Charge start date.", vbOKOnly + vbCritical, "Service Charge start date"
      tabLease.Tab = 4
      txtPayableFrom.SetFocus
      Exit Sub
   End If
   If Trim(txtChargingFigure.text) = "" Then
        txtChargingFigure.text = "0.00"
   End If
   If IsDate(txtSCNextDueDt.text) And IsDate(txtPayableFrom.text) Then
        If DateDiff("d", txtPayableFrom.text, txtSCNextDueDt.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the Service Charge start date", vbInformation, "Warning"
            txtSCNextDueDt.text = ""
            FocusControl txtSCNextDueDt
            Exit Sub
        End If
    End If
   If txtFreqSC.text = "" Then
      MsgBox "You must select a Frequency.", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 4
      FocusControl cmdFreqSC
      Exit Sub
   End If
   If txtSCNextDueDt.text = "" Then
      MsgBox "You must enter next due date.", vbOKOnly + vbCritical, "Next due date Required"
      tabLease.Tab = 4
      txtSCNextDueDt.SetFocus
      Exit Sub
   End If
    If chargingMonthValidateSC = False Then
        Exit Sub
   End If
    If chargingMonthValidateSCFreq = False Then
        Exit Sub
   End If
   
   
   If txtSCFundName.text = "" Then
      MsgBox "You must select a Service Charge Fund.", vbOKOnly + vbCritical, "Fund Required"
      tabLease.Tab = 4
      FocusControl cmdSCFund
      Exit Sub
   End If
   If txtFreqSC.text = "" Then
      MsgBox "You must select a Service Charge Frequency!", vbOKOnly + vbCritical, "Frequency Required"
      tabLease.Tab = 4
      cmdFreqSC.SetFocus
      Exit Sub
   End If
   If txtPayableFrom.text = "" Then
      MsgBox "You must enter a Service Charge Start Date!", vbOKOnly + vbCritical, "Start Date Required"
      tabLease.Tab = 4
      txtPayableFrom.SetFocus
      Exit Sub
   End If
   If txtSCDemandType.text = "" Then
      MsgBox "You must choose a Demand Type from the dropdown menu.", vbCritical + vbOKOnly, "Demand Type Required"
      tabLease.Tab = 4
      FocusControl cmdSCDemandType
      Exit Sub
   End If
   If txtChargingFigure.text = "" Then
      MsgBox "You must enter a Service Charge Amount.!", vbOKOnly + vbCritical, "Charge Amount Required"
      tabLease.Tab = 4
      txtChargingFigure.SetFocus
      Exit Sub
   End If
   If cboSCChargingMth.text = "" Then
      MsgBox "You must select a Service Charge Charging Method from the dropdown menu.", vbCritical + vbOKOnly, "Charge Method Required"
      tabLease.Tab = 4
      cboSCChargingMth.SetFocus
   End If
   If Not chkSCDes.Value And Trim(txtSCDesc.text) = "" Then
      MsgBox "You must enter a Service Charge Description.", vbOKOnly + vbCritical, "Description Required"
      tabLease.Tab = 4
      txtSCDesc.SetFocus
      Exit Sub
   End If
'   If Val(txtSCDueEachPeriod.text) < 0 Then
'      MsgBox "Service Charge Amount cannot be less than zero.", vbOKOnly + vbCritical, "Service Charge Amount Less Than Zero"
'      tabLease.Tab = 4
'      cboSCChargingMth.SetFocus
'      Exit Sub
'   End If

'      Dim szaTemp() As String
'      szaTemp = Split(cboFreqSC.text, "-")
''*** Saving data in the grid.
'       Add a row at the bottom of the grid
 'Modified by anol 04 Feb 2015
      
      Dim Conn1 As New ADODB.Connection
   If SERVICECHARGES_EDIT = 0 Then
      If flxSC.TextMatrix(flxSC.Rows - 1, 0) <> "" Then flxSC.AddItem ""

'         flxSC.TextMatrix(flxSC.Rows - 1, 0) = CByte(cboSCDemandType.ListIndex + 1)
      flxSC.TextMatrix(flxSC.Rows - 1, 0) = CInt(txtSCDemandType.Tag)
      flxSC.TextMatrix(flxSC.Rows - 1, 1) = txtSCDemandType.text
      flxSC.TextMatrix(flxSC.Rows - 1, 20) = txtSCDemandType.text
      flxSC.TextMatrix(flxSC.Rows - 1, 2) = Format(txtPayableFrom.text, "dd/mm/yyyy")
      flxSC.TextMatrix(flxSC.Rows - 1, 3) = txtFreqSC.text
      flxSC.TextMatrix(flxSC.Rows - 1, 4) = Format(txtSCNextDueDt.text, "dd/mm/yyyy")
      flxSC.TextMatrix(flxSC.Rows - 1, 5) = CInt(txtSCFundCode.Tag)
      flxSC.TextMatrix(flxSC.Rows - 1, 6) = txtSCFundCode.text
      flxSC.TextMatrix(flxSC.Rows - 1, 7) = IIf(cboSchedule.text = "", 0, cboSchedule.Value)
      flxSC.TextMatrix(flxSC.Rows - 1, 8) = cboSchedule.text
      flxSC.TextMatrix(flxSC.Rows - 1, 9) = cboSCChargingMth.Column(0)
      flxSC.TextMatrix(flxSC.Rows - 1, 10) = cboSCChargingMth.text
      flxSC.TextMatrix(flxSC.Rows - 1, 11) = Format(txtChargingFigure.text, "0.00000000")
      
      'Modified by anol 04 Feb 2015
      If Val(txtCapAmount.text) > 0 Then
         flxSC.TextMatrix(flxSC.Rows - 1, 12) = IIf(Val(txtSCTotalAmount.text) > Val(txtCapAmount.text), Val(txtCapAmount.text), Val(txtSCTotalAmount.text))
         Conn1.Open getConnectionString
         Rst1.Open "SELECT PARTOFYEAR FROM FREQUENCIES WHERE ID = " & txtFreqSC.Tag & ";", Conn1, adOpenStatic, adLockReadOnly
         flxSC.TextMatrix(flxSC.Rows - 1, 13) = Format((IIf(Val(txtSCTotalAmount.text) > Val(txtCapAmount.text), Val(txtCapAmount.text), Val(txtSCTotalAmount.text)) / CInt(Rst1!PartOfYear)), "0.00")
         Conn1.Close
      Else
         flxSC.TextMatrix(flxSC.Rows - 1, 12) = txtSCTotalAmount.text
         flxSC.TextMatrix(flxSC.Rows - 1, 13) = txtSCDueEachPeriod.text
      End If
      
      flxSC.TextMatrix(flxSC.Rows - 1, 14) = UniqueID()
      flxSC.TextMatrix(flxSC.Rows - 1, 15) = Trim(txtSCDesc.text)
      flxSC.TextMatrix(flxSC.Rows - 1, 17) = CInt(txtFreqSC.Tag)
      flxSC.TextMatrix(flxSC.Rows - 1, 18) = Trim(txtStopSC.text)
      flxSC.TextMatrix(flxSC.Rows - 1, 19) = txtSCFundName.text
   Else
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 0) = CInt(txtSCDemandType.Tag)
'         flxSC.TextMatrix(SERVICECHARGES_EDIT, 0) = CByte(cboSCDemandType.ListIndex + 1)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 1) = txtSCDemandType.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 20) = txtSCDemandType.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 2) = Format(txtPayableFrom.text, "dd/mm/yyyy")
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 3) = txtFreqSC.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 4) = Format(txtSCNextDueDt.text, "dd/mm/yyyy")
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 5) = CInt(txtSCFundCode.Tag)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 6) = txtSCFundCode.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 7) = IIf(cboSchedule.text = "", 0, cboSchedule.Value)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 8) = cboSchedule.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 9) = cboSCChargingMth.Column(0)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 10) = cboSCChargingMth.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 11) = Format(txtChargingFigure.text, "0.00000000")
       'Modified by anol 04 Feb 2015
      If Val(txtCapAmount.text) > 0 Then
         flxSC.TextMatrix(SERVICECHARGES_EDIT, 12) = IIf(Val(txtSCTotalAmount.text) > Val(txtCapAmount.text), Val(txtCapAmount.text), Val(txtSCTotalAmount.text))
         Conn1.Open getConnectionString
         Rst1.Open "SELECT PARTOFYEAR FROM FREQUENCIES WHERE ID = " & txtFreqSC.Tag & ";", Conn1, adOpenStatic, adLockReadOnly
         flxSC.TextMatrix(SERVICECHARGES_EDIT, 13) = Format((IIf(Val(txtSCTotalAmount.text) > Val(txtCapAmount.text), Val(txtCapAmount.text), Val(txtSCTotalAmount.text)) / CInt(Rst1!PartOfYear)), "0.00")
         Conn1.Close
      Else
         flxSC.TextMatrix(SERVICECHARGES_EDIT, 12) = txtSCTotalAmount.text
         flxSC.TextMatrix(SERVICECHARGES_EDIT, 13) = txtSCDueEachPeriod.text
      End If
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 14) = txtSCCharge.text
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 15) = Trim(txtSCDesc.text)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 17) = CInt(txtFreqSC.Tag)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 18) = Trim(txtStopSC.text)
      flxSC.TextMatrix(SERVICECHARGES_EDIT, 19) = txtSCFundName.text
   End If

   ControlsModeServiceCharges DefaultMode
   SERVICECHARGES_EDIT = 0
   ShowMsgInTaskBar "The service charge grid has been updated."
   Frame1(8).Enabled = True
'   End If
End Sub

Private Sub cmdSetBreachType_Click()
   Dim i As Integer

   frmSecondaryCode.PRIMARY_CODE_SHOW = "BTYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   Conn2.Open getConnectionString
   'Set the RDO Connections to the dataset
   szSQL = "SELECT SecondaryCode.Value as V, SecondaryCode.CODE as C  " & _
           "FROM   SecondaryCode " & _
           "WHERE  PrimaryCode = 'BTYP' " & _
           "ORDER BY Value;"
   Rst1.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
'#
   i = Rst1.RecordCount
   If i > 0 Then
      ReDim Data(1, i - 1) As String

      i = 0
      While Not Rst1.EOF
         Data(0, i) = CStr(Rst1!c)
         Data(1, i) = CStr(Rst1!V)
         Rst1.MoveNext
         i = i + 1
      Wend
      cboBreachType.Clear
      cboBreachType.Column() = Data()
   End If
   Rst1.Close
   Conn2.Close
End Sub

Private Sub cmdtenants_Click()
    fraList(1).Visible = False
    Dim iRow As Integer
    cmdPropertyList.Enabled = False
    cmdClientList.Enabled = False
'   cboTenant.Left = txtTenant.Left
'   cboTenant.Top = txtTenant.Top
'   cboTenant.Visible = True
'   cboTenant.SetFocus

   
   ' You cannot search tenant by client and property, on add new mode. This is the rule I set (anol), because we are activating a tenant against a new unit
   ' This is a add new mode for leasedetail, So I need to exclude showing current Lease that are in leasedetail
   ' I am only going to show Tenants that does not have any lease ( when chkexpire not ticked)
   ' I am only going to show Tenants that has lease and lease is expired ( when chkexpire ticked)
   
   '20180402
   txtClientList.text = "ALL"
   txtClientList.Tag = "ALL"
   txtPropertyList.text = "ALL"
   txtPropertyList.Tag = "ALL"
   
   
   Dim szWhere As String
   Dim szOrderby As String
   
   szOrderby = "LeaseDetails.SageAccountNumber ASC"
   
   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = ""
      
   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
      
   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' "
      
   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
                         
   
   txtSearchName.text = ""
   txtSearchTenant.text = ""
   txtSearchUnitName.text = ""
   
   
   fraList(1).Top = Frame1(16).Top + txtTenant.Top + txtTenant.Height + 5
   fraList(1).Left = Frame1(16).Left + txtTenant.Left + 5
   'fraList(1).Visible = True
   
   Conn2.Open getConnectionString
   If frmMMain.Leasee4_LesseList_isUptoDate = False Then
        TenantAccountBalance Conn2
        frmMMain.Leasee4_LesseList_isUptoDate = True
   End If
  
'issue 559
'2)  It is not showing the expired lessee Name in the list where we are creating
'New lease. by anol 20180326
Rem out on 20180510
' final SQl 0n 20180526
'If chkIncluseExlessee.Value = 1 Then
'    SQLStr2 = "SELECT Tenants.SageAccountNumber, CompanyName,'' AS UnitNumber,'' AS UnitName " & _
'             "FROM Tenants " & _
'             "LEFT JOIN " & _
'                 "(SELECT LeaseDetails.SageAccountNumber " & _
'                 "FROM LeaseDetails " & _
'                 "WHERE Status=True) AS X ON X.SageAccountNumber=Tenants.SageAccountNumber where " & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') AND X.SageAccountNumber IS NULL " & _
'             "ORDER BY Tenants.SageAccountNumber"
'Else
     SQLStr2 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName,'' as unitName,'' as Unitnumber " & _
            "FROM Tenants LEFT JOIN LeaseDetails ON Tenants.SageAccountNumber=LeaseDetails.SageAccountNumber " & _
            "WHERE LeaseDetails.SageAccountNumber is null  AND " & _
            "(Tenants.Comments IS NULL OR Tenants.Comments = '')  ORDER BY Tenants.SageAccountNumber"
'End If
' final SQl 0n 20180520
'    SQLStr2 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName,'' as unitName,'' as Unitnumber " & _
'"FROM Tenants LEFT JOIN LeaseDetails ON Tenants.SageAccountNumber=LeaseDetails.SageAccountNumber " & _
'"WHERE LeaseDetails.SageAccountNumber is null  AND " & _
'"(Tenants.Comments IS NULL OR Tenants.Comments = '')  ORDER BY Tenants.SageAccountNumber"
'   If chkExpLease.Value = 0 Then
'            'entering new rows in leasedetail table
'            'Here we should not show current lease.This SQL is deducting the current Lesse
'            'it is showing newly created Lesse + Does not show ex-lessee+it should not show exe lesse because chkExpLease is not ticked
'            'below SQL is showing exe-leessee and we don't want that so I remmed out 20180516
'            SQLStr2 = "SELECT Tenants.SageAccountNumber, CompanyName ,'' AS UnitNumber,'' AS UnitName " & _
'             "FROM Tenants " & _
'             "LEFT JOIN " & _
'                 "(SELECT LeaseDetails.SageAccountNumber " & _
'                 "FROM LeaseDetails where status=true " & _
'                 " ) AS X ON X.SageAccountNumber=Tenants.SageAccountNumber where" & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') AND X.SageAccountNumber IS NULL " & _
'             "ORDER BY Tenants.SageAccountNumber"
''
'              SQLStr2 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName ,'' AS UnitNumber,'' AS UnitName " & _
'             "FROM Tenants " & _
'             "LEFT JOIN " & _
'                 "(  Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
'                "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
'                "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false) AS X ON X.SageAccountNumber=Tenants.SageAccountNumber where" & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') AND X.SageAccountNumber IS NULL " & _
'             "ORDER BY Tenants.SageAccountNumber"
'
''                SQLStr2 = "Select SageAccountNumber, CompanyName , UnitNumber, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) as UnitName " & _
''        "FROM (  Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
''        "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
''        "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false) as IM ORDER BY SageAccountNumber"
'
'   Else   'Expired leases only
'          'This is showing max date 1 expired lease details. ?? why ?? because we are taking last status= false and that is the expires lease only
'        SQLStr2 = "Select SageAccountNumber, CompanyName , UnitNumber, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) as UnitName " & _
'        "FROM (  Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
'        "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
'        "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false) as IM ORDER BY SageAccountNumber"
'
''Below SQL is showing current lessee  as  well which is wrong 20180516
''          SQLStr2 = "SELECT LeaseID, LeaseDetails.SageAccountNumber as SageAccountNumber, " & _
''               "Tenants.CompanyName as CompanyName, UnitName, LeaseDetails.UnitNumber, LeaseDetails.Usage, " & _
''               "ClientName, PropertyName, Property.PropertyID " & _
''               "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
''               "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
''               "LeaseDetails.Status = " & IIf(chkExpLease.Value = 0, "True", "False") & " And " & _
''               "Units.PropertyId = Property.PropertyID And " & _
''               "Property.ClientID = Client.ClientID AND " & _
''               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
''               "" & szWhere & " " & _
''                "ORDER BY " & szOrderBy & ";"
'
'
'
'   End If
Rem out on 20180511
''   If chkExpLease.Value = 0 Then
''            SQLStr2 = "SELECT LeaseDetails.leaseID,LeaseDetails.SageAccountNumber,Tenants.CompanyName " & _
''                 "FROM LeaseDetails,Tenants where LeaseDetails.SageAccountNumber=Tenants.SageAccountNumber AND status=true " & _
''             "ORDER BY Tenants.SageAccountNumber"
''   Else   'Expired leases only
''            SQLStr2 = "SELECT LeaseDetails.leaseID,LeaseDetails.SageAccountNumber,Tenants.CompanyName " & _
''                 "FROM LeaseDetails,Tenants where LeaseDetails.SageAccountNumber=Tenants.SageAccountNumber AND status=false " & _
''             "ORDER BY Tenants.SageAccountNumber"
''   End If
                      
        ' Rst2.Close
   Rst2.Open SQLStr2, Conn2, adOpenStatic, adLockReadOnly
'   cboTenant.Clear
'   While Rst2.EOF = False
'       cboTenant.AddItem Rst2!SageAccountNumber & " / " & Rst2!CompanyName
'       Rst2.MoveNext
'   Wend

   ConfigFlxLeaseList
   fraList(1).Visible = True
   fraList(1).Refresh
   fraList(1).ZOrder 0
   
   flxLeaseList.Rows = Rst2.RecordCount + 1
'    flxLeaseList.RowHeight(1) = 0
''   flxLeaseList.Rows = 2
   iRow = 1
   While Not Rst2.EOF
      'Here in the first column I need to bring leaseID (anol)2018/05/11 .no we are adding new
      flxLeaseList.TextMatrix(iRow, 1) = "" 'Rst2!LeaseID & " / " & Rst2!LeaseID
      flxLeaseList.TextMatrix(iRow, 2) = Rst2!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 3) = Rst2!CompanyName
      flxLeaseList.TextMatrix(iRow, 4) = Rst2!UnitNumber '"" 'Unit Number
      flxLeaseList.TextMatrix(iRow, 5) = Rst2!UnitName '"" 'Unit Name
      flxLeaseList.TextMatrix(iRow, 6) = ""
      flxLeaseList.TextMatrix(iRow, 7) = ""
      flxLeaseList.TextMatrix(iRow, 8) = ""
      flxLeaseList.TextMatrix(iRow, 9) = ""
      Rst2.MoveNext
      iRow = iRow + 1
      If iRow = 11 Then
            fraList(1).Visible = True
            fraList(2).Visible = True
            Label16(11).Visible = True
            fraList(1).ZOrder 0
            fraList(2).ZOrder 0
            fraList(2).Refresh
            Label16(11).Refresh
      End If
   Wend
   fraList(2).Visible = False
   Label16(11).Visible = False
   UpdateBalance
   Rst2.Close
   Conn2.Close
   
   FocusControl txtSearchTenant
End Sub

Private Sub cmdTerminate_Click()
   If OpenedLeasePreviewForm Then
      MsgBox "Please close the Lease Preview form.", vbCritical + vbOKOnly, "Terminate lease"
      Exit Sub
   End If
   
   Dim Conn1 As New ADODB.Connection

   If strLeaseId = "" Then
       MsgBox "You must select a lease to terminate", vbOKOnly + vbCritical, "Lease"
       cmdLease.SetFocus
       Exit Sub
   End If

   Conn1.Open getConnectionString

   If Not Clearance4DeleteLease(strLeaseId, Conn1) Then
      MsgBox "This lease cannot be terminated. This leaseholder has demands that are unposted to demand history.", vbInformation + vbOKOnly, "Demand Termination"
      Conn1.Close
      Set Conn1 = Nothing
      Exit Sub
   End If

   If MsgBox("Are you sure you want to terminate the lease for tenant: " & txtTenant.text & "?", vbYesNo + vbQuestion, "Delete Lease") = vbYes Then
        var = ""
        frmTerminationDate.SourceOfCalling = ""
        DisplayDateform Me, "", Date
        cmdTerminate.Visible = False
   End If

End Sub
Public Sub terminate_lease_ERR2()
        Dim DTdate As Date
        On Error GoTo ErrorHandler
        DTdate = Format(var, "dd mmmm yyyy")
        terminate_lease DTdate
        var = ""
        Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
        DisplayDateform Me, "", Format(Date, "dd/mm/yyyy")
   Else
        cmdTerminate.Visible = True
   End If
        
End Sub
Private Sub DisplayDateform(frmMe As Form, szDate As String, szDefaultDate As String)


   frmTerminationDate.szCallingForm = frmMe.Name
   Load frmTerminationDate

   If szDate <> "" Then
      frmTerminationDate.txtTerminationDate.text = szDate
   Else
      frmTerminationDate.txtTerminationDate.text = szDefaultDate
   End If

   frmTerminationDate.Top = frmMe.Top + frmMe.Height / 2 - frmTerminationDate.Height / 2
   frmTerminationDate.Left = frmMe.Left + frmMe.Width / 2 - frmTerminationDate.Width / 2

   frmMe.Enabled = False
   frmTerminationDate.Show
End Sub
Private Sub terminate_lease_ERR()
        Dim DTdate As Date, var
        On Error GoTo ErrorHandler
        var = InputBox("Please type the termination date. (dd/mm/yyyy)", "termination date")
        If var = "" Then Exit Sub
        
        DTdate = Format(var, "dd mmmm yyyy")
        terminate_lease DTdate
        
        Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
        terminate_lease_ERR
   End If
        
End Sub
Private Function Clearance4DeleteLease(szLeaseID As String, adoConn As ADODB.Connection) As Boolean
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT * FROM DemandRecords, LeaseDetails, DemandSplitRecords " & _
           "WHERE LeaseDetails.LeaseID = '" & szLeaseID & "' AND " & _
                " DemandRecords.DemandHistory = False AND " & _
                "((DemandRecords.LeaseRef = LeaseDetails.LeaseID) OR " & _
                " (DemandRecords.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                "  DemandRecords.UnitNumber = LeaseDetails.UnitNumber)) AND " & _
                "DemandRecords.DemandID = DemandSplitRecords.DemandID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Clearance4DeleteLease = adoRST.EOF
   adoRST.Close
   Set adoRST = Nothing
End Function

'Private Sub cmdNDD_Cancel_Click()
'   If (tabLease.Tab = 1) Then
'      txtStopRC.text = flxRentCharges.TextMatrix(RENTCHARGES_EDIT, 16)
'   ElseIf (tabLease.Tab = 4) Then
'      txtStopSC.text = flxSC.TextMatrix(SERVICECHARGES_EDIT, 18)
'   ElseIf (tabLease.Tab = 8) Then
'      txtStopIC.text = flxIns.TextMatrix(INSURANCECHARGES_EDIT, 16)
'   End If
'
'   Frame4(0).Visible = False
'   tabLease.Enabled = True
'   txtNDD.text = ""
'End Sub

Private Sub cmdUsage_Click(Index As Integer)
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim SelUsage As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UUSE"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'UUSE'"
   SelUsage = IIf(cboUsage.text = "", "", cboUsage.Value)
   populateCombo adoConn, sSQLQuery, cboUsage
   cboUsage.text = SelUsage

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxIns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxIns.ToolTipText = flxIns.TextMatrix(flxIns.MouseRow, flxIns.MouseCol)
End Sub

Private Sub flxIns_RowColChange()
   'On Error Resume Next

   
End Sub

Public Sub LoadFormFromTree(szLeaseID As String)
   Dim adoConn As New ADODB.Connection
   Dim adoLease As New ADODB.Recordset
   
'   ConfigFlxLeaseList
   
   adoConn.Open getConnectionString
''
''   szSQL = "SELECT LeaseID, LeaseDetails.SageAccountNumber, " & _
''               "Tenants.CompanyName, UnitName, LeaseDetails.UnitNumber, LeaseDetails.Usage, " & _
''               "ClientName, PropertyName, Property.PropertyID " & _
''           "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
''           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
''               "Units.PropertyId = Property.PropertyID And " & _
''               "Property.ClientID = Client.ClientID AND " & _
''               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber AND " & _
''               "LeaseDetails.LeaseID = '" & szLeaseID & "';"
'''Debug.Print szSQL
''   adoLease.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'''i don't understand why he is loading this grid here anol 20180511
'''   flxLeaseList.Rows = 2
'''   flxLeaseList.TextMatrix(1, 1) = adoLease!LeaseID
'''   flxLeaseList.TextMatrix(1, 2) = adoLease!SageAccountNumber
'''   flxLeaseList.TextMatrix(1, 3) = adoLease!CompanyName
'''   flxLeaseList.TextMatrix(1, 4) = adoLease!UnitName
'''   flxLeaseList.TextMatrix(1, 5) = adoLease!UnitNumber
'''   flxLeaseList.TextMatrix(1, 6) = adoLease!ClientName
'''   flxLeaseList.TextMatrix(1, 7) = adoLease!PropertyName
'''   flxLeaseList.TextMatrix(1, 8) = adoLease!propertyID
'''   flxLeaseList.TextMatrix(1, 9) = IIf(IsNull(adoLease!Usage), "", adoLease!Usage)
''   'flxLeaseList.row = 1
''   Rem out calling this function putting the actual code
''   'LoadingForm adoConn
''   If adoLease.EOF Then
''        adoLease.Close
''        adoConn.Close
''        MsgBox "Could not load the lease information, lease Detail ID incorrect!!", vbInformation, "Warning!"
''        Exit Sub
''   End If
''   fraList(1).Visible = False
''
''   txtLeaseID.text = adoLease!LeaseID 'flxLeaseList.TextMatrix(flxLeaseList.row, 1)
''   txtTenant.Tag = adoLease!SageAccountNumber 'flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
''   txtTenant.text = adoLease!CompanyName 'flxLeaseList.TextMatrix(flxLeaseList.row, 3) 'CompanyName
''   Me.Caption = "Lease Information : " & adoLease!SageAccountNumber & " / " & adoLease!CompanyName
''   txtUnitName(0).text = adoLease!UnitName
''   txtUnitNumber.text = adoLease!UnitNumber 'flxLeaseList.TextMatrix(flxLeaseList.row, 5)
''   txtClient.text = adoLease!ClientName ' flxLeaseList.TextMatrix(flxLeaseList.row, 6)
''   txtProperty.text = adoLease!PropertyName ' flxLeaseList.TextMatrix(flxLeaseList.row, 7)
''   PROPERTY_ID = adoLease!propertyID 'flxLeaseList.TextMatrix(flxLeaseList.row, 8)
''   cboUsage.text = IIf(IsNull(adoLease!Usage), "", adoLease!Usage) 'flxLeaseList.TextMatrix(flxLeaseList.row, 9)

   strLeaseId = szLeaseID
   GetRecord adoConn 'this function loads form information based on txtLeaseID.text
   AllDemandType adoConn
   

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxLeaseList_Click()
    tabLease.Enabled = True
    Frame1(16).Enabled = True
    
    If flxLeaseList.TextMatrix(flxLeaseList.row, 1) = "" And Not cmdCancelNew.Visible Then
        MsgBox "This tenant does not have a lease ID", vbInformation, "Warning"
        fraList(1).Visible = False
        Exit Sub
    End If
    Label16(12).Caption = ""
    txtLeaseEndDate.ForeColor = vbBlack
    Dim adoConn As New ADODB.Connection
    ControlsModeRentCharges DefaultMode
    ControlsModeServiceCharges DefaultMode
    ControlsModeInsuranceCharges DefaultMode
    BreachButtonMode DefaultMode
    If cmdCancelNew.Visible = False Then 'This is nevigation mode on existing lease detail
       
       strLeaseId = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
       txtSageAccountNumber.text = flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
       
       fraList(1).Visible = False
       adoConn.Open getConnectionString
       GetRecord adoConn 'this function loads form information based on txtLeaseID.text
       AllDemandType adoConn
'       LoadDept adoconn 'Load the fund in all charges
'       txtUnitName(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 3)
'       LoadingForm adoConn' i have replaced it with above two functions
       If chkExpLease.Value = 1 Then
          cmdAddNew.Visible = False
    
          cmdTerminate.Visible = False
          'cmdDelete.Visible = False
          
          cmdDelete.Visible = True
         ' cmdCopy.Visible = False
       Else
          cmdAddNew.Visible = True
          cmdUnitNumber(0).Enabled = False
          cmdCopy.Visible = True
          cmdEdit.Visible = True
          cmdTerminate.Visible = True
          cmdDelete.Visible = True
       End If
    
       tabLease.Enabled = True
       cmdCopy.Visible = True
       'Added by anol 11 Nov 2015
    
       Dim rsCheck As New ADODB.Recordset
       If flxLeaseList.TextMatrix(flxLeaseList.row, 8) <> "" Then
                rsCheck.Open "SELECT G.* " & _
                "FROM (GlobalData AS G INNER JOIN Property AS P ON G.PropertyID = P.PropertyID) " & _
                "WHERE P.PropertyID = '" & flxLeaseList.TextMatrix(flxLeaseList.row, 8) & "';", adoConn, adOpenStatic, adLockReadOnly
                If Not rsCheck.EOF Then
                    txtYearEnd.text = IIf(IsNull(rsCheck.Fields("SCYearEnd").Value), "", rsCheck.Fields("SCYearEnd").Value)
                End If
                rsCheck.Close
       End If
       adoConn.Close
       Set adoConn = Nothing
  Else
      'Here you are creating new lease,hence you must clear unit
      'txtTenant.text = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
      txtTenant.Tag = flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
      txtSageAccountNumber.text = flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
      txtTenant.text = flxLeaseList.TextMatrix(flxLeaseList.row, 3) 'CompanyName
      txtUnitName(0).text = ""
      txtUnitNumber.text = ""
      strLeaseId = "" ' ID shall be produced automatically when you select a Unit
      
      fraList(1).Visible = False
      FocusControl cmdUnitNumber(0)
      'cmdUnitNumber.Set
  End If
 
End Sub
'Private Function FocusCommandButton(ctr As CommandButton)
''in general case when fails to focus the control it come up with an error
''Written by anol 20161114
'    On Error GoTo ERR
'        ctr.SetFocus
'    Exit Function
'ERR:
'End Function
'Private Sub LoadingForm(adoConn As ADODB.Connection) ' I have made this obsolte function anol 20180511
'   'Dim szaData(1, 0) As String
'
'   'Call EmptyBoxes
'
'   fraList(1).Visible = False
'
'   txtLeaseID.text = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
'   txtTenant.Tag = flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
'   txtTenant.text = flxLeaseList.TextMatrix(flxLeaseList.row, 3) 'CompanyName
'   Me.Caption = "Lease Information : " & flxLeaseList.TextMatrix(flxLeaseList.row, 2) & " / " & flxLeaseList.TextMatrix(flxLeaseList.row, 3)
'   txtUnitName(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 4)
''   szaData(0, 0) = flxLeaseList.TextMatrix(flxLeaseList.row, 5)
''   szaData(1, 0) = flxLeaseList.TextMatrix(flxLeaseList.row, 5)
''   cboUnit.Clear
''   cboUnit.Column() = szaData()
''   cboUnit.ListIndex = 0
'   txtUnitNumber.text = flxLeaseList.TextMatrix(flxLeaseList.row, 5)
'   'txtUnitName(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 5)
'
'
''MsgBox txtUnitNumber.text
'   txtClient.text = flxLeaseList.TextMatrix(flxLeaseList.row, 6)
'   txtProperty.text = flxLeaseList.TextMatrix(flxLeaseList.row, 7)
'   PROPERTY_ID = flxLeaseList.TextMatrix(flxLeaseList.row, 8)
'   cboUsage.text = flxLeaseList.TextMatrix(flxLeaseList.row, 9)
'
'   GetRecord adoConn ' this function loads form information based on txtLeaseID.text
'   AllDemandType adoConn
'End Sub

Private Sub flxRentAnalysis_RowColChange()
   'populateControl Me, flxRentAnalysis ' this is a very old and obsolate loading method

   If Val(flxRentAnalysis.TextMatrix(flxRentAnalysis.row, 1)) > 0 Then
      RentReviewButtonMode GridRowOnSelection
   End If
End Sub

Private Sub flxSC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'flxSC.ToolTipText = flxSC.TextMatrix(flxSC.MouseRow, flxSC.MouseCol)
   
    If flxSC.MouseCol = 6 Then
         flxSC.ToolTipText = flxSC.TextMatrix(flxSC.MouseRow, 19)
    Else
         flxSC.ToolTipText = ""
    End If
End Sub

Private Sub flxSC_RowColChange()
'   On Error Resume Next

   
'
'ErrHanlder:
'   MsgBox "Data Missing. Please check data!", vbCritical + vbOKOnly, "Data Missing.."
End Sub

'Private Sub UpdateDatabase_LeaseDetails(adoConn As ADODB.Connection)
'    On Error GoTo Err
'    Dim rst1 As New ADODB.Recordset
'    rst1.Open "Select IncreamentalID from LeaseDetails", adoConn
'    rst1.Close
'    Exit Sub
'Err:
'    adoConn.Execute "ALTER TABLE LeaseDetails ADD COLUMN IncrementalID Long;"
'    adoConn.Execute "Update LeaseDetails set IncrementalID=right(LeaseID,12)"
'End Sub
Private Sub LoadFunds()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1965
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup(0).Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup(0).Left + cmdGridUnitLookup(0).Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Fund Code"
   lblClientID(1).Caption = "Fund Name"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup(0).Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoConn.Open getConnectionString
   'szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoConn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
   Else
        iSel = 1
        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
                PROPERTY_ID & "' and ClientID='" & txtClient.Tag & "' and isDeleted=false"
   End If
   rsFundMatrix.Close
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxClientList.Clear
      flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundCode").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundName").Value
                    flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundID").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub Form_Load()
    Me.Height = 8535
    Me.Width = 18555
    Dim adoConn As New ADODB.Connection
    Frame1(1).BackColor = MODULEBACKCOLOR
    Frame1(2).BackColor = MODULEBACKCOLOR
    Frame1(3).BackColor = MODULEBACKCOLOR
    Frame1(4).BackColor = MODULEBACKCOLOR
    Frame1(5).BackColor = MODULEBACKCOLOR
    Frame1(6).BackColor = MODULEBACKCOLOR
    Frame1(7).BackColor = MODULEBACKCOLOR
    Frame1(8).BackColor = MODULEBACKCOLOR
    Frame1(9).BackColor = MODULEBACKCOLOR
    Frame1(10).BackColor = MODULEBACKCOLOR
    Frame1(11).BackColor = MODULEBACKCOLOR
    Frame1(12).BackColor = MODULEBACKCOLOR
    Frame1(13).BackColor = MODULEBACKCOLOR
    Frame1(14).BackColor = MODULEBACKCOLOR
    Frame1(15).BackColor = MODULEBACKCOLOR
    lblDefaultDescption(4).BackColor = MODULEBACKCOLOR
    chkSCDes.BackColor = MODULEBACKCOLOR
    chkSubLease.BackColor = MODULEBACKCOLOR
    chkHoldingOver.BackColor = MODULEBACKCOLOR
    chkOLED.BackColor = MODULEBACKCOLOR
    chkGPrataDmd.BackColor = MODULEBACKCOLOR
    Label8(0).BackColor = MODULEBACKCOLOR
    Label8(1).BackColor = MODULEBACKCOLOR
    Label8(6).BackColor = MODULEBACKCOLOR
    Label8(7).BackColor = MODULEBACKCOLOR
    lblRentCharges(0).BackColor = MODULEBACKCOLOR
    lblRentCharges(1).BackColor = MODULEBACKCOLOR
    lblRentCharges(2).BackColor = MODULEBACKCOLOR
    lblRentCharges(3).BackColor = MODULEBACKCOLOR
    lblRentCharges(4).BackColor = MODULEBACKCOLOR
    chkRentDes.BackColor = MODULEBACKCOLOR
    chkInsDes.BackColor = MODULEBACKCOLOR
    
    lblRentCharges(7).BackColor = MODULEBACKCOLOR
    lblRentCharges(8).BackColor = MODULEBACKCOLOR
    lblRentCharges(9).BackColor = MODULEBACKCOLOR
    Label17(0).BackColor = MODULEBACKCOLOR
    Label17(1).BackColor = MODULEBACKCOLOR
    Label17(2).BackColor = MODULEBACKCOLOR
    Label17(3).BackColor = MODULEBACKCOLOR
    Label17(4).BackColor = MODULEBACKCOLOR
    lblInc(0).BackColor = MODULEBACKCOLOR
    lblInc(1).BackColor = MODULEBACKCOLOR
    lblInc(2).BackColor = MODULEBACKCOLOR
    lblInc(3).BackColor = MODULEBACKCOLOR
    lblInc(4).BackColor = MODULEBACKCOLOR
    lblInc(5).BackColor = MODULEBACKCOLOR
    lblInc(6).BackColor = MODULEBACKCOLOR
    lblInc(7).BackColor = MODULEBACKCOLOR
    lblInc(8).BackColor = MODULEBACKCOLOR
    lblInc(9).BackColor = MODULEBACKCOLOR
    Label17(6).BackColor = MODULEBACKCOLOR
    Label17(7).BackColor = MODULEBACKCOLOR
    Label17(8).BackColor = MODULEBACKCOLOR
    Label17(5).BackColor = MODULEBACKCOLOR
    Label44(0).BackColor = MODULEBACKCOLOR
    Label44(1).BackColor = MODULEBACKCOLOR
    Label44(2).BackColor = MODULEBACKCOLOR
    Label44(3).BackColor = MODULEBACKCOLOR
    Label44(4).BackColor = MODULEBACKCOLOR
    Label44(5).BackColor = MODULEBACKCOLOR
    Label44(6).BackColor = MODULEBACKCOLOR
    optAutoIntCal.BackColor = MODULEBACKCOLOR
    optManIntCal.BackColor = MODULEBACKCOLOR
    Label10(3).BackColor = MODULEBACKCOLOR
    Label10(1).BackColor = MODULEBACKCOLOR
    optLSR.BackColor = MODULEBACKCOLOR
    optGIR.BackColor = MODULEBACKCOLOR
    Label10(0).BackColor = MODULEBACKCOLOR
    Label10(6).BackColor = MODULEBACKCOLOR
  
    szLeaseStatus = False
    adoConn.Open getConnectionString
    chkMultipleLH.Enabled = True
    chkExpLease.Enabled = True
    
    
    cmdTenants.Left = cmdLease.Left
    Label16(12).Caption = ""
    Label16(12).Tag = ""
    'Call UpdateDatabase_LeaseDetails(adoConn)
    If Not AllDemandType(adoConn) Then
        MsgBox "You have not defined any demand types. Please create demand types within Global Data.", vbInformation + vbOKOnly, "Demand Type"
        FormLoad = False
        
        adoConn.Close
        Set adoConn = Nothing
        Exit Sub
    End If

'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
  
   Me.BackColor = MODULEBACKCOLOR
   tabLease.BackColor = MODULEBACKCOLOR
   Frame1(16).BackColor = MODULEBACKCOLOR
   chkExpLease.BackColor = Frame1(16).BackColor
   chkMultipleLH.BackColor = MODULEBACKCOLOR

   tabLease.Tab = 0

   'On Error GoTo ErrorTrap

   ConfigFlxRentAnalysis
   ConfigFlxRentCharges
   ConfigFlxSC
   ConfigFlxInsurance
   ConfigGridBreach
   ConfigAssignmentGrid
   ConfigflxLicences

   Call EmptyBoxes
   Call DisableBoxes
   Call FillCbos(adoConn)
   LoadValues adoConn

   BreachButtonMode DefaultMode
   AssignmentButtonMode DefaultMode
   RentReviewButtonMode DefaultMode
   ControlsModeRentCharges DefaultMode
   ControlsModeServiceCharges DefaultMode
   ControlsModeInsuranceCharges DefaultMode

   FormLoad = True
'   LoadDept adoconn
'This mouse wheel . now this is causing a problem. runtime
   'Call WheelHook(Me.hWnd)

  'added by anol 22 Sep 2015
   Dim Rst1 As New ADODB.Recordset
Add_Column_Cap:
   On Error GoTo Missing_Column_Cap
   Rst1.Open "SELECT CapAmount FROM LeaseDetails;", adoConn, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo NormalOperation
Missing_Column_Cap:
   adoConn.Execute "ALTER TABLE LeaseDetails Add column CapAmount Number"
  
NormalOperation:

   adoConn.Close
   Set adoConn = Nothing
 If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
   Call WheelHook(Me.hWnd)
 End If
'   cboUnit.Column() = szaData()
Exit Sub
ErrorTrap:
   Set adoConn = Nothing
   
   If Err.Number > 0 Then
      If Err.Number = 40002 Then
         If MsgBox("DSN - " & Adsn & " not found. Please check with your system administrator.", vbRetryCancel + vbCritical, "DSN Set Up Error") = vbRetry Then
            Resume
         Else
            Exit Sub
         End If
      Else
         MsgBox Err.Number & " - " & Err.description
         Exit Sub
      End If
   End If
End Sub

Private Sub ConfigFlxInsurance()
   Dim szHeader As String

   flxIns.Cols = 20
   flxIns.Rows = 2
   flxIns.Clear
   szHeader$ = "lblInc|<InsuranceDemandType|<InsuranceStartDate|<InsuranceFrequency|<InsuranceNextDueDate" & _
               "|<InsuranceDept|<ChargingType|>ChargingFigure|>TotalYearlyInsurance|>InsuranceEachPeriod" & _
               "|<Department|<FrequencyID|<Demand type|Charging Method|Description|DELETE|<StopDate"
   flxIns.FormatString = szHeader$
   flxIns.RowHeight(0) = 0
   flxIns.ColWidth(0) = 0
   flxIns.ColWidth(1) = lblInc(1).Left - lblInc(0).Left                       'Demand Type
   flxIns.ColWidth(2) = lblInc(2).Left - lblInc(1).Left                       'Start Date
   flxIns.ColWidth(3) = lblInc(3).Left - lblInc(2).Left                       'frequency
   flxIns.ColWidth(4) = lblInc(4).Left - lblInc(3).Left                       'next due date
   flxIns.ColWidth(5) = lblInc(5).Left - lblInc(4).Left                       'Fund
   flxIns.ColWidth(6) = lblInc(6).Left - lblInc(5).Left                       'Changing method
   flxIns.ColWidth(7) = lblInc(7).Left - lblInc(6).Left                       'Amount
   flxIns.ColWidth(8) = lblInc(8).Left - lblInc(7).Left                       'yearly charge
   flxIns.ColWidth(9) = lblInc(9).Left - lblInc(8).Left                       'Each period
   flxIns.ColWidth(10) = 0                                                    'Department id
   flxIns.ColWidth(11) = 0                                                    'FrequencyID
   flxIns.ColWidth(12) = 0                                                    'InsuranceDemandType
   flxIns.ColWidth(13) = 0                                                    'Charging Method
   flxIns.ColWidth(14) = 0                                                    'Description
   flxIns.ColWidth(15) = 0                                                    'Delete
   flxIns.ColWidth(16) = flxIns.Left + flxIns.Width - lblInc(9).Left - 60     'Stop Date
   flxIns.ColWidth(17) = 0 'FundCode
   flxIns.ColWidth(18) = 0 'Fund Name
   flxIns.ColWidth(19) = 0 'Demand Type Description
End Sub

Private Sub LoadFlxRentAnalysis()
   Dim iRow As Integer

   Conn2.Open getConnectionString

   'get all sage account numbers and company names from tenants.
   SQLStr2 = "SELECT RentAnalysis.*, DemandTypes.Type " & _
             "FROM RentAnalysis, DemandTypes " & _
             "WHERE LeaseID = '" & flxLeaseList.TextMatrix(flxLeaseList.row, 1) & "' AND " & _
                  "RentAnalysis.RRDemandType = DemandTypes.ID " & _
             "ORDER BY RentAnalysis.ID ASC"
   Rst2.Open SQLStr2, Conn2, adOpenStatic, adLockReadOnly

   flxRentAnalysis.Clear
   ConfigFlxRentAnalysis

   flxRentAnalysis.Rows = 2
   If Not Rst2.EOF Then
      iRow = 1
      While Not Rst2.EOF
         flxRentAnalysis.TextMatrix(iRow, 1) = IIf(IsNull(Rst2!SerialNumber), "", Rst2!SerialNumber)
         flxRentAnalysis.TextMatrix(iRow, 2) = IIf(IsNull(Rst2!Type), "", Rst2!Type)
         flxRentAnalysis.TextMatrix(iRow, 3) = IIf(IsNull(Rst2!RRDemandType), "", Rst2!RRDemandType)
         flxRentAnalysis.TextMatrix(iRow, 4) = IIf(IsNull(Rst2!RentReviewDate), "", Format(Rst2!RentReviewDate, "dd/mm/yyyy"))
         flxRentAnalysis.TextMatrix(iRow, 5) = IIf(IsNull(Rst2!Comments), "", Rst2!Comments)
         flxRentAnalysis.TextMatrix(iRow, 6) = IIf(IsNull(Rst2!RentIncreaseDate), "", Format(Rst2!RentIncreaseDate, "dd/mm/yyyy"))
         'flxRentAnalysis.TextMatrix(iRow, 7) = IIf(IsNull(Rst2!RentIncreaseAmount), "", Rst2!RentIncreaseAmount)
         'resolved by BOSL
         '0000478: Lease Details Crash
         'Modified by anol 25 sep 2014
         flxRentAnalysis.TextMatrix(iRow, 7) = IIf(IsNull(Rst2!RentIncreaseAmount), "0.00", Format(Rst2!RentIncreaseAmount, "0.00"))
         flxRentAnalysis.TextMatrix(iRow, 8) = Rst2!Id
         flxRentAnalysis.TextMatrix(iRow, 9) = IIf(Rst2!Reminder_ID <> "", "YES", "NO")
         flxRentAnalysis.TextMatrix(iRow, 10) = IIf(IsNull(Rst2!Reminder_ID), "", Rst2!Reminder_ID)
         flxRentAnalysis.TextMatrix(iRow, 11) = IIf(IsNull(Rst2!RRStatus), "NO", IIf(Rst2!RRStatus = "Y", "YES", "NO"))
          flxRentAnalysis.TextMatrix(iRow, 12) = IIf(IsNull(Rst2!Status), "NO", IIf(Rst2!Status = "Y", "YES", "NO"))
         Rst2.MoveNext
         If Not Rst2.EOF Then flxRentAnalysis.AddItem ""
         iRow = iRow + 1
      Wend
   End If

   Rst2.Close
   Conn2.Close
   Set Rst2 = Nothing

   flxRentAnalysis.row = 0
   flxRentAnalysis.col = 0
End Sub

Private Sub LoadFlxRentCharges(adoConn As ADODB.Connection)
   Dim iRow As Integer, szaTemp() As String
   Dim adoRst1 As ADODB.Recordset

   Set adoRst1 = New ADODB.Recordset

'   get all sage account numbers and company names from tenants.
   szSQL = "SELECT RentCharges, RentChargeDept, BRFrequency, BRStartDate, " & _
               "BRNextDueDate, BRTotal, BRAmount, BRDemandType, RentDesc, " & _
               "Type, L.spare1, C.ChargingMethod, " & _
               "L.spare2, L.StopRC,RentChargeDept,F.FundCode,F.FundName,FR.Frequency " & _
             "FROM LRentCharges L, DemandTypes D, ChargingMethod C,Fund F ,Frequencies FR " & _
             "WHERE LeaseID = '" & strLeaseId & "' And (ISNULL(L.spare3) OR L.spare3 = '') AND " & _
               "L.BRDemandType = D.ID And F.FundID=Cint(L.RentChargeDept) AND  FR.ID=Cint(L.BRfrequency) AND " & _
               "C.ChargingMethodID = CINT(IIF(ISNULL(L.spare1),3,L.spare1)) And " & _
               "(ISNULL(L.spare3) OR L.spare3 = '') " & _
             "ORDER BY RentCharges ASC"
'Debug.Print szSQL
   adoRst1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   Call ConfigFlxRentCharges
   If Not adoRst1.EOF Then
      iRow = 1
      While Not adoRst1.EOF
         flxRentCharges.RowHeight(iRow) = 240
         flxRentCharges.TextMatrix(iRow, 0) = IIf(IsNull(adoRst1!BRDemandType), "", adoRst1!BRDemandType) 'BR demand Type ID
         flxRentCharges.TextMatrix(iRow, 1) = IIf(IsNull(adoRst1!BRDemandType), "", DemandType(adoConn, adoRst1!BRDemandType)) 'Demand Type name
         flxRentCharges.TextMatrix(iRow, 2) = IIf(IsNull(adoRst1!BRStartDate), "", Format(adoRst1!BRStartDate, "dd/mm/yyyy"))
         flxRentCharges.TextMatrix(iRow, 3) = IIf(IsNull(adoRst1!BRfrequency), "", adoRst1!BRfrequency)
         flxRentCharges.TextMatrix(iRow, 4) = IIf(IsNull(adoRst1!Frequency), "", adoRst1!Frequency)
         flxRentCharges.TextMatrix(iRow, 5) = IIf(IsNull(adoRst1!BRNextDueDate), "", Format(adoRst1!BRNextDueDate, "dd/mm/yyyy"))
         flxRentCharges.TextMatrix(iRow, 6) = IIf(IsNull(adoRst1!RentChargeDept), "", adoRst1!RentChargeDept) 'FundID
         flxRentCharges.TextMatrix(iRow, 7) = IIf(IsNull(adoRst1!FundCode), "", adoRst1!FundCode) 'IIf(IsNull(adoRst1!RentChargeDept), "", DeptName(IIf(IsNull(adoRst1!RentChargeDept) Or adoRst1!RentChargeDept = "", 999, adoRst1!RentChargeDept)))
         flxRentCharges.TextMatrix(iRow, 8) = IIf(IsNull(adoRst1!spare1) Or adoRst1!spare1 = "", "3", adoRst1!spare1)
         flxRentCharges.TextMatrix(iRow, 9) = IIf(IsNull(adoRst1!ChargingMethod) Or adoRst1!ChargingMethod = "", "Annual", adoRst1!ChargingMethod)
         'issue 478: Lease Details Crash
         'Modified by anol 25 Sep 2014
         flxRentCharges.TextMatrix(iRow, 10) = IIf(IsNull(adoRst1!spare2) Or adoRst1!spare2 = "", IIf(IsNull(adoRst1!spare2), "0.00", Format(adoRst1!spare2, "0.00")), Format(adoRst1!spare2, "0.00"))
         'flxRentCharges.TextMatrix(iRow, 10) = IIf(IsNull(adoRst1!spare2) Or adoRst1!spare2 = "", adoRst1!BRTotal, Format(adoRst1!spare2, "0.00")) 'original Code
         flxRentCharges.TextMatrix(iRow, 11) = IIf(IsNull(adoRst1!BRTotal), "0.00", Format(adoRst1!BRTotal, IIf(Val(flxRentCharges.TextMatrix(iRow, 10)) = 2, "0.0000", "0.00")))
         'End of modification
         flxRentCharges.TextMatrix(iRow, 12) = Format(IIf(IsNull(adoRst1!BRAmount), "0.00", (adoRst1!BRAmount)), "0.00")
         flxRentCharges.TextMatrix(iRow, 13) = adoRst1!RentCharges
         flxRentCharges.TextMatrix(iRow, 14) = IIf(IsNull(adoRst1!RentDesc), "", adoRst1!RentDesc)
         If Not IsNull(adoRst1!StopRC) Or adoRst1!StopRC <> "" Then
            flxRentCharges.TextMatrix(iRow, 16) = Format(adoRst1!StopRC, "dd/mm/yyyy")
         Else
            flxRentCharges.TextMatrix(iRow, 16) = ""
         End If
         flxRentCharges.TextMatrix(iRow, 17) = IIf(IsNull(adoRst1!FundName), "", adoRst1!FundName)
         flxRentCharges.TextMatrix(iRow, 18) = IIf(IsNull(adoRst1!Type), "", adoRst1!Type) 'Demand Type Name/Description
         flxRentCharges.TextMatrix(iRow, 19) = IIf(IsNull(adoRst1!Frequency), "", adoRst1!Frequency) 'Frequency description
         adoRst1.MoveNext
         If Not adoRst1.EOF Then flxRentCharges.AddItem ""
         iRow = iRow + 1
      Wend
   End If

   flxRentCharges.row = 0
   flxRentCharges.col = 0

   adoRst1.Close
   Set adoRst1 = Nothing
End Sub

Private Sub ConfigFlxRentAnalysis()
   Dim szFlxHeader As String

   flxRentAnalysis.RowHeight(0) = 0
   flxRentAnalysis.Clear
   flxRentAnalysis.Cols = 13
   szFlxHeader$ = "|<Serial|<DemandType|<RRDemandType|<RentReviewDate|<Comments|<RentIncreaseDate" & _
                  "|>RentIncreaseAmount|ID|<Alarm|<AlarmID|<RRStatus"
   flxRentAnalysis.FormatString = szFlxHeader$

   flxRentAnalysis.ColWidth(0) = 0                                      '
   flxRentAnalysis.ColWidth(1) = Label8(1).Left - Label8(0).Left        'Serial
   flxRentAnalysis.ColWidth(2) = Label8(2).Left - Label8(1).Left        'Demand Type
   flxRentAnalysis.ColWidth(3) = 0                                      'Demand Type ID
   flxRentAnalysis.ColWidth(4) = Label8(3).Left - Label8(2).Left        'RentReviewDate
   flxRentAnalysis.ColWidth(5) = Label8(4).Left - Label8(3).Left        'Comments
   flxRentAnalysis.ColWidth(6) = Label8(5).Left - Label8(4).Left        'RentIncreaseDate
   flxRentAnalysis.ColWidth(7) = Label8(6).Left - Label8(5).Left        'RentIncreaseAmount
   flxRentAnalysis.ColWidth(8) = 0                                      'Y/N
   flxRentAnalysis.ColWidth(9) = Label8(7).Left - Label8(6).Left        'Alarm
   flxRentAnalysis.ColWidth(10) = 0                                     'AlarmID
   flxRentAnalysis.ColWidth(11) = Label8(6).Width                       'Status
   flxRentAnalysis.ColWidth(12) = 0                                     'Demand runned Action status
End Sub

Private Sub ConfigFlxRentCharges()
   Dim szFlxHeader As String

   flxRentCharges.RowHeight(0) = 0
   flxRentCharges.Clear
   flxRentCharges.Rows = 2
   flxRentCharges.Cols = 20
   szFlxHeader$ = "BRDemandTypeID|<BRDemandType|<RentStartDate|<FreqBR|<Freq|<NextDueDate" & _
                  "|<RentChargeDeptID|<RentChargeDept||<ChargingMethod|>Amt|>TotalRentYear" & _
                  "|>RentDueEachPeriod|<LeaseIDEdit|<DELETED|<StopDate"
   flxRentCharges.FormatString = szFlxHeader$

   flxRentCharges.ColWidth(0) = 0                                                'Demand type
   flxRentCharges.ColWidth(1) = lblRentCharges(1).Left - lblRentCharges(0).Left  'Demand type
   flxRentCharges.ColWidth(2) = lblRentCharges(2).Left - lblRentCharges(1).Left  'Start date
   flxRentCharges.ColWidth(3) = 0                                                'Frequency
   flxRentCharges.ColWidth(4) = lblRentCharges(3).Left - lblRentCharges(2).Left  'Frequency
   flxRentCharges.ColWidth(5) = lblRentCharges(4).Left - lblRentCharges(3).Left  'Next due date
   flxRentCharges.ColWidth(6) = 0                                                'Fund ID
   flxRentCharges.ColWidth(7) = lblRentCharges(5).Left - lblRentCharges(4).Left  'Fund
   flxRentCharges.ColWidth(8) = 0                                                      'Charging Method ID
   flxRentCharges.ColWidth(9) = lblRentCharges(6).Left - lblRentCharges(5).Left - 10   'Charging Method
   flxRentCharges.ColWidth(10) = lblRentCharges(7).Left - lblRentCharges(6).Left + 600 'Amount / Precentage
   flxRentCharges.ColWidth(11) = lblRentCharges(8).Left - lblRentCharges(7).Left - 550  'Total for year
   flxRentCharges.ColWidth(12) = lblRentCharges(9).Left - lblRentCharges(8).Left - 90 'Due each period

   flxRentCharges.ColWidth(13) = 0
   flxRentCharges.ColWidth(14) = 0
   flxRentCharges.ColWidth(15) = 0        'DELETED NOTATION
   flxRentCharges.ColWidth(16) = flxRentCharges.Left + flxRentCharges.Width - lblRentCharges(9).Left - 60 'stop date
   flxRentCharges.ColWidth(17) = 0 'Fund Name
   flxRentCharges.ColWidth(18) = 0 'Demand Type ID
   flxRentCharges.ColWidth(19) = 0 'Frequency Description
    
End Sub

Public Sub SetAddNewMode()
   Dim temp As String
'   Dim Conn1 As New ADODB.Connection
'
'   Conn1.Open getConnectionString
'
'   szSQL = "SELECT SageAccountNumber, CompanyName " & _
'             "FROM LeaseDetails " & _
'             "ORDER BY SageAccountNumber"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   cboTenant.Clear
'   If Not Rst1.EOF Then
'      While Rst1.EOF = False
'          cboTenant.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
'          Rst1.MoveNext
'      Wend
'   Else
'      MsgBox "Please input new TENANT into tenant form.", vbInformation + vbOKOnly, "No Tenant"
'   End If
'   Rst1.Close
'   Conn1.Close

   cmdAddNew.Visible = True
   cmdUnitNumber(0).Enabled = False
   cmdAddNew.TabIndex = 25
   cmdTerminate.Visible = True
'   cmdDelete.Visible = True
   cmdTerminate.TabIndex = 26
   cmdEdit.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
End Sub

Public Sub GetTenantsWithoutLease()
   Dim temp As String
   Dim iRow As Integer
   Conn2.Open getConnectionString

   'get all sage account numbers and company names from tenants.
'   SQLStr2 = "SELECT SageAccountNumber, CompanyName " & _
'             "FROM Tenants " & _
'             "WHERE Tenants.SageAccountNumber NOT IN " & _
'                 "(SELECT LeaseDetails.SageAccountNumber " & _
'                 "FROM LeaseDetails " & _
'                 " ) AND " & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
'             "ORDER BY SageAccountNumber"
' issue 559 2)  It is not showing the expired lessee Name in the list where we are creating
'New lease. 20180326 here deducting the current lessee so that it shows ex-lessee+newly created lessee
'this SQL is wrong because it does not determine whether this is an ex-ticked or not
  SQLStr2 = "SELECT Tenants.SageAccountNumber, CompanyName " & _
             "FROM Tenants " & _
             "LEFT JOIN " & _
                 "(SELECT LeaseDetails.SageAccountNumber " & _
                 "FROM LeaseDetails where status=true " & _
                 " ) AS X ON X.SageAccountNumber=Tenants.SageAccountNumber where" & _
                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') AND X.SageAccountNumber IS NULL " & _
             "ORDER BY Tenants.SageAccountNumber"
             
             
   Rst2.Open SQLStr2, Conn2, adOpenStatic, adLockReadOnly
'   cboTenant.Clear
'   While Rst2.EOF = False
'       cboTenant.AddItem Rst2!SageAccountNumber & " / " & Rst2!CompanyName
'       Rst2.MoveNext
'   Wend
    ConfigFlxLeaseList
    flxLeaseList.Rows = Rst2.RecordCount + 1
'    flxLeaseList.RowHeight(1) = 0
''   flxLeaseList.Rows = 2
    iRow = 1
   While Not Rst2.EOF
        MsgBox "Debug!!"
      flxLeaseList.TextMatrix(iRow, 1) = Rst2!SageAccountNumber & " / " & Rst2!CompanyName
      flxLeaseList.TextMatrix(iRow, 2) = Rst2!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 3) = Rst2!CompanyName
      flxLeaseList.TextMatrix(iRow, 4) = ""
      flxLeaseList.TextMatrix(iRow, 5) = ""
      flxLeaseList.TextMatrix(iRow, 6) = ""
      flxLeaseList.TextMatrix(iRow, 7) = ""
      flxLeaseList.TextMatrix(iRow, 8) = ""
      flxLeaseList.TextMatrix(iRow, 9) = ""
      Rst2.MoveNext
      iRow = iRow + 1
   Wend
'WHERE Status=True
'comment out by anol 21 Mar 2016
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

Public Sub FillCbos(Conn1 As ADODB.Connection)
   Dim i As Integer, Data() As String

   'Fill the yes / no cbos
   cboIntCrgable.AddItem "No", 0
   cboIntCrgable.AddItem "Yes", 1
   cboBreakClause.AddItem "No", 0
   cboBreakClause.AddItem "Yes", 1

   szSQL = "SELECT * FROM Frequencies"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = Rst1.RecordCount
   If Rst1.EOF = False Then
      ReDim Preserve Data(1, i) As String
      i = 0
      While Rst1.EOF = False
         Data(0, i) = Rst1!Id
         Data(1, i) = Rst1!Frequency
         i = i + 1
         Rst1.MoveNext
      Wend
'      cboFreqBR.Clear
'      cboFreqSC.Clear
'      cboInsFreq.Clear

'      cboFreqBR.Column() = Data()
'      cboFreqSC.Column() = Data()
'      cboInsFreq.Column() = Data()
   End If

   '' Fill the Head leases
   Rst1.Close
   
''    szSQL = "SELECT LeaseID,UnitNumber FROM LeaseDetails"
''    Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''
''    Dim TotalRow As Integer
''    Dim TotalCol As Integer
''    Dim j As Integer
''    TotalRow = Rst1.RecordCount - 1
''    TotalCol = Rst1.Fields.count - 1
''
''    ReDim szaProperty(TotalCol, TotalRow) As String
''
''    For i = 0 To TotalRow
''       For j = 0 To TotalCol
''           szaProperty(j, i) = IIf(IsNull(Rst1.Fields(j).Value), "", Rst1.Fields(j).Value)
''       Next j
''       Rst1.MoveNext
''       If Rst1.EOF Then Exit For
''    Next i
''
''   cboHeadLease.Column() = szaProperty()
''
''   Rst1.Close

   'fill the type of store cbo.
   LoadType "LTYP", cboType

   'fill the break type cbo.
   cboBreak.AddItem "Landlord", 0
   cboBreak.AddItem "Tenant", 1
   cboBreak.AddItem "Mutual", 2

   'Set the RDO Connections to the dataset
   szSQL = "SELECT * FROM ChargingMethod"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = 0
   If Rst1.EOF = False Then
      While Rst1.EOF = False
         ReDim Preserve Data(1, Rst1!ChargingMethodID) As String
         Data(0, Rst1!ChargingMethodID - 1) = CStr(Rst1!ChargingMethodID)
         Data(1, Rst1!ChargingMethodID - 1) = CStr(Rst1!ChargingMethod)
         Rst1.MoveNext
      Wend
      cboSCChargingMth.Clear
      cboBRChargingMth.Clear
      cboIncCharMth.Clear
      cboSCChargingMth.Column() = Data()
      cboBRChargingMth.Column() = Data()
      cboIncCharMth.Column() = Data()
   End If

   Rst1.Close
   Set Rst1 = Nothing
   
'   FillComboByCode "ICRG", cboIncCharMth, Conn1

'*************************************************   BREACHES COMBO
   'Set the RDO Connections to the dataset
   szSQL = "SELECT SecondaryCode.Value as V, SecondaryCode.CODE as C  " & _
           "FROM   SecondaryCode " & _
           "WHERE  PrimaryCode = 'BTYP' " & _
           "ORDER BY Value;"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = Rst1.RecordCount
   If i > 0 Then
      ReDim Data(1, i - 1) As String

      i = 0
      While Not Rst1.EOF
         Data(0, i) = CStr(Rst1!c)
         Data(1, i) = CStr(Rst1!V)
         Rst1.MoveNext
         i = i + 1
      Wend
      cboBreachType.Clear
      cboBreachType.Column() = Data()
   End If
   Rst1.Close
'activate the following codes when you will have Schedule in the menu bar like sage version
'************************************* SCHEDULE COMBO
   szSQL = "SELECT * FROM Schedule;"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = Rst1.RecordCount
   If i > 0 Then
      ReDim Data(1, i) As String

      i = 0
      While Not Rst1.EOF
         Data(0, i) = CStr(Rst1!ScheduleID)
         Data(1, i) = CStr(Rst1!ScheduleName)
         Rst1.MoveNext
         i = i + 1
      Wend

      If i > 1 Then
         Data(0, i) = Data(0, i - 1) + 1
         Data(1, i) = "MULTIPLE"
      End If
      cboSchedule.Clear
      cboSchedule.Column() = Data()
   End If

   Rst1.Close
   Set Rst1 = Nothing
End Sub

Private Sub LoadType(szValue As String, conCombo As Control)
   If Conn2.State = 0 Then
   Conn2.Open getConnectionString
   End If

   szSQL = "SELECT CODE, VALUE " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = '" & szValue & "' " & _
           "ORDER BY Value;"
'   populateCombo Conn2, szSQL, conCombo
'
   Rst2.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   conCombo.Clear
   While Not Rst2.EOF
      conCombo.AddItem Rst2!Value
      Rst2.MoveNext
   Wend

   Rst2.Close
   Conn2.Close
   Set Rst2 = Nothing
End Sub

Public Sub GetRecord(adoConn1 As ADODB.Connection) 'this function loads form information based on txtLeaseID.text
   Dim i As Integer, szMsg As String
   Dim adoRst1 As New ADODB.Recordset

   ConfigFlxRentAnalysis
   ConfigFlxRentCharges
   ConfigFlxSC
   ConfigFlxInsurance
   ConfigGridBreach

'   Get record for selected Tenant.
   szSQL = "SELECT LeaseDetails.*, Property.PropertyName, Property.propertyID, Client.ClientName,Client.clientID,Units.UnitName,Units.UnitNumber " & _
             "FROM LeaseDetails, Property, Client, Units, Tenants " & _
             "WHERE LeaseDetails.LeaseID = '" & strLeaseId & "' AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "Client.ClientId = Property.ClientId AND " & _
                  "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber;"
'Debug.Print szSQL
 
               
               
        adoRst1.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
    'Loading Header section
        If adoRst1.EOF Then
             Label16(12).Caption = ""
             adoRst1.Close
             adoConn1.Close
             MsgBox "Could not load the lease information, lease Detail ID incorrect!!", vbInformation, "Warning!"
             Exit Sub
        End If
        strLeaseId = adoRst1!LeaseID 'flxLeaseList.TextMatrix(flxLeaseList.row, 1)
        txtTenant.Tag = adoRst1!SageAccountNumber 'flxLeaseList.TextMatrix(flxLeaseList.row, 2) 'sage account ID
        txtSageAccountNumber.text = adoRst1!SageAccountNumber
        txtTenant.text = adoRst1!CompanyName 'flxLeaseList.TextMatrix(flxLeaseList.row, 3) 'CompanyName
        Me.Caption = "Lease Information : " & adoRst1!SageAccountNumber & " / " & adoRst1!CompanyName
        txtUnitName(0).text = adoRst1!UnitName
        txtUnitNumber.text = adoRst1!UnitNumber 'flxLeaseList.TextMatrix(flxLeaseList.row, 5)
        txtClient.text = adoRst1!ClientName ' flxLeaseList.TextMatrix(flxLeaseList.row, 6)
        txtClient.Tag = adoRst1!ClientID
        szClientID = adoRst1!ClientID
        txtProperty.text = adoRst1!PropertyName ' flxLeaseList.TextMatrix(flxLeaseList.row, 7)
        PROPERTY_ID = adoRst1!propertyID 'flxLeaseList.TextMatrix(flxLeaseList.row, 8)
        cboUsage.text = IIf(IsNull(adoRst1!Usage), "", adoRst1!Usage) 'flxLeaseList.TextMatrix(flxLeaseList.row, 9)
        'issue 591 by anol 20180516
        Label16(12).Tag = ""
        If CBool(adoRst1!Status) = True Then
            szLeaseStatus = True
        Else
            szLeaseStatus = False
            cmdTerminate.Visible = False
        End If
        If CBool(adoRst1!Status) = True Then
            szMsg = " Current"
        ElseIf CBool(adoRst1!Status) = False And IsNull(adoRst1!TerminateDate) = False Then
            szMsg = " Terminated (" & adoRst1!TerminateDate & ")"
            Label16(12).Tag = adoRst1!TerminateDate
        ElseIf CBool(adoRst1!Status) = False And IsNull(adoRst1!TerminateDate) = True Then
            szMsg = " Expired (" & adoRst1!EndDate & ")"
        Else
            szMsg = ""
        End If
        Label16(12).Caption = szMsg
        Label16(12).FontUnderline = False
    'End of loading header section
    
'   Check for sub lease.
   Dim HeadLease As String
   If IsNull(adoRst1!HeadLease) = False Then
      HeadLease = adoRst1!HeadLease
      'issue 537
      'Modified by anol 13 Feb 2015
      If adoRst1!HeadLease <> "" Then
         txtUnitName(1).Tag = adoRst1!HeadLease
         txtUnitName(1).text = adoRst1!UnitNumber
         chkSubLease.Value = 1
      Else
         chkSubLease.Value = 0
      End If
      
   Else
      HeadLease = ""
      chkSubLease.Value = 0
   End If
   'Anol Set Cap Amount here
   '04 Feb 2015
   txtCapAmount.text = IIf(IsNull(adoRst1!CapAmount), "", adoRst1!CapAmount)
'   Fill text boxes with lease details.
   cboType.text = IIf(IsNull(adoRst1!TYPEOFSTORE), "", adoRst1!TYPEOFSTORE)
   txtLeaseStDt.text = IIf(IsNull(adoRst1!StartDate), "", adoRst1!StartDate)
   'Below line comment out by anol 11  Nov 2015
   'txtYearEnd.text = IIf(IsNull(adoRst1!YearEnd), "", adoRst1!YearEnd)
   txtLeaseEndDate.text = IIf(IsNull(adoRst1!EndDate), "", adoRst1!EndDate)
   szUndoLeaseEndDate = txtLeaseEndDate.text
   chkOLED.Value = adoRst1!OLED
   chkGPrataDmd.Value = adoRst1!GPrataDmd
   chkHoldingOver.Value = adoRst1!HoldingOver

'   ****************************************
'     Rent Charge tab
'   ****************************************
   If adoRst1!BRPayable = "Y" Then LoadFlxRentCharges adoConn1
'    LoadFlxRentCharges adoConn1
'   ****************************************
'     Service Charge tab
'   ****************************************
   If adoRst1!SCPayable = "Y" Then LoadFlxSC adoConn1
   ' LoadFlxSC adoConn1
'   ****************************************
'     Interest Charge tab
'   ****************************************
   If adoRst1!InterestChargeable = "Y" Then
      cboIntCrgable.text = "Yes"
      cboIntChargeDept.text = IIf(IsNull(adoRst1!IntChargeDept), "", DeptName(IIf(IsNull(adoRst1!IntChargeDept) Or adoRst1!IntChargeDept = "", 1, adoRst1!IntChargeDept)))
      If Not IsNull(adoRst1!IntDemandType) Then cboIntDemandType.ListIndex = CInt(adoRst1!IntDemandType) - 1

      If Not IsNull(adoRst1!AdditionalInterest) Or adoRst1!AdditionalInterest <> "" Then
         If Val(adoRst1!AdditionalInterest) = 0 Then
            optGIR.Value = True
            txtAdditionalIntRate.text = ""
         Else
            optLSR.Value = True
            txtAdditionalIntRate.text = Format(adoRst1!AdditionalInterest, "0.0000")
         End If
      Else
         optGIR.Value = True
         txtAdditionalIntRate.text = ""
      End If
'                        this ServiceChargeDept field is using by interest not by service charge any more
      If Not IsNull(adoRst1!ServiceChargeDept) Or adoRst1!ServiceChargeDept <> "" Then
         If UCase(adoRst1!ServiceChargeDept) = "AUTO" Then
            optAutoIntCal.Value = True
            txtIntPayableAfterDays.text = IIf(IsNull(adoRst1!DaysAfterInterestPayable), "", adoRst1!DaysAfterInterestPayable)
         End If
         If UCase(adoRst1!ServiceChargeDept) = "MANU" Then
            optManIntCal.Value = True
            txtAmtCrgIntOn.text = Format(IIf(IsNull(adoRst1!InterestChargedOn) Or adoRst1!InterestChargedOn = "", 0, adoRst1!InterestChargedOn), "0.00")
            txtNoIntDays.text = CInt(IIf(IsNull(adoRst1!SCTOLimit) Or adoRst1!SCTOLimit = "", 0, adoRst1!SCTOLimit))
            txtInt2bChrg.text = Format(IIf(IsNull(adoRst1!InterestAmount) Or adoRst1!InterestAmount = "", 0, adoRst1!InterestAmount), "0.00")
         End If
      End If

      txtInterestDescription.text = IIf(IsNull(adoRst1!Text1) Or adoRst1!Text1 = "", "", adoRst1!Text1)
   Else
      cboIntCrgable.text = "No"
   End If

'   ****************************************
'     Break Clause tab
'   ****************************************
   If adoRst1!BreakClause = "Y" Then
      cboBreakClause.text = "Yes"
      cboBreak.text = IIf(IsNull(adoRst1!BreakType) = True, "", adoRst1!BreakType)
      txtBreakDate.text = IIf(IsNull(adoRst1!BreakDate) = True, "", adoRst1!BreakDate) 'adoRst1!BreakDate
   Else
      cboBreakClause.text = "No"
   End If

'   ****************************************
'     Rent Review tab
'   ****************************************
   LoadFlxRentAnalysis

'   ****************************************
'     Supplementary tab
'   ****************************************
'   txtSupp1.text = IIf(IsNull(adoRst1!SuppText1) Or adoRst1!SuppText1 = "", "", adoRst1!SuppText1) 'adoRst1!SuppText1
'   txtSupp2.text = IIf(IsNull(adoRst1!SuppText2) Or adoRst1!SuppText2 = "", "", adoRst1!SuppText2) 'adoRst1!SuppText2
'   txtSupp3.text = IIf(IsNull(adoRst1!SuppText3) Or adoRst1!SuppText3 = "", "", adoRst1!SuppText3)  'adoRst1!SuppText3

'   lblSupplementary1.Caption = IIf(IsNull(adoRst1!SuppCaption1) Or adoRst1!SuppCaption1 = "", "", adoRst1!SuppCaption1) 'adoRst1!SuppCaption1
'   lblSupplementary2.Caption = IIf(IsNull(adoRst1!SuppCaption2) Or adoRst1!SuppCaption2 = "", "", adoRst1!SuppCaption2) ''adoRst1!SuppCaption2
'   lblSupplementary3.Caption = IIf(IsNull(adoRst1!SuppCaption3) Or adoRst1!SuppCaption3 = "", "", adoRst1!SuppCaption3) 'adoRst1!SuppCaption3

'   txtDtFlgDate.text = IIf(IsNull(adoRst1!DateFlagDate) = True, "", Format(adoRst1!DateFlagDate, "dd/mm/yyyy"))
'   txtDtFlgDesc.text = IIf(IsNull(adoRst1!DateFlagDescription) Or adoRst1!DateFlagDescription = "", "", adoRst1!DateFlagDescription) 'adoRst1!DateFlagDescription
'   txtDtFlgDt2.text = IIf(IsNull(adoRst1!DateFlagDt2) = True, "", Format(adoRst1!DateFlagDt2, "dd/mm/yyyy")) ' Format(adoRst1!DateFlagDt2, "dd/mm/yyyy")
'   txtDtFlgDesc2.text = IIf(IsNull(adoRst1!DateFlagDescription2) Or adoRst1!DateFlagDescription2 = "", "", adoRst1!DateFlagDescription2) 'adoRst1!DateFlagDescription2
'
'   txtDtFlgDt3.text = IIf(IsNull(adoRst1!DateFlagDt3) = True, "", Format(adoRst1!DateFlagDt3, "dd/mm/yyyy")) 'Format(adoRst1!DateFlagDt3, "dd/mm/yyyy")
'   txtDtFlgDesc3.text = IIf(IsNull(adoRst1!DateFlagDescription3) Or adoRst1!DateFlagDescription3 = "", "", adoRst1!DateFlagDescription3)
   'If Not IsNull(adoRst1!DateFlagDescription3) Then txtDtFlgDesc3.text = adoRst1!DateFlagDescription3

'   ****************************************
'     Memo & Attachment tab
'   ****************************************
   If Not IsNull(adoRst1!Notes) Then
        txtMemo.text = adoRst1!Notes
   Else
        txtMemo.text = ""
   End If
   Call LoadAttachmentFiles(cmbFiles, strLeaseId, "Lease")

'   ****************************************
'     Insurance tab
'   ****************************************
   If adoRst1!InsurancePayable = "Y" Then LoadFlxIns adoConn1

'   ****************************************
'     Hanlding the delete command button
   cmdDelete.Enabled = CheckProducedDemands(adoConn1, strLeaseId)

   adoRst1.Close
   
'   ****************************************
'     Breaches tab
'   ****************************************
        
   loadFlxBreach
   
'   ****************************************
'     Assignment tab
'   ****************************************
   PopulateAssignments
   
   If Not chkOLED.Value And txtLeaseEndDate.text <> "" Then
'      If Len(Label16(12).Tag) Then 'Label16(12).Tag Holds the termination date
'            If DateDiff("d", Format(Label16(12).Tag, "DD/MM/YYYY"), txtLeaseEndDate.text) >= 0 Then
'                 szMsg = "This lease was terminated on '" & Label16(12).Tag & "'."
'                 MsgBox szMsg, vbOKOnly + vbInformation, "Lease Terminated"
'                 Exit Sub
'            End If
'      End If
      If DateDiff("d", txtLeaseEndDate.text, Format(Date, "DD/MM/YYYY")) >= 0 Then
         szMsg = "This Lease expired on " & txtLeaseEndDate.text & "." + Chr$(10)
         szMsg = szMsg + "If you wish to extend this lease, please edit the lease and change the lease end date."
         MsgBox szMsg, vbOKOnly + vbInformation, "Lease Expired"
         txtLeaseEndDate.ForeColor = vbRed
         cmdCancelEdit.Visible = True
         cmdCancelEdit.Enabled = True
         FocusControl cmdEdit
      End If
   End If
End Sub

Private Function CheckProducedDemands(ByVal adoConn As ADODB.Connection, ByVal szLeaseID As String) As Boolean
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT DemandID " & _
           "FROM DemandRecords " & _
           "WHERE LeaseRef = '" & szLeaseID & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      CheckProducedDemands = True
   Else
      CheckProducedDemands = False
   End If

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Sub LoadFlxIns(adoConn As ADODB.Connection)
   Dim iRow As Integer
   Dim adoRst1 As ADODB.Recordset

   Set adoRst1 = New ADODB.Recordset

'   get all sage account numbers and company names from tenants.
   szSQL = "SELECT InsCharges, InsuranceFrequency, InsuranceStartDate, InsuranceDemandType, " & _
               "InsuranceEachPeriod, InsuranceNextDueDate, ChargingType, ChargingFigure, " & _
               "TotalYearlyInsurance, InsuranceDept, InsuranceDept, InsDesc, StopIC,FundCode,FundName,fundID,D.Type,FC.Frequency " & _
             "FROM LInsuranceCharges AS L, DemandTypes AS D , Fund AS F,Frequencies as FC " & _
             "WHERE LeaseID = '" & strLeaseId & "' And (ISNULL(L.spare3) OR L.spare3 = '') AND FC.ID = L.InsuranceFrequency AND " & _
               "L.InsuranceDemandType = D.ID AND f.fundid=L.InsuranceDept AND " & _
               "(ISNULL(L.spare3) OR L.spare3 = '') " & _
             "ORDER BY InsCharges ASC"
'Debug.Print szSQL
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Call ConfigFlxInsurance
   If Not adoRst1.EOF Then
      iRow = 1
      While Not adoRst1.EOF
         flxIns.RowHeight(iRow) = 240
         flxIns.TextMatrix(iRow, 0) = IIf(IsNull(adoRst1!InsCharges), "", adoRst1!InsCharges)
         flxIns.TextMatrix(iRow, 1) = IIf(IsNull(adoRst1!InsuranceDemandType), "", DemandType(adoConn, adoRst1!InsuranceDemandType))
         flxIns.TextMatrix(iRow, 2) = IIf(IsNull(adoRst1!InsuranceStartDate), "", Format(adoRst1!InsuranceStartDate, "dd/mm/yyyy"))
         flxIns.TextMatrix(iRow, 11) = IIf(IsNull(adoRst1!InsuranceFrequency), "", adoRst1!InsuranceFrequency)       'Frequency
         flxIns.TextMatrix(iRow, 3) = IIf(IsNull(adoRst1!Frequency), "", adoRst1!Frequency)
         flxIns.TextMatrix(iRow, 4) = IIf(IsNull(adoRst1!InsuranceNextDueDate), "", Format(adoRst1!InsuranceNextDueDate, "dd/mm/yyyy"))
         flxIns.TextMatrix(iRow, 5) = IIf(IsNull(adoRst1!FundCode), "", adoRst1!FundCode)
         flxIns.TextMatrix(iRow, 6) = IIf(IsNull(adoRst1!ChargingType), "", cboIncCharMth.Column(1, adoRst1!ChargingType - 1))
         flxIns.TextMatrix(iRow, 7) = IIf(IsNull(adoRst1!ChargingFigure), "", Format(adoRst1!ChargingFigure, IIf(adoRst1!ChargingType = 2, "0.0000", "0.00")))
         flxIns.TextMatrix(iRow, 8) = IIf(IsNull(adoRst1!TotalYearlyInsurance), "", Format(adoRst1!TotalYearlyInsurance, "0.00"))
         flxIns.TextMatrix(iRow, 9) = IIf(IsNull(adoRst1!InsuranceEachPeriod), "", Format(adoRst1!InsuranceEachPeriod, "0.00"))
         flxIns.TextMatrix(iRow, 10) = IIf(IsNull(adoRst1!InsuranceDept), "", adoRst1!InsuranceDept)                 'Department
         flxIns.TextMatrix(iRow, 12) = IIf(IsNull(adoRst1!InsuranceDemandType), "", adoRst1!InsuranceDemandType)     'Demand type
         flxIns.TextMatrix(iRow, 13) = IIf(IsNull(adoRst1!ChargingType), "", adoRst1!ChargingType)                   'Charging Method
         flxIns.TextMatrix(iRow, 14) = IIf(IsNull(adoRst1!InsDesc), "", adoRst1!InsDesc)                             'Description
         If Not IsNull(adoRst1!StopIC) Or adoRst1!StopIC <> "" Then
            flxIns.TextMatrix(iRow, 16) = Format(adoRst1!StopIC, "dd/mm/yyyy")
         Else
            flxIns.TextMatrix(iRow, 16) = ""
         End If
         flxIns.TextMatrix(iRow, 17) = IIf(IsNull(adoRst1!fundID), "", adoRst1!fundID)
         flxIns.TextMatrix(iRow, 18) = IIf(IsNull(adoRst1!FundName), "", adoRst1!FundName)
         flxIns.TextMatrix(iRow, 19) = IIf(IsNull(adoRst1!Type), "", adoRst1!Type) 'Demand Type Description
         adoRst1.MoveNext
         If Not adoRst1.EOF Then flxIns.AddItem ""
         iRow = iRow + 1
      Wend
   End If

   flxIns.row = 0
   flxIns.col = 0

   adoRst1.Close
   Set adoRst1 = Nothing
End Sub

Private Sub LoadFlxSC(adoConn As ADODB.Connection)
   Dim iRow As Integer
   Dim adoRst1 As New ADODB.Recordset

'   get all sage account numbers and company names from tenants.
   szSQL = "SELECT type,ServiceCharge, SCFrequency, SCPayableFrom, SCNextDueDate, " & _
               "ChargingMethod, CMFigure, SCTotal, SCAmount, " & _
               "SCTOLimit, SCDemandType, ServiceChargeDept, SCDesc, " & _
               "S.ScheduleID, S.StopSC, S.spare3,F.FundCode,F.FundID,F.FundName,FC.Frequency " & _
             "FROM LServiceCharges AS S, DemandTypes as D, Fund as F,Frequencies as FC " & _
             "WHERE LeaseID = '" & strLeaseId & "' And F.FundID=cint(S.ServiceChargeDept) AND FC.ID=Cint(S.SCFrequency) AND " & _
               "S.SCDemandType = D.ID AND " & _
               "(ISNULL(S.spare3) OR S.spare3 <> 'DELETED') " & _
             "ORDER BY ServiceCharge ASC"
'Debug.Print szSQL
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Call ConfigFlxSC
   If Not adoRst1.EOF Then
      iRow = 1
      While Not adoRst1.EOF
         flxSC.RowHeight(iRow) = 240 ' we are hiding rows when we r deleteing a row before we actually deleting.
         flxSC.TextMatrix(iRow, 0) = IIf(IsNull(adoRst1!SCDemandType), "", adoRst1!SCDemandType)
         flxSC.TextMatrix(iRow, 1) = IIf(IsNull(adoRst1!SCDemandType), "", DemandType(adoConn, adoRst1!SCDemandType))
         flxSC.TextMatrix(iRow, 2) = IIf(IsNull(adoRst1!SCPayableFrom), "", Format(adoRst1!SCPayableFrom, "dd/mm/yyyy"))
         flxSC.TextMatrix(iRow, 17) = IIf(IsNull(adoRst1!SCFrequency), "", adoRst1!SCFrequency) 'this column is prior beacuse it is using in next statement
         flxSC.TextMatrix(iRow, 3) = IIf(IsNull(adoRst1!Frequency), "", adoRst1!Frequency)
         flxSC.TextMatrix(iRow, 4) = IIf(IsNull(adoRst1!SCNextDueDate), "", Format(adoRst1!SCNextDueDate, "dd/mm/yyyy"))
         flxSC.TextMatrix(iRow, 5) = IIf(IsNull(adoRst1!ServiceChargeDept), "", adoRst1!ServiceChargeDept)
'         flxSC.TextMatrix(iRow, 6) = IIf(IsNull(adoRst1!ServiceChargeDept), "", DeptName(IIf(IsNull(adoRst1!ServiceChargeDept) Or adoRst1!ServiceChargeDept = "", 1, adoRst1!ServiceChargeDept)))
         flxSC.TextMatrix(iRow, 6) = IIf(IsNull(adoRst1!FundCode), "", adoRst1!FundCode) 'IIf(IsNull(adoRst1!ServiceChargeDept), "", DeptCode(IIf(IsNull(adoRst1!ServiceChargeDept) Or adoRst1!ServiceChargeDept = "", 1, adoRst1!ServiceChargeDept), adoConn))
         If Not IsNull(adoRst1!ScheduleID) Then
            flxSC.TextMatrix(iRow, 7) = adoRst1!ScheduleID
            szSQL = ScheduleName(CLng(adoRst1!ScheduleID), adoConn)
            flxSC.TextMatrix(iRow, 8) = IIf(IsNull(szSQL), "", szSQL)
         End If
         flxSC.TextMatrix(iRow, 9) = IIf(IsNull(adoRst1!ChargingMethod), "", adoRst1!ChargingMethod)
         flxSC.TextMatrix(iRow, 10) = IIf(IsNull(adoRst1!ChargingMethod), "", ChargingMethod(adoConn, adoRst1!ChargingMethod))
         'Below line is modified by anol 25 Feb 2015
         flxSC.TextMatrix(iRow, 11) = Format(IIf(IsNull(adoRst1!CMFigure), "", adoRst1!CMFigure), IIf(flxSC.TextMatrix(iRow, 9) = 2, "0.00000000", "0.00"))
         flxSC.TextMatrix(iRow, 12) = Format(IIf(IsNull(adoRst1!SCTotal), 0, adoRst1!SCTotal), "0.00")
         flxSC.TextMatrix(iRow, 13) = Format(IIf(IsNull(adoRst1!SCAmount), 0, adoRst1!SCAmount), "0.00")
         flxSC.TextMatrix(iRow, 14) = adoRst1!ServiceCharge
         flxSC.TextMatrix(iRow, 15) = IIf(IsNull(adoRst1!SCDesc), "", adoRst1!SCDesc)
'         flxSC.TextMatrix(iRow, 17) = IIf(IsNull(adoRst1!SCFrequency), "", adoRst1!SCFrequency)
         If Not IsNull(adoRst1!StopSC) Or adoRst1!StopSC <> "" Then
            flxSC.TextMatrix(iRow, 18) = Format(adoRst1!StopSC, "dd/mm/yyyy")
         Else
            flxSC.TextMatrix(iRow, 18) = ""
         End If
         flxSC.TextMatrix(iRow, 19) = IIf(IsNull(adoRst1!FundName), "", adoRst1!FundName)
         flxSC.TextMatrix(iRow, 20) = IIf(IsNull(adoRst1!Type), "", adoRst1!Type)
         adoRst1.MoveNext
         If Not adoRst1.EOF Then flxSC.AddItem ""
         iRow = iRow + 1
      Wend
   End If

   flxSC.row = 0
   flxSC.col = 0

   adoRst1.Close
   Set adoRst1 = Nothing
End Sub

Private Function ScheduleName(iScheduleID As Long, adoConn As ADODB.Connection) As String
   Dim adoRst1 As New ADODB.Recordset

   szSQL = "SELECT * FROM Schedule WHERE ScheduleID = " & iScheduleID & ";"

   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst1.EOF Then
      ScheduleName = ""
      adoRst1.Close
      Set adoRst1 = Nothing
      Exit Function
   End If

   ScheduleName = adoRst1.Fields.Item("ScheduleName").Value
   adoRst1.Close
   Set adoRst1 = Nothing
End Function

Private Sub ConfigFlxSC()
   Dim szFlxHeader As String

   flxSC.RowHeight(0) = 0
   flxSC.Clear
   flxSC.Rows = 2
   flxSC.Cols = 21
'   szFlxHeader$ = "DemandTypeID|<DemandType|<StartDate|<Frequency|<NextDueDate" & _
'                  "||<fund||<Schedule||<ChargingMethod|>Amt|>TotalYear" & _
'                  "|>EachPeriod|||<DELETED|<FreqID|<StopDate"
'   flxSC.FormatString = szFlxHeader$

   flxSC.ColWidth(0) = 0                                                   'demand type id
   flxSC.ColWidth(1) = lblSC(1).Left - lblSC(0).Left                       'demand type
   flxSC.ColWidth(2) = lblSC(2).Left - lblSC(1).Left                       'start date
   flxSC.ColWidth(3) = lblSC(3).Left - lblSC(2).Left                       'frequency
   flxSC.ColWidth(4) = lblSC(4).Left - lblSC(3).Left                       'next due date
   flxSC.ColWidth(5) = 0                                                   'fund id /  SC DEPT
   flxSC.ColWidth(6) = lblSC(5).Left - lblSC(4).Left                       'fund    /  SC DEPT
   flxSC.ColWidth(7) = 0                                                   'Schedule ID
   flxSC.ColWidth(8) = lblSC(6).Left - lblSC(5).Left                       'Schedule Method
   flxSC.ColWidth(9) = 0                                                   'Charging Method ID
   flxSC.ColWidth(10) = lblSC(7).Left - lblSC(6).Left - 150                    'Charging Method
   flxSC.ColWidth(11) = lblSC(8).Left - lblSC(7).Left + 200                    'amt
   flxSC.ColWidth(12) = lblSC(9).Left - lblSC(8).Left                      'total/yr
   flxSC.ColWidth(13) = lblSC(10).Left - lblSC(9).Left - 60                'each period
   flxSC.ColWidth(14) = 0                                                  'ServiceCharge ->id of LServiceCharge
   flxSC.ColWidth(15) = 0                                                  'SCDesc ->description
   flxSC.ColWidth(16) = 0                                                  'Marking ->deleted
   flxSC.ColWidth(17) = 0                                                  'Frequency ID
   flxSC.ColWidth(18) = lblSC(11).Left - lblSC(10).Left 'flxSC.Width + flxSC.Left - lblSC(10).Left - 200     'stop date
   flxSC.ColWidth(19) = 0
   flxSC.ColWidth(20) = 0 'Deman Type
End Sub

Public Sub loadFlxBreach()
   'Set the RDO Connections to the dataset
   Dim sSQLQuery_ As String

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   sSQLQuery_ = "SELECT " & _
         "SecondaryCode.value as BreachType,SecondaryCode.Code as code, " & _
         "LeaseBreaches.CommenceDate, LeaseBreaches.InitiatedBy, " & _
         "LeaseBreaches.Resolved, LeaseBreaches.DateReceived, " & _
         "LeaseBreaches.ReceivedBy, LeaseBreaches.BreachID,LeaseBreaches.LeaseMemo " & _
         "FROM LeaseBreaches, SecondaryCode " & _
         "WHERE LeaseBreaches.LeaseID = '" & strLeaseId & "' " & _
         "AND SecondaryCode.Code = LeaseBreaches.BreachType " & _
         "AND SecondaryCode.PrimaryCode = 'BTYP' AND  (LeaseBreaches.DELETEFLAG <> 'DELETED'  or isnull( LeaseBreaches.DELETEFLAG)) ORDER BY LeaseID ASC"
         

   'populateGridDefinedHeader adoConn, sSQLQuery_, gridBreach
   Dim adoRST As New ADODB.Recordset
   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockOptimistic

   Dim i As Integer, j As Integer
   Dim iRows As Integer
   iRows = 1
   ConfigGridBreach
   If adoRST.RecordCount > 0 Then
        gridBreach.Rows = adoRST.RecordCount + 1
   End If
   While Not adoRST.EOF
        gridBreach.RowHeight(gridBreach.row) = 240
        gridBreach.TextMatrix(iRows, 0) = IIf(IsNull(adoRST.Fields.Item("BreachType").Value), "", adoRST.Fields.Item("BreachType").Value)
        gridBreach.TextMatrix(iRows, 1) = IIf(IsNull(adoRST.Fields.Item("CommenceDate").Value), "", adoRST.Fields.Item("CommenceDate").Value)
        gridBreach.TextMatrix(iRows, 2) = IIf(IsNull(adoRST.Fields.Item("InitiatedBy").Value), "", adoRST.Fields.Item("InitiatedBy").Value)
        gridBreach.TextMatrix(iRows, 3) = IIf(IsNull(adoRST.Fields.Item("DateReceived").Value), "", adoRST.Fields.Item("DateReceived").Value)
        gridBreach.TextMatrix(iRows, 4) = IIf(IsNull(adoRST.Fields.Item("ReceivedBy").Value), "", adoRST.Fields.Item("ReceivedBy").Value) 'IIf(chkResolved.Value = 0, "No", "Yes")
        gridBreach.TextMatrix(iRows, 5) = IIf(IIf(IsNull(adoRST.Fields.Item("Resolved").Value), "0", adoRST.Fields.Item("Resolved").Value) = 0, "No", "Yes")
        gridBreach.TextMatrix(iRows, 6) = IIf(IsNull(adoRST.Fields.Item("BreachID").Value), "", adoRST.Fields.Item("BreachID").Value)
        gridBreach.TextMatrix(iRows, 7) = "" 'Delete flag
        gridBreach.TextMatrix(iRows, 8) = IIf(IsNull(adoRST.Fields.Item("code").Value), "", adoRST.Fields.Item("code").Value)
        gridBreach.TextMatrix(iRows, 9) = IIf(IsNull(adoRST.Fields.Item("LeaseMemo").Value), "", adoRST.Fields.Item("LeaseMemo").Value)
        iRows = iRows + 1
        adoRST.MoveNext
   Wend
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub PopulateAssignments()
   'Set the RDO Connections to the dataset
   Dim sSQLQuery_ As String, szHeader As String

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   sSQLQuery_ = "SELECT LeaseAssignments.AssignmentID, " & _
                  "LeaseAssignments.AssignDate, " & _
                  "LeaseAssignments.Assignee, " & _
                  "LeaseAssignments.Decp " & _
                "FROM LeaseAssignments " & _
                "WHERE LeaseAssignments.LeaseID = '" & strLeaseId & "' "

   ConfigAssignmentGrid
   szHeader$ = "<AssignmentID|<Assignment_Date|<Assignee|<Description"
   populateGridSimply adoConn, sSQLQuery_, gridAssignment, szHeader

   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub ConfigGridBreach()
   Dim szHeader As String
   Dim iFlxCol As Integer, i As Integer

   szHeader$ = "<BreachType|<CommenceDate|<InitiatedBy|<Resolved|<DateReceived|<ReceivedBy|<BreachID"
   gridBreach.FormatString = szHeader$

   iFlxCol = 0
   gridBreach.Clear
   gridBreach.Cols = 10
   gridBreach.Rows = 2
   
   For i = 0 To Label44.UBound - 1
      gridBreach.ColWidth(iFlxCol) = Label44(i + 1).Left - Label44(i).Left - 5
      iFlxCol = iFlxCol + 1
   Next i
   gridBreach.ColWidth(iFlxCol) = gridBreach.Width + gridBreach.Left - Label44(i).Left - 50
   gridBreach.ColWidth(6) = 0
   gridBreach.ColWidth(7) = 0 'delete flag
   gridBreach.ColWidth(8) = 0 'value
   gridBreach.ColWidth(9) = 1550 'Memo

   gridBreach.RowHeight(0) = 0
End Sub

Private Sub ConfigflxLicences()
   Dim iRow As Integer
   iRow = 1

   flxLicences.Clear
   flxLicences.Rows = 2
   flxLicences.Cols = 6
   flxLicences.RowHeight(0) = 0

   flxLicences.ColWidth(0) = 0
   flxLicences.ColWidth(1) = Label17(1).Left - Label17(0).Left
   flxLicences.ColWidth(2) = Label17(2).Left - Label17(1).Left
   flxLicences.ColWidth(3) = Label17(3).Left - Label17(2).Left
   flxLicences.ColWidth(4) = Label17(4).Left - Label17(3).Left
   flxLicences.ColWidth(5) = flxLicences.Left + flxLicences.Width - Label17(4).Left - 300
End Sub

Private Sub ConfigAssignmentGrid()
   Dim iRow As Integer
   iRow = 1

   gridAssignment.Clear
   gridAssignment.Rows = 2
   gridAssignment.Cols = 4
   gridAssignment.RowHeight(0) = 0

   gridAssignment.ColWidth(0) = 0
   gridAssignment.ColWidth(1) = txtAssignee.Left - txtAssignment_Date.Left
   gridAssignment.ColWidth(2) = txtDescription.Left - txtAssignee.Left
   gridAssignment.ColWidth(3) = txtDescription.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim X As Integer

   If (cmdSaveNew.Visible Or cmdSaveEdit.Visible) Then
      X = MsgBox("Do you wish to save before closing?", vbQuestion + vbYesNoCancel, "Lease Module")

      If X = vbYes And cmdSaveNew.Visible Then
            
         If ValidationSaveLease = True Then
            If SaveUpdateLease(True) Then
                ShowMsgInTaskBar "The new lease record has been saved successfully."
    '            ConfigFlxRentCharges
    '            ConfigFlxSC
    '            ConfigFlxInsurance
                ControlsModeRentCharges DefaultMode
                ControlsModeServiceCharges DefaultMode
                ControlsModeInsuranceCharges DefaultMode
            End If
'            ConfigFlxRentAnalysis
'            SetAssignmentGrid
         Else
            Cancel = 1
         End If
      End If
      If X = vbYes And cmdSaveEdit.Visible Then
          If ValidationSaveLease = True Then
            If SaveUpdateLease(False) Then 'parameter is false becuase now it shall update the existing charges
                MsgBox "The lease record has been updated", vbOKOnly + vbInformation, "Updated"
    '            ConfigFlxRentCharges
    '            ConfigFlxSC
    '            ConfigFlxInsurance
                ControlsModeRentCharges DefaultMode
                ControlsModeServiceCharges DefaultMode
                ControlsModeInsuranceCharges DefaultMode
            End If
'            ConfigFlxRentAnalysis
'            SetAssignmentGrid
         Else
            Cancel = 1
         End If
      End If

      If X = vbNo And cmdSaveNew.Visible Then
         Call EmptyBoxes
         Call SetAddNewMode
         Call DisableBoxes
      End If
      If X = vbNo And cmdSaveEdit.Visible Then cmdCancelEdit_Click

      'If X <> vbCancel Then frmMMain.fraCmdButton.Enabled = True

      If X = vbCancel Then Cancel = 1
   Else
      'frmMMain.fraCmdButton.Enabled = True
   End If
   strSessionClientID = ""
   strSessionPropertyID = ""
'   Call WheelUnHook(Me.hWnd)
End Sub



Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Frame4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

'Private Sub mnuDemands_Click()
'
''Load frmDemands1
''Unload Me
''frmDemands1.Show
'
'End Sub
'
'Private Sub mnuExit_Click()
'   Unload frmMMain
'End Sub
'
'Private Sub mnuGlobal_Click()
'
'Call EmptyBoxes
'
'Load frmGlobal
'Unload Me
'frmGlobal.Show
'
'End Sub
'
'Private Sub mnuMain_Click()
'   Call EmptyBoxes
'
'   Unload Me
'End Sub
'
'Private Sub mnuShopCentre_Click()
'
'Call EmptyBoxes
'
'Load frmShoppingCentre
'Unload Me
'frmShoppingCentre.Show
'
'End Sub
'
'Private Sub mnuTenants_Click()
'   Call EmptyBoxes
'End Sub
'
'Private Sub mnuUnits_Click()
'   Call EmptyBoxes
'   Unload Me
'End Sub
'
Private Sub gridAssignment_RowColChange()
   populateControl Me, gridAssignment
   AssignmentButtonMode GridRowOnSelection
End Sub

Private Sub gridBreach_Click()
   'BreachButtonMode GridRowOnSelection
   If gridBreach.TextMatrix(1, 0) = "" Then Exit Sub
   cboBreachType.text = gridBreach.TextMatrix(gridBreach.row, 0)
   txtCommenceDate.text = gridBreach.TextMatrix(gridBreach.row, 1)
   txtInitiatedBy.text = gridBreach.TextMatrix(gridBreach.row, 2)
   txtDateReceived.text = gridBreach.TextMatrix(gridBreach.row, 3)
   txtReceivedBy.text = gridBreach.TextMatrix(gridBreach.row, 4)
   txtMemo2.text = gridBreach.TextMatrix(gridBreach.row, 9)
   chkResolved.Value = Val(gridBreach.TextMatrix(gridBreach.row, 5))
   strBreachID = gridBreach.TextMatrix(gridBreach.row, 6)
   gridBreach_EDIT = gridBreach.row
   'ControlsMode GridRowOnSelection
    cmdBreachNew.Enabled = False
    cmdBreachEdit.Enabled = True
    cmdBreachSave.Enabled = False
    cmdBreachCancel.Enabled = True
    cmdDeleteBreaches.Enabled = True
End Sub



'Private Sub lblSupplementary1_DblClick()
'   txtSuppCaption1.Visible = True
'   txtSuppCaption1.Left = lblSupplementary1.Left
'   txtSuppCaption1.Top = lblSupplementary1.Top
'   txtSuppCaption1.text = lblSupplementary1.Caption
'   txtSuppCaption1.SetFocus
'End Sub

'Private Sub lblSupplementary2_DblClick()
'   txtSuppCaption2.Visible = True
'   txtSuppCaption2.Left = lblSupplementary2.Left
'   txtSuppCaption2.Top = lblSupplementary2.Top
'   txtSuppCaption2.text = lblSupplementary2.Caption
'   txtSuppCaption2.SetFocus
'End Sub

Private Sub SCGlobal()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String ', szaUnit() As String
   Dim sPpSF As Single, lArea As Long

   adoConn.Open getConnectionString

'   szaUnit = Split(cboUnit.text, " - ")

   lArea = GetUnitTA(adoConn, txtUnitNumber.text)         'Total Area of the unit
   If lArea = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      sPpSF = GetSC_PpSF(adoConn, txtUnitNumber.text, txtSCFundCode.Tag)
      If sPpSF < 0 Then
         MsgBox "There is no service charge defined against the fund.", vbCritical + vbOKOnly, "Service Charge"
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If

      txtChargingFigure.text = Format(sPpSF * lArea, "0.00")
      txtSCTotalAmount.text = txtChargingFigure.text

'      If cboFreqSC.ListIndex > -1 Then
'         szSQL = "SELECT PARTOFYEAR " & _
'                   "FROM FREQUENCIES " & _
'                   "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
'      Else
         szSQL = "SELECT PARTOFYEAR " & _
                   "FROM FREQUENCIES " & _
                   "WHERE ID = " & txtFreqSC.Tag & ";"
'      End If
      rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      txtSCDueEachPeriod.text = Format((CDbl(txtSCTotalAmount.text) / CInt(rstRst!PartOfYear)), "0.00")

      rstRst.Close
      Set rstRst = Nothing
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub BRGlobal()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String ', szaUnit() As String
   Dim sPpSF As Single, lUnitArea As Long

   adoConn.Open getConnectionString

'   szaUnit = Split(cboUnit.text, " - ")

   lUnitArea = GetUnitTA(adoConn, txtUnitNumber.text)       'Total Area of the unit
   If lUnitArea = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      sPpSF = GetRC_PpSF(adoConn, txtUnitNumber.text, txtRCFundCode.Tag)
      If sPpSF < 0 Then
         MsgBox "There is no rent charge defined against the fund.", vbCritical + vbOKOnly, "Rent Charge"
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If

      txtBRChargingFigure.text = Format(sPpSF * lUnitArea, "0.00")
      txtTotalRentYear.text = txtBRChargingFigure.text

'      If cboFreqBR.ListIndex > -1 Then
'         szSQL = "SELECT PARTOFYEAR " & _
'                   "FROM FREQUENCIES " & _
'                   "WHERE ID = " & (cboFreqBR.ListIndex + 1) & ";"
'      Else
         szSQL = "SELECT PARTOFYEAR " & _
                   "FROM FREQUENCIES " & _
                   "WHERE ID = " & txtFreqBR.Tag & ";"
'      End If
      rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      txtRentDueEachPeriod.text = Format((CDbl(txtTotalRentYear.text) / CInt(rstRst!PartOfYear)), "0.00")

      rstRst.Close
      Set rstRst = Nothing
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function GetIC_PpSF(rdoConn As ADODB.Connection, szUnitNumber As String, szFund As String) As Double
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHanlder

'Samrat 18/09/2014    Financial year has been implemented
   szSQL = "SELECT GlobalInsurance.PPSF as RC " & _
           "FROM Units, GlobalInsurance, Property " & _
           "WHERE " & _
                 "Units.PropertyID = GlobalInsurance.PropertyID AND " & _
                 "Units.PropertyID = Property.PropertyID AND " & _
                 "Property.CBY = GlobalInsurance.FinancialYear AND " & _
                 "Units.UnitNumber = '" & szUnitNumber & "' AND " & _
                 "GlobalInsurance.FundType = " & szFund & ";"
'Debug.Print szSQL
   rstRst.Open szSQL, rdoConn, adOpenStatic, adLockReadOnly

   If Not rstRst.EOF Then
      GetIC_PpSF = CDbl(rstRst!rc)
   Else
      GetIC_PpSF = -1
   End If

   rstRst.Close
   Set rstRst = Nothing
   Exit Function

ErrorHanlder:
   GetIC_PpSF = -1
   rstRst.Close
   Set rstRst = Nothing
End Function

Private Sub IBGlobal()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String ', szaUnit() As String
   Dim sPpSF As Single, lUnitArea As Long

   adoConn.Open getConnectionString

'   szaUnit = Split(cboUnit.text, " - ")

   lUnitArea = GetUnitTA(adoConn, txtUnitNumber.text)       'Total Area of the unit
   If lUnitArea = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      sPpSF = GetIC_PpSF(adoConn, txtUnitNumber.text, txtICFundCode.Tag)
      If sPpSF < 0 Then
         MsgBox "There is no insurance charge defined against the fund.", vbCritical + vbOKOnly, "insurance Charge"
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If

      txtInsPercentage.text = Format(sPpSF * lUnitArea, "0.00")
      txtTotalYearlyIns.text = txtInsPercentage.text

'      If cboInsFreq.ListIndex > -1 Then
'         szSQL = "SELECT PARTOFYEAR " & _
'                   "FROM FREQUENCIES " & _
'                   "WHERE ID = " & (cboInsFreq.ListIndex + 1) & ";"
'      Else
         szSQL = "SELECT PARTOFYEAR " & _
                   "FROM FREQUENCIES " & _
                   "WHERE ID = " & txtFreqIC.Tag & ";"
'      End If
      rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      txtInsEachPeriod.text = Format((CDbl(txtTotalYearlyIns.text) / CInt(rstRst!PartOfYear)), "0.00")

      rstRst.Close
      Set rstRst = Nothing
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

'Private Sub lblSupplementary3_DblClick()
'   txtSuppCaption3.Visible = True
'   txtSuppCaption3.Left = lblSupplementary3.Left
'   txtSuppCaption3.Top = lblSupplementary3.Top
'   txtSuppCaption3.text = lblSupplementary3.Caption
'   txtSuppCaption3.SetFocus
'End Sub

Private Sub optManIntCal_Click()
   If optManIntCal.Value Then
      txtIntPayableAfterDays.Locked = True
      Frame1(4).Enabled = True
      If txtAmtCrgIntOn.Enabled Then txtAmtCrgIntOn.SetFocus
   End If
End Sub

Private Sub optAutoIntCal_Click()
   If optAutoIntCal.Value Then
      txtIntPayableAfterDays.Locked = False
      Frame1(4).Enabled = False
      If tabLease.Tab = 5 Then txtIntPayableAfterDays.SetFocus
   End If
End Sub

Private Sub optGIR_Click()
   If optGIR.Value Then
      txtAdditionalIntRate.text = ""
      txtAdditionalIntRate.Locked = True
   End If
End Sub

Private Sub optLSR_Click()
   If optLSR.Value Then
      txtAdditionalIntRate.Locked = False
      txtAdditionalIntRate.SetFocus
   End If
End Sub

Private Sub fraList_KeyPress(KeyAscii As Integer, Index As Integer)
   If KeyAscii = 27 Then fraList(1).Visible = False
End Sub

Private Sub tabLease_Click(PreviousTab As Integer)
   Select Case PreviousTab
      Case 1:           'Rent Charges
         If cmdSaveRentCrg.Enabled Then
            If MsgBox("Do you what to save changes to Rent Charges?", vbQuestion + vbYesNo, "Rent Charges") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdSaveRentCrg.SetFocus
            End If
         End If

      Case 2:           'Rent Review
         If cmdSaveRentAnalysis.Enabled Then
            If MsgBox("Do you what to save changes to Rent Reviews?", vbQuestion + vbYesNo, "Rent Reviews") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdSaveRentAnalysis.SetFocus
            End If
         End If

      Case 4:           'Service Charges
         If cmdSCSave.Enabled Then
            If MsgBox("Do you what to save changes to Service Charges?", vbQuestion + vbYesNo, "Service Charges") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdSCSave.SetFocus
            End If
         End If

      Case 6:           'Breaches
         If cmdBreachSave.Enabled Then
            If MsgBox("Do you what to save changes to Breaches?", vbQuestion + vbYesNo, "Breaches") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdBreachSave.SetFocus
            End If
         End If

      Case 7:           'Assignemnt
         If cmdAssignmentSave.Enabled Then
            If MsgBox("Do you what to save changes to Assignments?", vbQuestion + vbYesNo, "Assignments") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdAssignmentSave.SetFocus
            End If
         End If

      Case 8:           'Insurance
         If cmdIncSave.Enabled Then
            If MsgBox("Do you what to save changes to Insurances?", vbQuestion + vbYesNo, "Insurances") = vbYes Then
               tabLease.Tab = PreviousTab
               cmdIncSave.SetFocus
            End If
         End If
   End Select
End Sub

Private Sub tabLease_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MousePointer = vbArrow
End Sub

Private Sub txtAdditionalIntRate_Change()
   If optManIntCal.Value And txtAmtCrgIntOn.text <> "" And txtNoIntDays.text <> "" Then
      txtNoIntDays_LostFocus
   End If
End Sub

Private Sub txtAdditionalIntRate_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtAdditionalIntRate, KeyAscii, 4
End Sub

Private Sub txtAmtCrgIntOn_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtAmtCrgIntOn, KeyAscii
End Sub

Private Sub txtAmtCrgIntOn_LostFocus()
   If txtAmtCrgIntOn.text <> "" Then
      txtAmtCrgIntOn.text = Format(txtAmtCrgIntOn.text, "0.00")
      If txtInt2bChrg.text <> "" Then txtNoIntDays_LostFocus
   End If
End Sub

Private Sub txtAssignment_Date_Change()
   TextBoxChangeDate txtAssignment_Date
End Sub

Private Sub txtAssignment_Date_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtAssignment_Date, KeyAscii
End Sub

Private Sub txtBRChargingFigure_GotFocus()
   If txtRentStartDate.Enabled Then Exit Sub

   If cboBRChargingMth.text = "" Then
      MsgBox "Please choose the charging method first.", vbInformation + vbOKOnly, "Charging Method"
      txtBRChargingFigure.text = ""
      cboBRChargingMth.SetFocus
   End If
End Sub

Private Sub txtBRChargingFigure_KeyPress(KeyAscii As Integer)
On Error GoTo Err
    If KeyAscii = 13 Then
        txtTotalRentYear.SetFocus
    End If
   If cboBRChargingMth.Column(0) <> 2 Then
      DigitTextKeyPress txtBRChargingFigure, KeyAscii
   Else
      DigitTextKeyPress txtBRChargingFigure, KeyAscii, 4
   End If
    Exit Sub
Err:
   If Err.Number = 381 Then
        MsgBox "Please select charging method correctly", vbInformation, "Warning"
        FocusControl cboBRChargingMth
   End If
End Sub

Private Sub txtCapAmount_KeyPress(KeyAscii As Integer)
   'added by anol 04 Feb 2015
   If KeyAscii = 13 Then
        FocusControl cmdSCSave
   End If
    DigitTextKeyPress txtCapAmount, KeyAscii
End Sub

Private Sub txtChargingFigure_GotFocus()
   If cmdSCNew.Enabled Then Exit Sub

   If cboSCChargingMth.text = "" Then
      MsgBox "Please choose the charging method first.", vbInformation + vbOKOnly, "Charging Method"
      txtChargingFigure.text = ""
      cboSCChargingMth.SetFocus
   End If
End Sub

Private Sub txtChargingFigure_KeyPress(KeyAscii As Integer)
    On Error GoTo Err
    If KeyAscii = 13 Then
        txtSCTotalAmount.SetFocus
    End If
   If cboSCChargingMth.Column(0) <> 2 Then
      DigitTextKeyPress txtChargingFigure, KeyAscii
   Else
      DigitTextKeyPress txtChargingFigure, KeyAscii, 6
   End If
    Exit Sub
Err:
   If Err.Number = 381 Then
        MsgBox "Please select charging method correctly", vbInformation, "Warning"
        FocusControl cboSCChargingMth
   End If
End Sub

Private Sub txtChargingFigure_LostFocus()
   If Trim(txtChargingFigure.text) = "" Then
        txtChargingFigure.text = "0.00"
        Exit Sub
   End If

   If cboSCChargingMth.text = "Price Per Sq Foot" Then SCPricePerSqFoot
   If cboSCChargingMth.text = "Percentage" Then SCPercentage
   If cboSCChargingMth.text = "Annual" Then SCAnnual
End Sub

Private Sub txtDateReceived_Change()
   'Added by Samrat. 16.01.2006
   TextBoxChangeDate txtDateReceived
End Sub

Private Sub txtDateReceived_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   If KeyAscii = 13 Then
        txtReceivedBy.SetFocus
   End If
   TextBoxKeyPrsDate txtDateReceived, KeyAscii
End Sub

Private Sub cboIncCharMth_Click()
   If cboIncCharMth.text = "Global" Then
      If txtICFundName.text = "" Then
         MsgBox "Please select the Fund before choosing the Global charging method.", vbCritical + vbOKOnly, "Fund"
         Exit Sub
      End If
      IBGlobal
      txtInsPercentage.text = txtTotalYearlyIns.text
      txtInsPercentage.Locked = True
   Else
      txtInsPercentage.text = ""
      txtTotalYearlyIns.text = "0.00"
      txtInsEachPeriod.text = "0.00"
      'Rem by anol 21 Jun 2016
      'txtInsPercentage.Locked = False
   End If
End Sub

'Private Sub txtDtFlgDate_Change()
'   TextBoxChangeDate txtDtFlgDate
'End Sub

'Private Sub txtDtFlgDate_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDate, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDt2_Change()
'   TextBoxChangeDate txtDtFlgDt2
'End Sub
'
'Private Sub txtDtFlgDt2_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDt2, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDt2_LostFocus()
'   TextBoxFormatDate txtDtFlgDt2
'End Sub
'
'Private Sub txtDtFlgDt3_Change()
'   TextBoxChangeDate txtDtFlgDt3
'End Sub
'
'Private Sub txtDtFlgDt3_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDt3, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDt3_LostFocus()
'   TextBoxFormatDate txtDtFlgDt3
'End Sub

Private Sub txtInsNextDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 8 Or KeyCode = 46 Then
'      txtSCNextDueDt.text = ""
'   Else
'      KeyCode = 0
'   End If
End Sub

Private Sub txtInsNextDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdICFundCode
    End If
'   If KeyAscii = 8 Or KeyAscii = 46 Then
'      txtSCNextDueDt.text = ""
'   Else
'      KeyAscii = 0
'   End If
End Sub

Private Sub txtInsPercentage_GotFocus()
   If cmdIncNew.Enabled Then Exit Sub

   If txtFreqIC.text = "" Then
      MsgBox "Please select the Insurance Frequency.", vbCritical + vbInformation, "Insurance Frequency"
      FocusControl cmdFreqIC
      Exit Sub
   End If
End Sub

Private Sub txtIntPayableAfterDays_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtIntPayableAfterDays, KeyAscii, 0
End Sub

Private Sub txtLeaseEndDate_GotFocus()
   If txtLeaseStDt.text = "" Then
      MsgBox "Please input the Lease Start date first.", vbCritical + vbOKOnly, "Lease Dates"
      txtLeaseStDt.SetFocus
   End If
End Sub

Private Sub txtNextDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 8 Or KeyCode = 46 Then
'      txtNextDueDate.text = ""
'   Else
'      KeyCode = 0
'   End If
End Sub

Private Sub txtNextDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRCFund.SetFocus
    End If
'   If KeyAscii = 8 Or KeyAscii = 46 Then
'      txtNextDueDate.text = ""
'   Else
'      KeyAscii = 0
'   End If
End Sub

Private Sub txtNoIntDays_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtNoIntDays, KeyAscii, 0
End Sub

Private Sub txtNoIntDays_LostFocus()
   If Val(txtNoIntDays.text) < 1 Then Exit Sub

   If txtAmtCrgIntOn.text = "" Or Val(txtAmtCrgIntOn.text) <= 0 Then
      MsgBox "Please input specified amount to charge interest on.", vbCritical + vbOKOnly, "Manual Interest Calculation"
      txtAmtCrgIntOn.SetFocus
      Exit Sub
   End If

   Conn2.Open getConnectionString

   If optGIR.Value Then          'GIR->Global Interest Rate
      CalIntGlbRate
   Else
      If Val(txtAdditionalIntRate.text) < 0 Then
         MsgBox "You must input lease specific additional interest rate", vbOKOnly + vbCritical, "Interest Charge - Interest Rate"
         txtAdditionalIntRate.SetFocus
         Exit Sub
      End If
      CalculateInterest
   End If

   Conn2.Close
'   txtInterestDescription.SetFocus
End Sub

Public Sub CalIntGlbRate()
   Dim sTotalIntRate As Single

   szSQL = "SELECT (csng(BaseRate) + csng(AdditionalRate)) as TotalIntRate " & _
             "FROM InterestRates " & _
             "WHERE PropertyID = '" & PROPERTY_ID & "' " & _
             "ORDER BY DateFrom ASC;"

   Rst1.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
   Rst1.MoveLast

   sTotalIntRate = CSng(Rst1!TotalIntRate)
   Rst1.Close
   Set Rst1 = Nothing

   txtInt2bChrg.text = Format(CDbl(txtAmtCrgIntOn.text) * (CDbl(sTotalIntRate) / 100) * CInt(txtNoIntDays.text) / 365, "0.00")
End Sub

Public Sub CalculateInterest()
   Dim sTotalIntRate As Single

   szSQL = "SELECT BaseRate FROM InterestRates " & _
             "WHERE PropertyID = '" & PROPERTY_ID & "' " & _
             "ORDER BY DateFrom ASC;"

   Rst1.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
   Rst1.MoveNext

   sTotalIntRate = CSng(Rst1!BaseRate) + CSng(txtAdditionalIntRate.text)
   Rst1.Close
   Set Rst1 = Nothing

   txtInt2bChrg.text = Format(CDbl(txtAmtCrgIntOn.text) * (CDbl(sTotalIntRate) / 100) * CInt(txtNoIntDays.text) / 365, "0.00")
End Sub

Private Sub txtPayableFrom_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtPayableFrom
End Sub

Private Sub txtPayableFrom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        FocusControl cmdFreqSC
   End If
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtPayableFrom, KeyAscii
   
End Sub

Private Sub txtPayableFrom_LostFocus()
   If cmdSCNew.Enabled Or Frame1(2).Enabled = False Then Exit Sub

   'Added By Asif. 13/01/2006
   TextBoxFormatDate txtPayableFrom
   If txtSCDemandType.text = "" Then
      MsgBox "Please select a demand type.", vbCritical + vbOKOnly, "Service Charge"
      FocusControl cmdSCDemandType
      Exit Sub
   End If
   Call cboFreqSC_LostFocus
   'cboFreqSC.SetFocus
End Sub

Private Sub txtRentIncreaseAmount_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   If KeyAscii = 13 And cmdSaveRentAnalysis.Enabled Then
        FocusControl cmdSaveRentAnalysis
   End If
   If bReviewLocked = True And txtRentIncreaseAmount.Locked Then
        ShowMsgInTaskBar "Rent increase has been applied, You cannot edit the line"
    End If
   DigitTextKeyPress txtRentIncreaseAmount, KeyAscii
   
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

Private Sub txtSCNextDueDt_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 8 Or KeyCode = 46 Then
'      txtSCNextDueDt.text = ""
'   Else
'      KeyCode = 0
'   End If
End Sub

Private Sub txtSCNextDueDt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        FocusControl cmdSCFund
   End If
'   If KeyAscii = 8 Or KeyAscii = 46 Then
'      txtSCNextDueDt.text = ""
'   Else
'      KeyAscii = 0
'   End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl flxClientList
    End If
End Sub

Private Sub txtSearchCompany_Change()
    Dim Filter As String
    Dim tempstr As String
'   Dim i As Integer
'
''   If Len(txtSearchUnitName.text) > 0 Then
''      txtSearchName.text = ""
''      txtSearchTenant.text = ""
''   End If
'
'   For i = 1 To flxLeaseList.Rows - 1
'      flxLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxLeaseList.TextMatrix(i, 4), Len(txtSearchUnitName.text))) <> UCase(txtSearchUnitName.text) Then
'         flxLeaseList.RowHeight(i) = 0
'      End If
'   Next i

    If Len(Trim(txtSearchCompany.text)) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
      txtSearchUnitName = ""
      tempstr = Replace(UCase(txtSearchCompany.text), "'", "''")
        Filter = " UnitNumber LIKE '%" + tempstr + "*'"
    Else
          Filter = ""
   End If
   If sText = "UnitNo" Then
       Call FilterTenantsList(Filter)
   End If
End Sub

Private Sub txtSearchCompany_GotFocus()
    sText = "UnitNo"
End Sub

Private Sub txtSearchCompany_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtSearchUnitName.SetFocus
    End If
End Sub

Private Sub txtSearchName_Change()
     Dim Filter As String
     Dim tempstr  As String
     
'   Dim i As Integer
'
'   If Len(txtSearchName.text) > 0 Then
'      txtSearchTenant.text = ""
'      txtSearchUnitName.text = ""
'   End If
'
'   For i = 1 To flxLeaseList.Rows - 1
'      flxLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxLeaseList.TextMatrix(i, 3), Len(txtSearchName.text))) <> UCase(txtSearchName.text) Then
'         flxLeaseList.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 26 Jul 2014
'   If Len(txtSearchName.text) > 0 Then
'      txtSearchTenant.text = ""
'      txtSearchUnitName.text = ""
'   End If
   If Len(Trim(txtSearchName.text)) > 0 Then
      txtSearchTenant.text = ""
      txtSearchUnitName.text = ""
      txtSearchCompany.text = ""
      tempstr = Replace(UCase(txtSearchName.text), "'", "''")
      Filter = " CompanyName LIKE '%" + tempstr + "*'"
   Else
       Filter = ""
   End If
   If sText = "TenantName" Then
        Call FilterTenantsList(Filter)
   End If
End Sub

Private Sub txtSearchName_GotFocus()
    sText = "TenantName"
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtSearchUnitName.SetFocus
'    End If
    If KeyAscii = 13 Then
        If Len(Trim(txtSearchUnitName.text)) > 0 Then
            FocusControl flxLeaseList
             If flxLeaseList.Rows > 1 Then
                flxLeaseList.row = 1
            End If
        Else
            FocusControl txtSearchCompany
        End If
    End If
End Sub

Private Sub txtSearchName_KeyUp(KeyCode As Integer, Shift As Integer)
'     FilterTenantsList
End Sub

Private Sub txtSearchTenant_Change()
     Dim Filter As String
     Dim tempstr  As String
'   Dim i As Integer
'
''   If Len(txtSearchTenant.text) > 0 Then
''      txtSearchName.text = ""
''      txtSearchUnitName.text = ""
''   End If
'
'   For i = 1 To flxLeaseList.Rows - 1
'      flxLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxLeaseList.TextMatrix(i, 2), Len(txtSearchTenant.text))) <> UCase(txtSearchTenant.text) Then
'         flxLeaseList.RowHeight(i) = 0
'      End If
'   Next i

   If Len(Trim(txtSearchTenant.text)) > 0 Then
      txtSearchName.text = ""
      txtSearchUnitName.text = ""
       txtSearchCompany.text = ""
      tempstr = Replace(UCase(txtSearchTenant.text), "'", "''")
      Filter = " SageAccountNumber LIKE '%" + tempstr + "*'"
   Else
      Filter = ""
   End If

    If sText = "TenantID" Then
           Call FilterTenantsList(Filter)
    End If

   
End Sub

Private Sub txtSearchTenant_GotFocus()
    sText = "TenantID"
End Sub

Private Sub txtSearchTenant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtSearchTenant.text)) > 0 Then
            FocusControl flxLeaseList
             If flxLeaseList.Rows > 1 Then
                flxLeaseList.row = 1
            End If
        Else
            FocusControl txtSearchName
        End If
    End If
End Sub

Private Sub txtSearchTenant_KeyUp(KeyCode As Integer, Shift As Integer)
'    FilterTenantsList
End Sub

Private Sub txtSearchUnitName_Change()
    Dim Filter As String
    Dim tempstr As String
'   Dim i As Integer
'
''   If Len(txtSearchUnitName.text) > 0 Then
''      txtSearchName.text = ""
''      txtSearchTenant.text = ""
''   End If
'
'   For i = 1 To flxLeaseList.Rows - 1
'      flxLeaseList.RowHeight(i) = 240
'      If UCase(Left(flxLeaseList.TextMatrix(i, 4), Len(txtSearchUnitName.text))) <> UCase(txtSearchUnitName.text) Then
'         flxLeaseList.RowHeight(i) = 0
'      End If
'   Next i

    If Len(Trim(txtSearchUnitName.text)) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
      txtSearchCompany.text = ""
      tempstr = Replace(UCase(txtSearchUnitName.text), "'", "''")
        Filter = " UnitName LIKE '%" + tempstr + "*'"
    Else
          Filter = ""
   End If
   If sText = "UnitName" Then
       Call FilterTenantsList(Filter)
   End If
End Sub

Private Sub txtSearchUnitName_GotFocus()
    sText = "UnitName"
End Sub

Private Sub txtSearchUnitName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxLeaseList.SetFocus
    End If
End Sub

Private Sub txtSearchUnitName_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not cmdSaveNew.Visible Then
'        FilterTenantsList
   End If
End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cboRRDemandType.SetFocus
        FocusControl cmdRentReviewDemandType
    End If
End Sub

Private Sub txtStopIC_Change()
   TextBoxChangeDate txtStopIC
End Sub

Private Sub txtStopIC_KeyPress(KeyAscii As Integer)
    'addded by anol 21 Jun 2016
    If KeyAscii = 13 And cmdSaveEdit.Enabled Then
        FocusControl cmdIncSave
    End If
   TextBoxKeyPrsDate txtStopIC, KeyAscii
End Sub

Private Sub txtStopIC_LostFocus()
   If txtStopIC.text <> "" Then TextBoxFormatDate txtStopIC

   If txtStopIC.text <> "" Then Exit Sub

   'If flxIns.TextMatrix(flxIns.row, 16) = "" Or flxIns.TextMatrix(flxSC.row, 16) = "StopDate" Then Exit Sub

'   Frame4(0).Top = tabLease.Top + flxIns.Top + 320
'   Frame4(0).Left = tabLease.Left + flxIns.Left + flxIns.Width - Frame4(0).Width + 60
'   Frame4(0).Visible = True
'   tabLease.Enabled = False
'   txtNDD.SetFocus
End Sub

Private Sub txtStopRC_Change()
   TextBoxChangeDate txtStopRC
End Sub

Private Sub txtStopRC_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl cmdNewRentCrg
'    End If
    If KeyAscii = 13 And cmdSaveEdit.Enabled Then
        FocusControl cmdSaveRentCrg
    End If
   TextBoxKeyPrsDate txtStopRC, KeyAscii
End Sub

Private Sub txtStopRC_LostFocus()
   If txtStopRC.text <> "" Then TextBoxFormatDate txtStopRC

   If txtStopRC.text <> "" Then Exit Sub

   If flxRentCharges.TextMatrix(flxRentCharges.row, 16) = "" Then Exit Sub

'   Frame4(0).Top = tabLease.Top + flxRentCharges.Top + 320
'   Frame4(0).Left = tabLease.Left + flxRentCharges.Left + flxRentCharges.Width - Frame4(0).Width + 60
'   Frame4(0).Visible = True
'   tabLease.Enabled = False
'   txtNDD.SetFocus
End Sub

Private Sub txtStopSC_Change()
   TextBoxChangeDate txtStopSC
End Sub

Private Sub txtStopSC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCapAmount.SetFocus
    End If
   TextBoxKeyPrsDate txtStopSC, KeyAscii
End Sub

Private Sub txtStopSC_LostFocus()
   If txtStopSC.text <> "" Then TextBoxFormatDate txtStopSC
   
   If txtStopSC.text <> "" Then Exit Sub

   If flxSC.TextMatrix(flxSC.row, 18) = "" Or flxSC.TextMatrix(flxSC.row, 18) = "StopDate" Then Exit Sub

'   Frame4(0).Top = tabLease.Top + flxSC.Top + 320
'   Frame4(0).Left = tabLease.Left + flxSC.Left + flxSC.Width - Frame4(0).Width + 60
'   Frame4(0).Visible = True
'   tabLease.Enabled = False
'   txtNDD.SetFocus
End Sub

'Private Sub txtSuppCaption1_GotFocus()
'   SelTxtInCtrl txtSuppCaption1
'End Sub
'
'Private Sub txtSuppCaption2_GotFocus()
'   SelTxtInCtrl txtSuppCaption2
'End Sub
'
'Private Sub txtSuppCaption3_GotFocus()
'   SelTxtInCtrl txtSuppCaption3
'End Sub




Private Sub txtTerminationDate_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
         tabLease.Enabled = True
         Frame1(16).Enabled = True
         fraList(3).Visible = False
         'terminate_lease
    End If
End Sub
Private Sub terminate_lease(dtTerminationDate As Date)
        Dim Conn1 As New ADODB.Connection
        If IsDate(dtTerminationDate) = False Then
            MsgBox "You have not entered any termination date, It will not terminate the lease", vbInformation, "Warning"
            Exit Sub
        End If
        'issue 600 by anol 20180615
        Conn1.Open getConnectionString
        szSQL = "UPDATE LeaseDetails " & _
                "SET Status = False, TerminateDate = #" & Format(dtTerminationDate, "dd mmmm yyyy") & "# " & _
                "WHERE LeaseID = '" & strLeaseId & "'"
                
'         szSQL = "UPDATE LeaseDetails " & _
'                "SET Status = False, TerminateDate = #" & Format(Date, "dd mmmm yyyy") & "# " & _
'                "WHERE LeaseID = '" & txtLeaseID.text & "'"
                

          Conn1.Execute szSQL
    
          szSQL = "UPDATE Units " & _
                    "SET Occupied = 'N' " & _
                    "WHERE UnitNumber = '" & txtUnitNumber.text & "'"

          Conn1.Execute szSQL
          GetRecord Conn1
          Conn1.Close
          Set Conn1 = Nothing
          Call SetAddNewMode
          cmdTerminate.Visible = False
          MsgBox "The lease have been terminated successfully on " & Format(dtTerminationDate, "dd/mm/yyyy") & ".", vbInformation + vbOKOnly, "Lease Terminate"
End Sub




Private Sub txtTotalRentYear_GotFocus()
   If txtRentStartDate.Enabled Then Exit Sub
   If txtFreqBR.text = "" Then
      MsgBox "Please select the Frequency before input the total.", vbCritical + vbOKOnly, "Frequency"
      txtTotalRentYear.text = ""
   End If
End Sub

Private Sub txtTotalRentYear_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   If KeyAscii = 13 Then
        txtRentDueEachPeriod.SetFocus
   End If
   DigitTextKeyPress txtTotalRentYear, KeyAscii
End Sub

Private Sub txtTotalYearlyIns_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInsEachPeriod.SetFocus
    End If
End Sub

Private Sub txtYearEnd_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtYearEnd
End Sub

Private Sub txtYearEnd_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtYearEnd, KeyAscii
End Sub

'Private Sub txtNDD_Change()
'   TextBoxChangeDate txtNDD
'End Sub

'Private Sub txtNDD_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtNDD, KeyAscii
'End Sub

'Private Sub txtNDD_LostFocus()
'   If txtNDD.text <> "" Then
'      If TextBoxFormatDate(txtNDD) Then cmdNDD_OK.SetFocus
'   End If
'End Sub

Private Sub txtRentStartDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtRentStartDate
End Sub

Private Sub txtRentStartDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   If KeyAscii = 13 Then
        FocusControl cmdFreqBR
   End If
   TextBoxKeyPrsDate txtRentStartDate, KeyAscii
End Sub

Private Sub txtInsPercentage_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   On Error GoTo Err
   If KeyAscii = 13 Then
   ''addded by anol 21 Jun 2016
        txtStopIC.SetFocus
   End If
   If cboIncCharMth.Column(0) = 1 Then
      DigitTextKeyPress txtInsPercentage, KeyAscii, 6
   Else
      DigitTextKeyPress txtInsPercentage, KeyAscii, 6
   End If
   Exit Sub
Err:
   If Err.Number = 381 Then
        MsgBox "Please select charging method correctly", vbInformation, "Warning"
        FocusControl cboIncCharMth
   End If
End Sub

Private Sub txtBRChargingFigure_LostFocus()
   If Trim(txtBRChargingFigure.text) = "" Then
        txtBRChargingFigure.text = "0.00"
        Exit Sub
   End If

   If cboBRChargingMth.text = "Price Per Sq Foot" Then RCPricePerSqFoot
   If cboBRChargingMth.text = "Percentage" Then RCPercentage
   If cboBRChargingMth.text = "Annual" Then RCAnnual
   
End Sub

Private Sub txtInsPercentage_LostFocus()
   If Trim(txtInsPercentage.text) = "" Then
        txtInsPercentage.text = "0.00"
        Exit Sub
   End If

   If cboIncCharMth.text = "Price Per Sq Foot" Then IBPricePerSqFoot
   If cboIncCharMth.text = "Percentage" Then IBPercentage
   If cboIncCharMth.text = "Annual" Then ICAnnual
'   If txtInsPercentage.text = "" Or txtInsPercentage.text = "0.00" Then
'      txtInsPercentage.text = "0.00"
'      txtTotalYearlyIns.text = "0.00"
'      txtInsEachPeriod.text = "0.00"
'      Exit Sub
'   End If
'
'   Dim Area As String, iPartOfYear As Integer
'   Dim szSQL As String
'   Dim TotalInsurance As Double
'   Dim Conn1 As New ADODB.Connection
'
'   Conn1.Open getConnectionString
'
'   szSQL = "SELECT PARTOFYEAR " & _
'             "FROM FREQUENCIES " & _
'             "WHERE ID = " & Val(txtFreqIC.Tag) & ";"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   iPartOfYear = CInt(Rst1!PartOfYear)
'   Rst1.Close
'
'   If cboIncCharMth.Column(0) = 2 Then
'      txtTotalYearlyIns.text = Format(IIf(txtInsPercentage.text = "", 0, txtInsPercentage.text), "0.00")
'      txtInsEachPeriod.text = Format(CDbl(txtInsPercentage.text) / iPartOfYear, "0.00")
'   Else
'      szSQL = "SELECT Amount " & _
'                "FROM GlobalInsurance, Units " & _
'                "WHERE Units.PropertyID = GlobalInsurance.PropertyID " & _
'                  "AND Units.UnitNumber = '" & Left(cboUnit.text, 8) & "' " & _
'                  "AND GlobalInsurance.DemandType = " & CByte(txtInsDemandType.tag) & ";"
'      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'      If Not Rst1.EOF Then
'         TotalInsurance = CDbl(Rst1!Amount)
'      Else
'         TotalInsurance = 0
'      End If
'      Rst1.Close
'
'      txtTotalYearlyIns.text = Format(TotalInsurance * (CDbl(txtInsPercentage.text) / 100), "0.00")
'      txtInsEachPeriod.text = Format((TotalInsurance * (CDbl(txtInsPercentage.text) / 100) / iPartOfYear), "0.00")
'   End If
'   Conn1.Close
End Sub

Private Sub SCAnnual()
   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection

   txtChargingFigure.text = Format(IIf(txtChargingFigure.text = "", 0, txtChargingFigure.text), "0.00")

   Total = CDbl(txtChargingFigure.text)
   txtSCTotalAmount.text = Format(Total, "0.00")

   Conn1.Open getConnectionString
'   If cboFreqSC.ListIndex > -1 Then
'      szSQL = "SELECT PARTOFYEAR " & _
'                "FROM FREQUENCIES " & _
'                "WHERE ID = " & (cboFreqSC.ListIndex + 1) & ";"
'   Else
'      Dim temp() As String
'
'      temp = Split(cboFreqSC.text, "-")
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqSC.Tag & ";"
'   End If
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   txtSCDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub RCAnnual()
   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection

   txtBRChargingFigure.text = Format(IIf(txtBRChargingFigure.text = "", 0, txtBRChargingFigure.text), "0.00")

   Total = CDbl(Val(txtBRChargingFigure.text))
   txtTotalRentYear.text = Format(Total, "0.00")

   Conn1.Open getConnectionString
'   If cboFreqBR.ListIndex > -1 Then
'      szSQL = "SELECT PARTOFYEAR " & _
'                "FROM FREQUENCIES " & _
'                "WHERE ID = " & (cboFreqBR.ListIndex + 1) & ";"
'   Else
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqBR.Tag & ";"
'   End If
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   txtRentDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub ICAnnual()
   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection
   If IsNull(txtFreqIC.Tag) Then
        MsgBox "Please select a valid frequency", vbInformation, "Warning"
        FocusControl cmdFreqIC
        Exit Sub
   End If
   txtInsPercentage.text = Format(IIf(txtInsPercentage.text = "", 0, txtInsPercentage.text), "0.00")

   Total = CDbl(txtInsPercentage.text)
   txtTotalYearlyIns.text = Format(Total, "0.00")

   Conn1.Open getConnectionString
'   If cboInsFreq.ListIndex > -1 Then
'      szSQL = "SELECT PARTOFYEAR " & _
'                "FROM FREQUENCIES " & _
'                "WHERE ID = " & (cboInsFreq.ListIndex + 1) & ";"
'   Else
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqIC.Tag & ";"
'   End If
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   txtInsEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
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
   If KeyAscii = 13 Then
      txtInitiatedBy.SetFocus
   End If
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

Private Sub txtInsStartDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtInsStartDate
End Sub

Private Sub txtInsStartDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   If KeyAscii = 13 Then
        FocusControl cmdFreqIC
   End If
   TextBoxKeyPrsDate txtInsStartDate, KeyAscii
End Sub

Private Sub txtInsStartDate_LostFocus()
   If cmdIncNew.Enabled Then Exit Sub

   TextBoxFormatDate txtInsStartDate
   If txtInsDemandType.text = "" Then
      MsgBox "Please select a demand type.", vbCritical + vbOKOnly, "Insurance Charge"
      FocusControl cmdInsDemandType
   End If
   Call cboInsFreq_LostFocus
End Sub

Private Function OnlyNumericString(szString As String) As String
   Dim i As Integer, X As Integer
   
   For i = 1 To Len(szString)
      X = Asc(Mid(szString, i, 1))
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

Private Sub RCPricePerSqFoot()
   If txtBRChargingFigure.Locked Then Exit Sub

   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection

   txtBRChargingFigure.text = Format(IIf(txtBRChargingFigure.text = "", 0, txtBRChargingFigure.text), "0.0000")

   Conn1.Open getConnectionString

   Area = GetUnitTA(Conn1, txtUnitNumber.text)

   If Area = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      Total = Area * CDbl(txtBRChargingFigure.text)
      txtTotalRentYear.text = Format(Total, "0.00")
   
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqBR.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'Debug.Print szSQL
      txtRentDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

      Rst1.Close
      Set Rst1 = Nothing
   End If

   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub IBPricePerSqFoot()
   If txtInsPercentage.Locked Then Exit Sub

   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection

   txtInsPercentage.text = Format(IIf(txtInsPercentage.text = "", 0, txtInsPercentage.text), "0.0000")

   Conn1.Open getConnectionString

   Area = GetUnitTA(Conn1, txtUnitNumber.text)

   If Area = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      Total = Area * CDbl(txtInsPercentage.text)
      txtTotalYearlyIns.text = Format(Total, "0.00")

      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqIC.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'Debug.Print szSQL
      txtInsEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

      Rst1.Close
      Set Rst1 = Nothing
   End If

   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub SCPricePerSqFoot()
   If txtChargingFigure.Locked Then Exit Sub

   Dim Area As String, Total As Double
   Dim Conn1 As New ADODB.Connection

   txtChargingFigure.text = Format(IIf(txtChargingFigure.text = "", 0, txtChargingFigure.text), "0.00")

   Conn1.Open getConnectionString

   Area = GetUnitTA(Conn1, txtUnitNumber.text)

   If Area = 0 Then
      MsgBox "      The Area of the unit has not been set." & (Chr(13) + Chr(10)) & _
             "Please enter the area of the unit in the Unit Screen.", vbInformation, "Unit Total Area"
   Else
      Total = Area * CDbl(txtChargingFigure.text)
      txtSCTotalAmount.text = Format(Total, "0.00")
   
'      Dim temp() As String
'      temp = Split(cboFreqSC.text, "-")
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqSC.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   
      txtSCDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")
      'Cap amount added by anol 12 Mar 2015
'      If Val(txtCapAmount.text) <> 0 And Val(txtSCTotalAmount.text) < Val(txtCapAmount.text) Then
'         txtSCTotalAmount.Tag = Format((txtCapAmount.text / CInt(Rst1!PartOfYear)), "0.00")
'      Else
'         txtSCTotalAmount.Tag = ""
'      End If
      
      Rst1.Close
      Set Rst1 = Nothing
   End If

   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub txtIntPayableAfterDays_LostFocus()
'   txtInterestDescription.SetFocus
End Sub

Private Sub txtAdditionalIntRate_LostFocus()
   If txtAdditionalIntRate.text <> "" Then
      txtAdditionalIntRate.text = Format(txtAdditionalIntRate.text, "0.00")
      If txtInt2bChrg.text <> "" Then txtNoIntDays_LostFocus
   End If
End Sub

Private Sub txtBreakDate_LostFocus()
   'Added By Asif. 13/01/2006
   ' Modified by Samrat 08/02/2006
   If txtBreakDate.text <> "" Then TextBoxFormatDate txtBreakDate
End Sub
'
'Private Sub txtRentReviewDt_LostFocus()
'   If txtRentReviewDt.text <> "" Then If CheckDate(txtRentReviewDt.text) = False Then txtRentReviewDt.text = ""
'End Sub
'
'Private Sub txtRentIncDt_LostFocus()
'   If txtRentIncDt.text <> "" Then If CheckDate(txtRentIncDt.text) = False Then txtRentIncDt.text = ""
'End Sub
'
'Private Sub txtRentIncAmt_LostFocus()
'   If txtRentIncAmt.text <> "" Then
'       If NumberCheck2(txtRentIncAmt.text) = False Then
'           txtRentIncAmt.text = ""
'       Else
'           txtRentIncAmt.text = Round(CDbl(txtRentIncAmt.text), 2)
'       End If
'   End If
'End Sub

'Private Sub txtDtFlgDate_LostFocus()
'   TextBoxFormatDate txtDtFlgDate
'End Sub

Private Sub txtLeaseStDt_LostFocus()
   ' Modified by Samrat 08/02/2006
   If txtLeaseStDt.text <> "" Then TextBoxFormatDate txtLeaseStDt
End Sub

Private Sub txtLeaseEndDate_LostFocus()
   Dim YN

   If txtLeaseStDt.text = "" Then Exit Sub
   If txtLeaseEndDate.text = "" Then Exit Sub
   If Not TextBoxFormatDate(txtLeaseEndDate) Then
      txtLeaseEndDate.ForeColor = vbBlack
      Exit Sub
   End If

   If DateDiff("d", txtLeaseStDt.text, txtLeaseEndDate.text) < 0 And Not chkOLED.Value Then
      MsgBox "Lease Start date can not be after Lease End date.", vbCritical + vbOKOnly, "Date Error"
      txtLeaseEndDate.text = ""
      Exit Sub
   End If
'  the follwoing code might need to change after discuss with SALIA.
'  this follwoing code conflicts with the terminating lease.
'  see the code & note at @*#*#*@.
   If DateDiff("d", Now, txtLeaseEndDate.text) < 0 Then
      YN = MsgBox("The lease end date is earlier than the current date. " & Chr$(10) & "Do you wish to accept this date?", vbQuestion + vbYesNo, "Lease Expired")
      If YN = vbYes Then
            If cmdSaveEdit.Visible = True Then
                    frmTerminationDate.SourceOfCalling = "Link"
                    frmTerminationDate.LeaseEndDate = txtLeaseEndDate.text
                    frmTerminationDate.LeaseOverRide = CBool(chkOLED.Value)
                    'frmTerminationDate.txtTerminationDate.text = Label16(12).Tag
                    'If Label16(12).Tag = "" Then
                        DisplayDateform Me, "", txtLeaseEndDate.text
                    'Else
                     '   DisplayDateform Me, "", Label16(12).Tag
                    'End If
            End If
         If cmdSaveEdit.Visible Then FocusControl cmdSaveEdit
         If cmdSaveNew.Visible Then FocusControl cmdSaveNew
         Exit Sub
      End If

      SelTxtInCtrl txtLeaseEndDate
      FocusControl txtLeaseEndDate
      tabLease.Tab = 0
      txtLeaseEndDate.ForeColor = vbRed
   End If
End Sub

Private Sub txtYearEnd_LostFocus()
   ' Added by Asif. 13/01/2006
   ' Modified by Samrat 08/02/2006
   If txtYearEnd.text <> "" Then TextBoxFormatDate txtYearEnd
End Sub

Private Sub txtRentStartDate_LostFocus()
    TextBoxFormatDate txtRentStartDate
   'If txtRentStartDate.Enabled Then Exit Sub

'   If txtRentStartDate.text <> "" And txtFreqBR.text <> "" Then 'And txtNextDueDate.text = "" removed by anol 04/04/2016
'      If TextBoxFormatDate(txtRentStartDate) And txtNextDueDate.text <> "" Then 'I have added this part And txtNextDueDate.text <> ""  20180405
'         NextDueDate txtFreqBR, txtRentStartDate, txtNextDueDate, txtBRDemandType.Tag
'      End If
'   End If
    Call cboFreqBR_LostFocus
   If txtBRDemandType.text = "" And tabLease.Tab = 1 Then
      MsgBox "Please select a demand type.", vbCritical + vbOKOnly, "Rent Charges"
      FocusControl cmdBRDemandType
   End If
End Sub

Private Sub txtTotalRentYear_LostFocus()
   If txtTotalRentYear.text <> "" Then
      If Not IsNumeric(txtTotalRentYear.text) Then
         MsgBox "Please enter valid amount.", vbCritical + vbYesNo, "Error numeric value"
         SelTxtInCtrl txtTotalRentYear
         Exit Sub
       Else
         txtTotalRentYear.text = Format(Round(CDbl(txtTotalRentYear.text), 2), "0.00")
       End If
       If txtRentStartDate.text <> "" Then Call RentDueEachPeriod
   End If
End Sub

Public Sub RentDueEachPeriod()
   If txtTotalRentYear.text <> "" Then 'If there is a Base Rate for Year Figure
      Select Case txtFreqBR.Tag
         Case 1: ' Weekly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 52), 2)
         Case 2: ' Weekly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 52), 2)
         Case 3: ' Fortnightly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 26), 2)
         Case 4: ' Fortnightly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 26), 2)
         Case 5: ' Monthly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 12), 2)
         Case 6: ' Monthly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 12), 2)
         Case 7: ' Quarterly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 4), 2)
         Case 8: ' Quarterly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 4), 2)
         Case 9: ' Half yearly in advance
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 2), 2)
         Case 10: ' Half yearly in arrears
            txtRentDueEachPeriod.text = Round((CDbl(txtTotalRentYear.text) / 2), 2)
         Case 11: ' Yearly in advance
            txtRentDueEachPeriod.text = Round(CDbl(txtTotalRentYear.text), 2)
         Case 12: ' Yearly in arrears
            txtRentDueEachPeriod.text = Round(CDbl(txtTotalRentYear.text), 2)
      End Select
   End If
End Sub

'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'/*/*/*/*/*/*/*/*/*/*/*/*/*   NEXT DUE DATE  /*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
Public Sub NextDueDate(cboFrequency As Control, txtTextStart As TextBox, txtTextNext As TextBox, szDemandType As String)
   If txtUnitNumber.text = "" Then
       MsgBox "You must select a unit!", vbOKOnly + vbCritical, "No Unit Selected"
       Exit Sub
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not GetGlobalDataPropertyWise(PROPERTY_ID, adoConn, szDemandType) Then Exit Sub

   adoConn.Close
   Set adoConn = Nothing

   If szGDYearly = "AUTOMATIC" Then
      Select Case txtFreqBR.Tag
         Case 1:                              'Weekly in advance
            txtTextNext.text = txtTextStart.text
            
         Case 2:                              'Weekly in arrears
            txtTextNext.text = DateAdd("d", 7, txtTextStart.text)
         Case 3:                              'Fortnightly in advance
            txtTextNext.text = txtTextStart.text
         Case 4:                              'Fortnightly in arrears
            txtTextNext.text = DateAdd("d", 14, txtTextStart.text)
         Case 5:                              'Monthly in advance
            txtTextNext.text = txtTextStart.text
         Case 6:                              'Monthly in arrears
            txtTextNext.text = DateAdd("m", 1, txtTextStart.text)
         Case 7:                              'Quarterly in advance
            txtTextNext.text = txtTextStart.text
         Case 8:                              'Quarterly in arrears
            txtTextNext.text = DateAdd("m", 3, txtTextStart.text)
         Case 9:                              'Half yearly in advance
            txtTextNext.text = txtTextStart.text
         Case 10:                              'Half yearly in arrears
            txtTextNext.text = DateAdd("m", 6, txtTextStart.text)
         Case 11:                             'yearly in advance
            txtTextNext.text = txtTextStart.text
         Case 12:                             'yearly in arrears
            txtTextNext.text = DateAdd("m", 12, txtTextStart.text)
         Case 13:                             'Daily
            txtTextNext.text = ""
         Case 14:                             '4 Weekly in advance
            txtTextNext.text = txtTextStart.text
         Case 15:                             '4 Weekly in arrears
            txtTextNext.text = DateAdd("d", 28, txtTextStart.text)
         Case 16:                             '4 Monrhly in advance
            txtTextNext.text = txtTextStart.text
         Case 17:                             '4 Monrhly in arrears
            txtTextNext.text = DateAdd("m", 4, txtTextStart)
      End Select

      Exit Sub
   End If

   Select Case cboFrequency.Tag
      Case 1:                              'Weekly in advance
         txtTextNext.text = txtTextStart.text
      Case 2:                              'Weekly in arrears
         txtTextNext.text = DateAdd("d", 7, txtTextStart.text)
      Case 3:                              'Fortnightly in advance
         txtTextNext.text = txtTextStart.text
      Case 4:                              'Fortnightly in arrears
         txtTextNext.text = DateAdd("d", 14, txtTextStart.text)
      Case 5:                              'Monthly in advance
         txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Monthly)
      Case 6:                              'Monthly in arrears
         txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Monthly)
      Case 7:                              'Quarterly in advance
         txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Quarterly)
      Case 8:                              'Quarterly in arrears
         txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Quarterly)
      Case 9:                              'Half yearly in advance
         txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Half_Yearly)
      Case 10:                              'Half yearly in arrears
         txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Half_Yearly)
      Case 11:                             'yearly in advance
         txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Yearly)
      Case 12:                             'yearly in arrears
         txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Yearly)
      Case 13:                             'Daily
         txtTextNext.text = ""
      Case 14:                             '4 Weekly in advance
         txtTextNext.text = txtTextStart.text
      Case 15:                             '4 Weekly in arrears
         txtTextNext.text = DateAdd("d", 28, txtTextStart.text)
      Case 16:                             '4 Monrhly in advance
         txtTextNext.text = txtTextStart.text
      Case 17:                             '4 Monrhly in arrears
         txtTextNext.text = DateAdd("m", 4, txtTextStart)
   End Select
End Sub

Private Sub txtRentIncreaseDate_Change()
   TextBoxChangeDate txtRentIncreaseDate
   
End Sub

Private Sub txtRentIncreaseDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRentIncreaseAmount.SetFocus
    End If
    If bReviewLocked = True And txtRentIncreaseDate.Locked Then
        ShowMsgInTaskBar "Rent increase has been applied, You cannot edit the line"
    End If
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
    If KeyAscii = 13 Then
        txtComments.SetFocus
    End If
    If bReviewLocked = True And txtRentReviewDate.Locked Then
        ShowMsgInTaskBar "Rent increase has been applied, You cannot edit the line"
    End If
   'Added By Samrat. 16/01/2006
   TextBoxKeyPrsDate txtRentReviewDate, KeyAscii
End Sub

Private Sub txtRentReviewDate_LostFocus()
   'Added By Asif. 13/01/2006
   'Modified By Samrat. 06/02/2006
   If Not txtRentReviewDate.Locked Then TextBoxFormatDate txtRentReviewDate
End Sub

Private Sub RCPercentage()
   Dim Total As Double, TotalRentCharge As String ', temp() As String
   Dim Conn1 As New ADODB.Connection

   ' The following code is to calculate SC payable according to percentage
   Conn1.Open getConnectionString

   TotalRentCharge = GetGlobalTotalRC(Conn1, txtUnitNumber.text, txtRCFundCode.Tag)
   If TotalRentCharge < 0 Then
      MsgBox "        A rent budget has not been set for this property." & (Chr(13) + Chr(10)) & _
             "Please enter a rent budget for this property before proceeding further.", vbInformation, "Property Rent Charge"
   Else
      Total = CDbl(TotalRentCharge) * (CDbl(Val(txtBRChargingFigure.text)) / 100)
      txtTotalRentYear.text = Format(Total, "0.00")
'      temp = Split(cboFreqBR.text, "-")
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqBR.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      txtRentDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

      Rst1.Close
      Set Rst1 = Nothing

      txtBRChargingFigure.text = Format(IIf(txtBRChargingFigure.text = "", 0, txtBRChargingFigure.text), "0.0000")
   End If

   Conn1.Close
   Set Conn1 = Nothing
   Exit Sub

ErrorHander:
   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub
Private Sub CheckGlobalDataBudgetYearSet()
'written by anol 20210930
    Dim strCheck As String
    Dim rsSQL As New ADODB.Recordset
    Dim adoConn1 As New ADODB.Connection
    adoConn1.Open getConnectionString
    Dim szSQL  As String
    szSQL = "SELECT P.CBY " & _
        "FROM Property P " & _
        "WHERE P.PropertyID = '" & PROPERTY_ID & "';"
    rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
    If Not rsSQL.EOF Then
        strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
    End If
    If strCheck = "" Then
        MsgBox "An Insurance charge budget year has not been set for this property." & (Chr(13) + Chr(10)) & " Please set an Insurance charge budget year in the global data screen", vbInformation, "Warning!!"
    End If
    adoConn1.Close
End Sub
Private Sub IBPercentage()
   Dim Total As Double, TotalInsCharge As String
   Dim Conn1 As New ADODB.Connection
   If txtICFundCode.text = "" Then
        MsgBox "Please select a fund", vbInformation, "Warning "
        FocusControl txtICFundCode
        Exit Sub
   End If
   ' The following code is to calculate SC payable according to percentage
   Conn1.Open getConnectionString

   TotalInsCharge = GetGlobalTotalIC(Conn1, txtUnitNumber.text, txtICFundCode.Tag)
   If TotalInsCharge < 0 Then
      Call CheckGlobalDataBudgetYearSet
      MsgBox "        An Insurance Budget has not been set for this property." & (Chr(13) + Chr(10)) & _
             "Please enter an Insurance Budget amount in the Insurance Budget screen.", _
             vbInformation, "Property Insurance Charge"
'      MsgBox "        An Insurance Budget has not been set for this property." & (Chr(13) + Chr(10)) & _
'             "Please enter an Insurance Budget amount in the Insurance Budget screen.", _
'             vbInformation, "Property Insurance Charge"
'                Call CheckGlobalDataBudgetYearSet
   Else
      Total = CDbl(TotalInsCharge) * (CDbl(Val(txtInsPercentage.text)) / 100)
      txtTotalYearlyIns.text = Format(Total, "0.00")

      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqIC.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      txtInsEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

      Rst1.Close
      Set Rst1 = Nothing

      txtInsPercentage.text = Format(IIf(txtInsPercentage.text = "", 0, txtInsPercentage.text), "0.000000")
   End If

   Conn1.Close
   Set Conn1 = Nothing
   Exit Sub

ErrorHander:
   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub SCPercentage()
   Dim Total As Double, TotalServiceCharge As String
   Dim Conn1 As New ADODB.Connection
  
   'Test for issue 471 Note 718
   ' The following code is to calculate SC payable according to percentage
   'issue 471 Note 718
   '05 Nov 2014
   Conn1.Open getConnectionString
   Dim rstRst1 As New ADODB.Recordset
   Dim szSQL12 As String
   Dim strflag As String
   szSQL12 = "SELECT CBY " & _
           "FROM Property,Units " & _
           "WHERE " & _
                 "Property.PropertyID = Units.PropertyID AND " & _
                 "Units.UnitNumber = '" & txtUnitNumber.text & "';"
'Debug.Print szSQL
   rstRst1.Open szSQL12, Conn1, adOpenStatic, adLockReadOnly
   If rstRst1.EOF = False Then
         strflag = IIf(IsNull(rstRst1("CBY").Value), "", rstRst1("CBY").Value)
   End If
   'strflag = Null
   If strflag = "" Then 'A service charge budget year has not been set for this property. Please set a service charge budget year in the global data screen
   
      MsgBox " A service charge budget year has not been set for this property" & (Chr(13) + Chr(10)) & _
             "Please set a service charge budget year in the global data screen.", vbInformation, "Global data"
'     MsgBox " A service charge budget year has not been set for this property" & (Chr(13) + Chr(10)) & _
'             "Please set a service charge budget year in the global data screen.", vbInformation, "Global data"
      'ControlsModeServiceCharges DefaultMode
      
      '****** Calcel service charge code is here
        ControlsModeServiceCharges ExpensesMode
      'added by anol 04 Feb 2015
      txtCapAmount.Locked = True
      txtCapAmount.Enabled = False
      'Added by anol 25 FEB 2015
      cmdSCDemandType.Enabled = False
      txtPayableFrom.Locked = True
      cmdFreqSC.Enabled = False
      txtSCNextDueDt.Locked = True
     cmdSCFund.Enabled = False
      cboSchedule.Locked = True
      cboSCChargingMth.Locked = True
      txtChargingFigure.Locked = True
      txtSCTotalAmount.Locked = True
      txtSCDueEachPeriod.Locked = True
      txtStopSC.Locked = True
  
      '*****************
      SERVICECHARGES_EDIT = 0
      Exit Sub
   End If
   'End of modification

   TotalServiceCharge = GetGlobalTotalSC(Conn1, txtUnitNumber.text, txtSCFundCode.Tag)
   'modified by anol 21 Sep 2014
   If TotalServiceCharge < 0 Then
'      MsgBox "       A service charge budget has not been set up for this property" & (Chr(13) + Chr(10)) & _
'             "Please enter the a service charge budget in the service charge budget screen.", vbInformation, "Property Service Charge"
            MsgBox "A service charge budget has not been set up for this property. " & (Chr(13) + Chr(10)) & "Please enter a service charge budget amount in the service charge budget screen"
            'ShowMsgInTaskBar "A service charge budget has not been set up for this property. Please enter a service charge budget amount in the service charge budget screen", "Y"
              'issue 471 Note 718
              
              'comment out by anol 18 Jan 2016
'            ControlsModeServiceCharges DefaultMode
'            SERVICECHARGES_EDIT = 0
   Else
      Total = CDbl(TotalServiceCharge) * (CDbl(Val(txtChargingFigure.text)) / 100)
      txtSCTotalAmount.text = Format(Total, "0.00")
      
      szSQL = "SELECT PARTOFYEAR " & _
                "FROM FREQUENCIES " & _
                "WHERE ID = " & txtFreqSC.Tag & ";"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      txtSCDueEachPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

      Rst1.Close
      Set Rst1 = Nothing
       'Below line is modified by anol 25 Feb 2015
      txtChargingFigure.text = Format(IIf(txtChargingFigure.text = "", 0, txtChargingFigure.text), "0.00000000")
   End If

   Conn1.Close
   Set Conn1 = Nothing
   Exit Sub

ErrorHander:
   Rst1.Close
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub
'
'Private Sub txtSuppCaption1_LostFocus()
'   txtSuppCaption1.Visible = False
'   lblSupplementary1.Caption = IIf(txtSuppCaption1.text = "", lblSupplementary1.Caption, txtSuppCaption1.text)
'End Sub
'
'Private Sub txtSuppCaption2_LostFocus()
'   txtSuppCaption2.Visible = False
'   lblSupplementary2.Caption = IIf(txtSuppCaption2.text = "", lblSupplementary2.Caption, txtSuppCaption2.text)
'End Sub
'
'Private Sub txtSuppCaption3_LostFocus()
'   txtSuppCaption3.Visible = False
'   lblSupplementary3.Caption = IIf(txtSuppCaption3.text = "", lblSupplementary3.Caption, txtSuppCaption3.text)
'End Sub
Private Function ValidateSaveBreaches() As Boolean
    If cboBreachType.text = "" Then
        MsgBox "Please select breach type", vbInformation, "Warning!"
        FocusControl cboBreachType
        Exit Function
    End If
    
    If txtCommenceDate.text = "" Then
        MsgBox "Please enter commence date", vbInformation, "Warning!"
        FocusControl txtCommenceDate
        Exit Function
    End If
    
    If txtInitiatedBy.text = "" Then
        MsgBox "Please enter Initiated By", vbInformation, "Warning!"
        FocusControl txtInitiatedBy
        Exit Function
    End If
    
     If txtDateReceived.text = "" Then
        MsgBox "Please enter Date Received", vbInformation, "Warning!"
        FocusControl txtDateReceived
        Exit Function
    End If
    
    
    If strBreachID = "" Then
        Exit Function
    End If
    ValidateSaveBreaches = True
End Function
Public Function SaveBreaches() As Boolean
    Dim conBreach As New ADODB.Connection
    Dim rstBreach As New ADODB.Recordset
    Dim sSQLQuery_ As String
    Dim sSQLDelete As String
    Dim sSQLFilter As String
    Dim iRowIndex As Integer

    sSQLFilter = ""

   ' On Error GoTo Exception
    'Set the RDO Connections to the dataset
'    conBreach.Open getConnectionString
'
'    If Not BREACH_NEW_ENTRY_ Then
'        sSQLFilter = "WHERE LeaseId = '" & txtLeaseID.text & "' AND BreachID = " & txtBreachID.text & ""
'    Else
'        sSQLFilter = ""
'    End If
'
'    sSQLQuery_ = "SELECT * " & _
'    "FROM LeaseBreaches " & sSQLFilter
'
'    rstBreach.Open sSQLQuery_, conBreach, adOpenDynamic, adLockOptimistic
'
'    'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
'    If BREACH_NEW_ENTRY_ Then rstBreach.AddNew
'
'    rstBreach!LeaseID = txtLeaseID.text
'    rstBreach!BreachType = cboBreachType.Column(0)
'    rstBreach!CommenceDate = IIf(txtCommenceDate.text = "", Null, txtCommenceDate.text)
'    rstBreach!InitiatedBy = txtInitiatedBy.text
'    If chkResolved.Value = 1 Then
'        rstBreach!Resolved = True
'    Else
'        rstBreach!Resolved = False
'    End If
'    rstBreach!DateReceived = IIf(txtDateReceived.text = "", Null, txtDateReceived.text)
'    rstBreach!ReceivedBy = txtReceivedBy.text
'    rstBreach.Update
'
'    rstBreach.Close
'    conBreach.Close
'    Set rstBreach = Nothing
'    Set conBreach = Nothing


    If Breach_EDIT = 0 Then
      If gridBreach.TextMatrix(gridBreach.Rows - 1, 0) <> "" Then gridBreach.AddItem ""
      gridBreach.TextMatrix(gridBreach.Rows - 1, 0) = cboBreachType.text
      gridBreach.TextMatrix(gridBreach.Rows - 1, 1) = Format(txtCommenceDate.text, "dd/mm/yyyy")
      gridBreach.TextMatrix(gridBreach.Rows - 1, 2) = txtInitiatedBy.text
      gridBreach.TextMatrix(gridBreach.Rows - 1, 3) = Format(txtDateReceived.text, "dd/mm/yyyy")
      gridBreach.TextMatrix(gridBreach.Rows - 1, 4) = txtReceivedBy.text
      gridBreach.TextMatrix(gridBreach.Rows - 1, 5) = IIf(chkResolved.Value = 0, "No", "Yes")
      gridBreach.TextMatrix(gridBreach.Rows - 1, 6) = strBreachID
      gridBreach.TextMatrix(gridBreach.Rows - 1, 8) = cboBreachType.Column(0)
      gridBreach.TextMatrix(gridBreach.Rows - 1, 9) = txtMemo2.text

   Else
      gridBreach.TextMatrix(gridBreach_EDIT, 0) = cboBreachType.text
      gridBreach.TextMatrix(gridBreach_EDIT, 1) = Format(txtCommenceDate.text, "dd/mm/yyyy")
      gridBreach.TextMatrix(gridBreach_EDIT, 2) = txtInitiatedBy.text
      gridBreach.TextMatrix(gridBreach_EDIT, 3) = Format(txtDateReceived.text, "dd/mm/yyyy")
      gridBreach.TextMatrix(gridBreach_EDIT, 4) = txtReceivedBy.text
      gridBreach.TextMatrix(gridBreach_EDIT, 5) = IIf(chkResolved.Value = 0, "No", "Yes")
      gridBreach.TextMatrix(gridBreach_EDIT, 6) = strBreachID
      gridBreach.TextMatrix(gridBreach_EDIT, 8) = cboBreachType.Column(0)
      gridBreach.TextMatrix(gridBreach_EDIT, 9) = txtMemo2.text
  End If

    SaveBreaches = True
    'loadFlxBreach
    Exit Function

Exception:

    MsgBox Err.Number & " - " & Err.description, vbOKOnly, "Error"
    SaveBreaches = False
End Function

Public Function SaveAssignment() As Boolean
    Dim conAssignment As New ADODB.Connection
    Dim rstAssignment As New ADODB.Recordset
    Dim sSQLQuery_ As String
    Dim sSQLDelete As String
    Dim sSQLFilter As String
    Dim iRowIndex As Integer

    sSQLFilter = ""

    On Error GoTo Exception
    'Set the RDO Connections to the dataset
    conAssignment.Open getConnectionString

    If Not ASSIGNMENT_NEW_ENTRY_ Then
        sSQLFilter = "WHERE LeaseId = '" & strLeaseId & "' AND AssignmentID = " & strAssignmentID & ""
    Else
        sSQLFilter = ""
    End If

    sSQLQuery_ = "SELECT * " & _
    "FROM LeaseAssignments " & sSQLFilter

    rstAssignment.Open sSQLQuery_, conAssignment, adOpenDynamic, adLockOptimistic

    'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
    If ASSIGNMENT_NEW_ENTRY_ Then rstAssignment.AddNew

    rstAssignment!LeaseID = strLeaseId
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
    
    MsgBox Err.Number & " - " & Err.description, vbOKOnly, "Error"
    SaveAssignment = False
End Function

Public Sub BreachButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control
   Select Case mode
   
   Case ComponentMode.DefaultMode
       cmdBreachNew.Enabled = False
       cmdBreachEdit.Enabled = False
       cmdBreachSave.Enabled = False
       cmdBreachCancel.Enabled = False
       cmdDeleteBreaches.Enabled = False

       gridBreach.Enabled = True

       cboBreachType.Enabled = False
       cmdSetBreachType.Enabled = False
       txtCommenceDate.Locked = True
       txtInitiatedBy.Locked = True
       chkResolved.Enabled = False
       txtDateReceived.Locked = True
       txtReceivedBy.Locked = True
       txtMemo2.Locked = True

Case ComponentMode.ExpensesMode
       cmdBreachNew.Enabled = True
       cmdBreachEdit.Enabled = False
       cmdBreachSave.Enabled = False
       cmdBreachCancel.Enabled = False
       cmdDeleteBreaches.Enabled = False

       gridBreach.Enabled = True

       cboBreachType.Enabled = False
       cmdSetBreachType.Enabled = False
       txtCommenceDate.Locked = True
       txtInitiatedBy.Locked = True
       chkResolved.Enabled = False
       txtDateReceived.Locked = True
       txtReceivedBy.Locked = True
       txtMemo2.Locked = True

   Case ComponentMode.GridRowOnSelection
       cmdBreachNew.Enabled = True
       cmdBreachEdit.Enabled = True
       cmdBreachSave.Enabled = False
       cmdBreachCancel.Enabled = False
        cmdDeleteBreaches.Enabled = True
   
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
       txtMemo2.Locked = False
       txtReceivedBy.text = ""
       txtMemo2.text = ""

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
       txtMemo2.Locked = False
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
         cmdSaveRentAnalysis.Enabled = False
         cmdCancelRentAnalysis.Enabled = False
         cmdDelRentAnalysis.Enabled = False

         flxRentAnalysis.Enabled = True
         flxRentAnalysis.row = 0

      Case ComponentMode.GridRowOnSelection
         cmdNewRentAnalysis.Enabled = True
         cmdEditRentAnalysis.Enabled = True
         cmdSaveRentAnalysis.Enabled = False
         cmdCancelRentAnalysis.Enabled = False
         cmdDelRentAnalysis.Enabled = True

         flxRentAnalysis.Enabled = True
'         flxRentAnalysis.row = 0

      Case ComponentMode.NewEntryMode
         cmdNewRentAnalysis.Enabled = False
         cmdEditRentAnalysis.Enabled = False
         cmdSaveRentAnalysis.Enabled = True
         cmdCancelRentAnalysis.Enabled = True
         cmdDelRentAnalysis.Enabled = False

         flxRentAnalysis.Enabled = False

      Case ComponentMode.EditMode
         cmdNewRentAnalysis.Enabled = False
         cmdEditRentAnalysis.Enabled = False
         cmdSaveRentAnalysis.Enabled = True
         cmdCancelRentAnalysis.Enabled = True
         cmdDelRentAnalysis.Enabled = False

         flxRentAnalysis.Enabled = False
   End Select
End Sub

Private Sub ControlsModeRentCharges(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         txtTotalRentYear.text = ""
         txtRentDueEachPeriod.text = ""
         txtBRChargingFigure.Locked = True
         cmdBRDemandType.Enabled = False
         txtNextDueDate.Locked = True
         txtRCFundCode.text = ""
         txtRCFundCode.Tag = ""
         txtRCFund.text = ""
         cmdRCFund.Enabled = False
         cboBRChargingMth.text = ""
         cboBRChargingMth.Locked = True
         txtBRChargingFigure.text = ""
         txtBRChargingFigure.Locked = True
         txtFreqBR.text = ""
         cmdFreqBR.Enabled = False
         txtRentStartDate.text = ""
         txtRentStartDate.Locked = True
         txtNextDueDate.text = ""
'         txtNextDueDate.Locked = True
         txtTotalRentYear.text = ""
         txtTotalRentYear.Locked = True
         txtRentDueEachPeriod.text = ""
         txtBRDemandType.text = ""
         txtStopRC.text = ""
         txtStopRC.Locked = True

         txtRentDesc.text = ""
         chkRentDes.Value = 1

         cmdNewRentCrg.Enabled = False
         cmdEditRentCrg.Enabled = False
         cmdSaveRentCrg.Enabled = False
         cmdCancelRentCrg.Enabled = False
         cmdDelRentCrg.Enabled = False

         flxRentCharges.Enabled = True
         flxRentCharges.row = 0
         flxRentCharges.col = 0

         chkRentDes.Visible = False
         txtRentDesc.Visible = False
         lblDefaultDescption(1).Visible = False
      Case ComponentMode.ExpensesMode ' THIS IS ECTUALLY BIG EDIT MODEand cacale mode rent as well
         txtTotalRentYear.text = ""
         txtRentDueEachPeriod.text = ""
         txtBRChargingFigure.Locked = True
         txtNextDueDate.Locked = True
         cmdBRDemandType.Enabled = False
         txtRCFundCode.text = ""
         txtRCFundCode.Tag = ""
         txtRCFund.text = ""
         cmdRCFund.Enabled = False
         cboBRChargingMth.text = ""
         cboBRChargingMth.Locked = True
         txtBRChargingFigure.text = ""
         txtBRChargingFigure.Locked = True
         txtFreqBR.text = ""
        cmdFreqBR.Enabled = False
         txtRentStartDate.text = ""
         txtRentStartDate.Locked = True
         txtNextDueDate.text = ""
'         txtNextDueDate.Locked = True
         txtTotalRentYear.text = ""
         txtTotalRentYear.Locked = True
         txtRentDueEachPeriod.text = ""
         txtBRDemandType.text = ""
         txtStopRC.text = ""
         txtStopRC.Locked = True

         txtRentDesc.text = ""
         chkRentDes.Value = 1

         cmdNewRentCrg.Enabled = True
         cmdEditRentCrg.Enabled = False
         cmdSaveRentCrg.Enabled = False
         cmdCancelRentCrg.Enabled = False
         cmdDelRentCrg.Enabled = False

         flxRentCharges.Enabled = True
         flxRentCharges.row = 0
         flxRentCharges.col = 0

         chkRentDes.Visible = False
         txtRentDesc.Visible = False
         lblDefaultDescption(1).Visible = False

      Case ComponentMode.EditMode
         txtBRChargingFigure.Locked = False
         cmdBRDemandType.Enabled = True
         cmdRCFund.Enabled = True
         cmdFreqBR.Enabled = True
         txtRentStartDate.Locked = False
         txtTotalRentYear.Locked = False
         cboBRChargingMth.Locked = False
         txtBRChargingFigure.Locked = False
         txtStopRC.Locked = False
         
         cmdNewRentCrg.Enabled = False
         cmdEditRentCrg.Enabled = False
         cmdSaveRentCrg.Enabled = True
         cmdCancelRentCrg.Enabled = True
         cmdDelRentCrg.Enabled = False

         flxRentCharges.Enabled = False
         txtNextDueDate.Locked = False

      Case ComponentMode.NewEntryMode
         
         txtBRChargingFigure.Locked = False
         txtNextDueDate.Locked = False
         cmdBRDemandType.Enabled = True
         txtRCFundCode.text = ""
         txtRCFundCode.Tag = ""
         txtRCFund.text = ""
         cmdRCFund.Enabled = True
         cboBRChargingMth.text = ""
         cboBRChargingMth.Locked = False
         txtBRChargingFigure.text = ""
         txtBRChargingFigure.Locked = False
         txtFreqBR.text = ""
         cmdFreqBR.Enabled = True
         txtRentStartDate.text = ""
         txtRentStartDate.Locked = False
         txtNextDueDate.text = ""
         txtTotalRentYear.text = ""
         txtTotalRentYear.Locked = False
         txtRentDueEachPeriod.text = ""
         txtBRDemandType.text = ""
         txtStopRC.text = ""
         txtStopRC.Locked = False

         cmdNewRentCrg.Enabled = False
         cmdEditRentCrg.Enabled = False
         cmdSaveRentCrg.Enabled = True
         cmdCancelRentCrg.Enabled = True
         cmdDelRentCrg.Enabled = False

         flxRentCharges.Enabled = False

      Case ComponentMode.GridRowOnSelection
         
         txtBRChargingFigure.Locked = True
         cmdNewRentCrg.Enabled = False
         cmdEditRentCrg.Enabled = True
         cmdSaveRentCrg.Enabled = False
         cmdCancelRentCrg.Enabled = False
         cmdDelRentCrg.Enabled = True
'         If flxRentCharges.TextMatrix(flxRentCharges.row, 15) = "DELETED" Then
'            cmdDelRentCrg.Caption = "Not &Delete Rent"
'         Else
'            cmdDelRentCrg.Caption = "&Delete Rent"
'         End If
   End Select
End Sub

Private Sub ControlsModeServiceCharges(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         txtSCDueEachPeriod.text = ""
         txtSCTotalAmount.text = ""
         txtChargingFigure.Locked = True
         txtSCFundCode.text = ""
        txtSCFundCode.Tag = ""
        txtSCFundName.text = ""
        cmdSCFund.Enabled = False
         txtFreqSC.text = ""
         cmdFreqSC.Enabled = False
         txtPayableFrom.text = ""
         txtPayableFrom.Locked = True
         txtSCNextDueDt.text = ""
'         txtSCNextDueDt.Locked = True
         txtSCTotalAmount.text = ""
         txtSCTotalAmount.Locked = True
         txtSCDueEachPeriod.text = ""
         txtSCDueEachPeriod.Locked = True
         txtSCDemandType.text = ""
         'cmdSCDemandType.Enabled=false
         cboSCChargingMth.text = ""
         cboSCChargingMth.Locked = True
         txtChargingFigure.text = ""
         txtChargingFigure.Locked = True
         cboSchedule.text = ""
         cboSchedule.Locked = True
         txtStopSC.text = ""
         txtStopSC.Locked = True

         txtSCDesc.text = ""
         chkSCDes.Value = 1

         cmdSCNew.Enabled = True
         cmdSCEdit.Enabled = False
         cmdSCSave.Enabled = False
         cmdSCCancel.Enabled = False
         cmdSCDelete.Enabled = False

         flxSC.Enabled = True
         flxSC.row = 0
         flxSC.col = 0

         chkSCDes.Visible = False
         txtSCDesc.Visible = False
         lblDefaultDescption(4).Visible = False
    Case ComponentMode.ExpensesMode 'cancel mode
        txtSCDueEachPeriod.text = ""
        txtSCTotalAmount.text = ""
        txtChargingFigure.Locked = True
        txtSCFundCode.text = ""
        txtSCFundCode.Tag = ""
        txtSCFundName.text = ""
        cmdSCFund.Enabled = False
         txtFreqSC.text = ""
         cmdFreqSC.Enabled = False
         txtPayableFrom.text = ""
         txtPayableFrom.Locked = True
         txtSCNextDueDt.text = ""
'         txtSCNextDueDt.Locked = True
         txtSCTotalAmount.text = ""
         txtSCTotalAmount.Locked = True
         txtSCDueEachPeriod.text = ""
         txtSCDueEachPeriod.Locked = True
         txtSCDemandType.text = ""
         'cmdSCDemandType.Enabled=false
         cboSCChargingMth.text = ""
         cboSCChargingMth.Locked = True
         txtChargingFigure.text = ""
         txtChargingFigure.Locked = True
         cboSchedule.text = ""
         cboSchedule.Locked = True
         txtStopSC.text = ""
         txtStopSC.Locked = True

         txtSCDesc.text = ""
         chkSCDes.Value = 1

         cmdSCNew.Enabled = True
         cmdSCEdit.Enabled = False
         cmdSCSave.Enabled = False
         cmdSCCancel.Enabled = False
         cmdSCDelete.Enabled = False

         flxSC.Enabled = True
         flxSC.row = 0
         flxSC.col = 0

         chkSCDes.Visible = False
         txtSCDesc.Visible = False
         lblDefaultDescption(4).Visible = False
      Case ComponentMode.EditMode
         txtChargingFigure.Locked = False
         cmdSCFund.Enabled = True
         cmdFreqSC.Enabled = True
         txtPayableFrom.Locked = False
         txtSCTotalAmount.Locked = False
         'cboSCDemandType.locked = False
         cboSCChargingMth.Locked = False
         txtChargingFigure.Locked = False
         cboSchedule.Locked = False
         txtStopSC.Locked = False
         
         cmdSCNew.Enabled = False
         cmdSCEdit.Enabled = False
         cmdSCSave.Enabled = True
         cmdSCCancel.Enabled = True
         cmdSCDelete.Enabled = False

         flxSC.Enabled = False

         chkSCDes.Visible = True

      Case ComponentMode.NewEntryMode
         txtChargingFigure.Locked = False
        txtSCFundCode.text = ""
        txtSCFundCode.Tag = ""
        txtSCFundName.text = ""
         cmdSCFund.Enabled = True
         txtFreqSC.text = ""
         cmdFreqSC.Enabled = True
         txtPayableFrom.text = ""
         txtPayableFrom.Locked = False
         txtSCNextDueDt.text = ""
         txtSCTotalAmount.text = ""
         txtSCTotalAmount.Locked = False
         txtSCDueEachPeriod.text = ""
         txtSCDemandType.text = ""
         'cboSCDemandType.locked = False
         cboSCChargingMth.text = ""
         cboSCChargingMth.Locked = False
         txtChargingFigure.text = ""
         txtChargingFigure.Locked = False
         cboSchedule.text = ""
         cboSchedule.Locked = False
         txtStopSC.text = ""
         txtStopSC.Locked = False

         cmdSCNew.Enabled = False
         cmdSCEdit.Enabled = False
         cmdSCSave.Enabled = True
         cmdSCCancel.Enabled = True
         cmdSCDelete.Enabled = False

         flxSC.Enabled = False
         chkSCDes.Visible = True

      Case ComponentMode.GridRowOnSelection
         txtChargingFigure.Locked = True
         cmdSCNew.Enabled = False
         cmdSCEdit.Enabled = True
         cmdSCSave.Enabled = False
         cmdSCCancel.Enabled = False
         cmdSCDelete.Enabled = True
          'added by anol 18 Jan 2016
         txtChargingFigure.Locked = True
         txtSCNextDueDt.Locked = True
'         If flxSC.TextMatrix(flxSC.row, 16) = "DELETED" Then
'            cmdSCDelete.Caption = "Not &Delete Charge"
'         Else
'            cmdSCDelete.Caption = "&Delete Charge"
'         End If
   End Select
End Sub

Private Sub ControlsModeInsuranceCharges(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         txtTotalYearlyIns.text = ""
         txtInsPercentage.text = ""
         txtInsEachPeriod.text = ""
         txtInsPercentage.Locked = True
         txtICFundCode.text = ""
         txtICFundCode.Tag = ""
         txtICFundName.text = ""
        
         cmdICFundCode.Enabled = False
         txtInsStartDate.text = ""
         txtInsStartDate.Locked = True
         txtFreqIC.text = ""
         cmdFreqIC.Enabled = False
         txtInsDemandType.text = ""
         txtInsDemandType.Tag = ""
         txtInsNextDueDate.Locked = True 'by anol 20199429
           'uncommented by anol 20160926
         cmdInsDemandType.Enabled = False
         txtInsNextDueDate.text = ""
'         txtInsNextDueDate.Locked = True
         cboIncCharMth.text = ""
         cboIncCharMth.Locked = True
         txtInsPercentage.text = ""
         txtInsPercentage.Locked = True
         txtTotalYearlyIns.text = ""
         txtInsEachPeriod.text = ""
         txtInsEachPeriod.Locked = True
         txtStopIC.text = ""
         txtStopIC.Locked = True

         txtInsDesc.text = ""
         chkInsDes.Value = 1

         cmdIncNew.Enabled = False
         cmdIncEdit.Enabled = False
         cmdIncSave.Enabled = False
         cmdIncCancel.Enabled = False
         cmdInsDelete.Enabled = False

         flxIns.Enabled = True
         flxIns.row = 0
         flxIns.col = 0
         chkInsDes.Visible = False
         txtInsDesc.Visible = False
         lblDefaultDescption(8).Visible = False
    Case ComponentMode.ExpensesMode 'this is also cancel mode for ins
         txtTotalYearlyIns.text = ""
         txtInsPercentage.text = ""
         txtInsPercentage.Locked = True
         txtInsNextDueDate.Locked = True
         txtICFundCode.text = ""
         txtICFundCode.Tag = ""
         txtICFundName.text = ""
         cmdICFundCode.Enabled = False
         txtInsStartDate.text = ""
         txtInsStartDate.Locked = True
         txtFreqIC.text = ""
         cmdFreqIC.Enabled = False
         txtInsDemandType.text = ""
         txtInsDemandType.Tag = ""
           'uncommented by anol 20160926
         cmdInsDemandType.Enabled = False
         txtInsNextDueDate.text = ""
'         txtInsNextDueDate.Locked = True
         cboIncCharMth.text = ""
         cboIncCharMth.Locked = True
         txtInsPercentage.text = ""
         txtInsPercentage.Locked = True
         txtTotalYearlyIns.text = ""
         txtInsEachPeriod.text = ""
         txtInsEachPeriod.Locked = True
         txtStopIC.text = ""
         txtStopIC.Locked = True

         txtInsDesc.text = ""
         chkInsDes.Value = 1

         cmdIncNew.Enabled = True
         cmdIncEdit.Enabled = False
         cmdIncSave.Enabled = False
         cmdIncCancel.Enabled = False
         cmdInsDelete.Enabled = False

         flxIns.Enabled = True
         flxIns.row = 0
         flxIns.col = 0
         chkInsDes.Visible = False
         txtInsDesc.Visible = False
         lblDefaultDescption(8).Visible = False
         '*******************************************
          'added by anol 03/08/2023
         txtFreqIC.Locked = True
         txtICFundCode.Locked = True
         txtICFundName.Locked = True
         '*******************************************
         

      Case ComponentMode.EditMode
         txtInsPercentage.Locked = False
         cmdICFundCode.Enabled = True
         cmdFreqIC.Enabled = True
         txtInsStartDate.Locked = False
         txtInsNextDueDate.Locked = False ''by anol 20199429
           'uncommented by anol 20160926
         cmdInsDemandType.Enabled = True
         cboIncCharMth.Locked = False
         txtInsPercentage.Locked = False
         cboSchedule.Locked = False
         txtStopIC.Locked = False
          '*******************************************
          'added by anol 03/08/2023
         txtFreqIC.Locked = True
         txtICFundCode.Locked = True
         txtICFundName.Locked = True
         '*******************************************
         
         cmdIncNew.Enabled = False
         cmdIncEdit.Enabled = False
         cmdIncSave.Enabled = True
         cmdIncCancel.Enabled = True
         cmdInsDelete.Enabled = False

         flxIns.Enabled = False
         chkInsDes.Visible = True

      Case ComponentMode.NewEntryMode
         txtInsPercentage.Locked = False
         txtInsNextDueDate.Locked = False ''by anol 20199429
         txtICFundCode.text = ""
         txtICFundCode.Tag = ""
         txtICFundName.text = ""
         cmdICFundCode.Enabled = True
         txtFreqIC.text = ""
         cmdFreqIC.Enabled = True
         txtInsStartDate.text = ""
         txtInsStartDate.Locked = False
         txtInsNextDueDate.text = ""
         txtTotalYearlyIns.text = ""
         txtInsEachPeriod.text = ""
         txtInsDemandType.text = ""
         txtInsDemandType.Tag = ""
         'uncommented by anol 20160926
         cmdInsDemandType.Enabled = True
         cboIncCharMth.text = ""
         cboIncCharMth.Locked = False
         txtInsPercentage.text = ""
         txtInsPercentage.Locked = False
         cboSchedule.text = ""
         cboSchedule.Locked = False
         txtStopIC.text = ""
         txtStopIC.Locked = False

         cmdIncNew.Enabled = False
         cmdIncEdit.Enabled = False
         cmdIncSave.Enabled = True
         cmdIncCancel.Enabled = True
         cmdInsDelete.Enabled = False
         
         '*******************************************
          'added by anol 03/08/2023
         txtFreqIC.Locked = True
         txtICFundCode.Locked = True
         txtICFundName.Locked = True
         '*******************************************
         

         flxIns.Enabled = False
         chkInsDes.Visible = True

      Case ComponentMode.GridRowOnSelection
         txtInsPercentage.Locked = True
         cmdIncNew.Enabled = False
         cmdIncEdit.Enabled = True
         cmdIncSave.Enabled = False
         cmdIncCancel.Enabled = False
         cmdInsDelete.Enabled = True
'         If flxIns.TextMatrix(flxIns.row, 15) = "DELETED" Then
'            cmdInsDelete.Caption = "Not &Delete Ins."
'         Else
'            cmdInsDelete.Caption = "&Delete Ins."
'         End If
   End Select
End Sub

Private Function AllDemandType(adoConn As ADODB.Connection) As Boolean
'   On Error GoTo ErrorHandler
   AllDemandType = True

'ErrorHandler:
   Dim adoRst1 As New ADODB.Recordset
   Dim szSQL As String, i As Integer, szaData() As String
   Dim iAllDemandType As Integer

   szSQL = "SELECT COUNT(*) AS C_I FROM DemandTypes"
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If adoRst1!C_I = 0 Then
      AllDemandType = False
      Exit Function
   End If

   adoRst1.Close

'  RENT           *****************************************************
   szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (CategoryCode = 1 OR CategoryCode = 4) AND " & _
                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iAllDemandType = adoRst1.RecordCount
   ReDim szaData(1, iAllDemandType) As String

   i = 0
   If Not adoRst1.EOF Then
      While Not adoRst1.EOF
         szaData(0, i) = adoRst1!Id
         szaData(1, i) = adoRst1!Type
         i = i + 1
         adoRst1.MoveNext
      Wend
   End If
'   cboBRDemandType.Clear
'   cboBRDemandType.Column() = szaData()
   adoRst1.Close

'  SERVICE CHARGE   *****************************************************
'Added Category 5 on 2020-12-02 issue 898
'   szSQL = "SELECT ID, Type FROM DemandTypes " & _
'             "WHERE (CategoryCode = 2 OR CategoryCode = 4 OR CategoryCode = 5) AND " & _
'                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
'   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   iAllDemandType = adoRst1.RecordCount
'   ReDim szaData(1, iAllDemandType) As String
'
'   i = 0
'   If Not adoRst1.EOF Then
'      While Not adoRst1.EOF
'         szaData(0, i) = adoRst1!Id
'         szaData(1, i) = adoRst1!Type
'         i = i + 1
'         adoRst1.MoveNext
'      Wend
'   End If
'   cboSCDemandType.Clear
'   adoRst1.Close
'   cboSCDemandType.Column() = szaData()

'  INSURANCE         *****************************************************
'   szSQL = "SELECT ID, Type FROM DemandTypes " & _
'             "WHERE (CategoryCode = 3 OR CategoryCode = 4) AND " & _
'                   "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
'   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   iAllDemandType = adoRst1.RecordCount
'   ReDim szaData(1, iAllDemandType) As String
'
'   i = 0
'   If Not adoRst1.EOF Then
'      While Not adoRst1.EOF
'         szaData(0, i) = adoRst1!Id
'         szaData(1, i) = adoRst1!Type
'         i = i + 1
'         adoRst1.MoveNext
'      Wend
'   End If
'   adoRst1.Close
'   cboInsDemandType.Clear
'   cboInsDemandType.Column() = szaData()

'  INTEREST   / ALL         *****************************************************
   szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iAllDemandType = adoRst1.RecordCount
   ReDim szaData(1, iAllDemandType) As String

   i = 0
   If Not adoRst1.EOF Then
      While Not adoRst1.EOF
         szaData(0, i) = adoRst1!Id
         szaData(1, i) = adoRst1!Type
         i = i + 1
         adoRst1.MoveNext
      Wend
   End If
   adoRst1.Close
   cboIntDemandType.Clear
   cboIntDemandType.Column() = szaData()

'  RENT REVIEW             ******************************************************
   szSQL = "SELECT DISTINCT ID, Type " & _
           "FROM LRentCharges, DemandTypes " & _
           "WHERE LRentCharges.BRDemandType = DemandTypes.ID AND " & _
                "(PropertyID = '" & PROPERTY_ID & "' OR PropertyID = 'ALL');"

'   szSQL = "SELECT ID, Type " & _
'           "FROM   LRentCharges, DemandTypes " & _
'           "WHERE  LRentCharges.BRDemandType = DemandTypes.ID AND " & _
'               "(PropertyID = '" & PROPERTY_ID & "' OR " & _
'               "PropertyID = 'ALL') " & _
'           "GROUP BY ID, Type;"
'Debug.Print szSQL
   adoRst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iAllDemandType = adoRst1.RecordCount
   ReDim szaData(1, iAllDemandType) As String

   i = 0
   If Not adoRst1.EOF Then
      While Not adoRst1.EOF
         szaData(0, i) = adoRst1!Id
         szaData(1, i) = adoRst1!Type
         i = i + 1
         adoRst1.MoveNext
      Wend
   End If
   adoRst1.Close
'   cboRRDemandType.Clear
  ' cboRRDemandType.Column() = szaData()

   AllDemandType = True
   Set adoRst1 = Nothing
End Function

'Private Sub LoadDept(adoConn As ADODB.Connection)
'   ' Error Handler
'   On Error GoTo Error_Handler
'   Dim iSel As Integer
'   Dim rRow As Integer, iRec As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim rsFundMatrix As New ADODB.Recordset
'   'added by anol 2020-11-08 issue 889
'   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoConn, adOpenStatic, adLockReadOnly
'   If rsFundMatrix("isfundAssign").Value = True Then
'        szSQL = "Select * from fundMatrix where PropertyID='" & PROPERTY_ID & "' and ClientID='" & txtClient.Tag & "' and isDeleted=false"
'        iSel = 1
'   Else
'        szSQL = "SELECT FundID, FundName, FundCode " & _
'           "FROM Fund;"
'           iSel = 2
'   End If
'   rsFundMatrix.Close
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'       If iSel = 2 Then
'            MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
'            txtRCFundCode.text = ""
'            txtRCFundCode.Tag = ""
'            txtRCFund.text = ""
'            cboSCDept.Clear
'            cboIntChargeDept.Clear
'       Else
'            MsgBox "Please assign a fund to this property.Please go to the fund form to assign it", vbExclamation, "Warning!"
'             txtRCFundCode.text = ""
'            txtRCFundCode.Tag = ""
'            txtRCFund.text = ""
'             cboSCDept.Clear
'             cboIntChargeDept.Clear
'       End If
'   Else
'      ReDim Data(2, adoRst.RecordCount) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
'         Data(1, rRow) = adoRst.Fields.Item("FundName").Value
'         Data(2, rRow) = adoRst.Fields.Item("FundCode").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'      txtRCFundCode.text = ""
'         txtRCFundCode.Tag = ""
'         txtRCFund.text = ""
'      cboRentChargeDept.Column() = Data()
'      cboSCDept.Clear
'      cboSCDept.Column() = Data()
'      cboIntChargeDept.Clear
'      cboIntChargeDept.Column() = Data()
'      cboInsDept.Clear
'      cboInsDept.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   MsgBox "Error in Loading fund.", vbExclamation, "Loading Fund"
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Public Sub DisableBoxes()
   cmdLease.Enabled = True
   cmdTenants.Enabled = False
'   cboTenant.Enabled = False
   
   
   cmdLease.Visible = True
   cmdTenants.Visible = False
   
   'cboTenant.Enabled = False
   

   
   cboUsage.Enabled = False
   cmdUsage(0).Enabled = False

'   Lease Details
   Frame2.Enabled = False

'   Rent Charges
   Frame1(1).Enabled = False

'   Rent Review
   Frame1(6).Enabled = False

'   Service Charges
'Modified by anol 20 Jan 2016
     Frame1(2).Enabled = False
     Frame1(12).Enabled = False
     
'    cmdSCDemandType.Enabled=false
'    txtPayableFrom.Locked = True
'    cmdFreqSC.Enabled= False
'    txtSCNextDueDt.Locked = True
'   cmdSCFund.enabled=false
'    cboSchedule.Locked = True
'    cboSCChargingMth.Locked = True
'    txtChargingFigure.Locked = True
'    txtSCTotalAmount.Locked = True
'    txtSCDueEachPeriod.Locked = True
'    txtStopSC.Locked = True
'    txtCapAmount.Locked = True
    
    'End of modififcation
'   Interest Charge
   cboIntCrgable.Enabled = False
   Frame3.Enabled = False

'   Break Clause
   cboBreakClause.Enabled = False
   Frame1(0).Enabled = False

'   Rent Review

'   Insurance
   Frame1(14).Enabled = False

'   Supplementary
'   Frame1(7).Enabled = False
'   Frame1(8).Enabled = False

'   txtMemo.Enabled = False
   Frame1(5).Enabled = False
'   txtSupp1.Enabled = False
'   txtSupp2.Enabled = False
'   txtSupp3.Enabled = False

   'Breaches
   Frame1(12).Enabled = False
End Sub

Public Sub EnableBoxes()
   cmdLease.Enabled = False
   cmdTenants.Enabled = True
   cmdLease.Visible = False
   cmdTenants.Visible = True

'   cboTenant.Enabled = True
   

   
   cboUsage.Enabled = True
   cmdUsage(0).Enabled = True

   'Lease details
   Frame2.Enabled = True

'   Rent Charges
   Frame1(1).Enabled = True

'   Rent Review
   Frame1(6).Enabled = True

'   Service Charges
   Frame1(2).Enabled = True
   Frame1(12).Enabled = True
'   cmdSCDemandType.Enabled=true
'    txtPayableFrom.Locked = False
'    cmdFreqSC.Enabled=  true
'    txtSCNextDueDt.Locked = False
'    cmdSCFund.enabled=true
'    cboSchedule.Locked = False
'    cboSCChargingMth.Locked = False
'    txtChargingFigure.Locked = False
'    txtSCTotalAmount.Locked = False
'    txtSCDueEachPeriod.Locked = False
'    txtStopSC.Locked = False
'    txtCapAmount.Locked = False
    
'   Interest Charge
   cboIntCrgable.Enabled = True
   Frame3.Enabled = IIf(cboIntCrgable.text = "Yes", True, False)

'   Break Clause
   cboBreakClause.Enabled = True
   Frame1(0).Enabled = True

'    Insurance
   Frame1(14).Enabled = True

'   Supplementary
   Frame1(7).Enabled = True
   Frame1(8).Enabled = True

'  Memo
   Frame1(5).Enabled = True

'   Breaches
   Frame1(12).Enabled = True
End Sub

Public Sub EmptyBoxes()
   txtUnitNumber.text = ""
   strLeaseId = ""
   txtSageAccountNumber.text = ""
   txtTenant.text = ""
   txtUnitName(0).text = ""

   txtClient.text = ""
   txtProperty.text = ""
   cboUsage.text = ""

   'Lease Details
'   cboHeadLease.ListIndex = -1
   txtUnitName(1).text = ""
   txtUnitName(1).Tag = ""
   chkSubLease.Value = 0
   cboType.text = ""
   txtYearEnd.text = ""
   txtLeaseStDt.text = ""
   txtLeaseEndDate.text = ""
   chkOLED.Value = False
   chkHoldingOver.Value = False
   chkGPrataDmd.Value = False

   'Breakes
   cboBreakClause.text = "No"
   txtBreakDate.text = ""
   cboBreak.text = ""

   'Service Charges
   txtPayableFrom.text = ""
   txtFreqSC.text = ""
   txtSCNextDueDt.text = ""
   txtSCDemandType.text = ""
   cboSchedule.text = ""

   'Interest Charges
   cboIntCrgable.text = "No"
   cboIntChargeDept.text = ""
   txtAdditionalIntRate.text = ""
   txtAmtCrgIntOn.text = ""
   txtNoIntDays.text = ""
   txtInt2bChrg.text = ""
   txtIntPayableAfterDays.text = ""
   cboIntDemandType.text = ""
   txtInterestDescription.text = ""

   'Breaches
   cboBreachType.text = ""
   txtCommenceDate.text = ""
   txtInitiatedBy.text = ""
   chkResolved.Value = 0
   txtDateReceived.text = ""
   txtReceivedBy.text = ""
   txtMemo2.text = ""

   'Assignment

   'Insurance

'   'Supplementary
'   txtDtFlgDate.text = ""
'   txtDtFlgDesc.text = ""
'   txtDtFlgDt2.text = ""
'   txtDtFlgDesc2.text = ""
'   txtDtFlgDt3.text = ""
'   txtDtFlgDesc3.text = ""
'   txtSupp1.text = ""
'   txtSupp2.text = ""
'   txtSupp3.text = ""
'
'   txtSuppCaption1.text = ""
'   txtSuppCaption2.text = ""
'   txtSuppCaption3.text = ""

   'Memo
   txtMemo.text = ""
End Sub

Private Sub LoadValues(adoConn As ADODB.Connection)
   Dim sSQLQuery As String
   
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'UUSE'"
   populateCombo adoConn, sSQLQuery, cboUsage
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

