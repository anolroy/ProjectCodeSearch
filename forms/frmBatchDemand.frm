VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBatchDemands 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Demands"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   3750
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchDemand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13170
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   12690
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   71
      Top             =   4860
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
         TabIndex        =   76
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   74
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7091
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
         TabIndex        =   73
         Top             =   375
         Width           =   4545
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   72
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   75
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "&Remove"
      Height          =   285
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Copy from lease"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtAmtIC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtPcgIC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtAmtSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtPcgSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtAmtRC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtPcgRC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Canc&el"
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
      Height          =   375
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Copy from lease"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Copy from lease"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame fraHeader 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton cmdFundIC 
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
         TabIndex        =   15
         Top             =   2070
         Width           =   300
      End
      Begin VB.CommandButton cmdFundSC 
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
         TabIndex        =   11
         Top             =   1710
         Width           =   300
      End
      Begin VB.CommandButton cmdFundRC 
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
         TabIndex        =   6
         Top             =   1350
         Width           =   300
      End
      Begin VB.CommandButton cmdDTIC 
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
         Left            =   3375
         TabIndex        =   14
         Top             =   2070
         Width           =   300
      End
      Begin VB.CommandButton cmdDTSC 
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
         Left            =   3375
         TabIndex        =   10
         Top             =   1710
         Width           =   300
      End
      Begin VB.CommandButton cmdDTRC 
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
         Left            =   3375
         TabIndex        =   5
         Top             =   1350
         Width           =   300
      End
      Begin VB.CommandButton cmdProperty 
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
         Left            =   4815
         TabIndex        =   1
         Top             =   540
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
         Left            =   4815
         TabIndex        =   0
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   4
         Top             =   525
         Width           =   1455
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   3
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txtDateIssue 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txtDescRC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1365
         Width           =   3375
      End
      Begin VB.CheckBox chkProIC 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7725
         TabIndex        =   16
         Top             =   2085
         Width           =   255
      End
      Begin VB.CheckBox chkProSC 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7725
         TabIndex        =   12
         Top             =   1725
         Width           =   255
      End
      Begin VB.CheckBox chkProRC 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7725
         TabIndex        =   8
         Top             =   1365
         Width           =   255
      End
      Begin VB.TextBox txtBudAmtRC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5925
         TabIndex        =   7
         Top             =   1365
         Width           =   1695
      End
      Begin VB.TextBox txtBudAmtSC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5925
         TabIndex        =   23
         Top             =   1725
         Width           =   1695
      End
      Begin VB.TextBox txtBudAmtIC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5925
         TabIndex        =   25
         Top             =   2085
         Width           =   1695
      End
      Begin VB.CheckBox chkIC 
         Height          =   255
         Left            =   1470
         TabIndex        =   24
         Top             =   2085
         Width           =   300
      End
      Begin VB.CheckBox chkSC 
         Height          =   255
         Left            =   1470
         TabIndex        =   22
         Top             =   1725
         Width           =   300
      End
      Begin VB.CheckBox chkRC 
         Height          =   255
         Left            =   1470
         TabIndex        =   21
         Top             =   1365
         Width           =   300
      End
      Begin VB.TextBox txtDescSC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1725
         Width           =   3375
      End
      Begin VB.TextBox txtDescIC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2085
         Width           =   3375
      End
      Begin MSForms.TextBox txtFundIC 
         Height          =   285
         Left            =   3780
         TabIndex        =   86
         Top             =   2070
         Width           =   1710
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3016;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtFundSC 
         Height          =   285
         Left            =   3780
         TabIndex        =   85
         Top             =   1710
         Width           =   1710
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3016;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtFundRC 
         Height          =   285
         Left            =   3780
         TabIndex        =   84
         Top             =   1350
         Width           =   1710
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3016;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDTIC 
         Height          =   285
         Left            =   1800
         TabIndex        =   83
         Top             =   2070
         Width           =   1575
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "2778;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDTSC 
         Height          =   285
         Left            =   1800
         TabIndex        =   82
         Top             =   1710
         Width           =   1575
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "2778;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDTRC 
         Height          =   285
         Left            =   1800
         TabIndex        =   81
         Top             =   1350
         Width           =   1575
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "2778;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   1080
         TabIndex        =   80
         Top             =   540
         Width           =   3735
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6588;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1080
         TabIndex        =   70
         Top             =   180
         Width           =   3735
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6588;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   7560
         TabIndex        =   69
         Top             =   165
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
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   60
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   59
         Top             =   165
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         Height          =   195
         Index           =   8
         Left            =   8805
         TabIndex        =   58
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date:"
         Height          =   195
         Index           =   7
         Left            =   5205
         TabIndex        =   57
         Top             =   165
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   9
         Left            =   3765
         TabIndex        =   56
         Top             =   1050
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   8325
         TabIndex        =   55
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Type"
         Height          =   195
         Index           =   0
         Left            =   1780
         TabIndex        =   54
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorata"
         Height          =   255
         Left            =   7725
         TabIndex        =   53
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Amount"
         Height          =   255
         Index           =   0
         Left            =   5925
         TabIndex        =   52
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Charge"
         Height          =   255
         Left            =   165
         TabIndex        =   51
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge"
         Height          =   255
         Left            =   165
         TabIndex        =   50
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Charge"
         Height          =   255
         Left            =   165
         TabIndex        =   49
         Top             =   2085
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1455
         Index           =   0
         Left            =   5865
         Top             =   1005
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         Height          =   195
         Index           =   19
         Left            =   8805
         TabIndex        =   48
         Top             =   165
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   11535
      End
   End
   Begin VB.TextBox txtInputGrid 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3600
      TabIndex        =   46
      Top             =   3240
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdPreviewDemands 
      Caption         =   "Preview Batch"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Copy from lease"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdBatchClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.OptionButton optSngBatchDemand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Single demand"
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   7845
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optConBatchDemand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consolidate demand"
      Height          =   375
      Left            =   2520
      TabIndex        =   32
      Top             =   7845
      Width           =   1815
   End
   Begin VB.CommandButton cmdBatchDemand 
      Caption         =   "Generate Demands"
      Height          =   375
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdLookupIC 
      Caption         =   "Apply %"
      Height          =   255
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Apply percentage from lease"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdLookupSC 
      Caption         =   "Apply %"
      Height          =   255
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Apply percentage from lease"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdLookupRC 
      Caption         =   "Apply %"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Apply percentage from lease"
      Top             =   3360
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemands 
      Height          =   3705
      Left            =   120
      TabIndex        =   45
      Top             =   3720
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6535
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
      BandDisplay     =   1
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   615
      Index           =   5
      Left            =   9200
      Top             =   3060
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   615
      Index           =   4
      Left            =   7160
      Top             =   3060
      Width           =   2060
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   615
      Index           =   3
      Left            =   5020
      Top             =   3060
      Width           =   2140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   61
      Top             =   7440
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   2
      Left            =   120
      Top             =   7800
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   3240
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   -120
      X2              =   13080
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount"
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   40
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "%"
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount"
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   38
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   36
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Charge"
      Height          =   375
      Index           =   8
      Left            =   9240
      TabIndex        =   43
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge"
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   34
      Top             =   3240
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   1
      X1              =   -120
      X2              =   13080
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rent Charge"
      Height          =   375
      Index           =   77
      Left            =   5040
      TabIndex        =   42
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBatchDemands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iLeft As Integer, iTop As Integer
Private szLeaseID_RC As String, szLeaseID_SC As String, szLeaseID_IC As String
Private szAmt_RC As String, szAmt_SC As String, szAmt_IC As String
Private szLease As String
Private iCurRow As Integer, iCurCol As Integer
Private szIC As String
Dim sTextBox As String
Private Sub LoadFunds()
   Dim adoRst As New ADODB.Recordset
   Dim conConnection As New ADODB.Connection
   Dim SQLStr1 As String
   conConnection.Open getConnectionString
   SQLStr1 = "SELECT FundID, FundCode, FundName FROM Fund;"
   adoRst.Open SQLStr1, conConnection, adOpenKeyset, adLockReadOnly

   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 2700
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3


   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 3600
   txtSearchClientID.Width = 1500
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   flxClient.Width = 5175
   
   cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   txtSearchClientName.Left = 1580
   txtSearchClientName.Width = 3600
   picClient.Height = 4095
   flxClient.Height = 3345
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
   lblClientID.Width = 1400
   lblClientID.Left = 250
   lblClientName.Width = 3600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   
   ReDim szaFundCode(adoRst.RecordCount, 2) As String
   
   Dim rRow As Integer
   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
       
               rRow = 1
'               flxClient.TextMatrix(rRow, 0) = ""
'               flxClient.TextMatrix(rRow, 1) = ""
'               flxClient.TextMatrix(rRow, 2) = ""
'               flxClient.RowHeight(rRow) = 280
'               flxClient.AddItem ""
'
'               rRow = 2
            While Not adoRst.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = "  " & adoRst.Fields.Item("FundID").Value
               flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("FundCode").Value
               flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("FundName").Value

                flxClient.RowHeight(rRow) = 280
               adoRst.MoveNext
               If Not adoRst.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   End If
   conConnection.Close
End Sub
Private Sub loadDemadtype()
    
   
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 2700
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.Clear
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3

   flxClient.ColWidth(0) = 200
   flxClient.ColWidth(1) = 0
   flxClient.ColWidth(2) = 2800
   picClient.Width = 3500
   cmdPicCLose.Left = 3200
   txtSearchClientID.Left = 145
   txtSearchClientID.Width = 3240
   

 
   flxClient.Rows = 2
   flxClient.Height = 2845
   If sTextBox = "1" Or sTextBox = "2" Then
        flxClient.Width = 5175
   Else
        flxClient.Width = 3200
   End If

   picClient.Height = 3595

   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Demand Type"
   lblClientName.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 250
   lblClientName.Width = 3600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   Dim rRow As Integer
   Dim szSQLStr As String
   Dim adoRst As New ADODB.Recordset
   Dim conConnection As New ADODB.Connection
   conConnection.Open getConnectionString
    If sTextBox = "3" Then
        szSQLStr = "SELECT ID, Type, CategoryCode " & _
              "FROM DemandTypes " & _
              "WHERE PropertyID = '" & txtProperty.Tag & "' AND CategoryCode=1;"
        adoRst.Open szSQLStr, conConnection, adOpenStatic, adLockReadOnly
    ElseIf sTextBox = "4" Then
        szSQLStr = "SELECT ID, Type, CategoryCode " & _
              "FROM DemandTypes " & _
              "WHERE PropertyID = '" & txtProperty.Tag & "' AND CategoryCode=2;"
        adoRst.Open szSQLStr, conConnection, adOpenStatic, adLockReadOnly
    ElseIf sTextBox = "5" Then
        szSQLStr = "SELECT ID, Type, CategoryCode " & _
              "FROM DemandTypes " & _
              "WHERE PropertyID = '" & txtProperty.Tag & "' AND CategoryCode=3;"
        adoRst.Open szSQLStr, conConnection, adOpenStatic, adLockReadOnly
    End If
   
  
        rRow = 1
'           flxClient.TextMatrix(rRow, 0) = "" 'adoRst.Fields.Item(0).Value
'           flxClient.TextMatrix(rRow, 1) = ""
'           flxClient.TextMatrix(rRow, 2) = ""
'           flxClient.RowHeight(rRow) = 280
'           flxClient.AddItem ""
'           rRow = 2
        While Not adoRst.EOF

           flxClient.TextMatrix(rRow, 0) = "" 'adoRst.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 280
           adoRst.MoveNext
           If Not adoRst.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   
   conConnection.Close
End Sub

Private Sub cmdDTIC_Click()
    picClient.Left = 1869.029
    picClient.Top = 1355.299
    sTextBox = "5"
    loadDemadtype
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdDTRC_Click()
    picClient.Left = 1869.029
    picClient.Top = 1355.299
    sTextBox = "3"
    loadDemadtype
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdDTSC_Click()
    picClient.Left = 1869.029
    picClient.Top = 1355.299
    sTextBox = "4"
    loadDemadtype
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdFundIC_Click()
    picClient.Left = 3869.029
    picClient.Top = 1355.299
    sTextBox = "8"
    LoadFunds
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdFundRC_Click()
    picClient.Left = 3869.029
    picClient.Top = 1355.299
    sTextBox = "6"
    LoadFunds
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdFundSC_Click()
    picClient.Left = 3869.029
    picClient.Top = 1355.299
    sTextBox = "7"
    LoadFunds
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    fraHeader.Enabled = True
    cmdClientList.SetFocus
End Sub

Private Sub LoadflxProperty()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45

   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   adoconn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'            rRow = 1
'            flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = "ALL"
'           flxClient.TextMatrix(rRow, 2) = "ALL Properties"
'           flxClient.RowHeight(rRow) = 280
'           flxClient.AddItem ""
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub cmdproperty_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "2"
    LoadflxProperty
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        picClient.Visible = False
        fraHeader.Enabled = True
       
         If sTextBox = "1" Then
            cmdClientList.SetFocus
'         ElseIf sTextBox = "2" Then
'            cmdBC.SetFocus
'         ElseIf sTextBox = "3" Then
'            cmdPropecrty.SetFocus
'         ElseIf sTextBox = "4" Then
'            cmdUnit.SetFocus
'         ElseIf sTextBox = "5" Then
'            cmdNC.SetFocus
'         ElseIf sTextBox = "6" Then
'            cmdFund.SetFocus
'         ElseIf sTextBox = "7" Then
'            cmdVATCode.SetFocus
         End If
        
    End If
    
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
Private Sub flxClient_Click()
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
       
        fraHeader.Enabled = True
        
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                cmdProperty.SetFocus
'                LoadBankAccountInCombo adoConn
'                LoadNCinCombo adoConn
'                LoadCboProperty adoConn
'                cmdBC.SetFocus
       
        ElseIf sTextBox = "2" Then
                txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                cboPropertyList_Click
                txtDateIssue.SetFocus
          ElseIf sTextBox = "3" Then
                txtDTRC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtDTRC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtDTRC.text = "" Then
                    chkRC.Value = vbUnchecked
                Else
                    chkRC.Value = vbChecked
                End If
                cmdFundRC.SetFocus
            ElseIf sTextBox = "4" Then
                txtDTSC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtDTSC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtDTSC.text = "" Then
                    chkSC.Value = vbUnchecked
                Else
                    chkSC.Value = vbChecked
                End If
                cmdFundSC.SetFocus
           ElseIf sTextBox = "5" Then
                txtDTIC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtDTIC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtDTIC.text = "" Then
                    chkIC.Value = vbUnchecked
                Else
                    chkIC.Value = vbChecked
                End If
                cmdFundIC.SetFocus
           ElseIf sTextBox = "6" Then
                txtFundRC.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                txtFundRC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtFundRC.text = "" And txtDTRC.text = "" Then
                    chkRC.Value = vbUnchecked
                Else
                    chkRC.Value = vbChecked
                End If
                txtBudAmtRC.SetFocus
          ElseIf sTextBox = "7" Then
                txtFundSC.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                txtFundSC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtFundSC.text = "" And txtDTSC.text = "" Then
                    chkSC.Value = vbUnchecked
                Else
                    chkSC.Value = vbChecked
                End If
              txtBudAmtSC.SetFocus
         ElseIf sTextBox = "8" Then
                txtFundIC.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                txtFundIC.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtFundIC.text = "" And txtDTSC.text = "" Then
                    chkIC.Value = vbUnchecked
                Else
                    chkIC.Value = vbChecked
                End If
                txtBudAmtIC.SetFocus
'        ElseIf sTextBox = "4" Then
'                txtUnit.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                txtUnit.text = flxClient.TextMatrix(flxClient.row, 2)
'                txtDetails.SetFocus
'        ElseIf sTextBox = "5" Then
'                txtNC.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                txtNC.text = flxClient.TextMatrix(flxClient.row, 2)
'                cmdFund.SetFocus
'         ElseIf sTextBox = "6" Then
'                txtFund.Tag = flxClient.TextMatrix(flxClient.row, 0)
'                txtFund.text = flxClient.TextMatrix(flxClient.row, 2)
'                txtDate.SetFocus
'        ElseIf sTextBox = "7" Then
'                Label1(24).Caption = flxClient.TextMatrix(flxClient.row, 1)
'                Label1(24).Tag = flxClient.TextMatrix(flxClient.row, 2)
'                txtVat_.text = Format(Val(txtNet.text) * (Val(flxClient.TextMatrix(flxClient.row, 2)) / 100), "0.00")
'                txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
'                cmdSave.SetFocus
        End If
        
        picClient.Visible = False
        adoconn.Close
        Set adoconn = Nothing
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new

   
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub

Private Sub flxDemands_Click()
    txtInputGrid.text = ""
    txtInputGrid.Visible = False
End Sub

Private Sub txtDescIC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
End Sub

Private Sub txtDescRC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDTSC.SetFocus
    End If
End Sub

Private Sub txtDescSC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDTIC.SetFocus
    End If
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
        If sTextBox = "1" Or sTextBox = "2" Then
            txtSearchClientName.SetFocus
        ElseIf sTextBox = "3" Then
            flxClient.SetFocus
        End If
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)

If KeyAscii = 27 Then
        picClient.Visible = False
        fraHeader.Enabled = True
        
         If sTextBox = "1" Then
         ElseIf sTextBox = "1" Then
            cmdClientList.SetFocus
'         ElseIf sTextBox = "2" Then
'            cmdBC.SetFocus
'         ElseIf sTextBox = "3" Then
'            cmdproperty.SetFocus
'         ElseIf sTextBox = "4" Then
'            cmdUnit.SetFocus
'
'         ElseIf sTextBox = "5" Then
'            cmdNC.SetFocus
'         ElseIf sTextBox = "6" Then
'            cmdFund.SetFocus
'         ElseIf sTextBox = "7" Then
'            cmdVATCode.SetFocus
         End If
'
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

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub



Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        picClient.Visible = False
        fraHeader.Enabled = True
       
         If sTextBox = "1" Then
            cmdClientList.SetFocus
'         ElseIf sTextBox = "2" Then
'            cmdBC.SetFocus
'         ElseIf sTextBox = "3" Then
'            cmdProperty.SetFocus
'         ElseIf sTextBox = "4" Then
'            cmdUnit.SetFocus
'
'         ElseIf sTextBox = "5" Then
'            cmdNC.SetFocus
'         ElseIf sTextBox = "6" Then
'            cmdFund.SetFocus
'         ElseIf sTextBox = "7" Then
'            cmdVATCode.SetFocus
         End If
        
    End If
End Sub
Private Sub FillCboType(conConnection As ADODB.Connection)
'   Dim adoRst     As New ADODB.Recordset
'   Dim szSQLStr   As String
'   Dim Data()     As String
'   Dim i          As Integer
'   Dim j          As Integer
'   Dim c          As Integer
'   Dim b          As Boolean
'
'   szSQLStr = "SELECT ID, Type, CategoryCode " & _
'              "FROM DemandTypes " & _
'              "WHERE PropertyID = '" & txtProperty.Tag & "';"
'   adoRst.Open szSQLStr, conConnection, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "No Demand types have been set up for this company.", vbExclamation, "Please create your demand type."
'   Else
''                                                     Rent Charge
'      ReDim Data(1, adoRst.RecordCount) As String
'
'      For i = 0 To adoRst.RecordCount - 1
'         b = False
'         For j = 0 To 1
'            If adoRst.Fields(2).Value = 1 Then
'               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'               b = True
'            End If
'         Next j
'         If b Then c = c + 1
'
'         adoRst.MoveNext
'         If adoRst.EOF Then Exit For
'      Next i
'
'      ReDim Preserve Data(1, c) As String
'
'      cmbDTRC.Column() = Data()
'
''                                                     Service Charge
'      ReDim Data(1, adoRst.RecordCount) As String
'      adoRst.MoveFirst
'      c = 0
'
'      For i = 0 To adoRst.RecordCount - 1
'         b = False
'         For j = 0 To 1
'            If adoRst.Fields(2).Value = 2 Then
'               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'               b = True
'            End If
'         Next j
'         If b Then c = c + 1
'
'         adoRst.MoveNext
'         If adoRst.EOF Then Exit For
'      Next i
'
'      ReDim Preserve Data(1, c) As String
'
'      cmbDTSC.Column() = Data()
'
''                                                     Insurance Charge
'      ReDim Data(1, adoRst.RecordCount) As String
'      adoRst.MoveFirst
'      c = 0
'
'      For i = 0 To adoRst.RecordCount - 1
'         b = False
'         For j = 0 To 1
'            If adoRst.Fields(2).Value = 3 Then
'               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'               b = True
'            End If
'         Next j
'         If b Then c = c + 1
'
'         adoRst.MoveNext
'         If adoRst.EOF Then Exit For
'      Next i
'
'      ReDim Preserve Data(1, c) As String
'
'      cmbDTIC.Column() = Data()
'   End If
'
'   adoRst.Close
'   Set adoRst = Nothing
End Sub

'Private Sub cboClientList_Change()
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString
'
'   LoadPropertyByClient adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub LoadPropertyByClient(adoConn As ADODB.Connection)
'   Dim iRec As Integer
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, szaData() As String
'
'   On Error GoTo Error_Handler
'
'   szSQL = "SELECT PropertyID, PropertyName " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY PropertyName;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim szaData(1, adoRst.RecordCount - 1) As String
'
'   While Not adoRst.EOF
'      szaData(0, iRec) = adoRst.Fields.Item("PropertyID").Value
'      szaData(1, iRec) = adoRst.Fields.Item("PropertyName").Value
'      iRec = iRec + 1
'      adoRst.MoveNext
'   Wend
'
'   cboPropertyList.Clear
'   cboPropertyList.Column() = szaData()
'   cboPropertyList.ListIndex = -1
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Private Sub cboFundIC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboFundRC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboFundSC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboPropertyList_Click()          'LoadFlxDemands
   Dim adoconn As New ADODB.Connection
   Dim adoRstLeaseDtl As New ADODB.Recordset ', adoRstSplitDemand As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   ConfigureFlxDemands

   adoconn.Open getConnectionString

   szSQLStr = "SELECT L.LeaseID, " & _
                  "L.SageAccountNumber, " & _
                  "L.CompanyName " & _
              "FROM LeaseDetails AS L, Units AS U " & _
              "WHERE L.Status = TRUE AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & txtProperty.Tag & "' " & _
              "ORDER BY L.SageAccountNumber;"

   adoRstLeaseDtl.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   iRow = 1
   While Not adoRstLeaseDtl.EOF
      flxDemands.TextMatrix(iRow, 0) = adoRstLeaseDtl.Fields.Item("SageAccountNumber").Value
      flxDemands.TextMatrix(iRow, 1) = adoRstLeaseDtl.Fields.Item("CompanyName").Value
      flxDemands.TextMatrix(iRow, 8) = adoRstLeaseDtl.Fields.Item("LeaseID").Value

      iRow = iRow + 1
      adoRstLeaseDtl.MoveNext
      If Not adoRstLeaseDtl.EOF Then flxDemands.AddItem ""
   Wend
   adoRstLeaseDtl.Close

   Set adoRstLeaseDtl = Nothing

'  Load Demand type. Demand type is property wise. therefore, they have to loaded after user selects the property
   FillCboType adoconn

   adoconn.Close
   Set adoconn = Nothing
   txtDateIssue.SetFocus
End Sub

Private Sub chkIC_Click()
   If chkIC.Value = 0 Then
      txtDTIC.text = ""
      txtFundIC.text = ""
      txtBudAmtIC.text = ""
      chkProIC.Value = 0
      txtDescIC.text = ""
   Else
      Label7.ForeColor = vbBlack
   End If
End Sub

Private Sub chkProIC_GotFocus()
   If txtBudAmtIC.text = "" Or Val(txtBudAmtIC) < 0 Then txtBudAmtIC.SetFocus
End Sub

Private Sub chkProRC_GotFocus()
'   If txtBudAmtRC.text = "" Or Val(txtBudAmtRC.text) < 0 Then txtBudAmtRC.SetFocus
   If txtBudAmtRC.text = "" Then txtBudAmtRC.SetFocus
End Sub

Private Sub chkProSC_GotFocus()
   If txtBudAmtSC.text = "" Or Val(txtBudAmtSC.text) < 0 Then txtBudAmtSC.SetFocus
End Sub

Private Sub GenConBtDmds()
   Dim BRcount As Integer, SCcount As Integer, szSQLStr As String
   Dim iSerial As Integer, lDemand As Long, ICcount As Integer, iVATCode As Integer, iSplitID As Integer
   Dim cAmount As Currency, sChargingFig As Single, Msg As String
   Dim sChargingFigFalse As Single
   Dim adoRstDemandRec As New ADODB.Recordset, adoDmdTypRC As ADODB.Recordset
   Dim adoRstLeaseDtl As New ADODB.Recordset, adoRstSplitDemand As New ADODB.Recordset
   Dim adoDmdTypSC As ADODB.Recordset, adoDmdTypIC As ADODB.Recordset

   If MsgBox("  Are you sure you wish to generate batch demands?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Generate Batch Demands") = vbNo Then Exit Sub

   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

'   Connect to Demands table to add new demands.
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   If chkRC.Value = 1 Then
      Set adoDmdTypRC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTRC.Tag
      adoDmdTypRC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If
   'temporarily rem by anol 09 Jun 2016
   If chkSC.Value = 1 Then
      Set adoDmdTypSC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTSC.Tag
      adoDmdTypSC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If
   If chkIC.Value = 1 Then
      Set adoDmdTypIC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTIC.Tag
      adoDmdTypIC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If

   'Resolved by BOSL
   'Issue No: 0000476
   'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
   'of the default control accounts set in the demand types records.
   'If not found, then exit the function before it updates or post any transaction to the database.
   'Modified By: Asif. 27 Sep 2014
   
   Dim OutputVATNominalCode As String
   Dim SalesLedgerNominalCode As String
   
   Dim OutputVATNominalName As String
   Dim SalesLedgerNominalName As String
   
   OutputVATNominalCode = ""
   SalesLedgerNominalCode = ""

   OutputVATNominalCode = GetNominalCodeForControlAccount(adoconn, "Output VAT", txtClientList.Tag)
   If (OutputVATNominalCode = "") Then
       Exit Sub
   Else
       OutputVATNominalName = GetNominalNameOfCode(adoconn, OutputVATNominalCode, txtClientList.Tag)
'       If MsgBox("There is no Nominal Code set for Output VAT Control Account. Do you want to continue", vbYesNo, "No Nominal Code set for Output VAT Control") = vbNo Then
'            Exit Function
'       End If
   End If
   
   SalesLedgerNominalCode = GetNominalCodeForControlAccount(adoconn, "Sales Ledger Control", txtClientList.Tag)
   If (SalesLedgerNominalCode = "") Then
       Exit Sub
   Else
       SalesLedgerNominalName = GetNominalNameOfCode(adoconn, SalesLedgerNominalCode, txtClientList.Tag)
'       If MsgBox("There is no Nominal Code set for Sales Ledger Control Account. Do you want to continue", vbYesNo, "No Nominal Code set for Sales Ledger Control") = vbNo Then
'            Exit Function
'       End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''
   
'   ListOfAllLessee

   szSQLStr = "SELECT L.*, Units.PropertyID " & _
              "FROM LeaseDetails AS L, Units " & _
              "WHERE L.UnitNumber = Units.UnitNumber AND " & _
                    "L.LeaseID IN (" & szLease & ");"
'Debug.Print szSQLStr
   adoRstLeaseDtl.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   iSerial = 1
   
   If adoRstLeaseDtl.EOF Then
      adoRstLeaseDtl.Close
      Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF
         iSplitID = 1
         iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoconn, sChargingFig)

'*********************************************************************************************************
'         Rent Charges Demands
'*********************************************************************************************************
'**** Insert the Header info in the DemandRecPreview table
         lDemand = NextRef(adoconn, "DEMAND_REF")
         With adoRstDemandRec
            .AddNew
            .Fields.Item("CreatedBy").Value = User
            .Fields.Item("CreatedDate").Value = Now
            .Fields.Item("DemandID").Value = lDemand
            .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
            .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
            .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
            .Fields.Item("Source").Value = 1
            .Fields.Item("TransactionType").Value = 1
            .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
            .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
            .Fields.Item("IsPrinted").Value = False
            .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoconn)
            .Fields.Item("Spare1").Value = adoDmdTypRC.Fields.Item("spare1").Value
            .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
            .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
            .Fields.Item("LastModifiedBy").Value = User
            .Fields.Item("LastModifiedDate").Value = Now
            .Update
         End With

         If IsRCSelected(adoRstLeaseDtl!LeaseID, cAmount) Then
            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !splitID = iSplitID '1'Modified by anol 08 June 2016
               iSplitID = iSplitID + 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypRC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypRC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypRC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypRC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypRC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypRC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sChargingFig / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypRC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTRC.Tag
               !description = txtDescRC.text ' modified by anol 2020-10-12
               !dateFrom = CDate(txtDateFrom.text)
               !DateTO = CDate(txtDateTo.text)
               'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
                !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !SageDepartment = txtFundRC.Tag

               .Update
            End With

            BRcount = BRcount + 1
            iSerial = iSerial + 1
         End If
'************************************************************************************************
'         Service Charge demands
'************************************************************************************************
         If IsSCSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFigFalse) Then 'Modified by anol 08 June 2016
         ' If IsSCSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !splitID = iSplitID ' 1 Modified by anol 08 June 2016
               iSplitID = iSplitID + 1
               !DEMANDID = lDemand
               !A_M = "B"
               
               !NominalCodeforAmount = adoDmdTypSC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypSC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypSC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypSC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypSC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypSC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sChargingFig / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypSC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTSC.Tag
               !description = txtDescSC.text '' modified by anol 2020-10-12
               !dateFrom = txtDateFrom.text
               !DateTO = txtDateTo.text
                'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
                  !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !SageDepartment = txtFundSC.Tag
               !ChargingFigure = sChargingFig
               .Update
            End With

            SCcount = SCcount + 1
            iSerial = iSerial + 1
         End If

'************************************************************************************************
'   Insurance Charge demands
'************************************************************************************************
         If IsICSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFigFalse) Then 'Modified by anol 08 June 2016
         'If IsICSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               '!SplitID = 1
                !splitID = iSplitID ' 1 Modified by anol 08 June 2016
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypIC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypIC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypIC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypIC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypIC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypIC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sChargingFig / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypIC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTIC.Tag
               !description = txtDescIC.text ' modified by anol 2020-10-12
               !dateFrom = txtDateFrom.text
               !DateTO = txtDateTo.text
                'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
                !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !SageDepartment = txtFundIC.Tag
               !ChargingFigure = sChargingFig
               .Update
            End With

            ICcount = ICcount + 1
            iSerial = iSerial + 1
         End If

         adoRstLeaseDtl.MoveNext
      Wend

      If chkRC.Value = 1 Then
         adoDmdTypRC.Close
         Set adoDmdTypRC = Nothing
         Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
      End If
      If chkSC.Value = 1 Then
         adoDmdTypSC.Close
         Set adoDmdTypSC = Nothing
         Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
      End If
      If chkIC.Value = 1 Then
         adoDmdTypIC.Close
         Set adoDmdTypIC = Nothing
         Msg = Msg & ICcount & " Demands for Insurance Charge were generated." & Chr(13)
      End If

      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close

      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If

   MousePointer = vbDefault

   Msg = Msg & "A total of " & BRcount + SCcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Batch Demand Generated"

'  Bring all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoconn

   adoconn.Close
   Set adoconn = Nothing
   Exit Sub

ErrH:
'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
   MsgBox Err.Number & " - (pcm_001)" & Err.description, vbOKOnly, "Error"

   Set adoconn = Nothing
End Sub

Private Sub ListOfLeaseID()
   Dim iRow As Integer

   For iRow = 1 To flxDemands.Rows - 1
      If Val(flxDemands.TextMatrix(iRow, 3)) > 0 Then
         If szLeaseID_RC = "" Then
            szLeaseID_RC = flxDemands.TextMatrix(iRow, 8)
            szAmt_RC = flxDemands.TextMatrix(iRow, 3)
         Else
            szLeaseID_RC = szLeaseID_RC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_RC = szAmt_RC & ", " & flxDemands.TextMatrix(iRow, 3)
         End If
      End If
      
      If Val(flxDemands.TextMatrix(iRow, 5)) > 0 Then
         If szLeaseID_SC = "" Then
            szLeaseID_SC = flxDemands.TextMatrix(iRow, 8)
            szAmt_SC = flxDemands.TextMatrix(iRow, 5)
         Else
            szLeaseID_SC = szLeaseID_SC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_SC = szAmt_SC & ", " & flxDemands.TextMatrix(iRow, 5)
         End If
      End If
      
      If Val(flxDemands.TextMatrix(iRow, 7)) > 0 Then
         If szLeaseID_IC = "" Then
            szLeaseID_IC = flxDemands.TextMatrix(iRow, 8)
            szAmt_IC = flxDemands.TextMatrix(iRow, 7)
         Else
            szLeaseID_IC = szLeaseID_IC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_IC = szAmt_IC & ", " & flxDemands.TextMatrix(iRow, 7)
         End If
      End If
   Next iRow

   szLease = ""
   For iRow = 1 To flxDemands.Rows - 1
      If (Val(flxDemands.TextMatrix(iRow, 3)) > 0 Or _
          Val(flxDemands.TextMatrix(iRow, 5)) > 0 Or _
          Val(flxDemands.TextMatrix(iRow, 7)) > 0) And _
          InStr(szLease, Len(flxDemands.TextMatrix(iRow, 8))) = 0 Then

         szLease = szLease + IIf(szLease = "", "", ", ") + "'" & flxDemands.TextMatrix(iRow, 8) & "'"
      End If
   Next iRow
End Sub

Private Sub chkRC_Click()
   If chkRC.Value = 0 Then
      txtDTRC.text = ""
      txtFundRC.text = ""
      txtBudAmtRC.text = ""
      chkProRC.Value = 0
      txtDescRC.text = ""
   Else
      Label5.ForeColor = vbBlack
   End If
End Sub

Private Sub chkSC_Click()
   If chkSC.Value = 0 Then
      txtDTSC.text = ""
      txtFundSC.text = ""
      txtBudAmtSC.text = ""
      chkProSC.Value = 0
      txtDescSC.text = ""
   Else
      Label6.ForeColor = vbBlack
   End If
End Sub

Private Sub cmbDTIC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmbDTRC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmbDTSC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdBatchClose_Click()
   Unload Me
End Sub

Private Sub cmdBatchDemand_Click()
   Dim iRow As Integer, iCol As Integer

   If cmdOK.Enabled Then
      cmdOK.SetFocus
      Exit Sub
   End If

   ListOfLeaseID

   If optSngBatchDemand.Value Then
      GenSngBtDmds
   Else
      GenConBtDmds
   End If

   szLeaseID_RC = ""
   szLeaseID_SC = ""
   szLeaseID_IC = ""

   For iRow = 1 To flxDemands.Rows - 1
      For iCol = 2 To 7
         flxDemands.TextMatrix(iRow, iCol) = ""
      Next iCol
   Next iRow
   cmdCancel_Click
End Sub

Private Sub GenSngBtDmds()
   Dim BRcount As Integer, SCcount As Integer, szSQLStr As String
   Dim iVATCode As Integer
   Dim iSerial As Integer, lDemand As Long, ICcount As Integer
   Dim cAmount As Currency, sChargingFig As Single, Msg As String, sVatRate As Single

   Dim adoRstDemandRec As New ADODB.Recordset, adoDmdTypRC As ADODB.Recordset
   Dim adoRstLeaseDtl As New ADODB.Recordset, adoRstSplitDemand As New ADODB.Recordset
   Dim adoDmdTypSC As ADODB.Recordset, adoDmdTypIC As ADODB.Recordset

   If MsgBox("  Are you sure you wish to generate batch demands?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Generate Batch Demands") = vbNo Then Exit Sub

   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

'   Connect to Demands table to add new demands.
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   If chkRC.Value = 1 Then
      Set adoDmdTypRC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTRC.Tag 'cmbDTRC.Column(0)
      adoDmdTypRC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If
   If chkSC.Value = 1 Then
      Set adoDmdTypSC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTSC.Tag
      adoDmdTypSC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If
   If chkIC.Value = 1 Then
      Set adoDmdTypIC = New ADODB.Recordset
      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & txtDTIC.Tag
      adoDmdTypIC.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
   End If

   'Resolved by BOSL
   'Issue No: 0000476
   'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
   'of the default control accounts set in the demand types records.
   'If not found, then exit the function before it updates or post any transaction to the database.
   'Modified By: Asif. 27 Sep 2014
   
   Dim OutputVATNominalCode As String
   Dim SalesLedgerNominalCode As String
   
   Dim OutputVATNominalName As String
   Dim SalesLedgerNominalName As String
   
   OutputVATNominalCode = ""
   SalesLedgerNominalCode = ""

   OutputVATNominalCode = GetNominalCodeForControlAccount(adoconn, "Output VAT", txtClientList.Tag)
   If (OutputVATNominalCode = "") Then
       Exit Sub
   Else
       OutputVATNominalName = GetNominalNameOfCode(adoconn, OutputVATNominalCode, txtClientList.Tag)
'       If MsgBox("There is no Nominal Code set for Output VAT Control Account. Do you want to continue", vbYesNo, "No Nominal Code set for Output VAT Control") = vbNo Then
'            Exit Function
'       End If
   End If
   
   SalesLedgerNominalCode = GetNominalCodeForControlAccount(adoconn, "Sales Ledger Control", txtClientList.Tag)
   If (SalesLedgerNominalCode = "") Then
       Exit Sub
   Else
       SalesLedgerNominalName = GetNominalNameOfCode(adoconn, SalesLedgerNominalCode, txtClientList.Tag)
'       If MsgBox("There is no Nominal Code set for Sales Ledger Control Account. Do you want to continue", vbYesNo, "No Nominal Code set for Sales Ledger Control") = vbNo Then
'            Exit Function
'       End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''
   
   szSQLStr = "SELECT LeaseDetails.*, Units.PropertyID " & _
              "FROM LeaseDetails, Units " & _
              "WHERE LeaseDetails.UnitNumber = Units.UnitNumber;"

   adoRstLeaseDtl.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic

   iSerial = 1

   If adoRstLeaseDtl.EOF Then
      adoRstLeaseDtl.Close
      Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF
'*********************************************************************************************************
'         Rent Charges Demands
'*********************************************************************************************************
         If IsRCSelected(adoRstLeaseDtl!LeaseID, cAmount) Then
            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoconn, sVatRate)

'**** Insert the Header info in the DemandRecPreview table
            lDemand = NextRef(adoconn, "DEMAND_REF")
            With adoRstDemandRec
               .AddNew
               .Fields.Item("DemandID").Value = lDemand
                .Fields.Item("CreatedBy").Value = User
               .Fields.Item("CreatedDate").Value = Now
               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields.Item("Source").Value = 1
               .Fields.Item("TransactionType").Value = 1
               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("IsPrinted").Value = False
               .Fields.Item("Spare1").Value = adoDmdTypRC.Fields.Item("spare1").Value
               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoconn)
               .Fields.Item("Details").Value = "Rent"
               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Fields.Item("LastModifiedBy").Value = User
               .Fields.Item("LastModifiedDate").Value = Now
               .Update
            End With

            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !splitID = 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypRC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypRC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypRC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypRC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypRC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypRC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sVatRate / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypRC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTRC.Tag 'cmbDTRC.Column(0)
               !dateFrom = CDate(txtDateFrom.text)
               !DateTO = CDate(txtDateTo.text)
               'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
                  !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !SageDepartment = txtFundRC.Tag
               !description = txtDescRC.text

               .Update
            End With

            BRcount = BRcount + 1
            iSerial = iSerial + 1
         End If
'************************************************************************************************
'         Service Charge demands
'************************************************************************************************
         If IsSCSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoconn, sVatRate)
'**** Insert the Header info in the DemandRecords table
            lDemand = NextRef(adoconn, "DEMAND_REF")        'GET THE NEXT DEMAND ID
            With adoRstDemandRec
               .AddNew
               .Fields.Item("DemandID").Value = lDemand
               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields.Item("Source").Value = 1
               .Fields.Item("TransactionType").Value = 1
               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("IsPrinted").Value = False
               .Fields.Item("Spare1").Value = adoDmdTypSC.Fields.Item("spare1").Value
               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoconn)
               .Fields.Item("Details").Value = "Service Charge"
               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Fields.Item("LastModifiedBy").Value = User
               .Fields.Item("LastModifiedDate").Value = Now
               .Update
            End With

            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !splitID = 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypSC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypSC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypSC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypSC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypSC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypSC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sVatRate / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypSC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTSC.Tag
               !dateFrom = txtDateFrom.text
               !DateTO = txtDateTo.text
               'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
                  !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !description = txtDescSC.text
               !SageDepartment = txtFundSC.Tag
               !ChargingFigure = sChargingFig
               .Update
            End With

            SCcount = SCcount + 1
            iSerial = iSerial + 1
         End If

'************************************************************************************************
'   Insurance Charge demands
'************************************************************************************************
         If IsICSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoconn, sVatRate)
   '**** Insert the Header info in the DemandRecords table
            lDemand = NextRef(adoconn, "DEMAND_REF")        'GET THE NEXT DEMAND ID
            With adoRstDemandRec
               .AddNew
               .Fields.Item("DemandID").Value = lDemand
               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields.Item("Source").Value = 1
               .Fields.Item("TransactionType").Value = 1
               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("IsPrinted").Value = False
               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoconn)
               .Fields.Item("Spare1").Value = adoDmdTypIC.Fields.Item("spare1").Value
               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
               .Fields.Item("Details").Value = "Insurance"
               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Update
            End With

            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !splitID = 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypIC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypIC.Fields.Item("NominalNameforAmount").Value
               
               'Resolved by BOSL
               'Issue No: 0000476
               'Retrieving the Sales Ledger Control and Output VAT Control Accounts from the Tools > Configuration instead
               'of the default control accounts set in the demand types records.
               'If not found, then exit the function before it updates or post any transaction to the database.
               'Modified By: Asif. 27 Sep 2014
   
'               !NominalCodeForVAT = adoDmdTypIC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypIC.Fields.Item("NominalNameforVAT").Value
               
               !NominalCodeForVAT = OutputVATNominalCode
               !NominalNameforVAT = OutputVATNominalName
               
'               !NominalCodeForTotal = adoDmdTypIC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypIC.Fields.Item("NominalNameforTotal").Value
   
               !NominalCodeForTotal = SalesLedgerNominalCode
               !NominalNameforTotal = SalesLedgerNominalName
               '''''''''''''''''''''''
               
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sVatRate / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
               'Below line is comment out by anol 20160523 It was a wrong assignment
              ' !TotalAmount = !amount
'               !SageRef = adoDmdTypIC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !Typeofdemand = txtDTIC.Tag
               !description = txtDescIC.text
               !VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoconn)
               !dateFrom = txtDateFrom.text
               !DateTO = txtDateTo.text
               'new requirement from WD they want  Startdate and end date as 2nd ref 2019-11-18
               !SageRef = Format(!dateFrom, "dd/mm/yy") & "-" & Format(!DateTO, "dd/mm/yy")
               !SageDepartment = txtFundIC.Tag
               !ChargingFigure = sChargingFig
               .Update
            End With

            ICcount = ICcount + 1
            iSerial = iSerial + 1
         End If

         adoRstLeaseDtl.MoveNext
      Wend

      If chkRC.Value = 1 Then
         adoDmdTypRC.Close
         Set adoDmdTypRC = Nothing
         Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
      End If
      If chkSC.Value = 1 Then
         adoDmdTypSC.Close
         Set adoDmdTypSC = Nothing
         Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
      End If
      If chkIC.Value = 1 Then
         adoDmdTypIC.Close
         Set adoDmdTypIC = Nothing
         Msg = Msg & ICcount & " Demands for Insurance Charge were generated." & Chr(13)
      End If

      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close

      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If

   MousePointer = vbDefault

   Msg = Msg & "A total of " & BRcount + SCcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Batch Demand Generated"

'  Bring all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoconn

   adoconn.Close
   Set adoconn = Nothing
   Exit Sub

ErrH:
'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
   MsgBox Err.Number & " - (pcm_001)" & Err.description, vbOKOnly, "Error"

   Set adoconn = Nothing
End Sub

Private Function IsRCSelected(ByVal szLeaseID As String, ByRef cAmount As Currency) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String

   szaLeaseID = Split(szLeaseID_RC, ", ")
   szaAmount = Split(szAmt_RC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsRCSelected = True
         cAmount = Val(szaAmount(i))
         Exit Function
      End If
   Next i

   IsRCSelected = False
   cAmount = 0
End Function

Private Function IsSCSelected(ByVal szLeaseID As String, ByRef cAmount As Currency, Optional sChargingFig As Single) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String, j As Integer

   szaLeaseID = Split(szLeaseID_SC, ", ")
   szaAmount = Split(szAmt_SC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsSCSelected = True
         cAmount = Val(szaAmount(i))
         
         For j = 1 To flxDemands.Rows - 1
            If flxDemands.TextMatrix(j, 8) = szLeaseID Then
               sChargingFig = IIf(IsNull(flxDemands.TextMatrix(j, 4)), Null, Val(flxDemands.TextMatrix(j, 4)))
               Exit For
            End If
         Next j
         
         Exit Function
      End If
   Next i

   IsSCSelected = False
   cAmount = 0
End Function

Private Function IsICSelected(ByVal szLeaseID As String, ByRef cAmount As Currency, ByRef sChargingFig As Single) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String, j As Integer

   szaLeaseID = Split(szLeaseID_IC, ", ")
   szaAmount = Split(szAmt_IC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsICSelected = True
         cAmount = Val(szaAmount(i))
         
         For j = 1 To flxDemands.Rows - 1
            If flxDemands.TextMatrix(j, 8) = szLeaseID Then
               sChargingFig = IIf(IsNull(flxDemands.TextMatrix(j, 6)), Null, Val(flxDemands.TextMatrix(j, 6)))
               Exit For
            End If
         Next j
         
         Exit Function
      End If
   Next i

   IsICSelected = False
   cAmount = 0
End Function
'
'Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
'   Dim adoRST As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
'
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRST.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRST.RecordCount - 1
'   TotalCol = adoRST.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
'       Next j
'       adoRST.MoveNext
'       If adoRST.EOF Then Exit For
'   Next i
'
'   cboClient.Column() = Data()
'   cboClient.ListIndex = 0
'   adoRST.Close
''*************************************** PROPERTY ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "ORDER BY PropertyID;"
'
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRST.EOF Then GoTo NoRes
'
'   TotalRow = adoRST.RecordCount
'   TotalCol = adoRST.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow - 1
'      For j = 0 To TotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
'      Next j
'      adoRST.MoveNext
'      If adoRST.EOF Then Exit For
'   Next i
'   cboProperty.Column() = Data()
''   cboProperty.ListIndex = 0
'
'NoRes:
'   adoRST.Close
'   Set adoRST = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRST.Close
'   Set adoRST = Nothing
'End Sub

Private Sub ConfigureFlxDemands()
   Dim szHeader As String, i As Integer

   flxDemands.Clear
   flxDemands.Cols = 9
   flxDemands.Rows = 2

   szHeader$ = "<ID|<Lessee|>RC Per|>RC Amt|>SC Per" & _
               "|>SC Amt|>IC Per|>IC Amt|LeaseID"
   flxDemands.FormatString = szHeader$

   For i = 1 To flxDemands.Cols - 2
      flxDemands.ColWidth(i - 1) = Label3(i).Left - Label3(i - 1).Left
   Next i
   flxDemands.ColWidth(i - 1) = flxDemands.Left + flxDemands.Width - Label3(i - 1).Left - 300
   flxDemands.ColWidth(i) = 0

   flxDemands.RowHeight(0) = 0
End Sub

Private Sub cmdCancel_Click()
   cmdOK.Enabled = True
   cmdCancel.Enabled = False
   fraHeader.Enabled = True
   chkRC.Value = 0
   chkSC.Value = 0
   chkIC.Value = 0
   txtInputGrid.text = ""
   txtInputGrid.Visible = False
   ClearGrid
End Sub

Private Sub ClearGrid()
   Dim iRow As Integer, iCol As Integer

   For iRow = 1 To flxDemands.Rows - 1
      For iCol = 2 To 7
         flxDemands.TextMatrix(iRow, iCol) = ""
      Next iCol
   Next iRow
   CalAllTotal
   szLeaseID_RC = ""
   szLeaseID_SC = ""
   szLeaseID_IC = ""
End Sub

Private Sub cmdClear_Click()
   If MsgBox("Do you want to clear the entries of the grid?", vbQuestion + vbYesNo, "Batch Demands") = vbNo Then Exit Sub

   ClearGrid
End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 255.299
    sTextBox = "1"
    LoadflxClient
    
    fraHeader.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdExclude_Click()
   If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Batch Demands") = vbNo Then Exit Sub

   Dim iRow As Integer, iCurRow As Integer, iCurCol As Integer

   iCurRow = flxDemands.row
   iCurCol = flxDemands.col
   
   flxDemands.col = 1
   iRow = flxDemands.Rows - 1
   Do
      flxDemands.row = iRow
      If flxDemands.CellBackColor = RGB(233, 232, 155) Then
         flxDemands.RemoveItem iRow
         If iRow = iCurRow Then iCurRow = 0
      End If

      iRow = iRow - 1
   Loop While iRow > 0

   flxDemands.row = iCurRow
   flxDemands.col = iCurCol
   CalAllTotal
End Sub

Private Sub cmdLookupIC_Click()
   If txtProperty.text = "" Then Exit Sub

   If chkIC.Value = 0 Then
      MsgBox "Insurance charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
      Label7.ForeColor = vbRed
      Exit Sub
   End If

   If Val(txtBudAmtIC.text) < 0 Or txtBudAmtIC.text = "" Then
      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"

      fraHeader.Enabled = True
      cmdOK.Enabled = True
      cmdCancel.Enabled = False

      txtBudAmtIC.SetFocus
      Exit Sub
   End If

   Dim adoconn As New ADODB.Connection
   Dim adoRstRC As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   adoconn.Open getConnectionString

   szSQLStr = "SELECT L.SageAccountNumber, R.ChargingFigure  " & _
              "FROM LeaseDetails AS L, Units AS U, LInsuranceCharges AS R " & _
              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.ChargingType = 1 AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & txtProperty.Tag & "';"

   adoRstRC.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQLStr
   While Not adoRstRC.EOF
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = adoRstRC.Fields.Item("SageAccountNumber").Value Then
            flxDemands.TextMatrix(iRow, 6) = adoRstRC.Fields.Item("ChargingFigure").Value
            flxDemands.TextMatrix(iRow, 7) = _
                  Format(Val(txtBudAmtIC.text) * (Val(flxDemands.TextMatrix(iRow, 6)) / 100), "0.00")
         End If
      Next iRow
      adoRstRC.MoveNext
   Wend
   adoRstRC.Close
   Set adoRstRC = Nothing

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub cmdLookupRC_Click()
   If txtProperty.text = "" Then Exit Sub

   If chkRC.Value = 0 Then
      MsgBox "Rent charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
      Label5.ForeColor = vbRed
      Exit Sub
   End If

   If Val(txtBudAmtRC.text) < 0 Or txtBudAmtRC.text = "" Then
      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"

      fraHeader.Enabled = True
      cmdOK.Enabled = True
      cmdCancel.Enabled = False

      txtBudAmtRC.SetFocus
      Exit Sub
   End If

   Dim adoconn As New ADODB.Connection
   Dim adoRstRC As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   adoconn.Open getConnectionString

   szSQLStr = "SELECT L.SageAccountNumber, R.spare2  " & _
              "FROM LeaseDetails AS L, Units AS U, LRentCharges AS R " & _
              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.spare1 = '2' AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & txtProperty.Tag & "';"

   adoRstRC.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQLStr
   While Not adoRstRC.EOF
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = adoRstRC.Fields.Item("SageAccountNumber").Value Then
            flxDemands.TextMatrix(iRow, 2) = adoRstRC.Fields.Item("spare2").Value
            flxDemands.TextMatrix(iRow, 3) = _
                  Format(Val(txtBudAmtRC.text) * (Val(flxDemands.TextMatrix(iRow, 2)) / 100), "0.00")
         End If
      Next iRow
      adoRstRC.MoveNext
   Wend
   adoRstRC.Close
   Set adoRstRC = Nothing

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub cmdLookupSC_Click()
   If txtProperty.text = "" Then Exit Sub

   If chkSC.Value = 0 Then
      MsgBox "Service charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
      Label6.ForeColor = vbRed
      Exit Sub
   End If

   If Val(txtBudAmtSC.text) < 0 Or txtBudAmtSC.text = "" Then
      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"

      fraHeader.Enabled = True
      cmdOK.Enabled = True
      cmdCancel.Enabled = False

      txtBudAmtSC.SetFocus
      Exit Sub
   End If

   Dim adoconn As New ADODB.Connection
   Dim adoRstSC As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   adoconn.Open getConnectionString

   szSQLStr = "SELECT L.SageAccountNumber, R.CMFigure  " & _
              "FROM LeaseDetails AS L, Units AS U, LServiceCharges AS R " & _
              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.ChargingMethod = 2 AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "R.ServiceChargeDept = '" & CStr(txtFundSC.Tag) & "' AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & txtProperty.Tag & "';"

   adoRstSC.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQLStr
   While Not adoRstSC.EOF
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = adoRstSC.Fields.Item("SageAccountNumber").Value Then
            flxDemands.TextMatrix(iRow, 4) = adoRstSC.Fields.Item("CMFigure").Value
            flxDemands.TextMatrix(iRow, 5) = _
                  Format(Val(txtBudAmtSC.text) * (Val(flxDemands.TextMatrix(iRow, 4)) / 100), "0.00")
         End If
      Next iRow
      adoRstSC.MoveNext
   Wend
   adoRstSC.Close
   Set adoRstSC = Nothing

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub cmdOK_Click()
   If Not CheckInput Then Exit Sub

   Dim iRow As Integer, dProrata As Double, dPercentage As Double

   If chkRC.Value = 1 Then
      If chkProRC.Value = 1 Then
         dProrata = Val(txtBudAmtRC.text) / (flxDemands.Rows - 1)
         dPercentage = 100 / (flxDemands.Rows - 1)

         For iRow = 1 To flxDemands.Rows - 1
            flxDemands.TextMatrix(iRow, 2) = Format(dPercentage, "0.0000")
            flxDemands.TextMatrix(iRow, 3) = Format(dProrata, "0.00")
         Next iRow
      End If
   End If

   If frmMMain.IsRibbonVersion And lblPostingDate.ToolTipText <> "" Then
      Dim adoconn As New ADODB.Connection
      Dim szSQL As String

      adoconn.Open getConnectionString

'      If IsPeriodStatus(lblPostingDate.ToolTipText, txtClientList.Tag, adoConn) <> 1 Then
'         ShowMsgInTaskBar "The transaction date falls within a closed period", "Y", "N"
'         txtDateIssue.SetFocus
'         adoConn.Close
'         Set adoConn = Nothing
'         Exit Sub
'      End If
      adoconn.Close
      Set adoconn = Nothing
   End If

   If chkSC.Value = 1 Then
      If chkProSC.Value = 1 Then
'         dProrata = RoundingNumber(Val(txtBudAmtSC.text) / (flxDemands.Rows - 1), 2)
         dPercentage = RoundingNumber(100 / (flxDemands.Rows - 1), 4)
         dProrata = RoundingNumber(Val(txtBudAmtSC.text) * (dPercentage / 100), 2)
'Debug.Print RoundingNumber(Val(txtBudAmtSC.text) * (dPercentage / 100), 2)
         For iRow = 1 To flxDemands.Rows - 1
            flxDemands.TextMatrix(iRow, 4) = Format(dPercentage, "0.0000")
            flxDemands.TextMatrix(iRow, 5) = Format(dProrata, "0.00")
         Next iRow
      End If
   End If

   If chkIC.Value = 1 Then
      If chkProIC.Value = 1 Then
         dProrata = Val(txtBudAmtIC.text) / (flxDemands.Rows - 1)
         dPercentage = 100 / (flxDemands.Rows - 1)

         For iRow = 1 To flxDemands.Rows - 1
            flxDemands.TextMatrix(iRow, 6) = Format(dPercentage, "0.0000")
            flxDemands.TextMatrix(iRow, 7) = Format(dProrata, "0.00")
         Next iRow
      End If
   End If

   fraHeader.Enabled = False
   cmdOK.Enabled = False
   cmdCancel.Enabled = True

   CalAllTotal
End Sub

Private Sub flxDemands_DblClick()
   If cmdOK.Enabled Then
      cmdOK.SetFocus
      Exit Sub
   End If

   If flxDemands.col = 0 Or flxDemands.col = 1 Then
      UMarkRowFlxGrid flxDemands, flxDemands.row
   End If

   Dim i As Integer, iFlxSPayCol As Integer

   If flxDemands.col < 2 Then Exit Sub
   If flxDemands.TextMatrix(flxDemands.row, 0) = "" Then Exit Sub

   If (flxDemands.col = 2 Or flxDemands.col = 3) And chkRC.Value = 0 Then
      MsgBox "Rent Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
      Exit Sub
   End If
   If (flxDemands.col = 2 And (txtBudAmtRC.text = "" Or Val(txtBudAmtRC.text) < 0)) And chkRC.Value = 0 Then
      MsgBox "Please input the budget amount for rent charge.", vbInformation + vbOKOnly, "Batch Payment"
      txtBudAmtRC.SetFocus
      Exit Sub
   End If
   
   If (flxDemands.col = 4 Or flxDemands.col = 5) And chkSC.Value = 0 Then
      MsgBox "Service Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
      Exit Sub
   End If
   If (flxDemands.col = 4 And (txtBudAmtSC.text = "" Or Val(txtBudAmtSC.text) < 0)) And chkSC.Value = 0 Then
      MsgBox "Please input the budget amount for service charge.", vbInformation + vbOKOnly, "Batch Payment"
      txtBudAmtSC.SetFocus
      Exit Sub
   End If
   
   If (flxDemands.col = 6 Or flxDemands.col = 7) And chkIC.Value = 0 Then
      MsgBox "Insurance Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
      Exit Sub
   End If
   If (flxDemands.col = 6 And (txtBudAmtIC.text = "" Or Val(txtBudAmtIC.text) < 0)) And chkIC.Value = 0 Then
      MsgBox "Please input the budget amount for insurance charge.", vbInformation + vbOKOnly, "Batch Payment"
      txtBudAmtIC.SetFocus
      Exit Sub
   End If

   If flxDemands.col = 2 Or flxDemands.col = 3 Then txtInputGrid.BackColor = Label3(77).BackColor
   If flxDemands.col = 4 Or flxDemands.col = 5 Then txtInputGrid.BackColor = Label4.BackColor
   If flxDemands.col = 6 Or flxDemands.col = 7 Then txtInputGrid.BackColor = Label3(8).BackColor

   txtInputGrid.Top = flxDemands.CellTop + flxDemands.Top
   txtInputGrid.Left = flxDemands.CellLeft + flxDemands.Left
   txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
   txtInputGrid.Height = flxDemands.RowHeight(flxDemands.row) - 15
   txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
   txtInputGrid.Visible = True
   txtInputGrid.SetFocus
   SelTxtInCtrl txtInputGrid

   iCurRow = flxDemands.row
   iCurCol = flxDemands.col
End Sub

Private Sub flxDemands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'   frmDemands3.Hide

   Me.Height = 8865
   Me.Width = 11865
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fraHeader.BackColor = MODULEBACKCOLOR
   chkRC.BackColor = MODULEBACKCOLOR
   chkSC.BackColor = MODULEBACKCOLOR
   chkIC.BackColor = MODULEBACKCOLOR
   chkProRC.BackColor = MODULEBACKCOLOR
   chkProSC.BackColor = MODULEBACKCOLOR
   chkProIC.BackColor = MODULEBACKCOLOR

   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

'   PrepareList adoConn, cboClientList, cboPropertyList
   LoadClients adoconn
   ConfigureFlxDemands

   LoadDept adoconn

   adoconn.Close
   Set adoconn = Nothing

   txtDateFrom.text = Format(Now, "dd/mm/yyyy")
   txtDateIssue.text = Format(Now, "dd/mm/yyyy")
   txtDateTo.text = Format(Now, "dd/mm/yyyy")

   txtPcgRC.Left = cmdLookupRC.Left
   txtPcgRC.Width = flxDemands.ColWidth(2)
   txtAmtRC.Left = txtPcgRC.Left + txtPcgRC.Width + 20
   txtAmtRC.Width = flxDemands.ColWidth(3)

   txtPcgSC.Left = cmdLookupSC.Left
   txtPcgSC.Width = flxDemands.ColWidth(4)
   txtAmtSC.Left = txtPcgSC.Left + txtPcgSC.Width + 20
   txtAmtSC.Width = flxDemands.ColWidth(5)

   txtPcgIC.Left = cmdLookupIC.Left
   txtPcgIC.Width = flxDemands.ColWidth(6)
   txtAmtIC.Left = txtPcgIC.Left + txtPcgIC.Width + 20
   txtAmtIC.Width = flxDemands.ColWidth(7)

   BATCH_DEMAND_PROCESS = True

   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadClients(adoconn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT   CLIENTID, CLIENTNAME " & _
           "FROM     CLIENT " & _
           "ORDER BY CLIENTID;"

    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
    End If
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientList.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub CalAllTotal()
   Dim iRow As Integer, cTotal As Currency

   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 2))
   Next iRow
   txtPcgRC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 3))
   Next iRow
   txtAmtRC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 4))
   Next iRow
   txtPcgSC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 5))
   Next iRow
   txtAmtSC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 6))
   Next iRow
   txtPcgIC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 7))
   Next iRow
   txtAmtIC.text = Format(cTotal, "0.00")
End Sub

Private Sub LoadDept(adoconn As ADODB.Connection)
'   Dim rRow As Integer, iRec As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT FundID, FundCode, FundName, CategoryCode FROM Fund;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
'   Else
'      ReDim Data(2, adoRst.RecordCount) As String
'
'      rRow = 0
'      adoRst.MoveFirst
'      While Not adoRst.EOF
'         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
'         Data(1, rRow) = adoRst.Fields.Item("FundCode").Value
'         Data(2, rRow) = adoRst.Fields.Item("FundName").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'      cboFundRC.Clear
'      cboFundRC.Column() = Data()
'      cboFundSC.Clear
'      cboFundSC.Column() = Data()
'      cboFundIC.Clear
'      cboFundIC.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ' Destroy Objects
'   Set adoRst = Nothing
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)

   If BATCH_DEMAND_PROCESS Then
      frmDemands3.cboPropertyList_Click
      frmDemands3.Show
      BATCH_DEMAND_PROCESS = False
   End If
End Sub

Private Sub fraHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    'Resolved by BOSL
    'issue 468
    'Modified by anol 01 Sep 2014
    If txtClientList.Tag = "" Then
        cmdClientList.SetFocus
        ShowMsgInTaskBar "Please select a client.", "Y"
        Exit Sub
    End If
'     If cboClientList.ListIndex = -1 Then Exit Sub
'     If frmMMain.IsRibbonVersion Then
'        Dim adoConn As New ADODB.Connection
'        Dim szSQL As String
'        adoConn.Open getConnectionString
'        If IsPeriodStatus(txtDateIssue.text, txtClientList.Tag, adoConn) = 0 Then
'            ShowMsgInTaskBar "The issue date cannot fall within a closed financial period", "Y", "N"
'            adoConn.Close
'            Exit Sub
'        ElseIf IsPeriodStatus(txtDateIssue.text, txtClientList.Tag, adoConn) = 9 Then
'            ShowMsgInTaskBar "The issue date does not fall in any existing financial period", "Y", "N"
'            adoConn.Close
'            Exit Sub
'        End If
'     End If
    'End of modification
    DispayCalendar Me, lblPostingDate.ToolTipText, txtDateIssue.text, txtClientList.Tag
End Sub

Private Sub txtBudAmtIC_GotFocus()
   If chkIC.Value = 0 Then
      chkIC.SetFocus
   Else
      SelTxtInCtrl txtBudAmtIC
   End If
End Sub

Private Sub txtBudAmtIC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescIC.SetFocus
    End If
   DigitTextKeyPress txtBudAmtIC, KeyAscii
End Sub

Private Sub txtBudAmtIC_LostFocus()
   txtBudAmtIC.text = Format(txtBudAmtIC.text, "0.00")
End Sub

Private Sub txtBudAmtRC_GotFocus()
   If chkRC.Value = 0 Then
      chkRC.SetFocus
   Else
      SelTxtInCtrl txtBudAmtRC
   End If
End Sub

Private Sub txtBudAmtRC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescRC.SetFocus
    End If
   If KeyAscii <> 45 Then
      DigitTextKeyPress txtBudAmtRC, KeyAscii
   Else
      txtBudAmtRC.text = Val(txtBudAmtRC.text) * (-1)
      If Val(txtBudAmtRC.text) < 0 Then KeyAscii = 0
   End If
End Sub

Private Sub txtBudAmtRC_LostFocus()
   If Val(txtBudAmtRC.text) > 0 And InStr(Me.Caption, "Credit") > 0 Then
      MsgBox "You cannot create an invoice and a credit note in the same batch run.", vbCritical + vbOKOnly, "Prestige Property Management"
      txtBudAmtRC.SetFocus
   End If

   txtBudAmtRC.text = Format(txtBudAmtRC.text, "0.00")
   If Val(txtBudAmtRC.text) < 0 Then
      If MsgBox("Do you wish to create a batch credit note?", vbQuestion + vbYesNo, "Batch Credit Note") = vbYes Then
         Me.Caption = "Batch Credit Demands"
      Else
         txtBudAmtRC.SetFocus
      End If
   End If
End Sub

Private Sub txtBudAmtSC_GotFocus()
   If chkSC.Value = 0 Then
      chkSC.SetFocus
   Else
      SelTxtInCtrl txtBudAmtSC
   End If
End Sub

Private Sub txtBudAmtSC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescSC.SetFocus
    End If
   DigitTextKeyPress txtBudAmtSC, KeyAscii
End Sub

Private Sub txtBudAmtSC_LostFocus()
   txtBudAmtSC.text = Format(txtBudAmtSC.text, "0.00")
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateTo.SetFocus
    End If
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   TextBoxFormatDate txtDateFrom
End Sub

Private Sub txtDateIssue_Change()
    TextBoxChangeDate txtDateIssue
    lblPostingDate.ToolTipText = txtDateIssue.text
End Sub

Private Sub txtDateIssue_GotFocus()
   SelTxtInCtrl txtDateIssue
End Sub

Private Sub txtDateIssue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateFrom.SetFocus
    End If
    TextBoxKeyPrsDate txtDateIssue, KeyAscii
End Sub

Private Sub txtDateIssue_LostFocus()
   If IsDate(txtDateIssue.text) Then
      If TextBoxFormatDate(txtDateIssue) Then
         lblPostingDate.ToolTipText = txtDateIssue.text
      End If
     'Modified by BOSL
     'issue 468
     'Modified by anol 02 Sep 2014
     If txtClientList.text = "" Then
            ShowMsgInTaskBar "Please select a client", "Y", "N"
            Exit Sub
     End If
     
     If frmMMain.IsRibbonVersion Then
        Dim adoconn As New ADODB.Connection
        Dim szSQL As String
        If IsDate(txtDateIssue.text) = False Then Exit Sub
        adoconn.Open getConnectionString
        If IsPeriodStatus(txtDateIssue.text, txtClientList.Tag, adoconn) = 0 Then
            ShowMsgInTaskBar "The issue date cannot fall within a closed financial period", "Y", "N"
            adoconn.Close
            Exit Sub
        ElseIf IsPeriodStatus(txtDateIssue.text, txtClientList.Tag, adoconn) = 9 Then
            ShowMsgInTaskBar "The issue date does not fall in any existing financial period", "Y", "N"
            adoconn.Close
            Exit Sub
        End If
     End If
    'End of modification
End If
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDTRC.SetFocus
    End If
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   TextBoxFormatDate txtDateTo
End Sub

Private Sub txtDescIC_GotFocus()
   If chkIC.Value = 0 Then chkIC.SetFocus
End Sub

Private Sub txtDescRC_GotFocus()
   If chkRC.Value = 0 Then chkRC.SetFocus
End Sub

Private Sub txtDescSC_GotFocus()
   If chkSC.Value = 0 Then chkSC.SetFocus
End Sub

Private Sub cboFundIC_GotFocus()
   'If chkIC.Value = 0 Then chkIC.Value = 1
End Sub

Private Sub cboFundRC_GotFocus()
'   If chkRC.Value = 0 Then chkRC.Value = 1
End Sub

Private Sub cboFundSC_GotFocus()
'   If chkSC.Value = 0 Then chkSC.Value = 1
End Sub

Private Sub cmbDTIC_GotFocus()
   'If chkIC.Value = 0 Then chkIC.Value = 1
End Sub

Private Sub cmbDTRC_GotFocus()
  ' If chkRC.Value = 0 Then chkRC.Value = 1
End Sub

Private Sub cmbDTSC_GotFocus()
  ' If chkSC.Value = 0 Then chkSC.Value = 1
End Sub

Private Function CheckInput() As Boolean
   CheckInput = False

   If txtProperty.text = "" Then
      MsgBox "Please select the property.", vbInformation + vbOKOnly, "Batch Demands"
      cmdProperty.SetFocus
      Exit Function
   End If
   If txtDateIssue.text = "" Then
      MsgBox "Please issue the from date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateIssue.SetFocus
      Exit Function
   End If
   If txtDateFrom.text = "" Then
      MsgBox "Please input the from date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateFrom.SetFocus
      Exit Function
   End If
   If txtDateTo.text = "" Then
      MsgBox "Please input the to date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateTo.SetFocus
      Exit Function
   End If

   If (chkRC.Value = 0) And (chkSC.Value = 0) And (chkIC.Value = 0) Then
      MsgBox "Please select atleast one charge.", vbInformation + vbOKOnly, "Batch Demands"
      chkRC.SetFocus
      Exit Function
   End If
   If chkRC.Value And txtDTRC.text = "" Then
      MsgBox "Please select the demand type for rent charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdDTRC.SetFocus
      Exit Function
   End If
   If chkSC.Value And txtDTSC.text = "" Then
      MsgBox "Please select the demand type for service charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdDTSC.SetFocus
      Exit Function
   End If
   If chkIC.Value And txtDTIC.text = "" Then
      MsgBox "Please select the demand type for insurance charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdDTIC.SetFocus
      Exit Function
   End If
   If chkRC.Value And txtFundRC.text = "" Then
      MsgBox "Please select the fund for rent charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdFundRC.SetFocus
      Exit Function
   End If
   If chkSC.Value And txtFundSC.text = "" Then
      MsgBox "Please select the fund for service charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdFundSC.SetFocus
      Exit Function
   End If
   If chkIC.Value And txtFundIC.text = "" Then
      MsgBox "Please select the fund for insurance charge.", vbInformation + vbOKOnly, "Batch Demands"
      cmdFundIC.SetFocus
      Exit Function
   End If

   CheckInput = True
End Function

Private Sub txtInputGrid_LostFocus()
   flxDemands.TextMatrix(iCurRow, iCurCol) = txtInputGrid.text

   CalAllTotal
End Sub

Private Sub txtInputGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      flxDemands.TextMatrix(flxDemands.row, flxDemands.col) = txtInputGrid.text

      InputGridLostFocus

      If MoveDownPosition Then
         txtInputGrid.Left = iLeft
         txtInputGrid.Top = iTop
         txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
         txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
         txtInputGrid.Visible = True
         SelTxtInCtrl txtInputGrid
      End If
   End If

   If KeyCode > 36 And KeyCode < 41 Then
      flxDemands.TextMatrix(flxDemands.row, flxDemands.col) = txtInputGrid.text
      InputGridLostFocus

      If KeyCode = 38 Then
         If MoveUpPosition Then                     'Up Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 40 Then
         If MoveDownPosition Then                   'Down key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 39 Then
         If MoveRightPosition Then                  'Right Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 37 Then
         If MoveLeftPosition Then                  'Left Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      SelTxtInCtrl txtInputGrid
   End If

   CalAllTotal
End Sub

Private Function MoveLeftPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.col Mod 2 = 1 Then
      flxDemands.col = flxDemands.col - 1
      iCurCol = flxDemands.col
   Else
      txtInputGrid.Visible = False
      MoveLeftPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveLeftPosition = True
End Function

Private Function MoveRightPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.col Mod 2 = 0 Then
      flxDemands.col = flxDemands.col + 1
      iCurCol = flxDemands.col
   Else
      txtInputGrid.Visible = False
      MoveRightPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveRightPosition = True
End Function

Private Function MoveUpPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.row > 1 Then
      flxDemands.row = flxDemands.row - 1
      iCurRow = flxDemands.row
   Else
      txtInputGrid.Visible = False
      MoveUpPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveUpPosition = True
End Function

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.row < flxDemands.Rows - 1 Then
      flxDemands.row = flxDemands.row + 1
      iCurRow = flxDemands.row
   Else
      txtInputGrid.Visible = False
      MoveDownPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveDownPosition = True
End Function

Private Sub InputGridLostFocus()
   If chkRC.Value = 1 And flxDemands.col = 2 And (txtBudAmtRC.text <> "" And Val(txtBudAmtRC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 3) = _
                 Format(Val(txtBudAmtRC.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 2)) / 100), "0.00")
   End If
   If chkRC.Value = 1 And flxDemands.col = 3 And (txtBudAmtRC.text <> "" And Val(txtBudAmtRC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 2) = _
                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 3)) / Val(txtBudAmtRC.text) * 100, "0.0000")
   End If

   If chkSC.Value = 1 And flxDemands.col = 4 And (txtBudAmtSC.text <> "" And Val(txtBudAmtSC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 5) = _
                 Format(Val(txtBudAmtSC.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 4)) / 100), "0.00")
   End If
   If chkSC.Value = 1 And flxDemands.col = 5 And (txtBudAmtSC.text <> "" And Val(txtBudAmtSC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 4) = _
                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 5)) / Val(txtBudAmtSC.text) * 100, "0.0000")
   End If

   If chkIC.Value = 1 And flxDemands.col = 6 And (txtBudAmtIC.text <> "" And Val(txtBudAmtIC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 7) = _
                 Format(Val(txtBudAmtIC.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 6)) / 100), "0.00")
   End If
   If chkIC.Value = 1 And flxDemands.col = 7 And (txtBudAmtIC.text <> "" And Val(txtBudAmtIC.text) > 0) Then
      flxDemands.TextMatrix(flxDemands.row, 6) = _
                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 7)) / Val(txtBudAmtIC.text) * 100, "0.00")
   End If
End Sub

Public Sub TestingCommand()
'   cboPropertyList.ListIndex = 2
'   chkRC.Value = 1
'   cmbDTRC.ListIndex = 0
'   cboFundRC.ListIndex = 0
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
