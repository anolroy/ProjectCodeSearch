VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManagementFees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Management Fees"
   ClientHeight    =   11685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   20130
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManagementFees.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11685
   ScaleWidth      =   20130
   Begin VB.PictureBox picAccounts 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   945
      ScaleHeight     =   3330
      ScaleWidth      =   6390
      TabIndex        =   122
      Top             =   11115
      Visible         =   0   'False
      Width           =   6420
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
         Left            =   6090
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2445
         Index           =   1
         Left            =   45
         TabIndex        =   124
         Top             =   855
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4313
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   13
         Left            =   5400
         Top             =   0
         Visible         =   0   'False
         Width           =   5580
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   2
         Left            =   1410
         TabIndex        =   136
         Top             =   540
         Width           =   2595
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4577;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   1
         Left            =   225
         TabIndex        =   135
         Top             =   540
         Width           =   1155
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2037;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   265
         Index           =   0
         Left            =   30
         TabIndex        =   134
         Top             =   540
         Visible         =   0   'False
         Width           =   900
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1587;467"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   3
         Left            =   1395
         TabIndex        =   133
         Top             =   300
         Width           =   1335
         VariousPropertyBits=   8388627
         Caption         =   "Account Name"
         Size            =   "2355;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   132
         Top             =   300
         Width           =   1095
         VariousPropertyBits=   8388627
         Caption         =   "Account ID"
         Size            =   "1931;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   131
         Top             =   300
         Visible         =   0   'False
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "A/C type"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   1
         Left            =   1515
         TabIndex        =   130
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   1
         Left            =   2115
         TabIndex        =   129
         Top             =   1200
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   4
         Left            =   4155
         TabIndex        =   128
         Top             =   300
         Width           =   870
         VariousPropertyBits=   8388627
         Caption         =   "A/C Balance"
         Size            =   "1535;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   6
         Left            =   4050
         TabIndex        =   127
         Top             =   540
         Width           =   1020
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1799;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   265
         Index           =   7
         Left            =   5085
         TabIndex        =   126
         Top             =   540
         Width           =   1065
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1879;467"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   9
         Left            =   5085
         TabIndex        =   125
         Top             =   315
         Width           =   1545
         VariousPropertyBits=   8388627
         Caption         =   "This Client"
         Size            =   "2725;344"
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
         Height          =   285
         Index           =   16
         Left            =   0
         Top             =   255
         Width           =   6345
      End
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   13365
      ScaleHeight     =   4290
      ScaleWidth      =   5310
      TabIndex        =   111
      Top             =   10575
      Visible         =   0   'False
      Width           =   5340
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
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   45
         Width           =   255
      End
      Begin VB.CheckBox chkShowBal 
         BackColor       =   &H80000009&
         Caption         =   "Show Bal"
         Height          =   195
         Left            =   4050
         TabIndex        =   112
         Top             =   405
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3570
         Index           =   0
         Left            =   45
         TabIndex        =   114
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6297
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
         Left            =   1350
         TabIndex        =   121
         Top             =   375
         Width           =   1215
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2143;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   120
         Top             =   375
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
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Left            =   3855
         TabIndex        =   119
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch1 
         Height          =   195
         Left            =   1875
         TabIndex        =   118
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   117
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   116
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   115
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   75
         Width           =   5355
      End
   End
   Begin VB.PictureBox picAccList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3220
      Left            =   7245
      ScaleHeight     =   3195
      ScaleWidth      =   5895
      TabIndex        =   99
      Top             =   10575
      Visible         =   0   'False
      Width           =   5925
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
         Index           =   2
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2535
         Index           =   2
         Left            =   15
         TabIndex        =   101
         Top             =   645
         Width           =   5860
         _ExtentX        =   10345
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   110
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "NotLoaded"
         Size            =   "1296;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   3
         Left            =   2115
         TabIndex        =   109
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   3
         Left            =   1515
         TabIndex        =   108
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   6
         Left            =   30
         TabIndex        =   107
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "A/C type"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   106
         Top             =   120
         Width           =   1095
         VariousPropertyBits=   8388627
         Caption         =   "Account ID"
         Size            =   "1931;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   105
         Top             =   120
         Width           =   1335
         VariousPropertyBits=   8388627
         Caption         =   "Account Name"
         Size            =   "2355;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   104
         Top             =   375
         Width           =   900
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1587;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   4
         Left            =   930
         TabIndex        =   103
         Top             =   375
         Width           =   1470
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2602;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   102
         Top             =   375
         Width           =   2415
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4260;450"
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
         Height          =   255
         Index           =   14
         Left            =   0
         Top             =   90
         Width           =   5580
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3390
      Left            =   45
      TabIndex        =   56
      Top             =   5445
      Width           =   1905
      Begin VB.CommandButton cmdPrintBreakdown 
         Caption         =   "Print Fee"
         Height          =   405
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   675
         Width           =   1620
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fix MgtF Marking"
         Height          =   405
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   2385
         Width           =   1485
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Fix Decimal"
         Height          =   405
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1845
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print List"
         Height          =   405
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1260
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E5E5E5&
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   0
         Left            =   225
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   1755
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   45
      Top             =   10620
      Visible         =   0   'False
      Width           =   5295
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
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   49
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5900
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   53
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   52
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   51
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1875
         TabIndex        =   46
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   47
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   48
         Top             =   375
         Width           =   3420
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6032;450"
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
         Height          =   255
         Index           =   15
         Left            =   0
         Top             =   80
         Width           =   5355
      End
   End
   Begin TabDlg.SSTab tabManagementFee 
      Height          =   11490
      Left            =   1980
      TabIndex        =   7
      Top             =   45
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   20267
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Management Fee"
      TabPicture(0)   =   "frmManagementFees.frx":1202
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Management Fee History"
      TabPicture(1)   =   "frmManagementFees.frx":121E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab2"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTab2 
         Height          =   10395
         Left            =   -74910
         TabIndex        =   58
         Top             =   270
         Width           =   16665
         Begin VB.TextBox txtDisplayMaxPurchaseHist 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13815
            MaxLength       =   80
            TabIndex        =   98
            Top             =   9900
            Width           =   1065
         End
         Begin VB.CommandButton cmdPrintListHistory 
            Caption         =   "Print List"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   9720
            Width           =   1440
         End
         Begin VB.CommandButton cmdRevHistory 
            Caption         =   "Reverse History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   9720
            Width           =   1455
         End
         Begin VB.CommandButton cmdOClientList 
            Caption         =   ". ."
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
            Left            =   2700
            TabIndex        =   64
            Top             =   285
            Width           =   315
         End
         Begin VB.CommandButton cmdOpenSupp 
            Caption         =   ". ."
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
            Left            =   10215
            TabIndex        =   63
            Top             =   315
            Width           =   315
         End
         Begin VB.CommandButton cmdSearchPurchaseHistory 
            Caption         =   "Sea&rch"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3375
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   9720
            Width           =   1170
         End
         Begin VB.CheckBox chkAllPurchaseHistory 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   135
            TabIndex        =   61
            Top             =   945
            Width           =   215
         End
         Begin VB.CheckBox chkPropertyHist 
            Caption         =   "Excl."
            Height          =   195
            Left            =   6390
            TabIndex        =   60
            Top             =   315
            Width           =   1185
         End
         Begin VB.TextBox txtPropertyIDHist 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   59
            Text            =   "ALL"
            Top             =   315
            Width           =   1545
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistory 
            Height          =   5400
            Left            =   120
            TabIndex        =   67
            Top             =   1185
            Width           =   16365
            _ExtentX        =   28866
            _ExtentY        =   9525
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   12
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistorySplit 
            Height          =   2565
            Left            =   45
            TabIndex        =   68
            Top             =   7110
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   4524
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   12
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   8985
            TabIndex        =   97
            Top             =   930
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSForms.ComboBox cmbPropertyHistory 
            Height          =   285
            Left            =   12285
            TabIndex        =   96
            Top             =   315
            Visible         =   0   'False
            Width           =   3525
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6218;503"
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
            Object.Width           =   "1058"
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   1
            Left            =   120
            Top             =   120
            Width           =   16395
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   675
            TabIndex        =   93
            Top             =   6780
            Width           =   690
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   1635
            TabIndex        =   92
            Top             =   6780
            Width           =   1125
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "N/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   3495
            TabIndex        =   91
            Top             =   6780
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   195
            TabIndex        =   90
            Top             =   6780
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   14985
            TabIndex        =   89
            Top             =   6780
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Job No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   6180
            TabIndex        =   88
            Top             =   6780
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   4635
            TabIndex        =   87
            Top             =   6780
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   7260
            TabIndex        =   86
            Top             =   6780
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   13800
            TabIndex        =   85
            Top             =   6780
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Net £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   12555
            TabIndex        =   84
            Top             =   6780
            Width           =   390
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   7875
            TabIndex        =   83
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   10035
            TabIndex        =   82
            Top             =   945
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   5490
            TabIndex        =   81
            Top             =   960
            Width           =   1035
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   2
            Left            =   120
            Top             =   120
            Width           =   16350
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   630
            TabIndex        =   80
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   8370
            TabIndex        =   79
            Top             =   945
            Width           =   255
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   14850
            TabIndex        =   78
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   77
            Top             =   975
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier A/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3540
            TabIndex        =   76
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2580
            TabIndex        =   75
            Top             =   975
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   74
            Top             =   975
            Width           =   345
         End
         Begin MSForms.TextBox txtClientIdlist 
            Height          =   255
            Left            =   1170
            TabIndex        =   73
            Top             =   285
            Width           =   1530
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "2699;450"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSupplierSearc 
            Height          =   255
            Left            =   8730
            TabIndex        =   72
            Top             =   315
            Width           =   1485
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "2619;450"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblDisplay 
            Caption         =   "Display : "
            Height          =   195
            Left            =   13005
            TabIndex        =   71
            Top             =   9900
            Width           =   690
         End
         Begin MSForms.CommandButton cmdOpPropertyHist 
            Height          =   285
            Left            =   5985
            TabIndex        =   70
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3555
            TabIndex        =   69
            Top             =   315
            Width           =   615
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   330
            Index           =   9
            Left            =   120
            TabIndex        =   95
            Top             =   6705
            Width           =   16335
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   240
            Index           =   0
            Left            =   315
            TabIndex        =   94
            Top             =   900
            Width           =   16140
         End
      End
      Begin VB.Frame fraTab0 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   11040
         Left            =   45
         TabIndex        =   8
         Top             =   360
         Width           =   17910
         Begin VB.TextBox txtRctTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   14625
            MaxLength       =   80
            TabIndex        =   138
            Top             =   6525
            Width           =   1200
         End
         Begin VB.PictureBox fmeLoading 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   5400
            ScaleHeight     =   450
            ScaleWidth      =   3195
            TabIndex        =   54
            Top             =   3105
            Visible         =   0   'False
            Width           =   3195
            Begin VB.Label lblLoading 
               BackStyle       =   0  'Transparent
               Caption         =   "Please wait while loading......"
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
               Height          =   330
               Left            =   405
               TabIndex        =   55
               Top             =   135
               Width           =   4590
            End
         End
         Begin VB.CheckBox chkSelectAllDemands 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   0
            TabIndex        =   13
            Top             =   825
            Width           =   215
         End
         Begin VB.TextBox txtIDClient 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   12
            Top             =   315
            Width           =   2175
         End
         Begin VB.TextBox txtPropID 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6075
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   11
            Top             =   315
            Width           =   2130
         End
         Begin VB.TextBox txtSupplier 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10935
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   10
            Top             =   315
            Width           =   1635
         End
         Begin VB.CheckBox chkProperty 
            Caption         =   "Excl."
            Height          =   195
            Left            =   8640
            TabIndex        =   9
            Top             =   360
            Width           =   780
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchaseSplit 
            Height          =   1185
            Left            =   45
            TabIndex        =   14
            Top             =   7155
            Width           =   17835
            _ExtentX        =   31459
            _ExtentY        =   2090
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   10
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchase 
            Height          =   5415
            Left            =   0
            TabIndex        =   15
            Top             =   1080
            Width           =   16530
            _ExtentX        =   29157
            _ExtentY        =   9551
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   12
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxReceipt 
            Height          =   2175
            Left            =   45
            TabIndex        =   143
            Top             =   8685
            Width           =   17835
            _ExtentX        =   31459
            _ExtentY        =   3836
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   10
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Invoice Details "
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   135
            TabIndex        =   162
            Top             =   6525
            Width           =   1770
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "ReceiptDate"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   14220
            TabIndex        =   161
            Top             =   10800
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "AgrPercentage"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   11970
            TabIndex        =   160
            Top             =   10800
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "MgtFeeAmt"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   13185
            TabIndex        =   159
            Top             =   10800
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "ReceiptAmount"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   10620
            TabIndex        =   158
            Top             =   10800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "ReceiptDate"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   9450
            TabIndex        =   157
            Top             =   10800
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "VATPercentage"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   15255
            TabIndex        =   156
            Top             =   10800
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "MgtFeeAmtTotal"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   16470
            TabIndex        =   155
            Top             =   10800
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "FundID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   7650
            TabIndex        =   154
            Top             =   10800
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "ChargeDate"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   8325
            TabIndex        =   153
            Top             =   10800
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "ChargingMethod"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   3195
            TabIndex        =   152
            Top             =   10800
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   2160
            TabIndex        =   151
            Top             =   10800
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Details"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   150
            Top             =   8415
            Width           =   1080
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Split ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1125
            TabIndex        =   148
            Top             =   10800
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "PropertyID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   6660
            TabIndex        =   147
            Top             =   10800
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   5445
            TabIndex        =   146
            Top             =   10800
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   145
            Top             =   10800
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   144
            Top             =   10800
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Index           =   12
            Left            =   14085
            TabIndex        =   139
            Top             =   6570
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   15435
            TabIndex        =   137
            Top             =   855
            Width           =   1005
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   6
            Left            =   0
            Top             =   45
            Width           =   16515
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   14175
            TabIndex        =   42
            Top             =   855
            Width           =   675
         End
         Begin MSForms.CommandButton cmdAccSel 
            Height          =   285
            Left            =   12585
            TabIndex        =   41
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Recoverable"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   10245
            TabIndex        =   40
            Top             =   6855
            Width           =   885
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Net £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   7530
            TabIndex        =   39
            Top             =   6855
            Width           =   390
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   8475
            TabIndex        =   38
            Top             =   6855
            Width           =   360
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   5790
            TabIndex        =   37
            Top             =   6855
            Width           =   840
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   4110
            TabIndex        =   36
            Top             =   6855
            Width           =   360
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Job No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   4860
            TabIndex        =   35
            Top             =   6855
            Width           =   510
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   9210
            TabIndex        =   34
            Top             =   6855
            Width           =   675
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   75
            TabIndex        =   33
            Top             =   6855
            Width           =   210
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "N/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   3450
            TabIndex        =   32
            Top             =   6855
            Width           =   285
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   1770
            TabIndex        =   31
            Top             =   6855
            Width           =   1125
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   465
            TabIndex        =   30
            Top             =   6855
            Width           =   690
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   945
            TabIndex        =   29
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   1740
            TabIndex        =   28
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "A/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   2700
            TabIndex        =   27
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   8370
            TabIndex        =   25
            Top             =   855
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Client ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   5970
            TabIndex        =   24
            Top             =   840
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   4290
            TabIndex        =   23
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   7155
            TabIndex        =   22
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   315
            Width           =   705
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   3
            Left            =   0
            Top             =   90
            Width           =   16560
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   10260
            TabIndex        =   20
            Top             =   315
            Width           =   855
         End
         Begin MSForms.CommandButton cmdOpClient 
            Height          =   285
            Left            =   2835
            TabIndex        =   19
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdOpProperty 
            Height          =   285
            Left            =   8235
            TabIndex        =   18
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   5355
            TabIndex        =   17
            Top             =   315
            Width           =   615
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Controls count"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   12690
            TabIndex        =   16
            Top             =   90
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   240
            Index           =   22
            Left            =   45
            TabIndex        =   43
            Top             =   810
            Width           =   16485
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Index           =   270
            Left            =   45
            TabIndex        =   44
            Top             =   6840
            Width           =   17865
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Index           =   20
            Left            =   180
            TabIndex        =   149
            Top             =   10710
            Visible         =   0   'False
            Width           =   17820
         End
      End
   End
   Begin VB.Frame fraGenerateDemands 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdPostohistory 
         Caption         =   "Post to history"
         Height          =   405
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4005
         Width           =   1620
      End
      Begin VB.CommandButton Command2 
         Caption         =   "View"
         Height          =   405
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3150
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Preview"
         Height          =   405
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   810
         Width           =   1620
      End
      Begin VB.CommandButton cmdPreViewGenDmds 
         Caption         =   "Generate Fee"
         Height          =   405
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1620
         Width           =   1620
      End
      Begin VB.Label lblGenerate 
         Alignment       =   2  'Center
         BackColor       =   &H00E5E5E5&
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Management Fee"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   435
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   210
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmManagementFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTextBox As String
Dim bSearchClientNameFocus As Boolean
Private Sub chkAllPurchaseHistory_Click()
     Dim i As Integer
    If chkAllPurchaseHistory.Value = 1 Then
        For i = 1 To flxPurchHistory.Rows - 1
            flxPurchHistory.TextMatrix(i, 1) = "X"
        Next
    Else
        For i = 1 To flxPurchHistory.Rows - 1
            flxPurchHistory.TextMatrix(i, 1) = ""
        Next
    End If
End Sub

Private Sub chkSelectAllDemands_Click()
    Dim i As Integer
    If chkSelectAllDemands.Value = 1 Then
        For i = 1 To flxPurchase.Rows - 1
            flxPurchase.TextMatrix(i, 1) = "X"
        Next
    Else
        For i = 1 To flxPurchase.Rows - 1
            flxPurchase.TextMatrix(i, 1) = ""
        Next
    End If
    
End Sub

Private Sub cmdOClientList_Click()
    sTextBox = "PIHistory"
    tabManagementFee.Enabled = False
'    tabPayment.Enabled = False
    chkShowBal.Visible = False
    picClient.Left = 2070
    picClient.Top = 800
    picClient.Visible = True
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub
Private Sub cmdGridUnitLookup_Click(Index As Integer)
   tabManagementFee.Enabled = True
   fraGenerateDemands.Enabled = True
    Frame1.Enabled = True
   If Index = 3 Then
        Call cmdClose_Click(0)
   End If
   If Index = 2 Then
      tabManagementFee.Enabled = True
      picAccList.Visible = False
      Exit Sub
   End If
   If Index = 1 Then
      tabManagementFee.Enabled = True
      picAccounts.Visible = False
      Exit Sub
   End If
   
   fraList.Visible = False

   tabManagementFee = True
   'tabPayment.Enabled = True
   If Index <> 0 Then
        fraList.Height = 2565
   End If
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Added by Anol 25 Mar 2015
    If Index = 0 Then
    
'        cmdACList(0).Enabled = True
'        If cmdACList(0).Enabled = True Then
'            cmdACList(0).SetFocus
'        End If
    End If
End Sub
Private Sub cmdClose_Click(Index As Integer)
'when flxPI.Tag = "Edited" and cmdEdit(1).Enabled=false msgbox do you want to save?
   Dim adoconn As New ADODB.Connection
   If flxPI.Tag = "EditedOrAdded" And cmdSavePI.Enabled = True Then
'        If MsgBox("Do you wish to save your changes?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
'           'If cmdSavePI.Enabled Then cmdSavePI.SetFocus
'           'Resolved by BOSL
'           'added by anol Date 21 Apr 2015
'           'issue 453 Note 5
'           'cmdSavePI_Click
'           'End of modification
'           'Exit Sub
'        End If
   Else ' you have done no changes
'         If cmdEdit(1).Enabled = False Then 'this is the left pane edit button. If this button is disabled that means invoice in edit mode
''            MsgBox "Returning from Editmode  non changes"
'                    adoConn.Open getConnectionString
'                    adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
'                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "' AND Module='Purchase Invoice'"
'                   flxPurchase.row = iPIEdit
'                    flxPurchase.col = 0
'                    flxPurchase.CellBackColor = vbWhite
'                   adoConn.Close
'                   Set adoConn = Nothing
'
'         End If
   End If
    flxPI.Tag = ""
'   If Not cmdEdit(Index).Enabled Then
'      If MsgBox("Do you want to add another transaction?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
'         If cmdSavePI.Enabled Then cmdSavePI.SetFocus
'         'Resolved by BOSL
'         'added by anol Date 21 Apr 2015
'         'issue 453 Note 5
'
'
'         'cmdSavePI_Click
'         'End of modification
'         Exit Sub
'      End If
'   End If
'   fraLay(0).Top = Me.Height + 1300
'   cmdEdit(1).Enabled = True
   
   txtSupplierName.text = ""
   txtSupplierID.text = ""
   txtProperty.text = ""
 'comment out by anol 20160912
''   Dim adoConn As New ADODB.Connection
'''   connect to database
''   adoConn.Open getConnectionString
''
''   LoadFlxPurchase adoConn
''
''   adoConn.Close
''   Set adoConn = Nothing
   ConfigFlxPI
   iPIEdit = 0
   FocusControl flxPurchase
End Sub
Private Sub ConfigFlxPI()
'issue 469
'modified by anol 28 Dec 2014
'On Error GoTo ERR
''   With flxPI
''      .Clear
''      .Cols = 27
''      .Rows = 2
''      .RowHeight(0) = 0     '                           Row Number of line
''      .ColWidth(0) = 700 ' SL NO
''      .ColAlignment(0) = vbLeftJustify
''      .ColWidth(1) = 0 '
''      .ColWidth(2) = 0 '
''      .ColWidth(3) = 0 '
''      .ColWidth(4) = 0 '
''      .ColWidth(5) = Label7(33).Left - Label7(32).Left  '  Unit No
''      .ColWidth(6) = Label7(34).Left - Label7(33).Left    'NominalCode
''      .ColWidth(7) = Label7(35).Left - Label7(34).Left    'Fund Code
''      .ColWidth(8) = 0  'Fund ID
''      .ColWidth(9) = Label7(36).Left - Label7(35).Left   'Job No
''      .ColWidth(10) = 0                     '"Cost Code"
''      .ColWidth(11) = Label7(37).Left - Label7(36).Left - 300 'Details
''      .ColWidth(12) = Label7(38).Left - Label7(37).Left   'Net
''      .ColWidth(13) = Label7(39).Left - Label7(38).Left   'T/C
''      .ColAlignment(13) = vbRightJustify
''      .ColWidth(14) = Label7(40).Left - Label7(39).Left   'VAT
''      .ColWidth(15) = 1150 '   'Total
''      .ColWidth(16) = 0                     '"Sage"
''      .ColWidth(17) = 0           'Stores PI Id hidenly
''      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
''      .ColWidth(19) = 0           'keep value 0 or 1 for edit
''      .ColWidth(20) = 0 'Label7(13).Left - Label7(12).Left           'Stores ScheduleId
''      .ColWidth(21) = 0           'Stores Unit ID
''      .ColWidth(22) = 0           '% Recoverable
''      .ColWidth(23) = 0           'ID
''      .ColWidth(24) = 0 '.Width - Label7(1).Left - 120           'FundCode
''      .ColWidth(25) = 0           'FundName
''      .ColWidth(26) = 0           'PO
''
''
'''      .ColWidth(0) = Label7(4).Left - .Left '"TransactionID" SL NO
'''      .ColWidth(1) = Label7(5).Left - Label7(4).Left '"A/C" SupplierId
'''      .ColWidth(2) = Label7(6).Left - Label7(5).Left '"Date"
'''      .ColWidth(3) = Label7(7).Left - Label7(6).Left '"Type" Property
'''      .ColWidth(4) = Label7(8).Left - Label7(7).Left '"Trans"
'''      .ColWidth(5) = Label7(9).Left - Label7(8).Left '"Unit ID + Name"
'''      .ColWidth(6) = Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
'''      .ColWidth(7) = 0                      '"N/C"
'''      .ColWidth(8) = 0                      '"Fund"
'''      .ColWidth(9) = 0                      '"Job No"
'''      .ColWidth(10) = 0                     '"Cost Code"
'''      .ColWidth(11) = Label7(11).Left - Label7(10).Left '"Details"
'''      .ColWidth(12) = Label7(12).Left - Label7(11).Left '"Net"
'''      .ColWidth(13) = Label7(13).Left - Label7(12).Left '"T/C"
'''      .ColWidth(14) = Label7(14).Left - Label7(13).Left '"VAT"
'''      .ColWidth(15) = .Width - Label7(14).Left - 120 '"Total"
'''      .ColWidth(16) = 0                     '"Sage"
'''      .ColWidth(17) = 0           'Stores PI Id hidenly
'''      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
'''      .ColWidth(19) = 0           'keep value 0 or 1 for edit
'''      .ColWidth(20) = 0           'Stores ScheduleId
'''      .ColWidth(21) = 0           'Stores Unit ID
'''      .ColWidth(22) = 0           '% Recoverable
'''      .ColWidth(23) = 0           'ID
'''      .ColWidth(24) = 0           'FundCode
'''      .ColWidth(25) = 0           'FundName
'''      .ColWidth(26) = 0           'PO
''      .row = 0
''   End With

'   txtPICNNet.Left = Label7(11).Left
'   txtPICNNet.Width = flxPI.ColWidth(12)
'   txtPICNVat.Left = Label7(13).Left
'   txtPICNVat.Width = flxPI.ColWidth(14)
'   txtPICNTotal.Left = Label7(14).Left
'   txtPICNTotal.Width = flxPI.ColWidth(15)
''   txtPICNNet.Left = Label7(8).Left - 20
''   txtPICNNet.Width = flxPI.ColWidth(12)
''   txtPICNVat.Left = Label7(10).Left - 20
''   txtPICNVat.Width = flxPI.ColWidth(14)
''   txtPICNTotal.Left = Label7(11).Left - 20
''   txtPICNTotal.Width = flxPI.ColWidth(15)
   Exit Sub
Err:
End Sub
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    tabManagementFee.Enabled = True
    fraGenerateDemands.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub cmdPostohistory_Click()
   Dim szTemp As String

'   frmPopUpMenu.Top = 500 ' frmMMain.fraCmdButton.Height + Me.Top + tabPurExp.Top + fraEditDemand.Top + cmdPostDemands.Top + 1140
'   frmPopUpMenu.Left = 1500 'frmMMain.tvwLandLord.Width + Me.Left + fraEditDemand.Left + tabPurExp.Left + cmdPostDemands.Left + 80
   frmPopUpMenu.CallingFrom "ManagementFee"

   szTemp = SelectedPurInvID()

   If szTemp <> "" Then
      frmPopUpMenu.optSelPO.Value = True
   Else
      frmPopUpMenu.optSelPI.Value = False
      frmPopUpMenu.optSelPI.Enabled = False
   End If
   chkProperty.Value = 0
   frmPopUpMenu.Left = 3800
   frmPopUpMenu.Top = 6000
   frmPopUpMenu.Show 1
End Sub
Private Function SelectedPurInvID() As String

   Dim i As Integer

  
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelectedPurInvID = SelectedPurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         SelectedPurInvID = SelectedPurInvID & ","
      End If
   Next i
   
  
   
   If Len(SelectedPurInvID) > 0 Then
      SelectedPurInvID = Left(SelectedPurInvID, Len(SelectedPurInvID) - 1)
   End If
End Function
Private Sub cmdPreViewGenDmds_Click()
    frmManagementFeeSelection.szCallingFrom = "ManagementFee"
    frmManagementFeeSelection.Caption = "Management Fee Generation Options"
    frmManagementFeeSelection.boldatechaged = False
    LoadForm frmManagementFeeSelection
End Sub
Private Sub ConfigflxReceipt()
   Dim szHeader As String, iCol As Integer
   flxReceipt.Clear
   flxReceipt.Cols = 20
   flxReceipt.Rows = 2
   'flxReceipt.RowHeight(0) = 0

   szHeader$ = "|<PI_ActualID|<Receipt No|<ReceiptSplitID|<ReceiptType|< ChargingMethod" & _
               "|<SageAccountNumber|>ReceiptDescription|<PropertyID|<FundCode|<ChargeDate|<ReceiptDate|<ReceiptTransactionID" & _
               "|<Receipt Amount|<AgrPercentage|<MgtFeeAmt|<VATPercentage|<VAT|<MgtFeeAmtTotal"
   flxReceipt.FormatString = szHeader$

   flxReceipt.ColWidth(0) = 0
   flxReceipt.ColWidth(1) = 0 'PI_ActualID
   flxReceipt.ColWidth(2) = 1200  'PISLNumber
   flxReceipt.ColWidth(3) = 1200  'ReceiptSplitID
   flxReceipt.ColWidth(4) = 0    'ReceiptType
   flxReceipt.ColWidth(5) = 1400    'ChargingMethod
   flxReceipt.ColWidth(6) = 1600    'SageAccountNumber
   flxReceipt.ColWidth(7) = 1600    'ReceiptTypeDescription
   flxReceipt.ColWidth(8) = 1200    'PropertyID
   flxReceipt.ColWidth(9) = 1200    'FundName
   flxReceipt.ColWidth(10) = 1200    'ChargeDate
   flxReceipt.ColWidth(11) = 1200    'ReceiptDate
   flxReceipt.ColWidth(12) = 0       'ReceiptTransactionID
   flxReceipt.ColWidth(13) = 1200    'ReceiptSplitID
   flxReceipt.ColWidth(14) = 1400    'Receipt Amount
   flxReceipt.ColWidth(15) = 1200    'AgrPercentage
   flxReceipt.ColWidth(16) = 1200    'MgtFeeAmt
   flxReceipt.ColWidth(17) = 1200    'VAT
   flxReceipt.ColWidth(18) = 1400    'MgtFeeAmtTotal
   flxReceipt.ColWidth(19) = 0    '
   
   
   

  
End Sub
Private Sub ConfigFlxPurchaseSplit()
   Dim szHeader As String, iCol As Integer
   Dim iLabel As Integer
   
   iLabel = 19

   flxPurchaseSplit.Clear
   flxPurchaseSplit.Cols = 12
   flxPurchaseSplit.Rows = 2
   flxPurchaseSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
               "|<Fund Code|<Job No|<Desc|>Net|>VAT|>Amount|>Recoverable"
   flxPurchaseSplit.FormatString = szHeader$

   flxPurchaseSplit.ColWidth(0) = 0
   flxPurchaseSplit.ColWidth(1) = lblPurchaseSplit(1 + iLabel).Left - flxPurchaseSplit.Left

   For iCol = 2 To flxPurchaseSplit.Cols - 2
      flxPurchaseSplit.ColWidth(iCol) = lblPurchaseSplit(iCol + iLabel).Left - lblPurchaseSplit(iCol - 1 + iLabel).Left
   Next iCol
   flxPurchaseSplit.ColWidth(iCol) = flxPurchaseSplit.Width + flxPurchaseSplit.Left - lblPurchaseSplit(iCol - 1 + iLabel).Left - 340
End Sub
Private Sub ConfigFlxPurchase()
   Dim szHeader As String, iCol As Integer

   flxPurchase.Clear
   'flxPurchase.Cols = 19
    ' I am adding 1 col . I shall use that for slnumber
   flxPurchase.Cols = 27
   flxPurchase.Rows = 2
   flxPurchase.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<ClientID|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID|>Outstanding|PostingDate|PO|PO_ID"

   flxPurchase.FormatString = szHeader$
   flxPurchase.ColWidth(0) = 0
   flxPurchase.ColWidth(1) = Label20(10).Left - flxPurchase.Left - 10
   '19-10=9
   '20-10=10
   '20-11=9
  
    flxPurchase.ColWidth(1 + 1) = Label20(11).Left - Label20(10).Left
    flxPurchase.ColWidth(2 + 1) = Label20(12).Left - Label20(11).Left
    flxPurchase.ColWidth(3 + 1) = Label20(13).Left - Label20(12).Left
    flxPurchase.ColWidth(4 + 1) = Label20(14).Left - Label20(13).Left
    flxPurchase.ColWidth(5 + 1) = Label20(15).Left - Label20(14).Left
    flxPurchase.ColWidth(6 + 1) = Label20(16).Left - Label20(15).Left
    flxPurchase.ColWidth(7 + 1) = Label20(17).Left - Label20(16).Left
    flxPurchase.ColWidth(8 + 1) = Label20(18).Left - Label20(17).Left
    flxPurchase.ColWidth(9 + 1) = Label20(19).Left - Label20(18).Left
    
    flxPurchase.ColWidth(10 + 1) = Label20(19).Left - Label20(18).Left
    flxPurchase.ColWidth(12) = 0
    flxPurchase.ColWidth(13) = 0
    flxPurchase.ColWidth(14) = 0
    flxPurchase.ColWidth(15) = 0
    flxPurchase.ColWidth(16) = 0
    flxPurchase.ColWidth(17) = 0
    flxPurchase.ColWidth(18) = 0
    flxPurchase.ColWidth(19) = 0
    flxPurchase.ColWidth(20) = 0
    flxPurchase.ColWidth(21) = 0
    flxPurchase.ColWidth(22) = 0
    flxPurchase.ColWidth(23) = 0
    flxPurchase.ColWidth(24) = 0
    flxPurchase.ColWidth(25) = 0
    flxPurchase.ColWidth(26) = 0

    
    
   
   'iCol = 10
'   flxPurchase.ColWidth(iCol) = 0                        'Client
'   flxPurchase.ColWidth(iCol + 1) = 0                    'Property
'   flxPurchase.ColWidth(iCol + 2) = 0                    'OS Amt
'   flxPurchase.ColWidth(iCol + 3) = 0                    'Due Date
'   flxPurchase.ColWidth(iCol + 4) = 0                    'Client ID
'   flxPurchase.ColWidth(iCol + 5) = 1700 'flxPurchase.Width + flxPurchase.Left - Label20(18).Left - 200  'Outstanding
'   flxPurchase.ColWidth(iCol + 6) = 0                    'Posting Date
'   flxPurchase.ColWidth(iCol + 7) = 0                    'PO
'   flxPurchase.ColWidth(iCol + 8) = 0                    'PO_ID
'   flxPurchase.ColWidth(iCol + 9) = 0                    'slnumber flxPurchase.col(19) shall be used for slnumber
'   flxPurchase.ColWidth(iCol + 10) = 0                   'Type col=20 type of supplier
'   flxPurchase.ColWidth(iCol + 11) = 0                   ' col=21 for tlbpayment transaction ID
'   flxPurchase.ColWidth(iCol + 12) = 0                   'col=22 for tlbpayment UserSessionID
'   flxPurchase.ColWidth(iCol + 13) = 0                   'col=22 for tlbpayment WindowsUserName
'   flxPurchase.ColWidth(iCol + 14) = 0                   'col=22 for tlbpayment MachineName
'   flxPurchase.ColWidth(iCol + 15) = 0                   'col=22 for tlbpayment Module
   
End Sub
Public Sub LoadFlxPurchase(adoconn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWherePropertyId As String
   Dim strWhereClient As String
   Dim dblRctTotal As Double
   ConfigFlxPurchase
   ConfigFlxPurchaseSplit
   If txtPropID.text <> "ALL" Then
        strWherePropertyId = " AND PI.PropertyID ='" & txtPropID.text & "'"
   End If
   If txtIDClient.text <> "ALL" Then
         strWhereClient = " AND PI.CL_ID='" & txtIDClient.text & "'"
   End If
   szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, " & _
               "Pt.OSAmount, PI.PO,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE  " & _
               "tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION,Supplier.Type,Pt.TransactionID,Pt.UserSessionID,Pt.WindowsUserName,Pt.MachineName ,Pt.Module  " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "Where PI.isManagementFee=true AND PI.History = False " & strWherePropertyId & strWhereClient & " AND (PI.TransactionType = 6 OR " & _
               "PI.TransactionType = 7) " & _
           "ORDER BY 3 , 2 Desc;"

   adoInv.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'Debug.Print time
    If Not adoInv.EOF Then
        fmeLoading.Visible = True
        fmeLoading.Refresh
    End If
   iKount = 1
   colTransactionIDOtherPIGrid = ""
   With flxPurchase
      .Rows = adoInv.RecordCount + 1
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
         .TextMatrix(iKount, 10) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         dblRctTotal = dblRctTotal + Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 12) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12 + 1) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 13 + 1) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 14 + 1) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 15 + 1) = adoInv.Fields.Item("PostingDate").Value
         .TextMatrix(iKount, 16 + 1) = IIf(IsNull(adoInv.Fields.Item("PO").Value), "", adoInv.Fields.Item("PO").Value)
'         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PO_ID").Value), "", adoInv.Fields.Item("PO_ID").Value)'No use I can see anol 20181118
         .TextMatrix(iKount, 17 + 1) = IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 18 + 1) = IIf(IsNull(adoInv.Fields.Item("Type").Value), "", adoInv.Fields.Item("Type").Value)
         .TextMatrix(iKount, 19 + 1) = IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value)
         .TextMatrix(iKount, 20 + 1) = IIf(IsNull(adoInv.Fields.Item("UserSessionID").Value), "", adoInv.Fields.Item("UserSessionID").Value)
         .TextMatrix(iKount, 21 + 1) = IIf(IsNull(adoInv.Fields.Item("WindowsUserName").Value), "", adoInv.Fields.Item("WindowsUserName").Value)
         .TextMatrix(iKount, 22 + 1) = IIf(IsNull(adoInv.Fields.Item("MachineName").Value), "", adoInv.Fields.Item("MachineName").Value)
         .TextMatrix(iKount, 23 + 1) = IIf(IsNull(adoInv.Fields.Item("Module").Value), "", adoInv.Fields.Item("Module").Value)
         If .TextMatrix(iKount, 22 + 1) <> "" Then
            .col = 1
            .row = iKount
            .CellBackColor = vbRed
            colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value) & ","
         End If
         
         'Debug.Print .TextMatrix(iKount, 20)
         'issue 316 by anol 20170221
         If iKount = 10 Then
            'frmPurchaseExpense.Refresh
            lblLoading.Caption = "Please wait while loading...."
            flxPurchase.Refresh
         End If
         If iKount = 17 Then
             lblLoading.Caption = "Please wait while loading....."
             lblLoading.Refresh
            flxPurchase.Refresh
         End If
      
         adoInv.MoveNext
         iKount = iKount + 1
         'If Not adoInv.EOF Then .AddItem ""
      Wend
      'Debug.Print time
   End With
   If Len(colTransactionIDOtherPIGrid) > 0 Then
            colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
   End If
    txtRctTotal.text = dblRctTotal
XX:
   adoInv.Close
   Set adoInv = Nothing
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

          'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
'                        If Not ctl Is picClient Then
'                            PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
'                        Else
                            bHandled = False
'                        End If

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

Private Sub SelPurHistory(ByRef SelPurHisID() As String)
   Dim i As Integer
   Dim j As Integer
   For i = 1 To flxPurchHistory.Rows - 1
      If flxPurchHistory.TextMatrix(i, 1) = "X" Then
'         SelPurHistory = SelPurHistory & "'" & CStr(flxPurchHistory.TextMatrix(i, 0)) & "'"
'         SelPurHistory = SelPurHistory & ","
         j = j + 1
      End If
   Next i
   ReDim SelPurHisID(j)
   j = 0
   For i = 1 To flxPurchHistory.Rows - 1
      If flxPurchHistory.TextMatrix(i, 1) = "X" Then
         SelPurHisID(j) = "'" & CStr(flxPurchHistory.TextMatrix(i, 0)) & "'"
         j = j + 1
      End If
   Next i
'   If Len(SelPurHistory) > 0 Then
'        SelPurHistory = Left(SelPurHistory, Len(SelPurHistory) - 1)
'   End If
End Sub

Private Sub cmdPrintBreakdown_Click()
   
   Dim iRow    As Integer
   Dim K       As Integer
  
   K = 0
   For iRow = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(iRow, 1) = "X" Then
         K = K + 1
         iSelRow = iRow
         flxPurchase.row = iRow
      End If
   Next iRow
   If K = 0 Then
      MsgBox "Please select a Management Fee", vbInformation, "Information"
      Exit Sub
   End If
   If K > 1 Then
      MsgBox "Please select only one Management Fee", vbInformation, "Information"
      For iRow = 1 To flxPurchase.Rows - 1
            flxPurchase.TextMatrix(iRow, 1) = ""
      Next iRow
      Exit Sub
   End If
   
   'frmNJ_Entry.lHeaderID = Mid(flxNJ.TextMatrix(flxNJ.row, 1), 4)
   
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ManagementFeeSingle.rpt")
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Report.ParameterFields(1).AddCurrentValue Val(flxPurchase.TextMatrix(flxPurchase.row, 18)) 'PI_ActualID
   Load frmReport
   frmReport.LoadReportViewer Report
   
End Sub

Private Sub cmdRevHistory_Click()
'Added by anol 10 Jun 2015
'issue 0000572: Purchases and expenses - Reverse postings to history not present
   Dim szSQL As String
   Dim iRow As Integer, szPurID As String
   Dim j As Long
   Dim K As Integer
   Dim SelPurHisID() As String
   Dim adoconn As New ADODB.Connection
   'in a grid if there is hundreds of records it cannot process with one SQL Need to seperate them or I can use an array to buld up Update query one Batch can hold upto 50
   'invoices
   'On Error GoTo Catch_Error
   Call SelPurHistory(SelPurHisID())
   j = UBound(SelPurHisID())
   If j = 0 Then
        MsgBox "Please select a Management Fee to reverese.", vbCritical + vbOKOnly, "Management Fee"
        Exit Sub
   End If
   If MsgBox("Are you sure you wish to reverse the selected transactions from history?", vbQuestion + vbYesNo, "Purchase Invoice History") = vbNo Then Exit Sub
  

   adoconn.Open getConnectionString
   K = CInt(j / 50)
   
   If K = j / 50 Then
        'No no need to do ceiling, this is fully divisible
        K = j / 50
   Else
        K = CInt(j / 50) + 1 'This is ceiling function
   End If
   For K = 0 To K - 1
           
           szPurID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurHisID())
           If szPurID = "" Then
                Exit For
           End If
       If Trim(szPurID) <> "" Then
           szSQL = "UPDATE tblPurInv " & _
           "SET    History = FALSE " & _
           "WHERE  My_ID IN (" & szPurID & ") ; "
            adoconn.Execute szSQL
       End If
   Next

'   szSQL = "UPDATE tblPurInv " & _
'           "SET    History = FALSE " & _
'           "WHERE  SlNumber IN (" & SelPurHistory & ") ; "
'    szSQL = "UPDATE tblPurInv " & _
'           "SET    History = FALSE " & _
'           "WHERE  My_ID IN (" & szPurID & ") ; "
           
'           AND " & _
'                  "  DemandID NOT IN (" & _
'                  "     SELECT DemandRef " & _
'                  "     From tlbReceipt " & _
'                  "     WHERE Type = 1 AND Amount > OSAmount);"
'Debug.Print szSQL
'   adoConn.Execute szSQL

   LoadFlxPurchHistory adoconn, ""
   LoadFlxPurchase adoconn
   fmeLoading.Visible = False
   adoconn.Close
   Set adoconn = Nothing
   chkAllPurchaseHistory.Value = 0
   ShowMsgInTaskBar "System has reversed " & j & " invoices.", "Y", "P"
   Exit Sub

Catch_Error:
   MsgBox "Select a purchase invoice to reverse.", vbCritical + vbOKOnly, "Purchase invoice"
End Sub
Private Sub Command1_Click()
    frmManagementFeeSelection.szCallingFrom = "ManagementFee Preview"
'    frmManagementFeeSelection.Show
'    frmManagementFeeSelection.ZOrder 0
    frmManagementFeeSelection.boldatechaged = False
    frmManagementFeeSelection.Caption = "Management Fee Preview"
    LoadForm frmManagementFeeSelection
End Sub



Private Sub Command3_Click()
     Dim Conn1 As New ADODB.Connection
    If MsgBox("are you sure you want to fix Mgt Fee marking?", vbYesNo, "Please confirm") = vbYes Then
        Conn1.Open getConnectionString
'        Conn1.Execute "Update tlbReceiptSplit Set isMgtFeeS=true where transactionID='22021511320600119795'"
'        Conn1.Execute "Update tlbReceiptSplit Set isMgtFeeS=true where transactionID='22020213314903374602'"
'        Conn1.Execute "Update tlbReceiptSplit Set isMgtFeeS=true where transactionID='22021712463703268516'"
'        Conn1.Execute "UPDATE tlbReceipt AS R, tlbReceiptsplit AS RS, tlbReceipt AS R1, rptTransactionsSPlit AS AL, DemandSplitRecords AS DS, Units AS U SET  RS.ISMGTFEES=true " & _
'                       "WHERE (((R1.DemandRef)=[DS].[DemandID]) AND ((AL.ToTran)=[R1].[TransactionID]) AND ((RS.SplitID)=[DS].[SPLITID]) AND ((AL.DeleteFlag)=False) AND " & _
'                       "((AL.TransactionID)=[RS].[RptTransactionsIDSplit]) AND ((R.TransactionID)=[RS].[RptHeader]) AND ((R.Type) In (3,4,23)) AND ((U.UnitNumber)=[R].[UnitID]) " & _
'                       "AND ((U.PropertyID)='45KSP') AND ((RS.FundID)=1) AND ((DS.TypeOfDemand) In (183)));"
             'Conn1.Execute "Update tlbReceipt set PIREFMGTFEE='22062812561601275332' where transactionID  in (3999,3998 )"
           'rem on 2023-02-17
           ' Conn1.Execute "Update tlbReceiptSplit SET ISMGTFeeS=true where transactionID ='22040111021600377020'"

        Conn1.Execute "Update tlbagreement T,ClientProAgr P set T.ReportClientID=P.ClientID,T.ReportPropertyID=P.PropertyID where T.CPA_ID=P.CPA_ID"
        Conn1.Execute "Update tlbReceiptSplit  RS,tlbAgreement G,rptTransactions A SET ISMGTFeeS=true where RS.PropertyID=G.ReportPropertyID " & _
        "AND G.CHARGE_METHOD='RE_ED' and A.AllocDate<=G.LastChargeDate AND A.FROMTran=RS.RptHeader AND cint(G.Fund)=RS.FundID AND A.DeleteFlag=False"
        MsgBox "Update completed"
        Conn1.Close
    End If
End Sub

Private Sub Command4_Click()
'    MsgBox "Print format not defined"
        Dim reportApp As New CRAXDRT.Application
        Dim Report As CRAXDRT.Report
        Dim rep As frmReport
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & "ManagementFeeListing.rpt")
        Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        Report.EnableParameterPrompting = False
        Report.DiscardSavedData
        Set rep = New frmReport
        Load rep
        rep.LoadReportViewer Report
End Sub

Private Sub Command6_Click()
     Dim Conn1 As New ADODB.Connection
    If MsgBox("are you sure you want to fix decimal places?", vbYesNo, "Please confirm") = vbYes Then
    
     Conn1.Open getConnectionString
    
      Update2DecimalPlace Conn1, "tlbPayment", "TransactionID", "Amount"
      Update2DecimalPlace Conn1, "tlbPayment", "TransactionID", "OSAmount"

      Update2DecimalPlace Conn1, "tlbPaymentSplit", "TransactionID", "Amount"
      Update2DecimalPlace Conn1, "tlbPaymentSplit", "TransactionID", "OSAmount"

      Update2DecimalPlace Conn1, "tblPurInv", "MY_ID", "TOTAL_AMOUNT"

      Update2DecimalPlace Conn1, "tblPurInvSRec", "MY_ID", "NET_AMOUNT"
      Update2DecimalPlace Conn1, "tblPurInvSRec", "MY_ID", "TOTAL_AMOUNT"
      
        MsgBox "Update done"
        Conn1.Close
    End If
End Sub
Private Sub Update2DecimalPlace(Conn1 As ADODB.Connection, szTable As String, szTableID As String, szField As String)
   On Error GoTo Err_Catch

   Dim adoS       As New ADODB.Recordset
   Dim adod       As New ADODB.Recordset
   Dim szSQL      As String

   szSQL = "SELECT " & szTableID & ", " & szField & " " & _
           "FROM " & szTable & ";"

   adoS.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   adod.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   While Not adod.EOF
'If adoD.Fields.Item(szTableID).Value = "1103181609414935206" Then
'MsgBox ""
'End If
'http://www.w3schools.com/ado/met_comm_createparameter.asp
      If adod.Fields.Item(szTableID).Type = 3 Then
         adoS.Find szTableID & " = " & adod.Fields.Item(szTableID).Value & "", , , 1
      Else
         adoS.Find szTableID & " = '" & adod.Fields.Item(szTableID).Value & "'", , , 1
      End If
'Debug.Print CCur(adoS.Fields.Item(szField).Value)
      adod.Fields.Item(szField).Value = RoundingNumber(adoS.Fields.Item(szField).Value, 2)
      adoS.MoveFirst
      adod.Update
      adod.MoveNext
   Wend

   adoS.Close
   Set adoS = Nothing
   adod.Close
   Set adod = Nothing

   Exit Sub
Err_Catch:
   Debug.Print Err.description
End Sub
Private Sub flxClient_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    tabManagementFee.Enabled = True
    fraGenerateDemands.Enabled = True
    Frame1.Enabled = True
      If sTextBox = "PIHistory" Then 'PIHistory
            txtClientIdlist.text = flxClient.TextMatrix(flxClient.row, 0)
            txtClientIdlist.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtPropertyIDHist.text = "ALL"
            chkPropertyHist.Value = 0
            LoadFlxPurchHistory adoconn, ""
            FocusControl cmdOClientList
        ElseIf sTextBox = "1" Then 'filter on client PI list
            txtIDClient.text = flxClient.TextMatrix(flxClient.row, 0)
            txtIDClient.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtPropID.text = "ALL"
            chkProperty.Value = 0
            Call LoadFlxPurchase(adoconn)
            fmeLoading.Visible = False
            FocusControl cmdOpClient
       End If
       adoconn.Close
       Set adoconn = Nothing
       picClient.Visible = False
        
End Sub

Private Sub flxSupplier_Click(Index As Integer)
    tabManagementFee.Enabled = True
    fraGenerateDemands.Enabled = True
    Frame1.Enabled = True
    Dim adoconn As New ADODB.Connection
    Dim rstVat As New ADODB.Recordset
    Dim szSQL As String
   
    
    If sTextBox = "PROPERTYFILTER" Then
        tabManagementFee.Enabled = True
        fraGenerateDemands.Enabled = True
        Frame1.Enabled = True
        txtPropID.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        txtPropID.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        
        fraList.Visible = False
        adoconn.Open getConnectionString
        Call LoadFlxPurchaseFilter(adoconn, "")
        adoconn.Close
        Set adoconn = Nothing
        fmeLoading.Visible = False
        Exit Sub
    End If
    fraList.Visible = False
End Sub
Public Sub LoadFlxPurchaseFilter(adoconn As ADODB.Connection, Filter As String)
   Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWhere As String
   Dim strWhereClient As String
   Dim strWhereProperty As String
   Dim tempstr As String
    Dim dblRctTotal As Double
   ConfigFlxPurchase
   ConfigFlxPurchaseSplit
    If Filter = "3" Then
         If txtSearchFromD.text <> "" Then
           strWhere = " AND PI.PostingDate =#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "#"
            If Len(txtSearchFromD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
    If txtPropID.text <> "ALL" Then
        strWhereProperty = " AND PI.PropertyID='" & txtPropID.text & "'"
    End If
    If txtIDClient.text <> "ALL" Then
         strWhereClient = " AND PI.CL_ID='" & txtIDClient.text & "'"
    End If
    If Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND PI.PostingDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND PI.PostingDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#"
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
   szSQL = "SELECT DISTINCT PI.MY_ID, (MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)& PI.SlNumber) AS INVNO, PI.SlNumber,PI.TransactionType, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, " & _
               "Pt.OSAmount, QQ.PO, QQ.PO_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE  " & _
               "tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION,Supplier.Type,Pt.TransactionID,Pt.UserSessionID,Pt.Module,Pt.WindowsUserName,Pt.MachineName " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "LEFT JOIN (" & _
                  "SELECT Q2.MY_ID, Q1.SLNumber AS PO, Q2.PO AS PO_ID " & _
                  "FROM ( " & _
                  "SELECT MY_ID, SLNumber " & _
                  "From tblPurInv " & _
                  "WHERE TransactionType = 25) AS Q1 INNER JOIN " & _
                  "(SELECT MY_ID, PO " & _
                  "From tblPurInv " & _
                  "WHERE PO <> '') AS Q2 ON Q1.MY_ID = Q2.PO " & _
               ") AS QQ ON PI.MY_ID = QQ.MY_ID " & _
           "Where  PI.isManagementFee=true AND PI.History = False " & strWhere & strWhereProperty & strWhereClient & " AND (PI.TransactionType = 6 OR " & _
               "PI.TransactionType = 7) " & _
           "ORDER BY 3 Desc, 2;"
'Debug.Print szSQL
   adoInv.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'adoInv.Close
'Exit Sub
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoInv.Filter = "INVNO Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoInv.Filter = "SupplierName Like '%" & tempstr & "%'"
        End If
    End If

    If Not adoInv.EOF Then
        fmeLoading.Visible = True
        fmeLoading.Refresh
    End If
   iKount = 1
   colTransactionIDOtherPIGrid = ""
   With flxPurchase
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVNO").Value
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
          dblRctTotal = dblRctTotal + Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 13) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 16) = adoInv.Fields.Item("PostingDate").Value
         .TextMatrix(iKount, 17) = IIf(IsNull(adoInv.Fields.Item("PO").Value), "", adoInv.Fields.Item("PO").Value)
         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PO_ID").Value), "", adoInv.Fields.Item("PO_ID").Value)
         .TextMatrix(iKount, 19) = IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 20) = IIf(IsNull(adoInv.Fields.Item("Type").Value), "", adoInv.Fields.Item("Type").Value)
         .TextMatrix(iKount, 21) = IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value)
         
         
         .TextMatrix(iKount, 22) = IIf(IsNull(adoInv.Fields.Item("UserSessionID").Value), "", adoInv.Fields.Item("UserSessionID").Value)
         If .TextMatrix(iKount, 22) <> "" Then
            colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value) & ","
         End If
         
         .TextMatrix(iKount, 23) = IIf(IsNull(adoInv.Fields.Item("WindowsUserName").Value), "", adoInv.Fields.Item("WindowsUserName").Value)
         .TextMatrix(iKount, 24) = IIf(IsNull(adoInv.Fields.Item("MachineName").Value), "", adoInv.Fields.Item("MachineName").Value)
         .TextMatrix(iKount, 25) = IIf(IsNull(adoInv.Fields.Item("Module").Value), "", adoInv.Fields.Item("Module").Value)

         If .TextMatrix(iKount, 22) <> "" Then
            .col = 1
            .row = iKount
            .CellBackColor = vbRed
         End If
         
         'issue 316 by anol 20170221
         If iKount = 10 Then
            frmPurchaseExpense.Refresh
            lblLoading.Caption = "Please wait while loading."
            flxPurchase.Refresh
         End If
         If iKount = 17 Then
             lblLoading.Caption = "Please wait while loading.."
             lblLoading.Refresh
            flxPurchase.Refresh
         End If
'         If iKount = 500 Then
'             lblLoading.Caption = "Please wait while loading..."
'             lblLoading.Refresh
'             flxPurchase.Refresh
''             GoTo XX
'         End If
'          If iKount = 1000 Then
'             lblLoading.Caption = "Please wait while loading...."
'             lblLoading.Refresh
'            flxPurchase.Refresh
'         End If
'          If iKount = 1500 Then
'             lblLoading.Caption = "Please wait while loading....."
'             lblLoading.Refresh
'            flxPurchase.Refresh
'         End If
'''######################################################################################################################
'''         Adding description of the header from the first split
''         szSQL = "SELECT DISTINCT * " & _
''                 "FROM tblPurInvSRec " & _
''                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
''                 "ORDER BY TRAN_ID;"
'''Debug.Print szSQL
''         adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
''
''         bFirstSp = True
''         If Not adoInvSp.EOF Then _
''            .TextMatrix(iKount, 8) = IIf(IsNull(adoInvSp.Fields.Item("DESCRIPTION").Value), "", adoInvSp.Fields.Item("DESCRIPTION").Value)
''
''         adoInvSp.Close
        .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
         adoInv.MoveNext
         iKount = iKount + 1
         If Not adoInv.EOF Then .AddItem ""
      Wend
      
   End With
   If Len(colTransactionIDOtherPIGrid) > 0 Then
            colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
   End If
   txtRctTotal.text = dblRctTotal
XX:
   adoInv.Close
   Set adoInv = Nothing
End Sub
Private Sub Form_Load()
    Me.Width = 20220
    Me.Height = 12165
    txtPropID.text = "ALL"
    txtIDClient.text = "ALL"
    txtClientIdlist.text = "ALL"
    txtSupplierSearc.text = "ALL"
   
    tabManagementFee.Tab = 0
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call LoadFlxPurchase(adoconn)
    fmeLoading.Visible = False
    adoconn.Close
    Call WheelHook(Me.hWnd)
    Set adoconn = Nothing
End Sub
Private Sub cmdOpClient_Click()
    sTextBox = "1"
'    chkShowBal.Visible = False
    tabManagementFee.Enabled = False
    fraGenerateDemands.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    picClient.Left = 2070
    picClient.Top = 800
    
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient(Filter As String)
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 6
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   
   txtSearchClientID.Width = 1530
   picClient.Width = 5295
   flxClient.Width = 5175
   cmdPicCLose.Left = 5010
   txtSearchClientName.Left = 1620
   lblClientName.Left = 1875
   txtSearchClientID.Left = 45
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3420
   picClient.Height = 4095
   flxClient.Height = 3345
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT, V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM ((CLIENT C INNER JOIN Supplier S ON C.ClientID=S.SupplierID) " & _
           "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) ORDER BY CLIENTID;"
     
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
            rstRec.Filter = Filter
   End If
   flxClient.Rows = rstRec.RecordCount + 2
   If tabManagementFee.Tab = 0 Then
        If sTextBox = "1" Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 240
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
                flxClient.row = 1
                flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
                flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
                flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("CT").Value), "", rstRec.Fields("CT").Value)
                flxClient.TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
                flxClient.TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "", rstRec.Fields("VAT_ID").Value)
                flxClient.TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "", rstRec.Fields("VAT_RATE").Value)
                rstRec.MoveNext
               rRow = rRow + 1
           Wend
        Else
           rRow = 1
            While Not rstRec.EOF
                flxClient.row = 1
                flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
                flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
                flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("CT").Value), "", rstRec.Fields("CT").Value)
                flxClient.TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
                flxClient.TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "", rstRec.Fields("VAT_ID").Value)
                flxClient.TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "", rstRec.Fields("VAT_RATE").Value)
               rstRec.MoveNext
               rRow = rRow + 1
            Wend
        End If
   End If
   If tabManagementFee.Tab = 1 Then
            rRow = 1
            While Not rstRec.EOF
               flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               rstRec.MoveNext
               rRow = rRow + 1
            Wend
   End If
   If tabManagementFee.Tab = 2 Or tabManagementFee.Tab = 3 Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 240
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               rstRec.MoveNext
               rRow = rRow + 1
           Wend
   End If

   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub

Private Sub LoadPropertyList(Filter As String)
   Dim rRow As Integer
   Dim szSQL As String
'Exit Sub
   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxSupplier(0).RowHeight(0) = 0
   flxSupplier(0).Cols = 6
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   flxSupplier(0).ColWidth(4) = 0
   flxSupplier(0).ColWidth(5) = 0

   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColAlignment(0) = vbLeftJustify
   flxSupplier(0).ColAlignment(1) = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1.Width = 1490
   txtSearch1.Left = 40
   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   lblSearch0(0).Caption = "Property ID"
   lblSearch1.Caption = "Property Name"
   lblSearch2.Visible = False
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoconn.Open getConnectionString

'   On Error Resume Next
    

   rRow = 1
   
    If sTextBox = "PROPERTY" Then
        rRow = 1
        If txtClientID.text <> "ALL" Then
            'Modification in SQL written by anol 2020-10-07
            szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as VAT_CODE1,G.VATRate  as  Rate,G.vatOptionEnabled " & _
                  "FROM ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN tlbVatCode V ON G.VATRate=V.VAT_ID) " & _
                  "WHERE ClientID = '" & txtClientID.text & "' " & _
                  "ORDER BY P.PropertyID;"
        Else
            szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as as VAT_CODE1,G.VATRate  as  Rate,vatOptionEnabled " & _
                  "FROM ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN  tlbVatCode V  ON G.VATRate=V.VAT_ID) " & _
                  "ORDER BY P.PropertyID;"
        End If
        
        rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Filter <> "" Then
            rstRec.Filter = Filter
        End If
        flxSupplier(0).Rows = rstRec.RecordCount + 2
        While Not rstRec.EOF
            flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
            If rRow = 1 Then
                Debug.Print rstRec.Fields.Item(0).Value
            End If
            flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
            If (IIf(IsNull(rstRec.Fields("vatOptionEnabled").Value), "", rstRec.Fields("vatOptionEnabled").Value)) = 1 Then
                flxSupplier(0).TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("Rate").Value), "", rstRec.Fields("Rate").Value) 'VAT_ID
                flxSupplier(0).TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("RateValue").Value), "0.00", rstRec.Fields("RateValue").Value)
                flxSupplier(0).TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value) 'like T9
                flxSupplier(0).TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("vatOptionEnabled").Value), "", rstRec.Fields("vatOptionEnabled").Value) 'like T9
            Else
                flxSupplier(0).TextMatrix(rRow, 2) = "" 'VAT_ID
                flxSupplier(0).TextMatrix(rRow, 3) = ""
                flxSupplier(0).TextMatrix(rRow, 4) = "" 'like T9
                flxSupplier(0).TextMatrix(rRow, 5) = "0" '0 means VAT disabled for the property
            End If
            rstRec.MoveNext
            rRow = rRow + 1
        Wend
         rstRec.Close
         Set rstRec = Nothing
 ElseIf sTextBox = "PROPERTYFILTER" Then 'Properties loading into the filter
        rRow = 1
        If txtIDClient.text <> "ALL" Then
            szSQL = "SELECT PropertyID, PropertyName " & _
                  "FROM Property " & _
                  "WHERE ClientID = '" & txtIDClient.text & "' " & _
                  "ORDER BY PropertyID;"
        Else
            szSQL = "SELECT PropertyID, PropertyName " & _
                  "FROM Property " & _
                  "ORDER BY PropertyID;"
        End If
        
        rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
         If Filter <> "" Then
            rstRec.Filter = Filter
        End If
        flxSupplier(0).Rows = rstRec.RecordCount + 3
        flxSupplier(0).TextMatrix(1, 0) = "ALL"
        flxSupplier(0).TextMatrix(1, 1) = "ALL Properties"
        'flxSupplier(0).AddItem ""
        rRow = 2
        While Not rstRec.EOF
            flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
            flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
            'flxSupplier(0).RowHeight(rRow) = 240
            rstRec.MoveNext
            'If Not rstRec.EOF Then flxSupplier(0).AddItem ""
            rRow = rRow + 1
        Wend
        rstRec.Close
  ElseIf sTextBox = "PROPERTYHIST" Then  'issue 629 creating new filter
            rRow = 1
            If txtClientIdlist.text <> "ALL" Then
               szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "WHERE ClientID = '" & txtClientIdlist.text & "' " & _
                       "ORDER BY PropertyID;"
            Else
                 szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "ORDER BY PropertyID;"
            End If
            rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            If Filter <> "" Then
                rstRec.Filter = Filter
            End If
            flxSupplier(0).Rows = rstRec.RecordCount + 3
            flxSupplier(0).TextMatrix(1, 0) = "ALL"
            flxSupplier(0).TextMatrix(1, 1) = "ALL Properties"
            'flxSupplier(0).AddItem ""

            rRow = 2
            While Not rstRec.EOF
               flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               'flxSupplier(0).RowHeight(rRow) = 240
               rstRec.MoveNext
               'If Not rstRec.EOF Then flxSupplier(0).AddItem ""
               rRow = rRow + 1
            Wend
              rstRec.Close
  End If
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub cmdOpProperty_Click()
'    chkShowBal.Visible = False
   sTextBox = "PROPERTYFILTER"
   LoadPropertyList ""

   tabManagementFee.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4855
   cmdGridUnitLookup(tabManagementFee.Tab).Left = fraList.Width - cmdGridUnitLookup(tabManagementFee.Tab).Width - 60
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(tabManagementFee.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 80
'   fraList.Left = txtProperty.Left + fraLay(0).Left + 100
   fraList.Left = txtPropID.Left - 400
   fraList.Top = 800
   fraList.Visible = True
   fraList.ZOrder 0
   
   
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
End Sub
Private Sub chkProperty_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    If chkProperty.Value = 0 Then
        txtPropID.text = "ALL"
        cmdOpProperty.Enabled = True
        LoadFlxPurchase adoconn
        
    Else
        txtPropID.text = ""
        cmdOpProperty.Enabled = False
        'SortTheGrid flxPurchase, txtIDClient, txtPropID, txtSupplier
        LoadFlxPurchase adoconn
    End If
    fmeLoading.Visible = False
    adoconn.Close
    Set adoconn = Nothing
End Sub
Private Sub LoadflxSupplier(ByVal adoconn As ADODB.Connection)
   ConfigFlxSupplier2

   Dim adoRst  As New ADODB.Recordset
'   Dim adoLL   As New ADODB.Recordset
   Dim adoC    As New ADODB.Recordset
   Dim adoMA   As New ADODB.Recordset

   Dim szSQL      As String
   Dim iTotalRow  As Integer
   Dim j          As Integer
   Dim i          As Integer
   Dim iTotalCol  As Integer
   Dim Data()     As String

   'On Error GoTo ErrorHandler


         szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' Or TYPE = 'AGENT'  Or TYPE = 'Client' Or TYPE = 'LLORD' " & _
           "ORDER BY TYPE,SupplierName;"
 
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If iTotalRow = 0 Then GoTo NoRes
   
   iTotalCol = adoRst.Fields.Count


   If adoRst.RecordCount > 0 Then adoRst.MoveFirst

   i = 1
   If Not adoRst.EOF Then
         flxSupplier(2).TextMatrix(i, 0) = "ALL"
         flxSupplier(2).TextMatrix(i, 1) = "ALL"
         flxSupplier(2).TextMatrix(i, 2) = "All Supplier"
         flxSupplier(2).AddItem ""
         i = i + 1
   End If
   While Not adoRst.EOF
      flxSupplier(2).TextMatrix(i, 0) = adoRst.Fields.Item(2).Value
      flxSupplier(2).TextMatrix(i, 1) = adoRst.Fields.Item(0).Value
      flxSupplier(2).TextMatrix(i, 2) = adoRst.Fields.Item(1).Value
      adoRst.MoveNext
      If Not adoRst.EOF Then flxSupplier(2).AddItem ""
      i = i + 1
   Wend


NoRes:

   adoRst.Close

   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   Set adoRst = Nothing
'   Set adoLL = Nothing
   Set adoC = Nothing
   Set adoMA = Nothing
End Sub
Private Sub ConfigFlxSupplier2()
   Dim szHeader As String

   flxSupplier(2).Clear
   flxSupplier(2).Rows = 2
   flxSupplier(2).Cols = 3

   szHeader$ = "<|<|<"
   flxSupplier(2).FormatString = szHeader$

   flxSupplier(2).RowHeight(0) = 0
   flxSupplier(2).ColWidth(0) = lblSearch0(7).Left - lblSearch0(6).Left
   txtAccountSearch(3).Width = lblSearch0(7).Left - lblSearch0(6).Left
   flxSupplier(2).ColWidth(1) = lblSearch0(8).Left - lblSearch0(7).Left
   txtAccountSearch(4).Width = lblSearch0(8).Left - lblSearch0(7).Left
   flxSupplier(2).ColWidth(2) = flxSupplier(2).Width - lblSearch0(8).Left - 300
   txtAccountSearch(5).Width = flxSupplier(2).Width - lblSearch0(8).Left - 300
   txtAccountSearch(3).Left = lblSearch0(6).Left
   txtAccountSearch(4).Left = lblSearch0(7).Left
   txtAccountSearch(5).Left = lblSearch0(8).Left
End Sub
Private Sub cmdAccSel_Click()
    Dim adoconn As New ADODB.Connection
    Dim iRow As Integer
     chkShowBal.Visible = False
    ' If lblSearch0(5).Caption = "NotLoaded" Then
    'Set the ADO Connections to the dataset
    If adoconn.State = 0 Then
        adoconn.Open getConnectionString
    End If
    sTextBox = "PILIST"
    txtAccountSearch(0).text = ""
    txtAccountSearch(1).text = ""
    txtAccountSearch(2).text = ""
    txtAccountSearch(6).text = ""
    txtAccountSearch(7).text = ""
    LoadflxSupplier adoconn
    lblSearch0(5).Caption = "Loaded"
    adoconn.Close
    'End If
    cmdGridUnitLookup(2).Left = 5640
    
    Set adoconn = Nothing
    Set adoconn = Nothing
    'Resolved by BOSL
    'Modified By anol 20 Aug 2014
    picAccList.Left = 8500
    picAccList.Top = 750
    picAccList.Visible = True
    tabManagementFee.Enabled = False
    picAccList.ZOrder 0
    txtAccountSearch(4).SetFocus
End Sub
Private Sub flxPurchase_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection

   If flxPurchase.TextMatrix(flxPurchase.row, 0) = "" Then Exit Sub
   If flxPurchase.RowHeight(flxPurchase.row) = 0 Then
      iPIEdit = 0
      Exit Sub
   End If

   adoconn.Open getConnectionString

'   HighLightRowFlxGrid flxPurchase, flxPurchase.row
   SelectFlxGridRowNocolor 1, flxPurchase, flxPurchase.row

   iPIEdit = flxPurchase.row

   ConfigFlxPurchaseSplit

   With flxPurchaseSplit
'         Adding the split of the header
        szSQL = "SELECT DISTINCT S.*, P.PropertyID, U.UnitNumber, " & _
                "P.PropertyName, U.UnitName,FundCode " & _
                "FROM ((tblPurInvSRec AS S " & _
                "LEFT JOIN  Property AS P ON S.TRANS = P.PropertyID) " & _
                "LEFT JOIN Units AS U ON S.UNIT_ID = U.UnitNumber) " & _
                "INNER JOIN Fund ON S.DEPT_ID = Fund.FundID " & _
                "WHERE S.ParentID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' " & _
                "ORDER BY TRAN_ID;"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = IIf(IsNull(adoInvSp.Fields.Item("UnitNumber").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyID").Value), "", _
                                 adoInvSp.Fields.Item("PropertyID").Value), adoInvSp.Fields.Item("UnitNumber").Value)
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("UnitName").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyName").Value), "", _
                                 adoInvSp.Fields.Item("PropertyName").Value), adoInvSp.Fields.Item("UnitName").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("FundCode").Value), "", adoInvSp.Fields.Item("FundCode").Value)
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 9) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 10) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("RecoverablePt").Value & "%"

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With
   
   lblPurchaseSplit(17).Caption = " Purchase Invoice Details " & flxPurchase.TextMatrix(flxPurchase.row, 2)
   Call ConfigflxReceipt
   With flxReceipt
    iRow = 0
'        Showing Receipt Informations
'      szSQL = "SELECT (MID(CONSTANT,4) & R.Slnumber) as NO,SplitID,SageAccountNumber,PropertyID,FundName,Ref,S.DESCRIPTION,S.amount as amount from tlbReceipt R, tlbReceiptSplit S,tlbTransactionTypes T, Fund F WHERE T.TYPE_ID=R.Type AND " & _
'              " S.rptHeader=R.TransactionID and F.FundID=S.FundID and S.PIRefMgtFees = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' "
 szSQL = "SELECT (MID(CONSTANT,4) & M.SRSLNumber) as NO,PI_ActualID,SRSLNumber,ReceiptSplitID,ReceiptType,ChargingMethod,SageAccountNumber,ReceiptTypeDescription, " & _
         "PropertyID,FundCode,ChargeDate,ReceiptDate,ReceiptTransactionID,ReceiptSplitID,ReceiptAmount,AgrPercentage,MgtFeeAmt,VATPercentage,VAT,MgtFeeAmtTotal " & _
         "from ManagementFee M,tlbTransactionTypes T,Fund F WHERE T.TYPE_ID=M.ReceiptType AND " & _
         "F.FundID=M.FundID and M.PI_ActualID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' "

      adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
      lblPurchaseSplit(5).Caption = " Receipt Details for " & flxPurchase.TextMatrix(flxPurchase.row, 2) & ""

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = ""
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("PI_ActualID").Value
         .TextMatrix(iRow, 2) = adoInvSp.Fields.Item("NO").Value
         .TextMatrix(iRow, 3) = adoInvSp.Fields.Item("ReceiptSplitID").Value
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("ReceiptType").Value
         .TextMatrix(iRow, 5) = adoInvSp.Fields.Item("ChargingMethod").Value
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("SageAccountNumber").Value), "", adoInvSp.Fields.Item("SageAccountNumber").Value)
         .TextMatrix(iRow, 7) = IIf(IsNull(adoInvSp.Fields.Item("ReceiptTypeDescription").Value), "", adoInvSp.Fields.Item("ReceiptTypeDescription").Value) 'adoInvSp.Fields.Item("ReceiptTypeDescription").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoInvSp.Fields.Item("PropertyID").Value), "", adoInvSp.Fields.Item("PropertyID").Value)
         .TextMatrix(iRow, 9) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 10) = adoInvSp.Fields.Item("ChargeDate").Value
         .TextMatrix(iRow, 11) = IIf(IsNull(adoInvSp.Fields.Item("ReceiptDate").Value), "", adoInvSp.Fields.Item("ReceiptDate").Value) 'adoInvSp.Fields.Item("ReceiptDate").Value
         '.TextMatrix(iRow, 12) = adoInvSp.Fields.Item("ReceiptSplitID").Value
         .TextMatrix(iRow, 13) = Format(adoInvSp.Fields.Item("ReceiptAmount").Value, "0.00")
         .TextMatrix(iRow, 14) = Format(adoInvSp.Fields.Item("AgrPercentage").Value, "0.00") & "%"
         .TextMatrix(iRow, 15) = Format(adoInvSp.Fields.Item("MgtFeeAmt").Value, "0.00")
         .TextMatrix(iRow, 16) = Format(adoInvSp.Fields.Item("VATPercentage").Value, "0.00")
         .TextMatrix(iRow, 17) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 18) = Format(adoInvSp.Fields.Item("MgtFeeAmtTotal").Value, "0.00")
         
         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With
   
   
   adoconn.Close
   Set adoInvSp = Nothing
   Set adoconn = Nothing
End Sub
Private Function SelectedPurInvIDArr(ByRef SelPurID() As String) As String
'This function is for storing Purchase Invoice ID into an array written by anol 20181113
   Dim i As Integer
   Dim j As Long

  
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelectedPurInvIDArr = SelectedPurInvIDArr & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         SelectedPurInvIDArr = SelectedPurInvIDArr & ","
         j = j + 1
      End If
   Next i
   
   ReDim SelPurID(j)
   j = 0
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelPurID(j) = "'" & CStr(flxPurchase.TextMatrix(i, 0)) & "'"
         j = j + 1
      End If
   Next i
   
   If Len(SelectedPurInvIDArr) > 0 Then
      SelectedPurInvIDArr = Left(SelectedPurInvIDArr, Len(SelectedPurInvIDArr) - 1)
   End If
End Function
 Private Function ReturnString(i As Long, j As Long, ByRef SelPurHisID() As String) As String
    On Error GoTo Err
    For j = i To j
        If SelPurHisID(j) = "" Then
                Exit For
        End If
        ReturnString = ReturnString & SelPurHisID(j)
        ReturnString = ReturnString & ","
    Next j
Err:
    If Len(ReturnString) > 0 Then
        ReturnString = Left(ReturnString, Len(ReturnString) - 1)
    End If
 End Function
Public Sub PostInvoice()
   Dim szPI_ID As String
   Dim szSQL As String
   Dim iPosted As Integer               'Finally posted
   Dim iIP     As Integer               'To be posted
   Dim adoconn As New ADODB.Connection
   Dim rsPI As New ADODB.Recordset
   Dim SelPurID() As String
   Dim j As Long
   Dim K As Integer

   adoconn.Open getConnectionString

   If frmPopUpMenu.optSelPO.Value Then
            szPI_ID = SelectedPurInvIDArr(SelPurID())
            j = UBound(SelPurID())
            If j = 0 Then
                    MsgBox "Please select a purchase invoice to post in history.", vbCritical + vbOKOnly, "Purchase invoice"
                    Exit Sub
            End If
            K = CInt(j / 50)
            If K = j / 50 Then
                 'No no need to do ceiling, this is fully divisible
                 K = j / 50
            Else
                 K = CInt(j / 50) + 1 'This is ceiling function
            End If
            For K = 0 To K - 1
                szPI_ID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurID())
                If szPI_ID = "" Then
                     Exit For
                End If
                If Trim(szPI_ID) <> "" Then
                    szSQL = "UPDATE tblPurInv " & _
                        "SET History = TRUE " & _
                        "WHERE MY_ID IN (" & szPI_ID & ");"
                     adoconn.Execute szSQL
                End If
            Next
            szPI_ID = ""
            GoTo XX
   End If
   If frmPopUpMenu.optPIDtRange.Value Then
            szPI_ID = DateRangePurInvID(CDate(frmPopUpMenu.txtDtRangeFrom.text), CDate(frmPopUpMenu.txtDtRangeTo.text), SelPurID())
            j = UBound(SelPurID())
            If j = 0 Then
                    MsgBox "Please select a purchase invoice to post in history.", vbCritical + vbOKOnly, "Purchase invoice"
                    Exit Sub
            End If
            K = CInt(j / 50)
            If K = j / 50 Then
                 'No no need to do ceiling, this is fully divisible
                 K = j / 50
            Else
                 K = CInt(j / 50) + 1 'This is ceiling function
            End If
            For K = 0 To K - 1
                szPI_ID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurID())
                If szPI_ID = "" Then
                     Exit For
                End If
                If Trim(szPI_ID) <> "" Then
                  szSQL = "UPDATE tblPurInv " & _
                    "SET History = TRUE " & _
                    "WHERE MY_ID IN (" & szPI_ID & ");"
            
                    adoconn.Execute szSQL
                End If
            Next

            szPI_ID = ""
            GoTo XX
        
   End If
   If frmPopUpMenu.optFP_PI.Value Then
        'szPI_ID = FullyPaidPurInvID(adoConn)
        szSQL = "Select P.* From tlbPayment AS P INNER JOIN tblPurInv AS I ON P.PI = I.MY_ID WHERE P.OSAmount = 0 AND (I.TransactionType = 6 Or I.TransactionType = 7) AND I.History = FALSE"
        rsPI.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        j = rsPI.RecordCount
        rsPI.Close
        Set rsPI = Nothing
        adoconn.Execute "Update tlbPayment AS P INNER JOIN tblPurInv AS I ON P.PI = I.MY_ID SET I.History = TRUE WHERE P.OSAmount = 0 AND (I.TransactionType = 6 Or I.TransactionType = 7)  AND I.History = FALSE"
   End If
   If frmPopUpMenu.optSlNoRange.Value Then
        
        szSQL = "Select P.* From tblPurInv AS P  WHERE slNumber>=" & StrDigitVal(frmPopUpMenu.txtPlRangeFrom.text) & " and  slNumber<=" & StrDigitVal(frmPopUpMenu.txtPlRangeTo.text) & "" & _
        " AND TransactionType=" & IIf(UCase(Left(frmPopUpMenu.txtPlRangeTo.text, 2)) = "PI", 6, 7) & " "
        rsPI.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        j = rsPI.RecordCount
        rsPI.Close
        Set rsPI = Nothing
        adoconn.Execute "Update tblPurInv Set History = TRUE WHERE slNumber>=" & StrDigitVal(frmPopUpMenu.txtPlRangeFrom.text) & " and  slNumber<=" & StrDigitVal(frmPopUpMenu.txtPlRangeTo.text) & "" & _
        " AND TransactionType=" & IIf(UCase(Left(frmPopUpMenu.txtPlRangeTo.text, 2)) = "PI", 6, 7) & " "
       ' szPI_ID = SlNoRangePurInv(frmPopUpMenu.txtPlRangeFrom.text, frmPopUpMenu.txtPlRangeTo.text)
            
   End If

'   If Len(szPI_ID) = 0 Then
'      ShowMsgInTaskBar "No invoice to post to history.", "Y", "N"
'
'      adoConn.Close
'      Set adoConn = Nothing
'      Exit Sub
'   End If
'
'   szSQL = "UPDATE tblPurInv " & _
'           "SET History = TRUE " & _
'           "WHERE MY_ID IN (" & szPI_ID & ");"
'
'   adoConn.Execute szSQL
XX:
   LoadFlxPurchase adoconn
   fmeLoading.Visible = False
   MousePointer = vbDefault

   adoconn.Close
   Set adoconn = Nothing
   ShowMsgInTaskBar "System has posted " & j & " invoice to history.", "Y", "P"
End Sub
Private Function SelectFlxGridRowNocolor(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
        ' conFlxGrid.CellBackColor = RGB(255, 255, 255)
      'Next iRow
      SelectFlxGridRowNocolor = -1
   Else
      conFlxGrid.TextMatrix(iSelRow, iColID) = "X"

      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
         'conFlxGrid.CellBackColor = RGB(174, 179, 233)
      'Next iRow
      SelectFlxGridRowNocolor = 1
   End If
End Function
Private Function DateRangePurInvID(dtFrom As Date, dtTo As Date, SelPurID() As String) As String
   Dim i As Integer
   Dim j As Long

   DateRangePurInvID = ""
   For i = 1 To flxPurchase.Rows - 1
      If CDate(flxPurchase.TextMatrix(i, 4)) >= dtFrom And CDate(flxPurchase.TextMatrix(i, 4)) <= dtTo Then
         DateRangePurInvID = DateRangePurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         DateRangePurInvID = DateRangePurInvID & ","
          j = j + 1
      End If
   Next i
   ReDim SelPurID(j)
   j = 0
   For i = 1 To flxPurchase.Rows - 1
      If CDate(flxPurchase.TextMatrix(i, 4)) >= dtFrom And CDate(flxPurchase.TextMatrix(i, 4)) <= dtTo Then
          SelPurID(j) = "'" & CStr(flxPurchase.TextMatrix(i, 0)) & "'"
          j = j + 1
      End If
   Next i
   If Len(DateRangePurInvID) > 0 Then
      DateRangePurInvID = Left(DateRangePurInvID, Len(DateRangePurInvID) - 1)
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm frmManagementFees
End Sub

Private Sub tabManagementFee_Click(PreviousTab As Integer)
     Dim adoconn As New ADODB.Connection
  
    If tabManagementFee.Tab = 1 Then
         adoconn.Open getConnectionString
         LoadFlxPurchHistory adoconn, ""
         adoconn.Close
         Set adoconn = Nothing
    End If
   
End Sub
Private Sub ConfigFlxPurchHeader(ctrHeader As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrHeader.Clear
   ctrHeader.Cols = 15
   ctrHeader.Rows = 2
   ctrHeader.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID"

   ctrHeader.FormatString = szHeader$
   ctrHeader.ColWidth(0) = 0
   ctrHeader.ColWidth(1) = Label20(1 + iLabel).Left - ctrHeader.Left
   For iCol = 2 To ctrHeader.Cols - 7
      ctrHeader.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrHeader.ColWidth(iCol) = ctrHeader.Width + ctrHeader.Left - Label20(iCol - 1 + iLabel).Left - 340
   ctrHeader.ColWidth(iCol + 1) = 0
   ctrHeader.ColWidth(iCol + 2) = 0
   ctrHeader.ColWidth(iCol + 3) = 0                   'OS Amt
   ctrHeader.ColWidth(iCol + 4) = 0                   'Due Date
   ctrHeader.ColWidth(iCol + 5) = 0                   'Client ID
End Sub
Private Sub flxPurchHistory_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection

   If flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) = "" Then Exit Sub
   If flxPurchHistory.RowHeight(flxPurchHistory.row) = 0 Then Exit Sub
   'below line has been added by anol
   'Date 10 Jun 2015
   'issue 0000572: Purchases and expenses - Reverse postings to history not present
   Call SelectFlxGridRow(1, flxPurchHistory, flxPurchHistory.RowSel)
   adoconn.Open getConnectionString

'   HighLightRowFlxGrid flxPurchHistory, flxPurchHistory.row

   ConfigFlxSplit flxPurchHistorySplit, 29

   With flxPurchHistorySplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyName AS XX " & _
              "FROM tblPurInvSRec AS S LEFT JOIN Property AS P ON S.TRANS = P.PropertyID " & _
              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' " & _
              "ORDER BY TRAN_ID;"
'
'      szSQL = szSQL + " UNION "
'
'      szSQL = szSQL + _
'              "SELECT DISTINCT S.*, U.UnitName AS XX " & _
'              "FROM tblPurInvSRec AS S, Units AS U " & _
'              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' AND " & _
'                  "S.TRANS = 'Unit' AND S.UNIT_ID = U.UnitNumber " & _
'              "ORDER BY TRAN_ID"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("XX").Value), "", adoInvSp.Fields.Item("XX").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = adoInvSp.Fields.Item("DEPT_ID").Value
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = adoInvSp.Fields.Item("NET_AMOUNT").Value
         .TextMatrix(iRow, 9) = adoInvSp.Fields.Item("VAT").Value
         .TextMatrix(iRow, 10) = adoInvSp.Fields.Item("TOTAL_AMOUNT").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoconn.Close
   Set adoInvSp = Nothing
   Set adoconn = Nothing
End Sub
Private Sub ConfigFlxSplit(ctrSplit As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrSplit.Clear
   ctrSplit.Cols = 11
   ctrSplit.Rows = 2
   ctrSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"
   ctrSplit.FormatString = szHeader$

   ctrSplit.ColWidth(0) = 0
   ctrSplit.ColWidth(1) = Label20(1 + iLabel).Left - ctrSplit.Left

   For iCol = 2 To ctrSplit.Cols - 2
      ctrSplit.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrSplit.ColWidth(iCol) = ctrSplit.Width + ctrSplit.Left - Label20(iCol - 1 + iLabel).Left - 340
End Sub
Private Sub LoadFlxPurchHistory(adoconn As ADODB.Connection, Filter As String) 'Load purchase history
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWhereProperty As String
   Dim strTopWhere As String
   Dim strWhere As String
   Dim strWhereSupplier As String
   Dim tempstr As String
   fmeLoading.Visible = True
   fmeLoading.Refresh
   ConfigFlxPurchHeader flxPurchHistory, 0
   ConfigFlxSplit flxPurchHistorySplit, 29
'   If Filter = "3" Then
'         If txtSearchFromD.text <> "" Then
'           strWhere = " AND PI.TRAN_DATE =#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "#"
'            If Len(txtSearchFromD.text) > 0 Then
'                 cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'            Else
'                 cmdSearchPurchaseHistory.Caption = "Sea&rch"
'            End If
'        End If
'    End If
    If txtSupplierSearc.text <> "ALL" Then
        strWhereSupplier = " AND Supplier.SupplierID='" & txtSupplierSearc.text & "'"
    End If
    If txtPropertyIDHist.text = "ALL" Then
    Else
        strWhereProperty = " AND PI.PropertyID='" & txtPropertyIDHist.text & "'"
    End If
    If Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND PI.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND PI.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#"
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
            Else
                 cmdSearchPurchaseHistory.Caption = "Sea&rch"
            End If
        End If
   End If
   If Filter = "" And Val(txtDisplayMaxPurchaseHist.text) > 0 Then
        strTopWhere = " Top " & txtDisplayMaxPurchaseHist.text
   End If
'   If txtClientIdlist.text = "ALL" Then
'        szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
'                    "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
'                    "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
'                "FROM ((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
'                    "INNER JOIN tblPurInvSRec AS S ON PI.MY_ID = S.ParentID) " & _
'                    "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
'                "WHERE History = YES " & strWhere & " AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
'                "ORDER BY PI.MY_ID DESC;"
'   Else
'        szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
'                    "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
'                    "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & PI.SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
'                "FROM ((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
'                    "INNER JOIN tblPurInvSRec AS S ON PI.MY_ID = S.ParentID) " & _
'                    "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
'                "WHERE History = YES " & strWhere & " AND PI.CL_ID='" & txtClientIdlist.text & "' AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
'                "ORDER BY PI.MY_ID DESC;"
'   End If
    ' I am removing purchase split table that  iss needed for showing zero values in the invoice issue 520
    If txtClientIdlist.text = "ALL" Then
         szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
                     "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
                     "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
                 "FROM (tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
                     "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
                 "WHERE PI.isManagementFee=true AND History = true " & strWhereProperty & strWhere & strWhereSupplier & " AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
                 "ORDER BY PI.transactionType ASC,PI.slnumber DESC;"
    Else
         szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
                     "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
                     "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & PI.SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
                 "FROM (tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
                     "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
                 "WHERE PI.isManagementFee=true AND History = true " & strWhereProperty & strWhere & strWhereSupplier & " AND PI.CL_ID='" & txtClientIdlist.text & "' AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
                 "ORDER BY PI.transactionType ASC,PI.slnumber DESC;"
    End If
'   Debug.Print szSQL
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoInv.Filter = "INVPur Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoInv.Filter = "INV_NO Like '%" & tempstr & "%'"
        End If
    End If
   adoInv.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   iKount = 1
   With flxPurchHistory
      .Rows = adoInv.RecordCount + 1
      While Not adoInv.EOF
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVPur").Value 'invoice number '''adoInv.Fields.Item("pf").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value) ' this is reference
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("CL_ID").Value), "", adoInv.Fields.Item("CL_ID").Value)
'######################################################################################################################
''         Adding the split of the header
'         szSQL = "SELECT DISTINCT * " & _
'                 "FROM tblPurInvSRec " & _
'                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
'                 "ORDER BY TRAN_ID;"
''Debug.Print szSQL
'         adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'         If Not adoInvSp.EOF Then
            .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
'         adoInvSp.Close

         adoInv.MoveNext
         iKount = iKount + 1
         'If Not adoInv.EOF Then .AddItem ""
      Wend
   End With

   adoInv.Close
   Set adoInv = Nothing
   fmeLoading.Visible = False
End Sub

Private Sub txtSearchClientID_Change()
    If bSearchClientNameFocus = False Then
          txtSearchClientName.text = ""
          Dim tempstr As String
'          If sTextBox = "BankAcPay" Then
'             If Trim(txtSearchClientID.text) = "" Then
'                  Call LoadflxBankAC("")
'                  Exit Sub
'             End If
'             tempstr = txtSearchClientID.text
'             tempstr = Replace(tempstr, "'", "''")
'             'Call LoadflxBankAC("tlbClientBanks.NominalCode Like '%" & tempstr & "%'")
'             Call LoadflxBankAC("BNC Like '%" & tempstr & "%'")
'             Exit Sub
'         End If
        
         If Trim(txtSearchClientID.text) = "" Then
              Call LoadflxClient("")
              Exit Sub
         End If
         tempstr = txtSearchClientID.text
         tempstr = Replace(tempstr, "'", "''")
         Call LoadflxClient("ClientID Like '%" & tempstr & "%'")
    End If
End Sub

Private Sub txtSearchClientID_GotFocus()
    bSearchClientNameFocus = False
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtSearchClientName
    End If
End Sub

Private Sub txtSearchClientName_Change()
   If bSearchClientNameFocus Then
        txtSearchClientID.text = ""
        Dim tempstr As String
'         If sTextBox = "BankAcPay" Then
'            If Trim(txtSearchClientName.text) = "" Then
'                 Call LoadflxBankAC("")
'                 Exit Sub
'            End If
'            tempstr = txtSearchClientName.text
'            tempstr = Replace(tempstr, "'", "''")
'            Call LoadflxBankAC("BNN Like '%" & tempstr & "%'")
'            Exit Sub
'        End If
        
        If Trim(txtSearchClientName.text) = "" Then
             Call LoadflxClient("")
             Exit Sub
        Else
             txtSearchClientID.text = ""
        End If
        tempstr = txtSearchClientName.text
        tempstr = Replace(tempstr, "'", "''")
        Call LoadflxClient("ClientName Like '%" & tempstr & "%'")
    End If
End Sub

Private Sub txtSearchClientName_GotFocus()
    bSearchClientNameFocus = True
End Sub
