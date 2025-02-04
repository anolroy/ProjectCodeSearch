VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCashbook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cashbook"
   ClientHeight    =   12045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16380
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12045
   ScaleWidth      =   16380
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   6210
      ScaleHeight     =   4470
      ScaleWidth      =   7470
      TabIndex        =   339
      Top             =   8190
      Visible         =   0   'False
      Width           =   7500
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3750
         Left            =   45
         TabIndex        =   15
         Top             =   675
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   6615
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
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   4590
         TabIndex        =   346
         Top             =   375
         Width           =   2835
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5001;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Left            =   4590
         TabIndex        =   345
         Top             =   135
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Caption"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   343
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   342
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   341
         Top             =   75
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1485
         TabIndex        =   340
         Top             =   90
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
         Left            =   135
         TabIndex        =   13
         Top             =   375
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1350
         TabIndex        =   14
         Top             =   375
         Width           =   3195
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5636;450"
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
         Left            =   45
         Top             =   75
         Width           =   7290
      End
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
      Left            =   4275
      TabIndex        =   0
      Top             =   90
      Width           =   300
   End
   Begin VB.PictureBox fmeLoading 
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
      Height          =   315
      Left            =   2205
      ScaleHeight     =   315
      ScaleWidth      =   2655
      TabIndex        =   336
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblLoading 
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
         Height          =   330
         Left            =   270
         TabIndex        =   337
         Top             =   45
         Width           =   2745
      End
   End
   Begin VB.ListBox lstBankStDates_ 
      Height          =   255
      ItemData        =   "frmCashbook.frx":1202
      Left            =   13680
      List            =   "frmCashbook.frx":1209
      TabIndex        =   335
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Transaction"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   334
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1995
      Index           =   0
      Left            =   11040
      TabIndex        =   318
      Top             =   9240
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optCbHRpt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Receipts && Payments"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   321
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optCbHRpt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All Receipts"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   319
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optCbHRpt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All Payments"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   320
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdCbHRptOk 
         Caption         =   "&OK"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   322
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCbHRptCancel 
         Caption         =   "&Cancel"
         Height          =   365
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   323
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   1215
         Index           =   4
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Height          =   1215
         Index           =   5
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   324
         Top             =   30
         Width           =   1110
      End
   End
   Begin VB.Frame fraBank 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1635
      Left            =   120
      TabIndex        =   161
      Top             =   9480
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdBankCancel 
         Caption         =   "&Cancel"
         Height          =   365
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBankOK 
         Caption         =   "&OK"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optPay 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bank Payment (BP)"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   163
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optReceipt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bank Receipt (BR)"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   162
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   18
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   166
         Top             =   30
         Width           =   1110
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Height          =   855
         Index           =   19
         Left            =   120
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1635
      Index           =   1
      Left            =   9600
      TabIndex        =   155
      Top             =   9480
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdAutoAllocSelCancel 
         Caption         =   "&Cancel"
         Height          =   365
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAutoAllocSel 
         Caption         =   "&OK"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optRIF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "The Recent invoices first"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   157
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optOIF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "The Oldest invoice first"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   8
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   160
         Top             =   30
         Width           =   1110
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Height          =   855
         Index           =   7
         Left            =   120
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox picLeaseList 
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
      Height          =   2535
      Left            =   3000
      ScaleHeight     =   2505
      ScaleWidth      =   6345
      TabIndex        =   141
      Top             =   8880
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   146
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            Height          =   195
            Index           =   4
            Left            =   3000
            TabIndex        =   150
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   149
            Top             =   0
            Width           =   465
         End
         Begin MSForms.ComboBox cboSrcProp 
            Height          =   315
            Index           =   0
            Left            =   3675
            TabIndex        =   148
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
         Begin MSForms.ComboBox cboSrcClient 
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   147
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
      Begin VB.TextBox txtTenantSearchUnitName 
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
         TabIndex        =   145
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox txtTenantSearchName 
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
         TabIndex        =   144
         Top             =   300
         Width           =   2460
      End
      Begin VB.TextBox txtTenantSearchID 
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
         Height          =   330
         Left            =   75
         TabIndex        =   143
         Top             =   300
         Width           =   1470
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
         Index           =   0
         Left            =   6080
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
         Height          =   1815
         Left            =   45
         TabIndex        =   151
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3201
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
         Caption         =   "Tenant ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   154
         Top             =   70
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Name"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   153
         Top             =   70
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   152
         Top             =   75
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   45
         Top             =   80
         Width           =   6015
      End
   End
   Begin VB.TextBox txtAccountName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   136
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox txtAcBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11745
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   120
      Width           =   1170
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   11340
      TabIndex        =   33
      Top             =   7800
      Width           =   1575
   End
   Begin TabDlg.SSTab tabCashbook 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Account Details"
      TabPicture(0)   =   "frmCashbook.frx":122A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape1(5)"
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(4)=   "Label1(7)"
      Tab(0).Control(5)=   "Label1(8)"
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(8)=   "Label1(11)"
      Tab(0).Control(9)=   "Label1(12)"
      Tab(0).Control(10)=   "Label1(13)"
      Tab(0).Control(11)=   "Shape1(0)"
      Tab(0).Control(12)=   "Shape1(1)"
      Tab(0).Control(13)=   "Shape1(3)"
      Tab(0).Control(14)=   "Label1(15)"
      Tab(0).Control(15)=   "Label1(16)"
      Tab(0).Control(16)=   "Label1(17)"
      Tab(0).Control(17)=   "Label1(18)"
      Tab(0).Control(18)=   "Label1(19)"
      Tab(0).Control(19)=   "Label1(20)"
      Tab(0).Control(20)=   "txtBankName"
      Tab(0).Control(21)=   "txtSortCode"
      Tab(0).Control(22)=   "txtAcNo"
      Tab(0).Control(23)=   "txtAcName"
      Tab(0).Control(24)=   "txtAcType"
      Tab(0).Control(25)=   "txtNC"
      Tab(0).Control(26)=   "txtNN"
      Tab(0).Control(27)=   "txtPaymentMethod"
      Tab(0).Control(28)=   "txtBACS"
      Tab(0).Control(29)=   "txtLastReconDt"
      Tab(0).Control(30)=   "txtWebsite"
      Tab(0).Control(31)=   "txteMail"
      Tab(0).Control(32)=   "txtFax"
      Tab(0).Control(33)=   "txtTel"
      Tab(0).Control(34)=   "txtContact"
      Tab(0).Control(35)=   "cmdContactEdit"
      Tab(0).Control(36)=   "cmdContactSave"
      Tab(0).Control(37)=   "cmdContactCancel"
      Tab(0).Control(38)=   "txtMobile"
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Payments && Receipts"
      TabPicture(1)   =   "frmCashbook.frx":1246
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBankTransfer"
      Tab(1).Control(1)=   "cmdBRP"
      Tab(1).Control(2)=   "cmdSupplierPayment"
      Tab(1).Control(3)=   "cmdTenantReceipt"
      Tab(1).Control(4)=   "tabPayRpt"
      Tab(1).Control(5)=   "Shape4(3)"
      Tab(1).Control(6)=   "Shape4(2)"
      Tab(1).Control(7)=   "Shape4(1)"
      Tab(1).Control(8)=   "Shape4(0)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Cashbook history"
      TabPicture(2)   =   "frmCashbook.frx":1262
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(35)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(36)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(37)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(38)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(39)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(40)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(41)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(42)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1(43)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Shape2(4)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label1(45)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label1(44)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label1(48)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1(49)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label1(52)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label1(28)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label1(53)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "flxCashBook"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtCBHDtFrm"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtCBHDtTo"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "cmdCBHFilter"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "cmdCbHReport"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Bank reconciliation"
      TabPicture(3)   =   "frmCashbook.frx":127E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command2"
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(2)=   "cmdHistoryReport"
      Tab(3).Control(3)=   "cmdAddTrans"
      Tab(3).Control(4)=   "cmdReconcile"
      Tab(3).Control(5)=   "cmdReconPrint"
      Tab(3).Control(6)=   "cmdReconcileAll"
      Tab(3).Control(7)=   "cmdReconAll"
      Tab(3).Control(8)=   "cmdReconSave"
      Tab(3).Control(9)=   "txtStValue"
      Tab(3).Control(10)=   "txtStOpenBal"
      Tab(3).Control(11)=   "txtProjClosingBal"
      Tab(3).Control(12)=   "txtStatementDate"
      Tab(3).Control(13)=   "optReconciliation(1)"
      Tab(3).Control(14)=   "optReconciliation(0)"
      Tab(3).Control(15)=   "flxStatementReconcile"
      Tab(3).Control(16)=   "Label1(51)"
      Tab(3).Control(17)=   "Label1(50)"
      Tab(3).Control(18)=   "Label1(47)"
      Tab(3).Control(19)=   "Label1(46)"
      Tab(3).Control(20)=   "lblBankRec(16)"
      Tab(3).Control(21)=   "lblBankRec(14)"
      Tab(3).Control(22)=   "lblClosingBalance"
      Tab(3).Control(23)=   "Label1(27)"
      Tab(3).Control(24)=   "Label1(26)"
      Tab(3).Control(25)=   "Label1(25)"
      Tab(3).Control(26)=   "Shape2(3)"
      Tab(3).Control(27)=   "Shape2(1)"
      Tab(3).Control(28)=   "lblBankRec(19)"
      Tab(3).Control(29)=   "lblBankRec(18)"
      Tab(3).Control(30)=   "lblBankRec(21)"
      Tab(3).Control(31)=   "lblBankRec(20)"
      Tab(3).Control(32)=   "lblBankRec(17)"
      Tab(3).Control(33)=   "lblBankRec(15)"
      Tab(3).Control(34)=   "lblBankRec(13)"
      Tab(3).Control(35)=   "Label1(24)"
      Tab(3).Control(36)=   "Label1(23)"
      Tab(3).Control(37)=   "Label1(21)"
      Tab(3).Control(38)=   "Shape2(0)"
      Tab(3).Control(39)=   "Shape2(2)"
      Tab(3).ControlCount=   40
      TabCaption(4)   =   "Memo && attachments"
      TabPicture(4)   =   "frmCashbook.frx":129A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtNote"
      Tab(4).Control(1)=   "cmdUnitMemoEdit"
      Tab(4).Control(2)=   "cmdUnitMemoSave"
      Tab(4).Control(3)=   "cmdUnitMemoCancel"
      Tab(4).Control(4)=   "Frame8(2)"
      Tab(4).Control(5)=   "Shape1(4)"
      Tab(4).Control(6)=   "Shape1(2)"
      Tab(4).ControlCount=   7
      Begin VB.CommandButton Command2 
         Caption         =   "FIX"
         Height          =   330
         Left            =   -74865
         TabIndex        =   347
         Top             =   5985
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Automatic"
         Height          =   375
         Left            =   -66405
         TabIndex        =   128
         Top             =   6345
         Width           =   1035
      End
      Begin VB.CommandButton cmdHistoryReport 
         Caption         =   "&History Report"
         Height          =   375
         Left            =   -67710
         TabIndex        =   127
         Top             =   6345
         Width           =   1300
      End
      Begin VB.CommandButton cmdAddTrans 
         Caption         =   "Add &Transaction"
         Height          =   375
         Left            =   -69030
         TabIndex        =   126
         Top             =   6345
         Width           =   1300
      End
      Begin VB.CommandButton cmdReconcile 
         Caption         =   "Reconcile"
         Height          =   375
         Left            =   -73650
         TabIndex        =   122
         Top             =   6345
         Width           =   1185
      End
      Begin VB.CommandButton cmdReconPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   -70080
         TabIndex        =   125
         Top             =   6345
         Width           =   1035
      End
      Begin VB.CommandButton cmdCbHReport 
         Caption         =   "Report &Print"
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
         Left            =   240
         TabIndex        =   317
         Top             =   6480
         Width           =   1300
      End
      Begin VB.CommandButton cmdCBHFilter 
         Caption         =   "Filter >>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   316
         Top             =   480
         Width           =   1300
      End
      Begin VB.TextBox txtCBHDtTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   315
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtCBHDtFrm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   314
         Text            =   "01/01/2000"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdBankTransfer 
         Caption         =   "Enter Bank Transfers"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71040
         TabIndex        =   310
         Top             =   5760
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CommandButton cmdBRP 
         Caption         =   "Enter Bank Receipts and Payments"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71040
         TabIndex        =   309
         Top             =   4125
         Width           =   4695
      End
      Begin VB.CommandButton cmdSupplierPayment 
         Caption         =   "Enter Supplier Payments"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71040
         TabIndex        =   308
         Top             =   2475
         Width           =   4695
      End
      Begin VB.CommandButton cmdTenantReceipt 
         BackColor       =   &H00808080&
         Caption         =   "Enter Tenant Receipts"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71040
         TabIndex        =   307
         Top             =   840
         Width           =   4695
      End
      Begin VB.CommandButton cmdReconcileAll 
         Caption         =   "Reconcile &All"
         Height          =   375
         Left            =   -72480
         TabIndex        =   123
         Top             =   6345
         Width           =   1185
      End
      Begin VB.CommandButton cmdReconAll 
         Caption         =   "&Reset All"
         Height          =   375
         Left            =   -74880
         TabIndex        =   121
         Top             =   6345
         Width           =   1185
      End
      Begin VB.CommandButton cmdReconSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   -71280
         TabIndex        =   124
         Top             =   6345
         Width           =   1185
      End
      Begin VB.TextBox txtStValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -68760
         TabIndex        =   110
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtStOpenBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -63585
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txtProjClosingBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -63585
         TabIndex        =   120
         Top             =   840
         Width           =   1170
      End
      Begin VB.TextBox txtStatementDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67170
         TabIndex        =   118
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optReconciliation 
         BackColor       =   &H00FFDFDF&
         Caption         =   "Show all transactions"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   113
         Top             =   800
         Width           =   1935
      End
      Begin VB.OptionButton optReconciliation 
         BackColor       =   &H00FFDFDF&
         Caption         =   "Show unreconciled transactions only"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   112
         Top             =   480
         Value           =   -1  'True
         Width           =   3660
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4890
         Width           =   3255
      End
      Begin VB.CommandButton cmdContactCancel 
         BackColor       =   &H00000000&
         Caption         =   "Cancel"
         Enabled         =   0   'False
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
         Left            =   -65880
         TabIndex        =   31
         Top             =   6015
         Width           =   1575
      End
      Begin VB.CommandButton cmdContactSave 
         BackColor       =   &H00000000&
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   -65880
         TabIndex        =   30
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmdContactEdit 
         Caption         =   "Edit"
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
         Left            =   -65880
         TabIndex        =   29
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtContact 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4485
         Width           =   3255
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   5295
         Width           =   3255
      End
      Begin VB.TextBox txteMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   5700
         Width           =   3255
      End
      Begin VB.TextBox txtWebsite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   6105
         Width           =   3255
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   3375
         Left            =   -74520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   104
         Top             =   960
         Width           =   11595
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         Height          =   435
         Left            =   -67560
         TabIndex        =   103
         Top             =   4500
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -65940
         TabIndex        =   102
         Top             =   4500
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -64320
         TabIndex        =   101
         Top             =   4500
         Width           =   1350
      End
      Begin VB.Frame Frame8 
         Caption         =   "Attachment Files:"
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
         Height          =   855
         Index           =   2
         Left            =   -74520
         TabIndex        =   96
         Top             =   5040
         Width           =   11535
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   435
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   435
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   435
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   100
            Top             =   360
            Width           =   4890
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8625;503"
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.TextBox txtLastReconDt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2985
         Width           =   2895
      End
      Begin VB.TextBox txtBACS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2475
         Width           =   2895
      End
      Begin VB.TextBox txtPaymentMethod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1980
         Width           =   2895
      End
      Begin VB.TextBox txtNN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1470
         Width           =   2895
      End
      Begin VB.TextBox txtNC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66600
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtAcType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2985
         Width           =   3255
      End
      Begin VB.TextBox txtAcName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2475
         Width           =   3255
      End
      Begin VB.TextBox txtAcNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1980
         Width           =   3255
      End
      Begin VB.TextBox txtSortCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1470
         Width           =   3255
      End
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin TabDlg.SSTab tabPayRpt 
         Height          =   7575
         Left            =   -64200
         TabIndex        =   41
         Top             =   1080
         Visible         =   0   'False
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   13361
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tenant Receipt"
         TabPicture(0)   =   "frmCashbook.frx":12B6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Line1(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblAllocating(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label3(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label10(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Line1(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label10(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Line1(1)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label10(3)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label10(4)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label3(2)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label19(10)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label19(11)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label19(12)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label19(13)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label19(14)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label19(15)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label19(16)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label19(17)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label19(18)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label19(19)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "flxTCrPoA"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "flxTReceipt"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtAllocatedDiff(0)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Frame5(4)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtCrReceipt"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtTReceipt"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Frame8(0)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "cmdRptAllocate"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "cmdTRClose"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Frame5(0)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).ControlCount=   31
         TabCaption(1)   =   "Supplier Payment"
         TabPicture(1)   =   "frmCashbook.frx":12D2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Line1(7)"
         Tab(1).Control(1)=   "Line1(4)"
         Tab(1).Control(2)=   "Line1(6)"
         Tab(1).Control(3)=   "lblAllocating(1)"
         Tab(1).Control(4)=   "Label3(5)"
         Tab(1).Control(5)=   "Label10(8)"
         Tab(1).Control(6)=   "Label10(9)"
         Tab(1).Control(7)=   "Line1(5)"
         Tab(1).Control(8)=   "Label3(4)"
         Tab(1).Control(9)=   "Label10(6)"
         Tab(1).Control(10)=   "Label10(7)"
         Tab(1).Control(11)=   "flxSCrPoA"
         Tab(1).Control(12)=   "flxSPayment"
         Tab(1).Control(13)=   "Frame8(1)"
         Tab(1).Control(14)=   "txtAllocatedDiff(1)"
         Tab(1).Control(15)=   "Frame5(1)"
         Tab(1).Control(16)=   "txtCrPayment"
         Tab(1).Control(17)=   "txtSPayment"
         Tab(1).Control(18)=   "cmdPayAllocate"
         Tab(1).Control(19)=   "cmdSPClose"
         Tab(1).Control(20)=   "Frame5(5)"
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "Bank Receipt and Payment"
         TabPicture(2)   =   "frmCashbook.frx":12EE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblBankRec(0)"
         Tab(2).Control(1)=   "lblBankRec(12)"
         Tab(2).Control(2)=   "lblBankRec(11)"
         Tab(2).Control(3)=   "lblBankRec(10)"
         Tab(2).Control(4)=   "lblBankRec(8)"
         Tab(2).Control(5)=   "lblBankRec(9)"
         Tab(2).Control(6)=   "lblBankRec(7)"
         Tab(2).Control(7)=   "lblBankRec(6)"
         Tab(2).Control(8)=   "lblBankRec(29)"
         Tab(2).Control(9)=   "lblBankRec(5)"
         Tab(2).Control(10)=   "lblBankRec(3)"
         Tab(2).Control(11)=   "lblBankRec(4)"
         Tab(2).Control(12)=   "Label3(0)"
         Tab(2).Control(13)=   "Label3(3)"
         Tab(2).Control(14)=   "flxBankPay(0)"
         Tab(2).Control(15)=   "cmdBankReceiptHistory"
         Tab(2).Control(16)=   "Frame5(2)"
         Tab(2).Control(17)=   "fraBkInput(0)"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "Bank Transfer"
         TabPicture(3)   =   "frmCashbook.frx":130A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5(6)"
         Tab(3).Control(1)=   "Frame5(3)"
         Tab(3).ControlCount=   2
         Begin VB.Frame fraBkInput 
            BackColor       =   &H00E0E0E0&
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
            Height          =   1815
            Index           =   0
            Left            =   -74640
            TabIndex        =   263
            Top             =   360
            Width           =   12495
            Begin VB.TextBox txtNCBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1680
               TabIndex        =   280
               Top             =   1440
               Width           =   1050
            End
            Begin VB.TextBox txtDeptBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7320
               TabIndex        =   279
               Top             =   120
               Width           =   1050
            End
            Begin VB.TextBox txtDateBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1680
               TabIndex        =   278
               Top             =   1116
               Width           =   3975
            End
            Begin VB.TextBox txtNetBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7320
               TabIndex        =   277
               Top             =   780
               Width           =   1300
            End
            Begin VB.TextBox txtDetailsBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7320
               MaxLength       =   254
               TabIndex        =   276
               Top             =   450
               Width           =   4395
            End
            Begin VB.TextBox txtVatBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7320
               Locked          =   -1  'True
               TabIndex        =   275
               Top             =   1110
               Width           =   1300
            End
            Begin VB.CommandButton cmdUpdateBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&OK"
               Height          =   375
               Index           =   0
               Left            =   10380
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   274
               Top             =   1350
               Width           =   1335
            End
            Begin VB.CommandButton cmdBkList 
               Caption         =   "..."
               Height          =   285
               Index           =   0
               Left            =   2760
               TabIndex        =   273
               Top             =   472
               Width           =   255
            End
            Begin VB.CommandButton cmdNCBk 
               Caption         =   "..."
               Height          =   285
               Index           =   0
               Left            =   2760
               TabIndex        =   272
               Top             =   1440
               Width           =   255
            End
            Begin VB.CommandButton cmdDeptBk 
               Caption         =   "..."
               Height          =   285
               Index           =   0
               Left            =   8385
               TabIndex        =   271
               Top             =   120
               Width           =   255
            End
            Begin VB.CommandButton cmdTaxListBk 
               Caption         =   "..."
               Height          =   285
               Index           =   0
               Left            =   8640
               TabIndex        =   270
               Top             =   1116
               Width           =   405
            End
            Begin VB.TextBox txtBkAc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1680
               TabIndex        =   269
               Top             =   472
               Width           =   1065
            End
            Begin VB.TextBox txtReference 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   268
               Top             =   794
               Width           =   3975
            End
            Begin VB.TextBox txtTotalBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7320
               TabIndex        =   267
               Text            =   "0.00"
               Top             =   1440
               Width           =   1305
            End
            Begin VB.TextBox txtNCNameBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   266
               Top             =   1440
               Width           =   2625
            End
            Begin VB.TextBox txtBkAcName 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   3030
               Locked          =   -1  'True
               TabIndex        =   265
               Top             =   472
               Width           =   2625
            End
            Begin VB.TextBox txtDeptBkName 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   8670
               Locked          =   -1  'True
               TabIndex        =   264
               Top             =   120
               Width           =   3045
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BankA/C:"
               Height          =   195
               Index           =   2
               Left            =   840
               TabIndex        =   291
               Top             =   472
               Width           =   630
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date:"
               Height          =   195
               Index           =   3
               Left            =   840
               TabIndex        =   290
               Top             =   1116
               Width           =   375
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fund:"
               Height          =   195
               Index           =   6
               Left            =   6600
               TabIndex        =   289
               Top             =   120
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N/C:"
               Height          =   195
               Index           =   7
               Left            =   840
               TabIndex        =   288
               Top             =   1440
               Width           =   315
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net:"
               Height          =   195
               Index           =   10
               Left            =   6600
               TabIndex        =   287
               Top             =   795
               Width           =   300
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   195
               Index           =   11
               Left            =   6600
               TabIndex        =   286
               Top             =   465
               Width           =   540
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VAT:"
               Height          =   195
               Index           =   12
               Left            =   6600
               TabIndex        =   285
               Top             =   1110
               Width           =   330
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Client"
               Height          =   195
               Index           =   5
               Left            =   840
               TabIndex        =   284
               Top             =   120
               Width           =   435
            End
            Begin MSForms.ComboBox cboBRPClient 
               Height          =   315
               Left            =   1680
               TabIndex        =   283
               Top             =   120
               Width           =   3975
               VariousPropertyBits=   1753237529
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "7011;556"
               TextColumn      =   2
               ColumnCount     =   8
               ListRows        =   0
               cColumnInfo     =   1
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1763"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference:"
               Height          =   195
               Index           =   1
               Left            =   840
               TabIndex        =   282
               Top             =   794
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
               Height          =   195
               Index           =   4
               Left            =   6600
               TabIndex        =   281
               Top             =   1440
               Width           =   390
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
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
            Height          =   4935
            Index           =   6
            Left            =   -74160
            TabIndex        =   234
            Top             =   960
            Width           =   11295
            Begin VB.ComboBox cboClientName 
               Height          =   315
               Left            =   1800
               TabIndex        =   223
               Top             =   240
               Width           =   3495
            End
            Begin VB.TextBox txtBkTrAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1800
               TabIndex        =   231
               Top             =   3990
               Width           =   1095
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   8
               Left            =   600
               TabIndex        =   252
               Top             =   240
               Width           =   480
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Client:"
               Size            =   "847;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   12
               Left            =   600
               TabIndex        =   251
               Top             =   720
               Width           =   675
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Client ID:"
               Size            =   "1191;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboFundBankTransf 
               Height          =   315
               Left            =   1800
               TabIndex        =   228
               Top             =   2550
               Width           =   3495
               VariousPropertyBits=   679495705
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "6165;556"
               TextColumn      =   2
               ColumnCount     =   2
               cColumnInfo     =   1
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "705"
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   5
               Left            =   5880
               TabIndex        =   250
               Top             =   240
               Width           =   390
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Date:"
               Size            =   "688;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   4
               Left            =   600
               TabIndex        =   249
               Top             =   3030
               Width           =   765
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Reference:"
               Size            =   "1349;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   9
               Left            =   600
               TabIndex        =   248
               Top             =   3510
               Width           =   885
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Description:"
               Size            =   "1561;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   10
               Left            =   600
               TabIndex        =   247
               Top             =   3990
               Width           =   450
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Value:"
               Size            =   "794;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   11
               Left            =   600
               TabIndex        =   246
               Top             =   2550
               Width           =   405
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Fund:"
               Size            =   "714;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtBkTrDate 
               Height          =   315
               Left            =   6480
               TabIndex        =   224
               Top             =   240
               Width           =   1215
               VariousPropertyBits=   746604571
               Size            =   "2143;556"
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtBkTrRef 
               Height          =   315
               Left            =   1800
               TabIndex        =   229
               Top             =   3000
               Width           =   3495
               VariousPropertyBits=   746604571
               Size            =   "6165;556"
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtBkTrDes 
               Height          =   345
               Left            =   1800
               TabIndex        =   230
               Top             =   3510
               Width           =   3525
               VariousPropertyBits=   -1400879077
               Size            =   "6218;609"
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   21
               Left            =   1800
               TabIndex        =   225
               Top             =   720
               Width           =   3495
               BackColor       =   16777215
               Size            =   "6165;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboANF 
               Bindings        =   "frmCashbook.frx":1326
               Height          =   315
               Left            =   1815
               TabIndex        =   226
               Top             =   1560
               Width           =   3495
               VariousPropertyBits=   679495705
               DisplayStyle    =   3
               Size            =   "6165;556"
               ColumnCount     =   2
               cColumnInfo     =   1
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1411"
            End
            Begin MSForms.Label Label13 
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   245
               Top             =   1560
               Width           =   1200
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Account From:"
               Size            =   "2117;450"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   6
               Left            =   5400
               TabIndex        =   244
               Top             =   1300
               Width           =   1005
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Nominal Code"
               Size            =   "1773;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   7
               Left            =   5400
               TabIndex        =   243
               Top             =   1560
               Width           =   1080
               BackColor       =   16777215
               Size            =   "1905;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   13
               Left            =   6480
               TabIndex        =   242
               Top             =   1560
               Width           =   1080
               BackColor       =   16777215
               Size            =   "1905;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   1
               Left            =   6480
               TabIndex        =   241
               Top             =   1300
               Width           =   720
               BackColor       =   16768960
               VariousPropertyBits=   276824083
               Caption         =   "Sort Code"
               Size            =   "1270;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   15
               Left            =   7560
               TabIndex        =   240
               Top             =   1560
               Width           =   3000
               BackColor       =   16777215
               Size            =   "5292;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   255
               Index           =   14
               Left            =   7560
               TabIndex        =   239
               Top             =   1300
               Width           =   1080
               BackColor       =   16768960
               VariousPropertyBits=   8388627
               Caption         =   "Account Name"
               Size            =   "1905;450"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboANT 
               Height          =   315
               Left            =   1815
               TabIndex        =   227
               Top             =   2040
               Width           =   3495
               VariousPropertyBits=   746604569
               DisplayStyle    =   3
               Size            =   "6165;556"
               ColumnCount     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   195
               Index           =   2
               Left            =   600
               TabIndex        =   238
               Top             =   2040
               Width           =   1335
               BackColor       =   16768960
               VariousPropertyBits=   8388627
               Caption         =   "Account To:"
               Size            =   "2355;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   20
               Left            =   6480
               TabIndex        =   237
               Top             =   2040
               Width           =   1080
               BackColor       =   16777215
               Size            =   "1905;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   19
               Left            =   7560
               TabIndex        =   236
               Top             =   2040
               Width           =   3000
               BackColor       =   16777215
               Size            =   "5292;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label13 
               Height          =   315
               Index           =   17
               Left            =   5400
               TabIndex        =   235
               Top             =   2040
               Width           =   1080
               BackColor       =   16777215
               Size            =   "1905;556"
               BorderStyle     =   1
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00C0FFFF&
               Height          =   1275
               Index           =   22
               Left            =   480
               Top             =   1200
               Width           =   10215
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Height          =   1275
               Index           =   23
               Left            =   480
               Top             =   1200
               Width           =   10215
            End
         End
         Begin VB.Frame Frame5 
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
            Height          =   915
            Index           =   3
            Left            =   -74160
            TabIndex        =   222
            Top             =   5880
            Width           =   11295
            Begin VB.CommandButton cmdBTSave 
               BackColor       =   &H00F0F0F0&
               Caption         =   "&Save"
               Height          =   400
               Left            =   4080
               Style           =   1  'Graphical
               TabIndex        =   232
               Top             =   330
               Width           =   1575
            End
            Begin VB.CommandButton cmdBTCancel 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Cancel"
               Height          =   400
               Left            =   5880
               Style           =   1  'Graphical
               TabIndex        =   233
               Top             =   330
               Width           =   1575
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Receipts:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   705
            Index           =   5
            Left            =   -74880
            TabIndex        =   200
            Top             =   6795
            Width           =   5415
            Begin VB.CommandButton cmdSPFull 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Pay in &Full"
               Height          =   400
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdSPayAll 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Pay &All"
               Height          =   400
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdSPSave 
               BackColor       =   &H00F0F0F0&
               Caption         =   "&Save"
               Enabled         =   0   'False
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdPaymentDiscard 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Clear"
               Height          =   400
               Left            =   4080
               Style           =   1  'Graphical
               TabIndex        =   178
               Top             =   225
               Width           =   1200
            End
         End
         Begin VB.CommandButton cmdSPClose 
            BackColor       =   &H00F0F0F0&
            Caption         =   "C&lose"
            Height          =   400
            Left            =   -63360
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   7020
            Width           =   1400
         End
         Begin VB.CommandButton cmdPayAllocate 
            BackColor       =   &H00F0F0F0&
            Caption         =   "All&ocation Only"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -66720
            Style           =   1  'Graphical
            TabIndex        =   199
            Top             =   7020
            Width           =   1700
         End
         Begin VB.TextBox txtSPayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   -67680
            TabIndex        =   198
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtCrPayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   -66120
            MaxLength       =   13
            TabIndex        =   197
            Top             =   5040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Allocation:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   705
            Index           =   1
            Left            =   -69600
            TabIndex        =   192
            Top             =   6720
            Visible         =   0   'False
            Width           =   3705
            Begin VB.CommandButton cmdPayAutomatic 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Automatic"
               Height          =   400
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   195
               Top             =   225
               Width           =   1080
            End
            Begin VB.CommandButton cmdPayAllocationDiscard 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Clear"
               Height          =   400
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   194
               Top             =   225
               Width           =   1080
            End
            Begin VB.CommandButton cmdPayAllocateSave 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Save"
               Enabled         =   0   'False
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   193
               Top             =   225
               Width           =   1080
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "allocation ref."
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   195
               Index           =   5
               Left            =   1440
               TabIndex        =   196
               Top             =   0
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.TextBox txtAllocatedDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   1
            Left            =   -64305
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "0.00"
            Top             =   6600
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   1
            Left            =   -74880
            TabIndex        =   167
            Top             =   360
            Width           =   12975
            Begin VB.Frame Frame8 
               BackColor       =   &H00DEDEDE&
               Caption         =   "Analysis:"
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
               Height          =   1215
               Index           =   3
               Left            =   10080
               TabIndex        =   180
               Top             =   0
               Width           =   2895
               Begin VB.TextBox txtDiffPay 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   183
                  Text            =   "0.00"
                  Top             =   855
                  Width           =   1215
               End
               Begin VB.TextBox txtPaymentEntered 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   182
                  Text            =   "0.00"
                  Top             =   527
                  Width           =   1215
               End
               Begin VB.TextBox txtPaymentTotal 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   181
                  Text            =   "0.00"
                  Top             =   200
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Difference                  £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   5
                  Left            =   120
                  TabIndex        =   186
                  Top             =   855
                  Width           =   1380
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Payment Entered £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   4
                  Left            =   120
                  TabIndex        =   185
                  Top             =   525
                  Width           =   1320
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Payment total        £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   3
                  Left            =   120
                  TabIndex        =   184
                  Top             =   195
                  Width           =   1290
               End
            End
            Begin VB.TextBox txtSPReference 
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
               Height          =   315
               Left            =   840
               MaxLength       =   12
               TabIndex        =   169
               Top             =   720
               Width           =   3195
            End
            Begin VB.TextBox txtSPaymentTotal 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   8760
               TabIndex        =   173
               Text            =   "0.00"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtSPDate 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   8760
               MaxLength       =   10
               TabIndex        =   174
               Top             =   720
               Width           =   1215
            End
            Begin MSForms.ComboBox cmbSPBankAc 
               Height          =   315
               Left            =   5160
               TabIndex        =   170
               Top             =   240
               Width           =   2535
               VariousPropertyBits=   1753237531
               DisplayStyle    =   3
               Size            =   "4471;556"
               BoundColumn     =   0
               TextColumn      =   2
               ColumnCount     =   3
               ListRows        =   20
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1058;3527"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank A/C"
               Height          =   195
               Index           =   2
               Left            =   4125
               TabIndex        =   211
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   210
               Top             =   240
               Width           =   600
            End
            Begin MSForms.ComboBox cmbSPSupplier 
               Height          =   315
               Left            =   855
               TabIndex        =   168
               Top             =   240
               Width           =   3195
               VariousPropertyBits=   1753237531
               DisplayStyle    =   3
               Size            =   "5636;556"
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
            Begin MSForms.CommandButton cmdSPAmtType 
               Height          =   315
               Left            =   7380
               TabIndex        =   172
               Top             =   720
               Width           =   315
               Caption         =   "- -"
               Size            =   "556;556"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.ComboBox cmbSPAmtType 
               Height          =   315
               Left            =   5160
               TabIndex        =   171
               Top             =   720
               Width           =   2205
               VariousPropertyBits=   1753237531
               DisplayStyle    =   3
               Size            =   "3889;556"
               BoundColumn     =   0
               TextColumn      =   2
               ColumnCount     =   3
               ListRows        =   20
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Type"
               Height          =   195
               Index           =   29
               Left            =   4125
               TabIndex        =   190
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference"
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   189
               Top             =   720
               Width           =   720
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Date"
               Height          =   195
               Index           =   7
               Left            =   7755
               TabIndex        =   188
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amt  £"
               Height          =   195
               Index           =   6
               Left            =   7755
               TabIndex        =   187
               Top             =   240
               Width           =   825
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
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
            Height          =   615
            Index           =   2
            Left            =   -74640
            TabIndex        =   89
            Top             =   6240
            Width           =   12615
            Begin VB.CommandButton cmdCancelBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               Height          =   400
               Index           =   0
               Left            =   3720
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdSaveBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               Height          =   400
               Index           =   0
               Left            =   7200
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   93
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdCloseBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C&lose"
               Height          =   400
               Index           =   0
               Left            =   10920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   92
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdNewBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Add"
               Height          =   400
               Index           =   0
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   91
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdEditBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Edit"
               Height          =   400
               Index           =   0
               Left            =   1920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   90
               Top             =   120
               Width           =   1450
            End
         End
         Begin VB.CommandButton cmdBankReceiptHistory 
            Appearance      =   0  'Flat
            Caption         =   "Bank Receipt &History"
            Height          =   420
            Left            =   -74760
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   6840
            Width           =   1580
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Receipts:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   705
            Index           =   0
            Left            =   120
            TabIndex        =   78
            Top             =   6800
            Width           =   5415
            Begin VB.CommandButton cmdTRFull 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Pay in &Full"
               Height          =   400
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdTRptAll 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Pay &All"
               Height          =   400
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdTRSave 
               BackColor       =   &H00F0F0F0&
               Caption         =   "&Save"
               Enabled         =   0   'False
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   225
               Width           =   1200
            End
            Begin VB.CommandButton cmdReceiptDiscard 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Clear"
               Height          =   400
               Left            =   4080
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   225
               Width           =   1200
            End
         End
         Begin VB.CommandButton cmdTRClose 
            BackColor       =   &H00F0F0F0&
            Caption         =   "C&lose"
            Height          =   400
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   7020
            Width           =   1400
         End
         Begin VB.CommandButton cmdRptAllocate 
            BackColor       =   &H00F0F0F0&
            Caption         =   "All&ocation Only"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   7020
            Width           =   1700
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   12975
            Begin VB.Frame Frame8 
               BackColor       =   &H00DEDEDE&
               Caption         =   "Analysis:"
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
               Height          =   1215
               Index           =   4
               Left            =   10080
               TabIndex        =   61
               Top             =   0
               Width           =   2895
               Begin VB.TextBox txtDifference 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   65
                  Text            =   "0.00"
                  Top             =   855
                  Width           =   1215
               End
               Begin VB.TextBox txtReceiptEntered 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   64
                  Text            =   "0.00"
                  Top             =   527
                  Width           =   1215
               End
               Begin VB.TextBox txtReceiptTotal 
                  Alignment       =   1  'Right Justify
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
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1540
                  Locked          =   -1  'True
                  TabIndex        =   63
                  Text            =   "0.00"
                  Top             =   200
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Difference                  £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   68
                  Top             =   855
                  Width           =   1380
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Receipt Entered    £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   67
                  Top             =   525
                  Width           =   1350
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Receipt total            £"
                  ForeColor       =   &H00004000&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   66
                  Top             =   195
                  Width           =   1350
               End
            End
            Begin VB.TextBox txtReceiptReference 
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
               Height          =   315
               Left            =   4560
               MaxLength       =   12
               TabIndex        =   52
               Top             =   240
               Width           =   2175
            End
            Begin VB.TextBox txtTReceiptTotal 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   8520
               TabIndex        =   55
               Text            =   "0.00"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtTRDate 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   8520
               MaxLength       =   10
               TabIndex        =   56
               Top             =   720
               Width           =   1215
            End
            Begin MSForms.CommandButton cmdTenantLookup 
               Height          =   255
               Left            =   3135
               TabIndex        =   51
               Top             =   720
               Width           =   255
               Caption         =   """"
               Size            =   "450;450"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox txtTenantID 
               Height          =   315
               Left            =   765
               TabIndex        =   75
               Top             =   720
               Width           =   2655
               VariousPropertyBits=   746604575
               BackColor       =   15858158
               Size            =   "4683;556"
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.CommandButton cmdSetAmtType 
               Height          =   315
               Left            =   6420
               TabIndex        =   54
               Top             =   720
               Width           =   315
               Caption         =   "- -"
               Size            =   "556;556"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.ComboBox cmbRptAmtType 
               Height          =   315
               Left            =   4560
               TabIndex        =   53
               Top             =   720
               Width           =   1845
               VariousPropertyBits=   1753237531
               DisplayStyle    =   3
               Size            =   "3254;556"
               BoundColumn     =   0
               TextColumn      =   2
               ColumnCount     =   3
               ListRows        =   20
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   6
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Receipt Type"
               Height          =   195
               Index           =   14
               Left            =   3600
               TabIndex        =   74
               Top             =   720
               Width           =   915
            End
            Begin MSForms.ComboBox cboRptPropertyList 
               Height          =   315
               Left            =   765
               TabIndex        =   50
               Top             =   240
               Width           =   2655
               VariousPropertyBits=   1753237531
               DisplayStyle    =   3
               Size            =   "4683;556"
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Property"
               Height          =   195
               Index           =   31
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference"
               Height          =   195
               Index           =   30
               Left            =   3600
               TabIndex        =   72
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Receipt Date"
               Height          =   195
               Index           =   34
               Left            =   7035
               TabIndex        =   71
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Receipt Amt    £"
               Height          =   195
               Index           =   33
               Left            =   7035
               TabIndex        =   70
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tenant"
               Height          =   195
               Index           =   32
               Left            =   120
               TabIndex        =   69
               Top             =   720
               Width           =   495
            End
         End
         Begin VB.TextBox txtTReceipt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   6840
            TabIndex        =   48
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtCrReceipt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   8880
            MaxLength       =   13
            TabIndex        =   47
            Top             =   5040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Allocation:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   705
            Index           =   4
            Left            =   5040
            TabIndex        =   43
            Top             =   6720
            Visible         =   0   'False
            Width           =   3705
            Begin VB.CommandButton cmdRptAutomatic 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Automatic"
               Height          =   400
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   225
               Width           =   1080
            End
            Begin VB.CommandButton cmdAllocationDiscard 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Clear"
               Height          =   400
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   225
               Width           =   1080
            End
            Begin VB.CommandButton cmdRptAllocateSave 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Save"
               Enabled         =   0   'False
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   225
               Width           =   1080
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "allocation ref."
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   195
               Index           =   2
               Left            =   1440
               TabIndex        =   46
               Top             =   0
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.TextBox txtAllocatedDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   0
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   6600
            Visible         =   0   'False
            Width           =   1200
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTReceipt 
            Height          =   2295
            Left            =   120
            TabIndex        =   79
            Top             =   2040
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   4048
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            GridLinesFixed  =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   1.5
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTCrPoA 
            Height          =   1800
            Left            =   120
            TabIndex        =   80
            Top             =   4770
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   3175
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   1.5
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPay 
            Height          =   3555
            Index           =   0
            Left            =   -74640
            TabIndex        =   95
            Top             =   2520
            Width           =   12555
            _ExtentX        =   22146
            _ExtentY        =   6271
            _Version        =   393216
            Cols            =   17
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
            _Band(0).Cols   =   17
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   201
            Top             =   1800
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   4471
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   13553358
            BackColorSel    =   12648384
            ForeColorSel    =   12582912
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
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
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCrPoA 
            Height          =   2040
            Left            =   -74880
            TabIndex        =   202
            Top             =   4530
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   3598
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648384
            ForeColorSel    =   12582912
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "  Bank Receipt  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Index           =   3
            Left            =   -69360
            TabIndex        =   305
            Top             =   6960
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "  Bank Payment  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Index           =   0
            Left            =   -69360
            TabIndex        =   304
            Top             =   6960
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Type"
            Height          =   240
            Index           =   4
            Left            =   -74160
            TabIndex        =   302
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Bank"
            Height          =   240
            Index           =   3
            Left            =   -74640
            TabIndex        =   303
            Top             =   2295
            Width           =   375
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Date"
            Height          =   255
            Index           =   5
            Left            =   -73320
            TabIndex        =   301
            Top             =   2295
            Width           =   495
         End
         Begin VB.Label lblBankRec 
            Height          =   255
            Index           =   29
            Left            =   -72120
            TabIndex        =   300
            Top             =   2160
            Width           =   15
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Client"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   -72360
            TabIndex        =   299
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblBankRec 
            Caption         =   "N/C"
            Height          =   255
            Index           =   7
            Left            =   -70800
            TabIndex        =   298
            Top             =   2295
            Width           =   375
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Fund"
            Height          =   255
            Index           =   9
            Left            =   -68520
            TabIndex        =   297
            Top             =   2295
            Width           =   495
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Ref"
            Height          =   255
            Index           =   8
            Left            =   -70200
            TabIndex        =   296
            Top             =   2295
            Width           =   615
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Details"
            Height          =   255
            Index           =   10
            Left            =   -67680
            TabIndex        =   295
            Top             =   2295
            Width           =   615
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Net"
            Height          =   255
            Index           =   11
            Left            =   -65040
            TabIndex        =   294
            Top             =   2295
            Width           =   615
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Tax"
            Height          =   255
            Index           =   12
            Left            =   -63960
            TabIndex        =   293
            Top             =   2295
            Width           =   375
         End
         Begin VB.Label lblBankRec 
            Caption         =   "Total"
            Height          =   255
            Index           =   0
            Left            =   -63240
            TabIndex        =   292
            Top             =   2295
            Width           =   375
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt"
            Height          =   195
            Index           =   19
            Left            =   10800
            TabIndex        =   221
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O/S Amount"
            Height          =   195
            Index           =   18
            Left            =   9600
            TabIndex        =   220
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   195
            Index           =   17
            Left            =   8880
            TabIndex        =   219
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
            Height          =   195
            Index           =   16
            Left            =   6600
            TabIndex        =   218
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   195
            Index           =   15
            Left            =   4560
            TabIndex        =   217
            Top             =   1800
            Width           =   225
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   14
            Left            =   3480
            TabIndex        =   216
            Top             =   1800
            Width           =   345
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit ID"
            Height          =   195
            Index           =   13
            Left            =   2040
            TabIndex        =   215
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Index           =   12
            Left            =   1320
            TabIndex        =   214
            Top             =   1800
            Width           =   345
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   195
            Index           =   11
            Left            =   720
            TabIndex        =   213
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sign"
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   212
            Top             =   1800
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "allocating row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   7
            Left            =   -63240
            TabIndex        =   205
            Top             =   1560
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "credit row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   6
            Left            =   -62880
            TabIndex        =   204
            Top             =   4320
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "  Allocation View  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Index           =   4
            Left            =   -69360
            TabIndex        =   203
            Top             =   4260
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   5
            X1              =   -74400
            X2              =   -61935
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   9
            Left            =   -74880
            TabIndex        =   209
            Top             =   1590
            Width           =   465
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   8
            Left            =   -74880
            TabIndex        =   208
            Top             =   4350
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total allocation difference:"
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
            Index           =   5
            Left            =   -66720
            TabIndex        =   207
            Top             =   6600
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label lblAllocating 
            BackStyle       =   0  'Transparent
            Caption         =   "Allocating..."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   210
            Index           =   1
            Left            =   -62880
            TabIndex        =   206
            Top             =   6600
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "  Allocation View  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Index           =   2
            Left            =   5640
            TabIndex        =   81
            Top             =   4260
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "credit row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   4
            Left            =   12120
            TabIndex        =   82
            Top             =   4320
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "allocating row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   3
            Left            =   11760
            TabIndex        =   83
            Top             =   1560
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   1
            X1              =   600
            X2              =   13065
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   1590
            Width           =   465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   2
            X1              =   675
            X2              =   13060
            Y1              =   4425
            Y2              =   4425
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   4350
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total allocation difference:"
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
            Index           =   1
            Left            =   8280
            TabIndex        =   85
            Top             =   6600
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label lblAllocating 
            BackStyle       =   0  'Transparent
            Caption         =   "Allocating..."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   210
            Index           =   0
            Left            =   12120
            TabIndex        =   84
            Top             =   6600
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0FFC0&
            BorderWidth     =   2
            Index           =   3
            X1              =   675
            X2              =   13060
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0FFFF&
            BorderWidth     =   2
            Index           =   0
            X1              =   600
            X2              =   13065
            Y1              =   1680
            Y2              =   1695
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0FFFF&
            BorderWidth     =   2
            Index           =   6
            X1              =   -74400
            X2              =   -61935
            Y1              =   1680
            Y2              =   1695
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   4
            X1              =   -74325
            X2              =   -61940
            Y1              =   4425
            Y2              =   4425
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0FFC0&
            BorderWidth     =   2
            Index           =   7
            X1              =   -74325
            X2              =   -61940
            Y1              =   4440
            Y2              =   4440
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCashBook 
         Height          =   5115
         Left            =   240
         TabIndex        =   253
         Top             =   1200
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   9022
         _Version        =   393216
         Cols            =   10
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
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxStatementReconcile 
         Height          =   4425
         Left            =   -74880
         TabIndex        =   306
         Top             =   1515
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   7805
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   53
         Left            =   11520
         TabIndex        =   349
         Top             =   6435
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   10710
         TabIndex        =   348
         Top             =   6435
         Width           =   630
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   5640
         TabIndex        =   333
         Top             =   6435
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Index           =   51
         Left            =   -68400
         TabIndex        =   331
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   50
         Left            =   -65160
         TabIndex        =   330
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   49
         Left            =   8400
         TabIndex        =   329
         Top             =   6435
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   48
         Left            =   6960
         TabIndex        =   328
         Top             =   6435
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   47
         Left            =   -66480
         TabIndex        =   327
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   46
         Left            =   -67680
         TabIndex        =   326
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -71040
         TabIndex        =   325
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date From:"
         Height          =   255
         Index           =   44
         Left            =   360
         TabIndex        =   312
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
         Height          =   255
         Index           =   45
         Left            =   3960
         TabIndex        =   313
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00B3B2B2&
         FillColor       =   &H00FFDFDF&
         FillStyle       =   0  'Solid
         Height          =   525
         Index           =   4
         Left            =   240
         Top             =   360
         Width           =   7815
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   255
         Index           =   14
         Left            =   -73560
         TabIndex        =   311
         Top             =   1320
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   735
         Index           =   3
         Left            =   -71040
         Top             =   5760
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   735
         Index           =   2
         Left            =   -71040
         Top             =   4125
         Width           =   4695
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   735
         Index           =   1
         Left            =   -71040
         Top             =   2475
         Width           =   4695
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   735
         Index           =   0
         Left            =   -71040
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Statement  Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   10920
         TabIndex        =   262
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reconciled"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   9840
         TabIndex        =   261
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   41
         Left            =   8760
         TabIndex        =   260
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   40
         Left            =   7560
         TabIndex        =   259
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         Height          =   195
         Index           =   39
         Left            =   5280
         TabIndex        =   258
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   38
         Left            =   3720
         TabIndex        =   257
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   2640
         TabIndex        =   256
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   1500
         TabIndex        =   255
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   480
         TabIndex        =   254
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblClosingBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Left            =   -63885
         TabIndex        =   140
         Top             =   6400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   195
         Index           =   27
         Left            =   -63285
         TabIndex        =   139
         Top             =   7155
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Balance:"
         Height          =   195
         Index           =   26
         Left            =   -65040
         TabIndex        =   138
         Top             =   6400
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Uncleared:"
         Height          =   195
         Index           =   25
         Left            =   -64440
         TabIndex        =   137
         Top             =   7155
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFDFDF&
         BorderColor     =   &H00FFC0C0&
         FillColor       =   &H00FFDFDF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   3
         Left            =   -65160
         Top             =   6315
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFDFDF&
         BorderColor     =   &H00FFC0C0&
         FillColor       =   &H00FFDFDF&
         FillStyle       =   0  'Solid
         Height          =   855
         Index           =   1
         Left            =   -74880
         Top             =   360
         Width           =   3810
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Value"
         Height          =   255
         Index           =   19
         Left            =   -66480
         TabIndex        =   135
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Value"
         Height          =   255
         Index           =   18
         Left            =   -67680
         TabIndex        =   134
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Reconciled"
         Height          =   255
         Index           =   21
         Left            =   -63960
         TabIndex        =   133
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Value"
         Height          =   255
         Index           =   20
         Left            =   -65280
         TabIndex        =   132
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   255
         Index           =   17
         Left            =   -69480
         TabIndex        =   131
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -72840
         TabIndex        =   130
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -74625
         TabIndex        =   129
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Projected Closing Balance:"
         Height          =   195
         Index           =   24
         Left            =   -65850
         TabIndex        =   117
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Opening Balance:"
         Height          =   195
         Index           =   23
         Left            =   -65850
         TabIndex        =   116
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Date:"
         Height          =   195
         Index           =   21
         Left            =   -68370
         TabIndex        =   114
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Mobile:"
         Height          =   255
         Index           =   20
         Left            =   -71400
         TabIndex        =   111
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tel:"
         Height          =   255
         Index           =   19
         Left            =   -71400
         TabIndex        =   109
         Top             =   4485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fax:"
         Height          =   255
         Index           =   18
         Left            =   -71400
         TabIndex        =   108
         Top             =   5295
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "e-Mail:"
         Height          =   255
         Index           =   17
         Left            =   -71400
         TabIndex        =   107
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Contact:"
         Height          =   255
         Index           =   16
         Left            =   -71400
         TabIndex        =   106
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Website:"
         Height          =   255
         Index           =   15
         Left            =   -71400
         TabIndex        =   105
         Top             =   6105
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Height          =   2775
         Index           =   3
         Left            =   -73560
         Top             =   3840
         Width           =   10095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C000&
         Height          =   5445
         Index           =   4
         Left            =   -74760
         Top             =   720
         Width           =   12105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000013&
         BorderWidth     =   2
         Height          =   3015
         Index           =   1
         Left            =   -73560
         Top             =   600
         Width           =   10095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   3015
         Index           =   0
         Left            =   -73560
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label Label1 
         Caption         =   "Last reconcile Statement Date:"
         Height          =   255
         Index           =   13
         Left            =   -68760
         TabIndex        =   38
         Top             =   2985
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "BACS Reference:"
         Height          =   255
         Index           =   12
         Left            =   -68760
         TabIndex        =   37
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payment Method:"
         Height          =   255
         Index           =   11
         Left            =   -68760
         TabIndex        =   36
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nominal A/C Name:"
         Height          =   255
         Index           =   10
         Left            =   -68760
         TabIndex        =   35
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nominal A/C:"
         Height          =   255
         Index           =   9
         Left            =   -68760
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Account Type:"
         Height          =   255
         Index           =   8
         Left            =   -73320
         TabIndex        =   32
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Name:"
         Height          =   255
         Index           =   7
         Left            =   -73320
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Account Name:"
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   25
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Account No.:"
         Height          =   255
         Index           =   4
         Left            =   -73320
         TabIndex        =   23
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sort Code:"
         Height          =   255
         Index           =   3
         Left            =   -73320
         TabIndex        =   21
         Top             =   1470
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Height          =   5445
         Index           =   2
         Left            =   -74760
         Top             =   720
         Width           =   12105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Height          =   2775
         Index           =   5
         Left            =   -73560
         Top             =   3840
         Width           =   10095
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   -74880
         Top             =   1320
         Width           =   12495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         FillColor       =   &H00FFDFDF&
         FillStyle       =   0  'Solid
         Height          =   855
         Index           =   2
         Left            =   -68400
         Top             =   360
         Width           =   6015
      End
   End
   Begin MSAdodcLib.Adodc adoBank 
      Height          =   330
      Left            =   12960
      Top             =   4320
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
      Caption         =   "Bank Account"
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
   Begin VB.CommandButton cmdBC 
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
      Left            =   8505
      TabIndex        =   1
      Top             =   90
      Width           =   300
   End
   Begin MSForms.TextBox txtBC 
      Height          =   285
      Left            =   5805
      TabIndex        =   344
      Top             =   90
      Width           =   2790
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4921;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   855
      TabIndex        =   338
      Top             =   90
      Width           =   3420
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6032;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   345
      Left            =   13080
      TabIndex        =   332
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
      ForeColor       =   16777215
      BackColor       =   8421504
      VariousPropertyBits=   268435483
      Caption         =   "Refresh"
      Size            =   "1905;609"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   22
      Left            =   8835
      TabIndex        =   115
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Balance:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   6
      Left            =   10845
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   1
      Left            =   4680
      TabIndex        =   17
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmCashbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const btRecColNo As Byte = 10
'Private iDataEntryRow As Integer
Private cTotalSI As Currency, cTotalAdjI As Currency
Dim iTotalTran As Integer, cGridRptTotal As Currency, cGridSPTotal As Currency
Dim cTempReceiptAmt As Currency, cTempPaymentAmt  As Currency
Dim bChangesMade As Boolean, baChangesMade() As Boolean
Dim bTotalRptTyped As Boolean, bTotalPayTyped As Boolean
Private curOpeningBal As Currency, yWorkingTabPayRpt As Byte
Dim szAllBankBalance As String
Dim szUndoText As String, iCurRow As Integer
Dim iCrPoARowSel As Integer
Private BANK_TYPE As String
Private sTextBox As String, yWorkingTabCashBook As Byte
Private iSelected As Integer
Private iCurEditRow As Integer
Private lLastID As Long
Private nTaxCode As Double
Private NEW_TYPE As String
Dim bSortingCol1 As Boolean, bSortingCol2 As Boolean, bSortingCol3 As Boolean, bSortingCol4 As Boolean, bSortingCol5 As Boolean
Dim bBRCol1 As Boolean, bBRCol2 As Boolean, bBRCol3 As Boolean, bBRCol4 As Boolean, bBRCol5 As Boolean
Dim bLoad As Boolean
Dim szStatementReconcile() As String 'it is holding the invoice number for all loaded receords
Dim colTransactionIDOtherReceipt As String
Dim colTransactionIDOtherPayment As String
Dim colTransactionIDOtherBankReceipt As String
Dim UserSessionID As String

Dim OtherScnsessionIDP As String
Dim OtherScnIPP As String
Dim OtherMechineNameP As String
Dim OtherWindowsUserNameP As String

Dim OtherScnsessionIDR As String
Dim OtherScnIPR As String
Dim OtherMechineNameR As String
Dim OtherWindowsUserNameR As String

Dim OtherScnsessionIDB As String
Dim OtherScnIPB As String
Dim OtherMechineNameB As String
Dim OtherWindowsUserNameB As String

Dim otherPcsCashbookIsOpen As Boolean
Dim haveYouLockedAnyReccord As Boolean
Public SelectedConBankID As Integer
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
           If isConsolidateExists(adoConn) Then
                    rRow = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "Consolidated"
                    flxClient.TextMatrix(rRow, 2) = "Consolidated"
                    flxClient.RowHeight(rRow) = 280
                    flxClient.AddItem ""
                    rRow = 2
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
          Else
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
          End If
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub

Private Sub cmdBC_Click()
    picClient.Left = 5355.029
    picClient.Top = 140
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    sTextBox = "2"
    If txtClientList.text = "Consolidated" Then
         ConfigConsolidateBank
    Else
         ConfigureFlxBank
    End If
    szAllBankBalance = BankAndBalance(adoConn)
    adoConn.Close
    Set adoConn = Nothing
    
    tabCashbook.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    Label2.Caption = ""
    TextBox1.Visible = False
    LoadflxClient
    tabCashbook.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    tabCashbook.Enabled = True
   
    FocusControl cmdClientList
End Sub





Private Sub Command2_Click()
    Dim adoConn As New ADODB.Connection
    Dim szSQL As String
     If MsgBox("Do you want to fix entries?", vbYesNo + vbInformation, "Bank reconciliation") = vbNo Then Exit Sub
    adoConn.Open getConnectionString
        adoConn.Execute "Update tlbPayment Set  ReconNow =null,Reconciled=null"
        adoConn.Execute "Update tlbReceipt Set  ReconNow =null,Reconciled=null "
        adoConn.Execute "Update tlbBankPayment Set  ReconNow =null,Reconciled=null"
    
    adoConn.Close
    Set adoConn = Nothing
    MsgBox "Fixed"
    
End Sub

Private Sub flxClient_Click()
            tabCashbook.Enabled = True
            Dim adoConn As New ADODB.Connection
           ' On Error GoTo ErrorHandler
            adoConn.Open getConnectionString
            Dim rsCheck As New ADODB.Recordset
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1) 'we shall update this field when sTextBox = "2" and tag to consolidated bank code
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    lblClosingBalance.Caption = "0.00"
                    txtStatementDate.text = ""
                    ClearForm
                    If txtClientList.text = "" Then Exit Sub
                    If txtClientList.text = "Consolidated" Then
                        Label1(22).Caption = "Bank AC"
                        Label1(1).Caption = "Bank Name"
'                        fraConinfo.Visible = True
'                        txtAccountName.Visible = False
'                        Label1(22).Visible = False
                        ConfigConsolidateBank
                    Else
                        Label1(1).Caption = "Bank Account"
                        Label1(22).Caption = "Bank Code"
'                        fraConinfo.Visible = False
'                        txtAccountName.Visible = True
'                        Label1(22).Visible = True
                        ConfigureFlxBank
                    End If
                    
                    'rem by anol 20190412 bcoz when u click client grid you are not loading bankgrid
                    LoadAdoBank 'we are using this when loading general bank info tab 1
                   ' ConfigureFlxBank
                    szAllBankBalance = BankAndBalance(adoConn)
                    ConfigFlxStatementReconcile
                    Label1(46).Caption = "0.00"
                    Label1(47).Caption = "0.00"
                    Label1(50).Caption = "0.00"
                    
                    txtBC.Tag = ""
                    txtBC.text = ""
                    cmdBC.Tag = ""
                    txtAccountName.Tag = ""
                    txtStOpenBal.text = "0.00"
                    FocusControl cmdBC
            ElseIf sTextBox = "2" Then 'bank accounts
                'Here the code after selecting the bank accounts
                   If txtClientList.text = "Consolidated" Then
                        txtBC.text = flxClient.TextMatrix(flxClient.row, 1) 'BankName
                        txtAccountName.text = flxClient.TextMatrix(flxClient.row, 2) 'BankACNumber
                        
                        cmdBC.Tag = flxClient.TextMatrix(flxClient.row, 4) 'conBankID"
                        SelectedConBankID = flxClient.TextMatrix(flxClient.row, 4)
                        
                        txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 5) 'BankCode
                        txtBC.Tag = flxClient.TextMatrix(flxClient.row, 5) 'BankCode
                        txtAccountName.Tag = flxClient.TextMatrix(flxClient.row, 5) 'BankCode
                        
                        txtStatementDate.text = flxClient.TextMatrix(flxClient.row, 6) 'StatementDate
                        'txtProjClosingBal.text = flxClient.TextMatrix(flxClient.row, 7) 'ClosingBal
                        txtStOpenBal.text = flxClient.TextMatrix(flxClient.row, 8) 'SOB
                        cmdReconSave.Enabled = True
                        Call LoadConsolidatedBankTransactions
                   Else
                        txtBC.Tag = flxClient.TextMatrix(flxClient.row, 1) 'BANK Nominal CODE
                        txtBC.text = flxClient.TextMatrix(flxClient.row, 2) 'BANK ACCOUNT nominal NAME
                        cmdBC.Tag = flxClient.TextMatrix(flxClient.row, 3) ' tlbclientbank MYID( Bank ID)
                        txtAccountName.Tag = flxClient.TextMatrix(flxClient.row, 4) 'clientID
                        cboBC_Click ' this is the main fucntion for  loading  bank reconcilitation
                        
                   End If
                   FocusControl tabCashbook
            End If
            adoConn.Close
            Set adoConn = Nothing
            picClient.Visible = False
    Exit Sub

ErrorHandler:
    MsgBox Err.description & "::" & Err.Number
    If adoConn.State = 1 Then
        adoConn.Close
        Set adoConn = Nothing
    End If
    picClient.Visible = False
        
End Sub
Private Sub ConfigureFlxBank()
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment = vbLeftJustify
   lblClientID.Caption = "Bank Code"
   lblClientName.Caption = "Bank Name"
   Label2.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
End Sub
Private Sub ConfigConsolidateBank()
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 9
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 3000
   flxClient.ColWidth(3) = 3000
   flxClient.ColWidth(4) = 0
   flxClient.ColWidth(5) = 0
   flxClient.ColWidth(6) = 0
   flxClient.ColWidth(7) = 0
   flxClient.ColWidth(8) = 0
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment = vbLeftJustify
   lblClientID.Caption = "Bank Name"
   lblClientName.Caption = "Bank AC number"
   Label2.Caption = "Sort Code"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
End Sub
Private Sub flxClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And flxClient.row = 1 Then
        FocusControl txtSearchClientID
     End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            flxClient_Click
'            tabCashbook.Enabled = True
'            If sTextBox = "1" Then
'                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
'                    lblClosingBalance.Caption = "0.00"
'                    txtStatementDate.text = ""
'                    ClearForm
'                    If txtClientList.text = "" Then Exit Sub
'                    Dim adoconn As New ADODB.Connection
'                    On Error GoTo ErrorHandler
'                    adoconn.Open getConnectionString
'                    If txtClientList.Tag = "Con" Then
'                        ConfigConsolidateBank
'                    Else
'                        ConfigureFlxBank
'                    End If
'                    szAllBankBalance = BankAndBalance(adoconn)
'                    adoconn.Close
'                    Set adoconn = Nothing
'                    txtBC.Tag = ""
'                    txtBC.text = ""
'                    cmdBC.Tag = ""
'                    txtAccountName.Tag = ""
'                    txtStOpenBal.text = "0.00"
'                     FocusControl cmdBC
'            ElseIf sTextBox = "2" Then
'                    txtBC.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                    txtBC.text = flxClient.TextMatrix(flxClient.row, 2)
'                    cmdBC.Tag = flxClient.TextMatrix(flxClient.row, 3) ' Bank ID
'                    txtAccountName.Tag = flxClient.TextMatrix(flxClient.row, 4) 'clientID
'                    cboBC_Click
'                    FocusControl tabCashbook
'            End If
'    picClient.Visible = False
'    Exit Sub
'
'ErrorHandler:
'            MsgBox Err.description & "::" & Err.Number
'            adoconn.Close
'            Set adoconn = Nothing
'            picClient.Visible = False
    End If
    If KeyAscii = 27 Then
         picClient.Visible = False
          tabCashbook.Enabled = True
         
          If sTextBox = "1" Then
                 FocusControl cmdBC
           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           End If
    End If
End Sub

Private Sub flxStatementReconcile_RowColChange() 'this procedure is written by anol as salia wanted to check if some new transaction come from the other screen
    
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim adoRST As New ADODB.Recordset
    Dim szSQL As String
    
    
    Dim i As Integer, j As Integer, iHeaderRow As Integer
    Dim szaTran() As String
   
   
    Debug.Print time
    If optReconciliation(0).Value = True Then 'Show unreconciled transactions only
  'Details has been changes to extref on first query with tlbReceipt as it was not showing correct reference
          'By anol 30 Apr 2015
            szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                       "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                       "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, 'tlbReceipt' as TableName, " & _
                       "R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID " & _
                   "FROM tlbReceipt AS R, tlbTransactionTypes AS T, Units AS U, Property AS P " & _
                   "WHERE R.BankCode = '" & Trim(txtBC.Tag) & "' AND U.PropertyID = P.PropertyID AND " & _
                       "U.UnitNumber = R.UnitID AND P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                       "(isnull(R.ReconNow) or right(R.ReconNow,5)='Saved') ;"
        
           szSQL = szSQL + " UNION "
        
           szSQL = szSQL + _
                   "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                       "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                       "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName,  " & _
                       "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID " & _
                   "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
                   "WHERE P.BankCode = '" & Trim(txtBC.Tag) & "' AND " & _
                       "P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                       "(isnull(P.ReconNow) or right(P.ReconNow,5)='Saved') ;"
        
        
           szSQL = szSQL + " UNION "
        
           szSQL = szSQL + _
                   "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                       "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                       "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN,'tlbBankPayment' as TableName, " & _
                        "BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID " & _
                   "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
                   "WHERE BP.BANK_AC = '" & Trim(txtBC.Tag) & "' AND " & _
                       "BP.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 AND " & _
                       "(isnull(BP.ReconNow) or right(BP.ReconNow ,5)='Saved')  " & _
                   "ORDER BY 1;"
        
        Else 'Show all transactions
           'Debug.Print "Hi"
           szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                       "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                       "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
                       "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID " & _
                   "FROM tlbReceipt AS R, tlbTransactionTypes AS T, Units AS U, Property AS P " & _
                   "WHERE R.BankCode = '" & Trim(txtBC.Tag) & "' AND U.PropertyID = P.PropertyID AND " & _
                       "U.UnitNumber = R.UnitID AND P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23)"
        
           szSQL = szSQL + " UNION "
        
           szSQL = szSQL + _
                   "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                       "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                       "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
                       "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID " & _
                   "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
                   "WHERE P.BankCode = '" & Trim(txtBC.Tag) & "' AND " & _
                       "P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24)"
        
           szSQL = szSQL + " UNION "
        
           szSQL = szSQL + _
                   "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                       "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                       "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
                       "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID " & _
                   "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
                   "WHERE BP.BANK_AC = '" & Trim(txtBC.Tag) & "' AND " & _
                       "BP.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                       "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 " & _
                   "ORDER BY 1;"
        End If
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        If flxStatementReconcile.Rows = adoRST.RecordCount + 2 Then
            Debug.Print "same records, no need to refresh"
        Else
            'adoRst.Filter = "TID='PP803'"
            'adoRst.Delete adAffectCurrent
            'MsgBox adoRst.RecordCount
            Debug.Print "Need to add which is not present in the grid"
           
            While Not adoRST.EOF
                With flxStatementReconcile
                        i = flxStatementReconcile.Rows - 1
                        If IsInArray(adoRST("TID").Value, szStatementReconcile) = False Then 'comparing SQl invoice with loaded invoice if something new invoice has come then add to grid
                    'MsgBox "Add in the grid"
                            
                            If Not IsNull(adoRST.Fields.Item("ReconNow")) Then
                                szaTran = Split(adoRST.Fields.Item("ReconNow"), "#")
                             Else
                                ReDim szaTran(1) As String
                                szaTran(1) = ""
                             End If
                    
                             If szaTran(1) = "Full" Then
                              '  .RowHeight(i) = 0
                             Else                                            '--------------------------- Part or not Reconciled
                                .RowHeight(i) = 240
                             End If
                    
                             If i > 1 And adoRST.Fields.Item("REF").Value <> "" And .TextMatrix(i - 1, 5) <> "" And _
                                   adoRST.Fields.Item("Details").Value = "BATCH RECEIPT" And _
                                   adoRST.Fields.Item("REF").Value = .TextMatrix(i - 1, 5) And _
                                   adoRST.Fields.Item("TD").Value = .TextMatrix(i - 1, 1) Then
                    
                    '                       Batch Receipts ->
                                If .TextMatrix(i - 1, 0) <> "-" Then
                    '              defining the parent row
                                   iHeaderRow = i - 1
                    '            duplicate the Previous row and create the previous row as header
                                   For j = 1 To .Cols - 1
                                      .TextMatrix(i, j) = .TextMatrix(iHeaderRow, j)
                                   Next j
                    
                                   .TextMatrix(iHeaderRow, 0) = "+"
                                   .TextMatrix(i, 0) = "-"
                                   .RowHeight(i) = 0
                                   i = i + 1
                                   .TextMatrix(iHeaderRow, 4) = "BATCH RECEIPT"
                                   .AddItem ""
                                End If
                                .TextMatrix(i, 0) = "-"
                                .RowHeight(i) = 0
                    
                                If .TextMatrix(iHeaderRow, 3) = "Sales Receipt" Or _
                                      .TextMatrix(iHeaderRow, 3) = "Sales Receipt on Account" Or _
                                      .TextMatrix(iHeaderRow, 3) = "Bank Receipt" Or _
                                      .TextMatrix(iHeaderRow, 3) = "Purchase Payment Refund" Then
                                   .TextMatrix(iHeaderRow, 6) = Format(adoRST.Fields.Item("AMT").Value + _
                                                                       Val(.TextMatrix(iHeaderRow, 6)), "0.00")
                            
                                    If szaTran(1) = "Full" Then
                                         .TextMatrix(iHeaderRow, 8) = Format(adoRST.Fields.Item("AMT").Value + _
                                                                       Val(.TextMatrix(iHeaderRow, 8)), "0.00")
                                    End If
                            
                                Else
                                   .TextMatrix(iHeaderRow, 7) = Format(adoRST.Fields.Item("AMT").Value + _
                                                                       Val(.TextMatrix(iHeaderRow, 7)), "0.00")
                                End If
                             End If
                    
                             .TextMatrix(i, 1) = adoRST.Fields.Item("TD").Value
                             .TextMatrix(i, 2) = adoRST.Fields.Item("TID").Value
                             'this array shall be used for recheck if some new transaction has come
                              ReDim Preserve szStatementReconcile(UBound(szStatementReconcile) + 1)
                             szStatementReconcile(i) = .TextMatrix(i, 2)
                             ''GoTo XX
                             .TextMatrix(i, 3) = adoRST.Fields.Item("TT").Value
                             If .TextMatrix(i, 4) = "" Then .TextMatrix(i, 4) = adoRST.Fields.Item("ACN").Value
                    
                             .TextMatrix(i, 5) = IIf(IsNull(adoRST.Fields.Item("REF").Value), "", _
                                                                                 adoRST.Fields.Item("REF").Value)
                             If adoRST.Fields.Item("TT").Value = "Sales Receipt" Or _
                                   adoRST.Fields.Item("TT").Value = "Sales Receipt on Account" Or _
                                   adoRST.Fields.Item("TT").Value = "Bank Receipt" Or _
                                   adoRST.Fields.Item("TT").Value = "Purchase Payment Refund" Then
                    
                    '            .TextMatrix(i, 6) = Format(adoRst.Fields.Item("AMT").Value, "0.00")
                                If szaTran(1) = "Part" Then
                                   .TextMatrix(i, 6) = Format(Val(adoRST.Fields.Item("AMT").Value) - _
                                                           Val(adoRST.Fields.Item("Reconciled").Value), "0.00")
                                Else
                                   .TextMatrix(i, 6) = Format(adoRST.Fields.Item("AMT").Value, "0.00")
                                End If
                             Else
                    '            .TextMatrix(i, 7) = Format(adoRst.Fields.Item("AMT").Value, "0.00")
                                If szaTran(1) = "Part" Then
                                   .TextMatrix(i, 7) = Format(Val(adoRST.Fields.Item("AMT").Value) + _
                                                           Val(adoRST.Fields.Item("Reconciled").Value), "0.00")
                                Else
                                   .TextMatrix(i, 7) = Format(adoRST.Fields.Item("AMT").Value, "0.00")
                                End If
                             End If
                    
                             .TextMatrix(i, 8) = IIf(IsNull(adoRST.Fields.Item("Reconciled").Value), "", _
                                                     Format(adoRST.Fields.Item("Reconciled").Value, "0.00"))
                    '         .TextMatrix(i, 8) = "0.00"
                    
                             .TextMatrix(i, 9) = szaTran(1)
                             .TextMatrix(i, 12) = szaTran(0)
                             If szaTran(1) = "Saved" Then .TextMatrix(i, 11) = "M"
                             .TextMatrix(i, btRecColNo) = adoRST.Fields.Item("TransactionID").Value
                             .TextMatrix(i, 18) = IIf(IsNull(adoRST!TableName), "", adoRST!TableName)
                             If .TextMatrix(i, 18) = "tlbPayment" And Len(colTransactionIDOtherPayment) > 0 And otherPcsCashbookIsOpen = True Then
                                'Lock for other screen
                                .col = 0
                                .row = i
                                .CellBackColor = RGB(255, 0, 0)
                                adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & OtherScnsessionIDP & "',WindowsUserName='" & _
                                OtherWindowsUserNameP & "',MachineName='" & OtherMechineNameP & "'," & _
                                "PrestigeUserName='" & User & "',ServerIPaddress='" & OtherScnIPP & "' where TransactionID =" & adoRST.Fields.Item("TransactionID").Value & ""
                                haveYouLockedAnyReccord = True
                             Else
                                'lock for this screen
                                adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                SystemUser & "',MachineName='" & WS_Name & "'," & _
                                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID =" & adoRST.Fields.Item("TransactionID").Value & ""
                                haveYouLockedAnyReccord = True
                             End If
                             If .TextMatrix(i, 18) = "tlbReceipt" And Len(colTransactionIDOtherReceipt) > 0 And otherPcsCashbookIsOpen = True Then
                               'Lock for other screen
                                .col = 0
                                .row = i
                                .CellBackColor = RGB(255, 0, 0)
                                adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & OtherScnsessionIDR & "',WindowsUserName='" & _
                                    OtherWindowsUserNameR & "',MachineName='" & OtherMechineNameR & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & OtherScnIPR & "' where TransactionID =" & adoRST.Fields.Item("TransactionID").Value & ""
                                    haveYouLockedAnyReccord = True
                             Else
                                'lock for this screen
                                   adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                    SystemUser & "',MachineName='" & WS_Name & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID =" & adoRST.Fields.Item("TransactionID").Value & ""
                                    haveYouLockedAnyReccord = True
                             End If
                             If .TextMatrix(i, 18) = "tlbBankPayment" And Len(colTransactionIDOtherBankReceipt) > 0 And otherPcsCashbookIsOpen = True Then
                               'Lock for other screen
                                .col = 0
                                .row = i
                                .CellBackColor = RGB(255, 0, 0)
                                adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & OtherScnsessionIDB & "',WindowsUserName='" & _
                                    OtherWindowsUserNameB & "',MachineName='" & OtherMechineNameB & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & OtherScnIPB & "' where MY_ID ='" & adoRST.Fields.Item("TransactionID").Value & "'"
                                    haveYouLockedAnyReccord = True
                             Else
                                'lock for this screen
                                adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                    SystemUser & "',MachineName='" & WS_Name & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID ='" & adoRST.Fields.Item("TransactionID").Value & "'" 'TransactionID
                                    haveYouLockedAnyReccord = True
                             End If
                            .AddItem ""
                             i = i + 1
                             
                End If

                adoRST.MoveNext
                End With
            Wend
             'Debug.Print time
            txtAcBal.text = Format(BankAccBalance(adoConn, txtBC.Tag, txtClientList.Tag), "0.00")
        End If
   End If
   adoRST.Close
   adoConn.Close
   Set adoConn = Nothing
   Call InstantUnLock
End Sub
'Public Function IsInArray(FindValue As Variant, arrSearch As _
'   Variant) As Boolean
'
'    On Error GoTo LocalError
'    If Not IsArray(arrSearch) Then Exit Function
'    IsInArray = InStr(1, vbNullChar & Join(arrSearch, _
'     vbNullChar) & vbNullChar, vbNullChar & FindValue & _
'     vbNullChar) > 0
'
'Exit Function
'LocalError:
'    'Justin (just in case)
'End Function
'Private Function ValueExistsInArray() As Boolean
'
'End Function
Private Sub Form_Terminate()
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    adoconn.Execute "Delete FROM Recordlocking where screen='Cashbook' and clientID='" & txtClientList.Tag & "' And BankCode='" & txtBC.Tag & "'"
'    adoconn.Close
End Sub

Private Sub txtBC_Change()
'    If bLoad = True Then
'        Dim adocon As New ADODB.Connection
'        adocon.Open getConnectionString
'        adocon.Execute "Delete FROM Recordlocking where screen='Cashbook' and clientID='" & txtClientList.Tag & "' And BankCode='" & txtBC.Tag & "'"
'        adocon.Close
'    End If
End Sub

Private Sub txtBC_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            FocusControl cmdBC
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdClientList
    End If
End Sub

Private Sub txtProjClosingBal_Change()
'    If IsNumeric(txtProjClosingBal.text) = False Then
'        txtProjClosingBal.text = ""
'    End If
     
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
           FocusControl flxClient
    End If
    If KeyCode = 13 Then
           FocusControl txtSearchClientName
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
            tabCashbook.Enabled = True
         
          'If sTextBox = "1" Then
           FocusControl cmdClientList
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
         FocusControl flxClient
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
           FocusControl flxClient
        End If
    End If
End Sub
Private Sub cboANF_Change()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iRec As Integer

   On Error GoTo ErrorHandler
   adoConn.Open getConnectionString
 
   szSQL = "SELECT DISTINCT BANK_SC, Bank_AC_Name " & _
           "FROM tlbClientBanks  " & _
           "WHERE CLIENT_ID = '" & Label13(21).Caption & "' AND BANK_AC_NUM = '" & cboANF.text & "'; "

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Label13(13).Caption = ""
   Label13(15).Caption = ""

   If adoRST.EOF Then GoTo NoRes
   Label13(13).Caption = IIf(IsNull(adoRST!BANK_SC), "", adoRST!BANK_SC)
   Label13(15).Caption = IIf(IsNull(adoRST!Bank_AC_Name), "", adoRST!Bank_AC_Name)
   Label13(7).Caption = cboANF.Column(1)

   adoRST.Close
   szSQL = "SELECT DISTINCT BANK_AC_NUM, NominalCode " & _
           "FROM tlbClientBanks  " & _
           "WHERE CLIENT_ID = '" & Label13(21).Caption & "'; "

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   cboANT.Clear

   cboANT.Enabled = True
   ReDim szaData(1, adoRST.RecordCount - 1) As String
   iRec = 0
   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item("BANK_AC_NUM").Value
      szaData(1, iRec) = adoRST.Fields.Item("NominalCode").Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend

   cboANT.Column() = szaData()
   FocusControl cboANT

NoRes:
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cboANT_Change()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler
   adoConn.Open getConnectionString

   szSQL = "SELECT BANK_SC, Bank_AC_Name " & _
           "FROM tlbClientBanks  " & _
           "WHERE CLIENT_ID = '" & Label13(21).Caption & "' AND BANK_AC_NUM = '" & cboANT.text & "'; "

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Label13(20).Caption = ""
   Label13(19).Caption = ""
   Label13(17).Caption = ""
   cboFundBankTransf.Enabled = True

   If adoRST.EOF Then GoTo NoRes
   Label13(20).Caption = IIf(IsNull(adoRST!BANK_SC), "", adoRST!BANK_SC)
   Label13(19).Caption = IIf(IsNull(adoRST!Bank_AC_Name), "", adoRST!Bank_AC_Name)
   Label13(17).Caption = cboANT.Column(1)

NoRes:
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

'Private Function ClientBank(szClientID As String) As Integer
'   Dim i As Integer
'
'   While i < cboClientID.ListCount
'      If szClientID = cboClientID.Column(0, i) Then
'         ClientBank = i
'         Exit Function
'      End If
'      i = i + 1
'   Wend
'End Function

Private Sub ClearForm()
   txtBankName.text = ""
   txtAccountName.text = ""
   txtSortCode.text = ""
   txtAcNo.text = ""
   txtAcName.text = ""
   txtAcType.text = ""
   txtNC.text = ""
   txtNN.text = ""
   txtPaymentMethod.text = ""
   txtBACS.text = ""
   txtLastReconDt.text = ""
   txtContact.text = ""
   txtTel.text = ""
   txtFax.text = ""
   txtEmail.text = ""
   txtWebsite.text = ""
   txtMobile.text = ""
   txtNote.text = ""
   txtProjClosingBal.text = ""
   txtAcBal.text = ""
   flxStatementReconcile.Clear
   flxCashBook.Clear
   txtNote.text = ""
   cmbFiles.Clear
End Sub



Public Sub cboBC_Click()
   fmeLoading.Top = 4305
   fmeLoading.Left = 4508
   fmeLoading.Visible = True
   fmeLoading.Refresh
   
   Dim szaBankBal() As String
   Dim adoConn As New ADODB.Connection
   If txtBC.Tag = "" Then
        MsgBox "Please select a Bank Code to view your Bank Transactions in the Cashbook", vbInformation, "Warning"
        
        Exit Sub
   End If
'issue 523
'closing balance was not refreshing on client change
   lblClosingBalance.Caption = "0.00"
   txtStOpenBal.Locked = False
   'End of modification
   szaBankBal = Split(szAllBankBalance, " # ")

   If txtBC.text <> "" Then
        txtAccountName.text = txtBC.Tag
        adoConn.Open getConnectionString
        Dim rsFramevisible As New ADODB.Recordset
        rsFramevisible.Open "Select * from tlbClientBanks where client_Id='" & txtClientList.Tag & "' and NominalCode='" & txtAccountName.text & "' and consolidated=true", adoConn, adOpenStatic, adLockReadOnly
        If Not rsFramevisible.EOF Then
            Label1(21).Visible = False
            Label1(23).Visible = False
            Label1(24).Visible = False
            Shape2(2).Visible = False
            txtStOpenBal.Visible = False
            txtProjClosingBal.Visible = False
            txtStatementDate.Visible = False
            cmdReconSave.Enabled = False
        Else
            Label1(21).Visible = True
            Label1(23).Visible = True
            Label1(24).Visible = True
            Shape2(2).Visible = True
            txtStOpenBal.Visible = True
            txtProjClosingBal.Visible = True
            txtStatementDate.Visible = True
            cmdReconSave.Enabled = True
        End If
        rsFramevisible.Close
        Set rsFramevisible = Nothing
        txtAcBal.text = Format(BankAccBalance(adoConn, txtBC.Tag, txtClientList.Tag), "0.00")
        'label change tlbBankReconcilation
        Dim rstlbBankReconcilation As New ADODB.Recordset
        rstlbBankReconcilation.Open "select * from tlbBankReconcilation where clientID='" & txtClientList.Tag & "' AND AccountNum='" & txtBC.Tag & "'", adoConn, adOpenKeyset, adLockReadOnly
        If rstlbBankReconcilation.EOF Then
            Label1(23).Caption = "Reconciled Cashbook Balance:"
            Label1(24).Caption = "Reconciled Statement Balance:"
            txtStOpenBal.Locked = False
        Else
            Label1(23).Caption = "Statement Opening Balance:"
            Label1(24).Caption = "Projected Closing Balance:"
            txtStOpenBal.Locked = True
        End If
        rstlbBankReconcilation.Close
        'end of showing label changes
        ConfigFlxStatementReconcile
        'Clear all locks when changing the bank accounts
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        LoadFlxStatementReconcile adoConn
        '      LoadLstBankStDates adoConn

      With adoBank.Recordset
         .Find ("NominalCode = '" & txtBC.Tag & "'"), , , 1
         If .EOF Or .BOF Then
            'MsgBox "Bank Id is missing. Please contact PCM!!", vbInformation, "Nominal Code is: " & txtBC.Tag
            fmeLoading.Visible = False
            Exit Sub
         End If
         txtBankName.text = .Fields("BANK_NAME").Value
         txtSortCode.text = .Fields("SORT_CODE").Value
         txtAcNo.text = IIf(IsNull(.Fields("BANK_AC_NUM").Value), "", .Fields("BANK_AC_NUM").Value)
         txtAcName.text = .Fields("Bank_AC_Name").Value
         txtAcType.text = IIf(IsNull(.Fields("AccountType").Value), "", .Fields("AccountType").Value)
         txtNC.text = .Fields("NominalCode").Value
         txtNN.text = .Fields("Name").Value
         txtPaymentMethod.text = IIf(IsNull(.Fields("PaymentMethod").Value), "", .Fields("PaymentMethod").Value)
         txtBACS.text = IIf(IsNull(.Fields("BacsRef").Value), "", .Fields("BacsRef").Value)
         txtLastReconDt.text = IIf(IsNull(.Fields("spare2").Value), "", .Fields("spare2").Value)
         txtContact.text = IIf(IsNull(.Fields("Contact").Value), "", .Fields("Contact").Value)
         txtTel.text = IIf(IsNull(.Fields("Tel").Value), "", .Fields("Tel").Value)
         txtFax.text = IIf(IsNull(.Fields("Fax").Value), "", .Fields("Fax").Value)
         txtEmail.text = IIf(IsNull(.Fields("eMail").Value), "", .Fields("eMail").Value)
         txtWebsite.text = IIf(IsNull(.Fields("Website").Value), "", .Fields("Website").Value)
         txtMobile.text = IIf(IsNull(.Fields("Mobile").Value), "", .Fields("Mobile").Value)
         txtNote.text = IIf(IsNull(.Fields("BankMemo").Value), "", .Fields("BankMemo").Value)
         txtProjClosingBal.text = ""
         If IsNull(.Fields("PCB").Value) Then
            txtBankName.text = .Fields("BANK_NAME").Value
            txtSortCode.text = .Fields("SORT_CODE").Value
            txtAcNo.text = IIf(IsNull(.Fields("BANK_AC_NUM").Value), "", .Fields("BANK_AC_NUM").Value)
            txtAcName.text = .Fields("Bank_AC_Name").Value
            txtAcType.text = IIf(IsNull(.Fields("AccountType").Value), "", .Fields("AccountType").Value)
            txtNC.text = .Fields("NominalCode").Value
            txtNN.text = .Fields("Name").Value
            txtPaymentMethod.text = IIf(IsNull(.Fields("PaymentMethod").Value), "", .Fields("PaymentMethod").Value)
            txtBACS.text = IIf(IsNull(.Fields("BacsRef").Value), "", .Fields("BacsRef").Value)
            txtLastReconDt.text = IIf(IsNull(.Fields("spare2").Value), "", .Fields("spare2").Value)
            txtContact.text = IIf(IsNull(.Fields("Contact").Value), "", .Fields("Contact").Value)
            txtTel.text = IIf(IsNull(.Fields("Tel").Value), "", .Fields("Tel").Value)
            txtFax.text = IIf(IsNull(.Fields("Fax").Value), "", .Fields("Fax").Value)
            txtEmail.text = IIf(IsNull(.Fields("eMail").Value), "", .Fields("eMail").Value)
            txtWebsite.text = IIf(IsNull(.Fields("Website").Value), "", .Fields("Website").Value)
            txtMobile.text = IIf(IsNull(.Fields("Mobile").Value), "", .Fields("Mobile").Value)
            txtNote.text = IIf(IsNull(.Fields("BankMemo").Value), "", .Fields("BankMemo").Value)
         Else
            If Val(.Fields("PCB").Value) <> 0 Then
               txtProjClosingBal.text = .Fields("PCB").Value
            End If
         End If

         txtStatementDate.text = IIf(.Fields("spare2").Value = "" Or _
                                     IsNull(.Fields("spare2").Value), _
                                     Format(Now, "dd/mm/yyyy"), Format(.Fields("spare2").Value, _
                                     "dd/mm/yyyy"))
      End With
      LoadFlxCashBook adoConn
      CalDrCrCBHistory

      adoConn.Close
      Set adoConn = Nothing

      If tabCashbook.Tab = 3 Then
         statefocus 'written by anol 20161013
         txtStatementDate.SelLength = Len(txtStatementDate.text)
      End If
   End If
   fmeLoading.Visible = False
   'resolved by BOSL
      'added by anol 19 Jan 2015
      'txtStOpenBal.Locked = False
      'txtStOpenBal.text = txtAcBal.text
      'end of modification
End Sub
Private Function BankAccBalanceConsolidated(adoConn As ADODB.Connection, ByVal SelectedConBankID As String, szClientID As String) As Currency
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0 " & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND " & _
                 "P.Amount > 0 " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRST.EOF
      If adoRST.Fields.Item("T").Value = "3" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "4" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "8" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "9" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "23" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRST.Fields.Item("AMT").Value
      If adoRST.Fields.Item("T").Value = "24" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRST.Fields.Item("AMT").Value

      adoRST.MoveNext
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Function
Public Sub LoadConsolidatedBankTransactions()
   fmeLoading.Top = 4305
   fmeLoading.Left = 4508
   fmeLoading.Visible = True
   fmeLoading.Refresh
   
   Dim szaBankBal() As String
   Dim adoConn As New ADODB.Connection
'issue 523
'closing balance was not refreshing on client change
   lblClosingBalance.Caption = "0.00"
   txtStOpenBal.Locked = False
   'End of modification
   szaBankBal = Split(szAllBankBalance, " # ")

   If txtBC.text <> "" Then
'        txtAccountName.text = txtBC.Tag
        adoConn.Open getConnectionString
        Dim rsFramevisible As New ADODB.Recordset
        ' Here txtClientList.Tag is a consolidated Bank ID
        rsFramevisible.Open "Select * from tlbClientBanks where client_Id='" & txtClientList.Tag & "' and NominalCode='" & txtAccountName.text & "' and consolidated=true", adoConn, adOpenStatic, adLockReadOnly
        If Not rsFramevisible.EOF Then
            Label1(21).Visible = False
            Label1(23).Visible = False
            Label1(24).Visible = False
            Shape2(2).Visible = False
            txtStOpenBal.Visible = False
            txtProjClosingBal.Visible = False
            txtStatementDate.Visible = False
        Else
            Label1(21).Visible = True
            Label1(23).Visible = True
            Label1(24).Visible = True
            Shape2(2).Visible = True
            txtStOpenBal.Visible = True
            txtProjClosingBal.Visible = True
            txtStatementDate.Visible = True
        End If
        rsFramevisible.Close
        Set rsFramevisible = Nothing
        'this Balance is  not a normal balance its a combined balance. Need to work it out
        txtAcBal.text = Format(BankAccBalanceConsolidated(adoConn, SelectedConBankID, txtClientList.Tag), "0.00")
        'label change tlbBankReconcilation
        Dim rstlbBankReconcilation As New ADODB.Recordset
        rstlbBankReconcilation.Open "select * from tlbBankReconcilation where clientID='" & txtClientList.Tag & "' AND AccountNum='" & txtBC.Tag & "'", adoConn, adOpenKeyset, adLockReadOnly
        If rstlbBankReconcilation.EOF Then
            Label1(23).Caption = "Reconciled Cashbook Balance:"
            Label1(24).Caption = "Reconciled Statement Balance:"
            txtStOpenBal.Locked = False
        Else
            Label1(23).Caption = "Statement Opening Balance:"
            Label1(24).Caption = "Projected Closing Balance:"
            txtStOpenBal.Locked = True
        End If
        rstlbBankReconcilation.Close
        'end of showing label changes
        ConfigFlxStatementReconcile
        'Clear all locks when changing the bank accounts
        
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
       If cmdBC.Tag = "" Then Exit Sub
       Debug.Print time & "StatementReconcileConsolid1"
       Call LoadFlxStatementReconcileConsolidated(adoConn, cmdBC.Tag) 'Show unreconciled transactions only
       Debug.Print time & "StatementReconcileConsolid2"
 
      Dim rsStatementdate As New ADODB.Recordset
      rsStatementdate.Open "Select * from consolidatedBankList where conBankID=" & SelectedConBankID & "", adoConn, adOpenStatic, adLockReadOnly
      If Not rsStatementdate.EOF Then
            txtStatementDate.text = IIf(IsNull(rsStatementdate("StatementDate").Value), "", rsStatementdate("StatementDate").Value)
            If txtStatementDate.text <> "" Then
                txtStatementDate.text = Format(txtStatementDate.text, "dd/mm/yyyy")
            End If
      End If
      rsStatementdate.Close
      Set rsStatementdate = Nothing
      If cmdBC.Tag = "" Then Exit Sub 'cmdBC.Tag contains consolidated bank account
      Call LoadFlxCashBookConsolidated(adoConn, cmdBC.Tag)  'Here I am loading transaction for consolidated Banks
      CalDrCrCBHistory

      adoConn.Close
      Set adoConn = Nothing

      If tabCashbook.Tab = 3 Then
         statefocus 'written by anol 20161013
         txtStatementDate.SelLength = Len(txtStatementDate.text)
      End If
   End If
   fmeLoading.Visible = False
   'resolved by BOSL
      'added by anol 19 Jan 2015
      'txtStOpenBal.Locked = False
      'txtStOpenBal.text = txtAcBal.text
      'end of modification
End Sub
Private Sub statefocus()
'   'written by anol 20161013
     On Error GoTo ERRR
     txtStatementDate.SetFocus
     Exit Sub
ERRR:
End Sub
' This method is called from UpdatingCB in the module modFoms
Public Sub QuickRefresh(adoConn As ADODB.Connection)
   ConfigFlxStatementReconcile
   LoadFlxStatementReconcile adoConn
'   LoadLstBankStDates adoConn

   LoadFlxCashBook adoConn
   CalDrCrCBHistory

'  Update Account Balance
   txtAcBal.text = Format(BankAccBalance(adoConn, txtBC.Tag, txtClientList.Tag), "0.00")
End Sub

Public Sub CalDrCrCBHistory()
   Dim iRow As Integer
   On Error Resume Next

   Label1(48).Caption = "0.00"
   For iRow = 1 To flxCashBook.Rows - 1
      If flxCashBook.RowHeight(iRow) <> 0 Then _
         Label1(48).Caption = Val(Label1(48).Caption) + CCur(flxCashBook.TextMatrix(iRow, 6))
   Next iRow

   Label1(49).Caption = "0.00"
   For iRow = 1 To flxCashBook.Rows - 1
      If flxCashBook.RowHeight(iRow) <> 0 Then _
         Label1(49).Caption = Val(Label1(49).Caption) + CCur(flxCashBook.TextMatrix(iRow, 7))
   Next iRow
   Label1(53).Caption = Format(Val(Label1(48).Caption) - Val(Label1(49).Caption), "0.00")
End Sub

Public Function CalDrCrAcBalance(dtStart As Date, dtEnd As Date) As Double
   Dim iRow As Integer, cDr As Currency, cCr As Currency

   On Error Resume Next

   For iRow = 1 To flxCashBook.Rows - 1
      If CDate(flxCashBook.TextMatrix(iRow, 1)) >= dtStart And _
            CDate(flxCashBook.TextMatrix(iRow, 1)) <= dtEnd Then
         cDr = cDr + CCur(flxCashBook.TextMatrix(iRow, 6))
         cCr = cCr + CCur(flxCashBook.TextMatrix(iRow, 7))
      End If
   Next iRow

   CalDrCrAcBalance = cDr - cCr
End Function

Public Sub LoadFlxCashBook(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset

'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
'                    ^           ^         ^           ^       ^      ^          ^           ^
'Resolved by BOSL
'Modified by anol 20 Apr 2015
'Issue 0000530: Batch receipts not working correctly
'Note 1014When the user processes a multiple batch receipt, the reference shown should be the reference entered by the user in batch receipts with multiple. This should be displayed in
'1/ Cashbook history
'I have chaged Extref to Ref for tlbReceipt
'I have chaged Extref to Ref for tlbPayment
    If txtClientList.text = "Consolidated" Then
            szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
                      "Type as Type1, R.EXTRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
               "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
               "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                      "TT.TYPE_ID = R.Type AND " & _
                      "R.BankCode = B.NominalCode AND " & _
                      "U.UnitNumber = R.UnitID AND " & _
                      "U.PropertyID = P.PropertyID AND " & _
                      "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                      "B.NominalCode = R.BankCode AND " & _
                      "B.CLIENT_ID = P.ClientID " & _
               "UNION "
            szSQL = szSQL & _
                    "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                           "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
                           "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
                           "BP.BANK_AC AS ACC " & _
                    "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                           "BP.BANK_AC =B.NominalCode AND BP.TransactionType = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID " & _
                    "UNION "
            szSQL = szSQL & _
                    "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
                           "P.Amount, P.Type AS Type1, P.EXTRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
                    "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                           "P.BankCode = B.NominalCode AND P.Type = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID " & _
                    "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"
    Else
            szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
                  "Type as Type1, R.EXTRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & txtBC.Tag & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID " & _
           "UNION "
        szSQL = szSQL & _
                "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                       "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
                       "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
                       "BP.BANK_AC AS ACC " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & txtBC.Tag & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & txtClientList.Tag & "' AND " & _
                       "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID " & _
                "UNION "
        szSQL = szSQL & _
                "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
                       "P.Amount, P.Type AS Type1, P.EXTRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                       "P.BankCode = '" & txtBC.Tag & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & txtClientList.Tag & "' AND " & _
                       "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID " & _
                "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"
    End If
   

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Issue 523
            'added by anol 21 Jan 2015
            If adoRST.EOF Then
               txtStOpenBal.Locked = True
            End If
   i = 1
   flxCashBook.Clear
   
'   Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'   Removed using the AddItem of FlexGrid as its slower. Instead set the number of rows in the grid
'   before populating the records.
'
'   flxCashBook.Rows = 2
   flxCashBook.Rows = adoRST.RecordCount + 1

   
   While Not adoRST.EOF
      flxCashBook.TextMatrix(i, 1) = adoRST.Fields.Item("RDate").Value                            'Date

      flxCashBook.TextMatrix(i, 2) = adoRST.Fields.Item("Type2").Value & _
                                     adoRST.Fields.Item("T_ID").Value                             'Type
      flxCashBook.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                              'Account
      flxCashBook.TextMatrix(i, 4) = IIf(IsNull(adoRST.Fields.Item("Rfn").Value), "", _
                                     adoRST.Fields.Item("Rfn").Value)
      flxCashBook.TextMatrix(i, 5) = adoRST.Fields.Item("Details").Value                          'Details
      If adoRST.Fields.Item("Type1").Value = "3" Or _
         adoRST.Fields.Item("Type1").Value = "4" Or _
         adoRST.Fields.Item("Type1").Value = "12" Or _
         adoRST.Fields.Item("Type1").Value = "24" Then
         flxCashBook.TextMatrix(i, 6) = Format(adoRST.Fields.Item("Amount").Value, "0.00")         'Debit
      Else
         flxCashBook.TextMatrix(i, 7) = Format(adoRST.Fields.Item("Amount").Value, "0.00")         'Credit
      End If
      flxCashBook.TextMatrix(i, 8) = IIf(Val(adoRST.Fields.Item("Amount").Value) - _
                                         IIf(adoRST.Fields.Item("Reconciled").Value < 0, _
                                         adoRST.Fields.Item("Reconciled").Value * (-1), _
                                         adoRST.Fields.Item("Reconciled").Value) = 0, "YES", _
                                         IIf(IsNull(adoRST.Fields.Item("Reconciled").Value), _
                                         "NO", "PART"))                                            'Reconcialid
      If Not IsNull(adoRST.Fields.Item("SDate").Value) Then
        If adoRST.Fields.Item("SDate").Value <> "" Then
            szaTemp() = Split(adoRST.Fields.Item("SDate").Value, "#")
            flxCashBook.TextMatrix(i, 9) = IIf(szaTemp(1) = "Saved", "", szaTemp(0))                     'Statement Date
            flxCashBook.TextMatrix(i, 8) = IIf(szaTemp(1) = "Saved", "NO", flxCashBook.TextMatrix(i, 8))
         End If
      End If
       flxCashBook.TextMatrix(i, 10) = adoRST.Fields.Item("Type2").Value
      adoRST.MoveNext
      'If Not adoRst.EOF Then flxCashBook.AddItem ""
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

         
   'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
   'Removed the following code as changing the cell color for all the records takes significant time

'   For i = 1 To flxCashBook.Rows - 1
'      flxCashBook.row = i
'      If flxCashBook.TextMatrix(i, 8) = "YES" Then
'         For r = 1 To flxCashBook.Cols - 1
'            flxCashBook.col = r
'            flxCashBook.CellBackColor = RGB(162, 185, 224)
'            'issue 523
'            'added by anol 21 Jan 2015
'            txtStOpenBal.Locked = True
'         Next r
'      End If
'   Next i
   flxCashBook.row = 0
End Sub
Public Sub LoadFlxCashBookConsolidated(adoConn As ADODB.Connection, lngConBankID As Long)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset

'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
'                    ^           ^         ^           ^       ^      ^          ^           ^
'Resolved by BOSL
'Modified by anol 20 Apr 2015
'Issue 0000530: Batch receipts not working correctly
'Note 1014When the user processes a multiple batch receipt, the reference shown should be the reference entered by the user in batch receipts with multiple. This should be displayed in
'1/ Cashbook history
'I have chaged Extref to Ref for tlbReceipt
'I have chaged Extref to Ref for tlbPayment
   szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
                  "Type as Type1, R.EXTRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
                  "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
                  "BP.BANK_AC AS ACC " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "BP.TransactionType = TT.TYPE_ID AND " & _
                  "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
                  "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
                  "P.Amount, P.Type AS Type1, P.EXTRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                  "P.Type = TT.TYPE_ID AND " & _
                  "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
                  "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID " & _
           "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"

'Debug.Print time & "cashbook1 fix this join as this taking long time"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
Debug.Print time & "2"
'Issue 523
            'added by anol 21 Jan 2015
            If adoRST.EOF Then
               txtStOpenBal.Locked = True
            End If
   i = 1
   flxCashBook.Clear
   
'   Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'   Removed using the AddItem of FlexGrid as its slower. Instead set the number of rows in the grid
'   before populating the records.
'
'   flxCashBook.Rows = 2
   flxCashBook.Rows = adoRST.RecordCount + 1

   
   While Not adoRST.EOF
      flxCashBook.TextMatrix(i, 1) = adoRST.Fields.Item("RDate").Value                            'Date

      flxCashBook.TextMatrix(i, 2) = adoRST.Fields.Item("Type2").Value & _
                                     adoRST.Fields.Item("T_ID").Value                             'Type
      flxCashBook.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                              'Account
      flxCashBook.TextMatrix(i, 4) = IIf(IsNull(adoRST.Fields.Item("Rfn").Value), "", _
                                     adoRST.Fields.Item("Rfn").Value)
      flxCashBook.TextMatrix(i, 5) = adoRST.Fields.Item("Details").Value                          'Details
      If adoRST.Fields.Item("Type1").Value = "3" Or _
         adoRST.Fields.Item("Type1").Value = "4" Or _
         adoRST.Fields.Item("Type1").Value = "12" Or _
         adoRST.Fields.Item("Type1").Value = "24" Then
         flxCashBook.TextMatrix(i, 6) = Format(adoRST.Fields.Item("Amount").Value, "0.00")         'Debit
      Else
         flxCashBook.TextMatrix(i, 7) = Format(adoRST.Fields.Item("Amount").Value, "0.00")         'Credit
      End If
      flxCashBook.TextMatrix(i, 8) = IIf(Val(adoRST.Fields.Item("Amount").Value) - _
                                         IIf(adoRST.Fields.Item("Reconciled").Value < 0, _
                                         adoRST.Fields.Item("Reconciled").Value * (-1), _
                                         adoRST.Fields.Item("Reconciled").Value) = 0, "YES", _
                                         IIf(IsNull(adoRST.Fields.Item("Reconciled").Value), _
                                         "NO", "PART"))                                            'Reconcialid
      If Not IsNull(adoRST.Fields.Item("SDate").Value) Then
        If adoRST.Fields.Item("SDate").Value <> "" Then
                szaTemp() = Split(adoRST.Fields.Item("SDate").Value, "#")
                flxCashBook.TextMatrix(i, 9) = IIf(szaTemp(1) = "Saved", "", szaTemp(0))                     'Statement Date
                flxCashBook.TextMatrix(i, 8) = IIf(szaTemp(1) = "Saved", "NO", flxCashBook.TextMatrix(i, 8))
         End If
      End If
       flxCashBook.TextMatrix(i, 10) = adoRST.Fields.Item("Type2").Value
      adoRST.MoveNext
      'If Not adoRst.EOF Then flxCashBook.AddItem ""
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

         
   'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
   'Removed the following code as changing the cell color for all the records takes significant time

'   For i = 1 To flxCashBook.Rows - 1
'      flxCashBook.row = i
'      If flxCashBook.TextMatrix(i, 8) = "YES" Then
'         For r = 1 To flxCashBook.Cols - 1
'            flxCashBook.col = r
'            flxCashBook.CellBackColor = RGB(162, 185, 224)
'            'issue 523
'            'added by anol 21 Jan 2015
'            txtStOpenBal.Locked = True
'         Next r
'      End If
'   Next i
   flxCashBook.row = 0
End Sub
Public Sub LoadFlxStatementReconcile(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, j As Integer, iHeaderRow As Integer
   Dim adoSI As New ADODB.Recordset
   Dim szaTran() As String
   Dim rsUserSessionID As String ' This variable shall store the value of timestamp from the recordset
   Dim colTransactionIDReceipt As String 'this variable shall hold all the locked transaction number which is locked by this screen
   Dim colTransactionIDPayment As String 'this variable shall hold all the locked transaction number which is locked by this screen
   Dim colTransactionIDBankPayment As String 'this variable shall hold all the locked transaction number which is locked by this screen
   colTransactionIDOtherReceipt = ""
   colTransactionIDOtherPayment = ""
   colTransactionIDOtherBankReceipt = ""
   otherPcsCashbookIsOpen = False
'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'Load the data from the database every time when the unreconciled/all transactions option is selected
'As this is faster than the existing process of loading all the data and then hiding from the records
'from the grid.
If txtClientList.text = "Consolidated" Then
        If optReconciliation(0).Value = True Then 'Show unreconciled transactions only
  
                        szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                                       "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                                       "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
                                       "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress, " & _
                                       "(Select P.PropertyID from Units AS U,Property as P where P.PropertyID=U.propertyID AND R.UnitID=U.UnitNumber) as PROPID  " & _
                                   "FROM tlbReceipt AS R, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
                                   "WHERE  R.BankCode = B.NominalCode  AND " & _
                                       " R.ClientID = B.CLIENT_ID AND B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                                       "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                                       "(isnull(R.ReconNow) or right(R.ReconNow,5)='Saved') ;"
                        'R.BankCode = B.NominalCode  AND
                           szSQL = szSQL + " UNION "
                        
                           szSQL = szSQL + _
                                   "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                                       "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                                       "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
                                       "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress,(Select PropertyID from property U where U.PropertyID=P.UnitID) as PROPID  " & _
                                   "FROM tlbPayment AS P, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
                                   "WHERE P.BankCode = B.NominalCode AND " & _
                                       "P.ClientID = B.Client_ID AND " & _
                                       "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                                       "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                                       "(isnull(P.ReconNow) or right(P.ReconNow,5)='Saved') ;"
                        'P.BankCode = B.NominalCode AND
                           szSQL = szSQL + " UNION "
                        
                           szSQL = szSQL + _
                                   "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                                       "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                                       "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
                                       "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress,BP.PropertyID as PROPID " & _
                                   "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T , tlbClientBanks AS B " & _
                                   "WHERE BP.BANK_AC = B.NominalCode AND " & _
                                       "BP.ClientID = B.Client_ID  AND " & _
                                       "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                                       "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 AND " & _
                                       "(isnull(BP.ReconNow) or right(BP.ReconNow ,5)='Saved')  " & _
                                   "ORDER BY 1;"
                
        Else 'Show all transactions
                   'Debug.Print "Hi"
                   szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                               "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                               "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
                               "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress " & _
                           "FROM tlbReceipt AS R, tlbTransactionTypes AS T, Units AS U, Property AS P ,tlbClientBanks CB " & _
                           "WHERE R.BankCode = CB.NominalCode AND U.PropertyID = P.PropertyID AND " & _
                              "CB.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                               "U.UnitNumber = R.UnitID AND P.ClientID =  CB.CLIENT_ID AND " & _
                               "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23)"
                
                   szSQL = szSQL + " UNION "
                
                   szSQL = szSQL + _
                           "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                               "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                               "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
                               "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress " & _
                           "FROM tlbPayment AS P, tlbTransactionTypes AS T ,tlbClientBanks CB " & _
                           "WHERE P.BankCode =  CB.NominalCode AND " & _
                               "P.ClientID =  CB.CLIENT_ID AND " & _
                               "CB.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                               "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24)"
                
                   szSQL = szSQL + " UNION "
                
                   szSQL = szSQL + _
                           "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                               "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                               "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
                               "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress " & _
                           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T ,tlbClientBanks CB " & _
                           "WHERE BP.BANK_AC = CB.NominalCode AND " & _
                               "BP.ClientID  =  CB.CLIENT_ID  AND " & _
                               "CB.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                               "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 " & _
                           "ORDER BY 1;"
                End If
Else

        If optReconciliation(0).Value = True Then 'Show unreconciled transactions only
              'Details has been changes to extref on first query with tlbReceipt as it was not showing correct reference
              'By anol 30 Apr 2015
                szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                           "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                           "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, 'tlbReceipt' as TableName, " & _
                           "R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress " & _
                       "FROM tlbReceipt AS R, tlbTransactionTypes AS T, Units AS U, Property AS P " & _
                       "WHERE R.BankCode = '" & Trim(txtBC.Tag) & "' AND U.PropertyID = P.PropertyID AND " & _
                           "U.UnitNumber = R.UnitID AND P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                           "(isnull(R.ReconNow) or right(R.ReconNow,5)='Saved') ;"
            
               szSQL = szSQL + " UNION "
            
               szSQL = szSQL + _
                       "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                           "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                           "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName,  " & _
                           "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress " & _
                       "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
                       "WHERE P.BankCode = '" & Trim(txtBC.Tag) & "' AND " & _
                           "P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                           "(isnull(P.ReconNow) or right(P.ReconNow,5)='Saved') ;"
            
            
               szSQL = szSQL + " UNION "
            
               szSQL = szSQL + _
                       "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                           "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                           "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN,'tlbBankPayment' as TableName, " & _
                            "BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress " & _
                       "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
                       "WHERE BP.BANK_AC = '" & Trim(txtBC.Tag) & "' AND " & _
                           "BP.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 AND " & _
                           "(isnull(BP.ReconNow) or right(BP.ReconNow ,5)='Saved')  " & _
                       "ORDER BY 1;"
            
            Else 'Show all transactions
               'Debug.Print "Hi"
               szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
                           "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
                           "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
                           "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress " & _
                       "FROM tlbReceipt AS R, tlbTransactionTypes AS T, Units AS U, Property AS P " & _
                       "WHERE R.BankCode = '" & Trim(txtBC.Tag) & "' AND U.PropertyID = P.PropertyID AND " & _
                           "U.UnitNumber = R.UnitID AND P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23)"
            
               szSQL = szSQL + " UNION "
            
               szSQL = szSQL + _
                       "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
                           "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
                           "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
                           "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress " & _
                       "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
                       "WHERE P.BankCode = '" & Trim(txtBC.Tag) & "' AND " & _
                           "P.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24)"
            
               szSQL = szSQL + " UNION "
            
               szSQL = szSQL + _
                       "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
                           "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
                           "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
                           "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress " & _
                       "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
                       "WHERE BP.BANK_AC = '" & Trim(txtBC.Tag) & "' AND " & _
                           "BP.ClientID = '" & Trim(txtClientList.Tag) & "' AND " & _
                           "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 " & _
                       "ORDER BY 1;"
            End If
End If


'Debug.Print szSQL

'END OF MODIFICATION
   adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   Label1(27).Caption = "0.00"
   lblClosingBalance.Caption = "0.00"
   
'   Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
'   Removed using of AddItem of FlexGrid as its slower. Instead set the number of rows in the grid
'   before populating the records.

   'Below line was modified by anol 22 Apr 2015 to remove the bug subscript out of range errror
   flxStatementReconcile.Rows = adoSI.RecordCount + 2
    'flxStatementReconcile.Rows = adoSI.RecordCount + 1
''END
   ReDim szStatementReconcile(adoSI.RecordCount + 2) As String
   
   With flxStatementReconcile
      While Not adoSI.EOF
 
         If Not IsNull(adoSI.Fields.Item("ReconNow")) Then
            If adoSI.Fields.Item("ReconNow") = "" Then
                 ReDim szaTran(1) As String
                 szaTran(1) = ""
            Else
                szaTran = Split(adoSI.Fields.Item("ReconNow"), "#")
            End If
            
         Else
            ReDim szaTran(1) As String
            szaTran(1) = ""
         End If

         If szaTran(1) = "Full" Then
          '  .RowHeight(i) = 0
         Else                                            '--------------------------- Part or not Reconciled
            .RowHeight(i) = 240
         End If

'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'Modified the condition so that batch receipt transactions are grouped according to transaction date.


         If i > 1 And adoSI.Fields.Item("REF").Value <> "" And .TextMatrix(i - 1, 5) <> "" And _
               adoSI.Fields.Item("Details").Value = "BATCH RECEIPT" And _
               adoSI.Fields.Item("REF").Value = .TextMatrix(i - 1, 5) And _
               adoSI.Fields.Item("TD").Value = .TextMatrix(i - 1, 1) Then
'END OF MODIFICATION

'                       Batch Receipts ->
            If .TextMatrix(i - 1, 0) <> "-" Then

'              defining the parent row
               iHeaderRow = i - 1
'Note : Anol 12 Feb this is issue needs to look with Mr.Asif
'            duplicate the Previous row and create the previous row as header
               For j = 1 To .Cols - 1
                  .TextMatrix(i, j) = .TextMatrix(iHeaderRow, j)
               Next j

               .TextMatrix(iHeaderRow, 0) = "+"

               .TextMatrix(i, 0) = "-"
               .RowHeight(i) = 0
               ReDim Preserve szStatementReconcile(UBound(szStatementReconcile) + 1)
               i = i + 1
               .TextMatrix(iHeaderRow, 4) = "BATCH RECEIPT"
               .AddItem ""
            End If
            .TextMatrix(i, 0) = "-"
            .RowHeight(i) = 0

            If .TextMatrix(iHeaderRow, 3) = "Sales Receipt" Or _
                  .TextMatrix(iHeaderRow, 3) = "Sales Receipt on Account" Or _
                  .TextMatrix(iHeaderRow, 3) = "Bank Receipt" Or _
                  .TextMatrix(iHeaderRow, 3) = "Purchase Payment Refund" Then
               .TextMatrix(iHeaderRow, 6) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 6)), "0.00")
                                                   
        '   Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
        '   Calcaluting the total batch receipts that are reconcilced and assigning to the batch
        '   receipt header that is grouped based on the transaction date
        
                If szaTran(1) = "Full" Then
                     .TextMatrix(iHeaderRow, 8) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 8)), "0.00")
                End If
        '   END
        
            Else
               .TextMatrix(iHeaderRow, 7) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 7)), "0.00")
            End If
         End If

         .TextMatrix(i, 1) = adoSI.Fields.Item("TD").Value 'transaction date
         .TextMatrix(i, 2) = adoSI.Fields.Item("TID").Value 'Invoice number
         'this array shall be used for recheck if some new transaction has come
         szStatementReconcile(i) = .TextMatrix(i, 2) ' MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID=transaction number
         .TextMatrix(i, 3) = adoSI.Fields.Item("TT").Value
         If .TextMatrix(i, 4) = "" Then .TextMatrix(i, 4) = adoSI.Fields.Item("ACN").Value

         .TextMatrix(i, 5) = IIf(IsNull(adoSI.Fields.Item("REF").Value), "", _
                                                             adoSI.Fields.Item("REF").Value)
         If adoSI.Fields.Item("TT").Value = "Sales Receipt" Or _
               adoSI.Fields.Item("TT").Value = "Sales Receipt on Account" Or _
               adoSI.Fields.Item("TT").Value = "Bank Receipt" Or _
               adoSI.Fields.Item("TT").Value = "Purchase Payment Refund" Then

'            .TextMatrix(i, 6) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            If szaTran(1) = "Part" Then
               .TextMatrix(i, 6) = Format(Val(adoSI.Fields.Item("AMT").Value) - _
                                       Val(adoSI.Fields.Item("Reconciled").Value), "0.00")
            Else
               .TextMatrix(i, 6) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            End If
         Else
'            .TextMatrix(i, 7) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            If szaTran(1) = "Part" Then
               .TextMatrix(i, 7) = Format(Val(adoSI.Fields.Item("AMT").Value) + _
                                       Val(adoSI.Fields.Item("Reconciled").Value), "0.00")
            Else
               .TextMatrix(i, 7) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            End If
         End If

         .TextMatrix(i, 8) = IIf(IsNull(adoSI.Fields.Item("Reconciled").Value), "", _
                                 Format(adoSI.Fields.Item("Reconciled").Value, "0.00"))
'         .TextMatrix(i, 8) = "0.00"

         .TextMatrix(i, 9) = szaTran(1)
         .TextMatrix(i, 12) = szaTran(0)
         If szaTran(1) = "Saved" Then .TextMatrix(i, 11) = "M"
         .TextMatrix(i, btRecColNo) = adoSI.Fields.Item("TransactionID").Value
         'Locking mechanism starts
         
          rsUserSessionID = IIf(IsNull(adoSI!UserSessionID), "", adoSI!UserSessionID)
         .TextMatrix(i, 13) = IIf(IsNull(adoSI!UserSessionID), "", adoSI!UserSessionID) 'Keeping the USersesssionID to check the lock
         .TextMatrix(i, 14) = IIf(IsNull(adoSI!WindowsUserName), "", adoSI!WindowsUserName) 'BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID
         .TextMatrix(i, 15) = IIf(IsNull(adoSI!MachineName), "", adoSI!MachineName)
         .TextMatrix(i, 16) = IIf(IsNull(adoSI!Module), "", adoSI!Module)
         .TextMatrix(i, 17) = IIf(IsNull(adoSI!ClientID), "", adoSI!ClientID)
         .TextMatrix(i, 18) = IIf(IsNull(adoSI!TableName), "", adoSI!TableName)
        If UCase(.TextMatrix(i, 16)) = "CASHBOOK" And otherPcsCashbookIsOpen = False Then
            otherPcsCashbookIsOpen = True 'this shall now know that other pc has cashbook open
        End If
        
         If adoSI.Fields.Item("TableName").Value = "tlbPayment" Then
                 If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                    colTransactionIDOtherPayment = colTransactionIDOtherPayment & adoSI.Fields.Item("TransactionID").Value & ","
                    OtherScnsessionIDP = .TextMatrix(i, 13)
                    OtherWindowsUserNameP = .TextMatrix(i, 14)
                    OtherMechineNameP = .TextMatrix(i, 15)
                    OtherScnIPP = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                 Else
                    colTransactionIDPayment = colTransactionIDPayment & adoSI.Fields.Item("TransactionID").Value & ","
                 End If
         ElseIf adoSI.Fields.Item("TableName").Value = "tlbReceipt" Then
                 If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                     colTransactionIDOtherReceipt = colTransactionIDOtherReceipt & adoSI.Fields.Item("TransactionID").Value & ","
                    OtherScnsessionIDR = .TextMatrix(i, 13)
                    OtherWindowsUserNameR = .TextMatrix(i, 14)
                    OtherMechineNameR = .TextMatrix(i, 15)
                    OtherScnIPR = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                 Else
                     colTransactionIDReceipt = colTransactionIDReceipt & adoSI.Fields.Item("TransactionID").Value & ","
                 End If
         ElseIf adoSI.Fields.Item("TableName").Value = "tlbBankPayment" Then
                If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                    colTransactionIDOtherBankReceipt = colTransactionIDOtherBankReceipt & "'" & adoSI.Fields.Item("TransactionID").Value & "',"
                    OtherScnsessionIDB = .TextMatrix(i, 13)
                    OtherWindowsUserNameB = .TextMatrix(i, 14)
                    OtherMechineNameB = .TextMatrix(i, 15)
                    OtherScnIPB = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                Else
                    colTransactionIDBankPayment = colTransactionIDBankPayment & "'" & adoSI.Fields.Item("TransactionID").Value & "',"
                End If
         End If
         'Locking mechanism Ends

         i = i + 1

         adoSI.MoveNext
'Removed by Asif. Issue 0000523. Refer to comment above
'         If Not adoSI.EOF Then .AddItem ""
      Wend
   End With

   Label1(27).Caption = Format(UnclearedBalance, "0.00")
   Sum_RptPayVal

   adoSI.Close
    If Len(colTransactionIDOtherPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherPayment = Left(colTransactionIDOtherPayment, Len(colTransactionIDOtherPayment) - 1)
   End If
    If Len(colTransactionIDOtherReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherReceipt = Left(colTransactionIDOtherReceipt, Len(colTransactionIDOtherReceipt) - 1)
   End If
    If Len(colTransactionIDOtherBankReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherBankReceipt = Left(colTransactionIDOtherBankReceipt, Len(colTransactionIDOtherBankReceipt) - 1)
   End If
   If Len(colTransactionIDPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDPayment = Left(colTransactionIDPayment, Len(colTransactionIDPayment) - 1)
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionIDPayment & ")"
                        haveYouLockedAnyReccord = True
   End If
   If Len(colTransactionIDReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDReceipt = Left(colTransactionIDReceipt, Len(colTransactionIDReceipt) - 1)
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionIDReceipt & ")"
                        haveYouLockedAnyReccord = True
   End If
   If Len(colTransactionIDBankPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDBankPayment = Left(colTransactionIDBankPayment, Len(colTransactionIDBankPayment) - 1)
        adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID in (" & colTransactionIDBankPayment & ")"
                        haveYouLockedAnyReccord = True
   End If
'  SOB --> Statement Opening Balance
                If txtClientList.text = "Consolidated" Then
                    szSQL = "SELECT  * from ConsolidatedBankList where conBankID=" & SelectedConBankID & ";"
                       adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                    
                       If adoSI.EOF Then
                          txtStOpenBal.text = "0.00"
                       Else
                             txtStOpenBal.text = Format(adoSI.Fields.Item("SOB").Value, "0.00")
                       End If
                       If IsNull(adoSI.Fields.Item("ClosingBal").Value) Then
                             lblClosingBalance.Caption = "0.00"
                       Else
                            lblClosingBalance.Caption = Format(adoSI.Fields.Item("ClosingBal").Value, "0.00")
                       End If
                       adoSI.Close
                Else
                       szSQL = "SELECT ClosingBal, SOB FROM tlbClientBanks " & _
                               "WHERE NominalCode = '" & txtBC.Tag & "' AND " & _
                                     "CLIENT_ID = '" & txtClientList.Tag & "';"
                       adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                    
                       If adoSI.EOF Then
                          txtStOpenBal.text = "0.00"
                       Else
                          If IsNull(adoSI.Fields.Item("ClosingBal").Value) Then
                             lblClosingBalance.Caption = "0.00"
                             txtStOpenBal.Locked = False
                          Else
                             txtStOpenBal.Locked = True
                             lblClosingBalance.Caption = Format(adoSI.Fields.Item("ClosingBal").Value, "0.00")
                             'added by anol 05052016
                             If lblClosingBalance.Caption = "0.00" Then
                                txtStOpenBal.Locked = False
                             End If
                          End If
                          If IsNull(adoSI.Fields.Item("SOB").Value) Then
                             txtStOpenBal.text = lblClosingBalance.Caption
                             'resolved by BOSL
                             'issue 523 bank reconciliation is now showing correct opening and closing balance.
                             'Modified by anol 19 Jan 2015
                             'UpdateBankOpeningBalance adoConn, CCur(txtStOpenBal.text), txtbc.Tag
                             adoConn.Execute "UPDATE tlbClientBanks " & _
                                       "SET SOB = " & CCur(txtStOpenBal.text) & " " & _
                                       "WHERE NominalCode = '" & txtBC.Tag & "' and Client_ID= '" & txtClientList.Tag & "';"
                          Else
                             txtStOpenBal.text = Format(adoSI.Fields.Item("SOB").Value, "0.00")
                          End If
               End If
   End If

   Set adoSI = Nothing
End Sub
Public Sub LoadFlxStatementReconcileConsolidated(adoConn As ADODB.Connection, lngConBankID) 'Show unreconciled transactions only
   Dim szSQL As String, i As Integer, j As Integer, iHeaderRow As Integer
   Dim adoSI As New ADODB.Recordset
   Dim szaTran() As String
   Dim rsUserSessionID As String ' This variable shall store the value of timestamp from the recordset
   Dim colTransactionIDReceipt As String 'this variable shall hold all the locked transaction number which is locked by this screen
   Dim colTransactionIDPayment As String 'this variable shall hold all the locked transaction number which is locked by this screen
   Dim colTransactionIDBankPayment As String 'this variable shall hold all the locked transaction number which is locked by this screen
   colTransactionIDOtherReceipt = ""
   colTransactionIDOtherPayment = ""
   colTransactionIDOtherBankReceipt = ""
   otherPcsCashbookIsOpen = False
'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'Load the data from the database every time when the unreconciled/all transactions option is selected
'As this is faster than the existing process of loading all the data and then hiding from the records
'from the grid.

If optReconciliation(0).Value = True Then 'Show unreconciled transactions only, other option is handled by separate function
  'Details has been changes to extref on first query with tlbReceipt as it was not showing correct reference

   szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
               "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
               "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress, " & _
               "(Select P.PropertyID from Units AS U,Property as P where P.PropertyID=U.propertyID AND R.UnitID=U.UnitNumber) as PROPID  " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
           "WHERE  R.BankCode = B.NominalCode  AND " & _
               " R.ClientID = B.CLIENT_ID AND B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
               "(isnull(R.ReconNow) or right(R.ReconNow,5)='Saved') ;"
'R.BankCode = B.NominalCode  AND
   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
               "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
               "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress,(Select PropertyID from property U where U.PropertyID=P.UnitID) as PROPID  " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND " & _
               "P.ClientID = B.Client_ID AND " & _
               "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
               "(isnull(P.ReconNow) or right(P.ReconNow,5)='Saved') ;"
'P.BankCode = B.NominalCode AND
   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
               "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
               "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
               "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress,BP.PropertyID as PROPID " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T , tlbClientBanks AS B " & _
           "WHERE BP.BANK_AC = B.NominalCode AND " & _
               "BP.ClientID = B.Client_ID  AND " & _
               "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 AND " & _
               "(isnull(BP.ReconNow) or right(BP.ReconNow ,5)='Saved')  " & _
           "ORDER BY 1;"

Else    'Show all transaction for consolidated transactions writetn by anol 2022-06-10
     szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, R.extref as Ref, R.Amount AS AMT, R.Reconciled, " & _
               "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, " & _
               "'tlbReceipt' as TableName,R.UserSessionID,R.WindowsUserName,R.MachineName,R.Module,R.ClientID,R.ServerIPaddress, " & _
               "(Select P.PropertyID from Units AS U,Property as P where P.PropertyID=U.propertyID AND R.UnitID=U.UnitNumber) as PROPID  " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
           "WHERE  R.BankCode = B.NominalCode  AND " & _
               " R.ClientID = B.CLIENT_ID AND B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "R.Type = T.TYPE_ID AND R.Amount > 0 AND (R.Type = 3 OR R.Type = 4 OR R.Type = 23) " & _
               ";"
'R.BankCode = B.NominalCode  AND
   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
               "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'tlbPayment' as TableName, " & _
               "P.UserSessionID,P.WindowsUserName,P.MachineName,P.Module,P.ClientID,P.ServerIPaddress,(Select PropertyID from property U where U.PropertyID=P.UnitID) as PROPID  " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS T, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND " & _
               "P.ClientID = B.Client_ID AND " & _
               "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "P.Type = T.TYPE_ID AND P.Amount > 0 AND (P.Type = 8 OR P.Type = 9 OR P.Type = 24) " & _
               ";"
'P.BankCode = B.NominalCode AND
   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
               "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
               "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, BP.NOMINAL_CODE AS ACN, " & _
               "'tlbBankPayment' as TableName,BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID,BP.ServerIPaddress,BP.PropertyID as PROPID " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T , tlbClientBanks AS B " & _
           "WHERE BP.BANK_AC = B.NominalCode AND " & _
               "BP.ClientID = B.Client_ID  AND " & _
               "B.ConsolidatedBankID = " & lngConBankID & " AND " & _
               "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 " & _
               "" & _
           "ORDER BY 1;"
               


End If
'Debug.Print szSQL
If szSQL = "" Then
        MsgBox "SQL is not set", vbInformation, "Warning"
        Exit Sub
End If
'END OF MODIFICATION
   adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   Label1(27).Caption = "0.00"
   lblClosingBalance.Caption = "0.00"
   
'   Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
'   Removed using of AddItem of FlexGrid as its slower. Instead set the number of rows in the grid
'   before populating the records.

   'Below line was modified by anol 22 Apr 2015 to remove the bug subscript out of range errror
   flxStatementReconcile.Rows = adoSI.RecordCount + 2
    'flxStatementReconcile.Rows = adoSI.RecordCount + 1
''END
   ReDim szStatementReconcile(adoSI.RecordCount + 2) As String
   
   With flxStatementReconcile
      While Not adoSI.EOF
 
         If Not IsNull(adoSI.Fields.Item("ReconNow")) Then
            szaTran = Split(adoSI.Fields.Item("ReconNow"), "#")
         Else
            ReDim szaTran(1) As String
            szaTran(1) = ""
         End If

         If szaTran(1) = "Full" Then
          '  .RowHeight(i) = 0
         Else                                            '--------------------------- Part or not Reconciled
            .RowHeight(i) = 240
         End If

'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'Modified the condition so that batch receipt transactions are grouped according to transaction date.


         If i > 1 And adoSI.Fields.Item("REF").Value <> "" And .TextMatrix(i - 1, 5) <> "" And _
               adoSI.Fields.Item("Details").Value = "BATCH RECEIPT" And _
               adoSI.Fields.Item("REF").Value = .TextMatrix(i - 1, 5) And _
               adoSI.Fields.Item("TD").Value = .TextMatrix(i - 1, 1) Then
'END OF MODIFICATION

'                       Batch Receipts ->
            If .TextMatrix(i - 1, 0) <> "-" Then

'              defining the parent row
               iHeaderRow = i - 1
'Note : Anol 12 Feb this is issue needs to look with Mr.Asif
'            duplicate the Previous row and create the previous row as header
               For j = 1 To .Cols - 1
                  .TextMatrix(i, j) = .TextMatrix(iHeaderRow, j)
               Next j

               .TextMatrix(iHeaderRow, 0) = "+"

               .TextMatrix(i, 0) = "-"
               .RowHeight(i) = 0
               ReDim Preserve szStatementReconcile(UBound(szStatementReconcile) + 1)
               i = i + 1
               .TextMatrix(iHeaderRow, 4) = "BATCH RECEIPT"
               .AddItem ""
            End If
            .TextMatrix(i, 0) = "-"
            .RowHeight(i) = 0

            If .TextMatrix(iHeaderRow, 3) = "Sales Receipt" Or _
                  .TextMatrix(iHeaderRow, 3) = "Sales Receipt on Account" Or _
                  .TextMatrix(iHeaderRow, 3) = "Bank Receipt" Or _
                  .TextMatrix(iHeaderRow, 3) = "Purchase Payment Refund" Then
               .TextMatrix(iHeaderRow, 6) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 6)), "0.00")
                                                   
        '   Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
        '   Calcaluting the total batch receipts that are reconcilced and assigning to the batch
        '   receipt header that is grouped based on the transaction date
        
                If szaTran(1) = "Full" Then
                     .TextMatrix(iHeaderRow, 8) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 8)), "0.00")
                End If
        '   END
        
            Else
               .TextMatrix(iHeaderRow, 7) = Format(adoSI.Fields.Item("AMT").Value + _
                                                   Val(.TextMatrix(iHeaderRow, 7)), "0.00")
            End If
         End If

         .TextMatrix(i, 1) = adoSI.Fields.Item("TD").Value 'transaction date
         .TextMatrix(i, 2) = adoSI.Fields.Item("TID").Value 'Invoice number
         'this array shall be used for recheck if some new transaction has come
         szStatementReconcile(i) = .TextMatrix(i, 2) ' MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID=transaction number
         .TextMatrix(i, 3) = adoSI.Fields.Item("TT").Value
         If .TextMatrix(i, 4) = "" Then .TextMatrix(i, 4) = adoSI.Fields.Item("ACN").Value

         .TextMatrix(i, 5) = IIf(IsNull(adoSI.Fields.Item("REF").Value), "", _
                                                             adoSI.Fields.Item("REF").Value)
         If adoSI.Fields.Item("TT").Value = "Sales Receipt" Or _
               adoSI.Fields.Item("TT").Value = "Sales Receipt on Account" Or _
               adoSI.Fields.Item("TT").Value = "Bank Receipt" Or _
               adoSI.Fields.Item("TT").Value = "Purchase Payment Refund" Then

'            .TextMatrix(i, 6) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            If szaTran(1) = "Part" Then
               .TextMatrix(i, 6) = Format(Val(adoSI.Fields.Item("AMT").Value) - _
                                       Val(adoSI.Fields.Item("Reconciled").Value), "0.00")
            Else
               .TextMatrix(i, 6) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            End If
         Else
'            .TextMatrix(i, 7) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            If szaTran(1) = "Part" Then
               .TextMatrix(i, 7) = Format(Val(adoSI.Fields.Item("AMT").Value) + _
                                       Val(adoSI.Fields.Item("Reconciled").Value), "0.00")
            Else
               .TextMatrix(i, 7) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
            End If
         End If

         .TextMatrix(i, 8) = IIf(IsNull(adoSI.Fields.Item("Reconciled").Value), "", _
                                 Format(adoSI.Fields.Item("Reconciled").Value, "0.00"))
'         .TextMatrix(i, 8) = "0.00"

         .TextMatrix(i, 9) = szaTran(1)
         .TextMatrix(i, 12) = szaTran(0)
         If szaTran(1) = "Saved" Then .TextMatrix(i, 11) = "M"
         .TextMatrix(i, btRecColNo) = adoSI.Fields.Item("TransactionID").Value
         'Locking mechanism starts
         
          rsUserSessionID = IIf(IsNull(adoSI!UserSessionID), "", adoSI!UserSessionID)
         .TextMatrix(i, 13) = IIf(IsNull(adoSI!UserSessionID), "", adoSI!UserSessionID) 'Keeping the USersesssionID to check the lock
         .TextMatrix(i, 14) = IIf(IsNull(adoSI!WindowsUserName), "", adoSI!WindowsUserName) 'BP.UserSessionID,BP.WindowsUserName,BP.MachineName,BP.Module,BP.ClientID
         .TextMatrix(i, 15) = IIf(IsNull(adoSI!MachineName), "", adoSI!MachineName)
         .TextMatrix(i, 16) = IIf(IsNull(adoSI!Module), "", adoSI!Module)
         .TextMatrix(i, 17) = IIf(IsNull(adoSI!ClientID), "", adoSI!ClientID)
         .TextMatrix(i, 18) = IIf(IsNull(adoSI!TableName), "", adoSI!TableName)
         .TextMatrix(i, 19) = "" ' For future use
         .TextMatrix(i, 20) = IIf(IsNull(adoSI!PROPID), "", adoSI!PROPID)
        If UCase(.TextMatrix(i, 16)) = "CASHBOOK" And otherPcsCashbookIsOpen = False Then
            otherPcsCashbookIsOpen = True 'this shall now know that other pc has cashbook open
        End If
        
         If adoSI.Fields.Item("TableName").Value = "tlbPayment" Then
                 If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                    colTransactionIDOtherPayment = colTransactionIDOtherPayment & adoSI.Fields.Item("TransactionID").Value & ","
                    OtherScnsessionIDP = .TextMatrix(i, 13)
                    OtherWindowsUserNameP = .TextMatrix(i, 14)
                    OtherMechineNameP = .TextMatrix(i, 15)
                    OtherScnIPP = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                 Else
                    colTransactionIDPayment = colTransactionIDPayment & adoSI.Fields.Item("TransactionID").Value & ","
                 End If
         ElseIf adoSI.Fields.Item("TableName").Value = "tlbReceipt" Then
                 If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                     colTransactionIDOtherReceipt = colTransactionIDOtherReceipt & adoSI.Fields.Item("TransactionID").Value & ","
                    OtherScnsessionIDR = .TextMatrix(i, 13)
                    OtherWindowsUserNameR = .TextMatrix(i, 14)
                    OtherMechineNameR = .TextMatrix(i, 15)
                    OtherScnIPR = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                 Else
                     colTransactionIDReceipt = colTransactionIDReceipt & adoSI.Fields.Item("TransactionID").Value & ","
                 End If
         ElseIf adoSI.Fields.Item("TableName").Value = "tlbBankPayment" Then
                If Len(rsUserSessionID) > 0 And UserSessionID <> rsUserSessionID Then
                    .col = 0
                    .row = i
                    .CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                    colTransactionIDOtherBankReceipt = colTransactionIDOtherBankReceipt & "'" & adoSI.Fields.Item("TransactionID").Value & "',"
                    OtherScnsessionIDB = .TextMatrix(i, 13)
                    OtherWindowsUserNameB = .TextMatrix(i, 14)
                    OtherMechineNameB = .TextMatrix(i, 15)
                    OtherScnIPB = IIf(IsNull(adoSI!ServerIPaddress), "", adoSI!ServerIPaddress)
                Else
                    colTransactionIDBankPayment = colTransactionIDBankPayment & "'" & adoSI.Fields.Item("TransactionID").Value & "',"
                End If
         End If
         'Locking mechanism Ends

         i = i + 1

         adoSI.MoveNext
'Removed by Asif. Issue 0000523. Refer to comment above
'         If Not adoSI.EOF Then .AddItem ""
      Wend
   End With

   Label1(27).Caption = Format(UnclearedBalance, "0.00")
   Sum_RptPayVal

   adoSI.Close
    If Len(colTransactionIDOtherPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherPayment = Left(colTransactionIDOtherPayment, Len(colTransactionIDOtherPayment) - 1)
   End If
    If Len(colTransactionIDOtherReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherReceipt = Left(colTransactionIDOtherReceipt, Len(colTransactionIDOtherReceipt) - 1)
   End If
    If Len(colTransactionIDOtherBankReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherBankReceipt = Left(colTransactionIDOtherBankReceipt, Len(colTransactionIDOtherBankReceipt) - 1)
   End If
 
   If Len(colTransactionIDPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDPayment = Left(colTransactionIDPayment, Len(colTransactionIDPayment) - 1)
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionIDPayment & ")"
                        haveYouLockedAnyReccord = True
   End If
   If Len(colTransactionIDReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDReceipt = Left(colTransactionIDReceipt, Len(colTransactionIDReceipt) - 1)
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionIDReceipt & ")"
                        haveYouLockedAnyReccord = True
   End If
   If Len(colTransactionIDBankPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDBankPayment = Left(colTransactionIDBankPayment, Len(colTransactionIDBankPayment) - 1)
        adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                        SystemUser & "',MachineName='" & WS_Name & "'," & _
                        "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID in (" & colTransactionIDBankPayment & ")"
                        haveYouLockedAnyReccord = True
   End If
'  SOB --> Statement Opening Balance
   szSQL = "SELECT ClosingBal, SOB FROM ConsolidatedBankList " & _
           "WHERE conBankID = " & SelectedConBankID & ""
   adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoSI.EOF Then
      txtStOpenBal.text = "0.00"
   Else
      If IsNull(adoSI.Fields.Item("ClosingBal").Value) Then
         lblClosingBalance.Caption = "0.00"
         txtStOpenBal.Locked = False
      Else
         txtStOpenBal.Locked = True
         lblClosingBalance.Caption = Format(adoSI.Fields.Item("ClosingBal").Value, "0.00")
         'added by anol 05052016
         If lblClosingBalance.Caption = "0.00" Then
            txtStOpenBal.Locked = False
         End If
      End If
      If IsNull(adoSI.Fields.Item("SOB").Value) Then
         txtStOpenBal.text = lblClosingBalance.Caption
         'resolved by BOSL
         'issue 523 bank reconciliation is now showing correct opening and closing balance.
         'Modified by anol 19 Jan 2015
         'UpdateBankOpeningBalance adoConn, CCur(txtStOpenBal.text), txtbc.Tag
         adoConn.Execute "UPDATE ConsolidatedBankList " & _
                   "SET SOB = " & CCur(txtStOpenBal.text) & " " & _
                   "WHERE  conBankID = " & SelectedConBankID & ""
      Else
         txtStOpenBal.text = Format(adoSI.Fields.Item("SOB").Value, "0.00")
      End If
   End If

   Set adoSI = Nothing
End Sub
'Private Sub cboBC_GotFocus()
'    SelTxtInCtrl cboBC
'   If txtClientList.text = "" Then
'      MsgBox "Please select a client first.", vbOKOnly + vbExclamation, "Cashbook"
'      cmdClientList.SetFocus
'      Exit Sub
'   End If
'End Sub
 
'Private Sub cboBC_LostFocus()
'   If cboBC.ListIndex = -1 And Len(txtbc.Text) > 0 Then
'      txtbc.Text = ""
'      Exit Sub
'   End If
'End Sub

'Private Sub cboClientID_Click()
'   'issue 523
'   'closing balance was not refreshing on client change
'   lblClosingBalance.Caption = "0.00"
'   txtStatementDate.text = ""
'   'End of modification
'
'   ClearForm
'
'   If cboClientID.ListIndex < 0 Or txtClientList.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'
'   On Error GoTo ErrorHandler
'
'   adoConn.Open getConnectionString
'
'   szAllBankBalance = BankAndBalance(adoConn)
'
'NoRes:
'   adoConn.Close
'   Set adoConn = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Function BankAndBalance(adoConn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim rRow As Integer
   If txtClientList.text = "Consolidated" Then    'we are never loading ALL into the clientlist rather we are loading con.
        szSQL = "SELECT  * from ConsolidatedBankList order by conBankID;"
   Else
         szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
   End If
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
   'Modified by anol 22 Feb 2015
            If txtClientList.text <> "Consolidated" Then
                    MsgBox "Please setup your Client Bank Accounts." & Chr(13) & _
                           "Please also check the nominal chart of account for the client."
             End If
             
   Else
        If txtClientList.text = "Consolidated" Then ' this if part is for consolidated options
              rRow = 1
                While Not adoRST.EOF
                    flxClient.row = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = adoRST.Fields.Item("BankName").Value
                    flxClient.TextMatrix(rRow, 2) = adoRST.Fields.Item("BankACNumber").Value
                    flxClient.TextMatrix(rRow, 3) = adoRST.Fields.Item("SortCode").Value
                    flxClient.TextMatrix(rRow, 4) = adoRST.Fields.Item("conBankID").Value
                    flxClient.TextMatrix(rRow, 5) = IIf(IsNull(adoRST.Fields.Item("BankCode").Value), "", adoRST.Fields.Item("BankCode").Value)
                    flxClient.TextMatrix(rRow, 6) = IIf(IsNull(adoRST.Fields.Item("StatementDate").Value), "", adoRST.Fields.Item("StatementDate").Value)
                    flxClient.TextMatrix(rRow, 7) = IIf(IsNull(adoRST.Fields.Item("ClosingBal").Value), "0", adoRST.Fields.Item("ClosingBal").Value)
                    flxClient.TextMatrix(rRow, 8) = IIf(IsNull(adoRST.Fields.Item("SOB").Value), "0", adoRST.Fields.Item("SOB").Value)
                    flxClient.RowHeight(rRow) = 280
                    adoRST.MoveNext
                    If Not adoRST.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
        Else
                rRow = 1
                While Not adoRST.EOF
                    flxClient.row = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = adoRST.Fields.Item("BNC").Value
                    flxClient.TextMatrix(rRow, 2) = adoRST.Fields.Item("BNN").Value
                    flxClient.TextMatrix(rRow, 3) = adoRST.Fields.Item("ID").Value 'this is tlbclientbank MY_ID field
                    flxClient.TextMatrix(rRow, 4) = adoRST.Fields.Item("CLIENT_ID").Value

                    flxClient.RowHeight(rRow) = 280
                    adoRST.MoveNext
                    If Not adoRST.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
         End If
   End If

   ' Destroy Objects
   Set adoRST = Nothing

   LoadAdoBank 'yes we are using this here for view bank general infor

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Function

Public Function LoadAdoBank()
   Dim szSQL As String
' I am not loading all into the bank list
   If txtClientList.Tag = "ALL" Then
      szSQL = "SELECT CB.NominalCode, CB.Bank_AC_Name, CB.BankMemo, " & _
                  "B.BANK_NAME, B.SORT_CODE, CB.BANK_AC_NUM, CB.AccountType, " & _
                  "CB.AccountType, N.Name, CB.PaymentMethod, CB.BacsRef, CB.LRSD, " & _
                  "B.Contact, B.Tel, B.Fax, B.Mobile, B.eMail, B.Website, B.BANK_ID, " & _
                  "CB.PCB, CB.spare2 " & _
              "FROM tlbClientBanks AS CB, tlbBank AS B, NominalLedger AS N " & _
              "WHERE CB.BANK_ID = B.BANK_ID AND " & _
                  "CB.NominalCode = N.Code " & _
              "AND N.clieNtID=CB.Client_ID ORDER BY NominalCode;"
   Else
      szSQL = "SELECT CB.NominalCode, CB.Bank_AC_Name, CB.BankMemo, " & _
                  "B.BANK_NAME, B.SORT_CODE, CB.BANK_AC_NUM, CB.AccountType, " & _
                  "CB.AccountType, N.Name, CB.PaymentMethod, CB.BacsRef, CB.LRSD, " & _
                  "B.Contact, B.Tel, B.Fax, B.Mobile, B.eMail, B.Website, B.BANK_ID, " & _
                  "CB.PCB, CB.spare2 " & _
              "FROM tlbClientBanks AS CB, tlbBank AS B, NominalLedger AS N " & _
              "WHERE CLIENT_ID = '" & txtClientList.Tag & "' AND " & _
                  "CB.BANK_ID = B.BANK_ID AND " & _
                  "CB.NominalCode = N.Code " & _
              "AND N.clieNtID=CB.Client_ID ORDER BY NominalCode;"
   End If
'Debug.Print szSQL
   adoBank.ConnectionString = getConnectionString
   adoBank.RecordSource = szSQL
   adoBank.CommandType = adCmdText
   adoBank.Refresh
End Function

Public Function PopulateTenantLookup(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRST As New ADODB.Recordset

   adoRST.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   While Not adoRST.EOF
      flxLeaseList.TextMatrix(iRow, 1) = adoRST!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 2) = adoRST!Name
      flxLeaseList.TextMatrix(iRow, 3) = adoRST!UnitNumber

      iRow = iRow + 1
      adoRST.MoveNext

      If Not adoRST.EOF Then flxLeaseList.AddItem ""
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Function

'Private Sub cboClientID_GotFocus()
'    SelTxtInCtrl txtClientList
'End Sub

'Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cmdbc.SetFocus
'    End If
'End Sub

Private Sub cboClientName_Change()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String, iRec As Integer

   On Error GoTo ErrorHandler
   adoConn.Open getConnectionString

   szSQL = "SELECT DISTINCT tlbClientBanks.CLIENT_ID, Client.ClientName, " & _
                           "tlbClientBanks.BANK_AC_NUM, NominalCode " & _
           "FROM tlbClientBanks, Client " & _
           "WHERE Client.ClientName = '" & cboClientName.text & "' AND " & _
                 "tlbClientBanks.CLIENT_ID = Client.ClientID " & _
           "ORDER BY CLIENT_ID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   cboANF.Clear

   If adoRST.RecordCount < 2 Then
      MsgBox "This Client has only one account. It's not possible make the transfer.", vbCritical + vbOKOnly, "Account Number Selection"
      FocusControl cboClientName
      adoRST.Close
      adoConn.Close
      Set adoRST = Nothing
      Set adoConn = Nothing
      Exit Sub
   Else
      cboANF.Enabled = True
      Label13(21).Caption = IIf(IsNull(adoRST!CLIENT_ID), "", adoRST!CLIENT_ID)

      ReDim szaData(1, adoRST.RecordCount - 1) As String
      iRec = 0
      While Not adoRST.EOF
         szaData(0, iRec) = adoRST.Fields.Item("BANK_AC_NUM").Value
         szaData(1, iRec) = adoRST.Fields.Item("NominalCode").Value
         iRec = iRec + 1
         adoRST.MoveNext
      Wend

      cboANF.Column() = szaData()
   End If
   FocusControl cboANF
NoRes:
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmbSPSupplier_Click()
   Frame5(5).Enabled = True

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadFlxSPayment adoConn
   LoadFlxSCrPoA adoConn

   adoConn.Close
   Set adoConn = Nothing

   ReDim baChangesMade(flxSPayment.Rows) As Boolean

   Frame5(0).Enabled = True
End Sub

Private Sub cmdAddTrans_Click()
   tabCashbook.Tab = 1
End Sub

Private Sub cmdBankCancel_Click()
   fraBank.Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True
   FocusControl cmdNewBk(0)
End Sub

Private Sub cmdBankReceiptHistory_Click()
   If Not BANK_PAYMENT_HISTORY_LOADED Then
      Load frmBankPaymentHistory
   End If

   frmBankPaymentHistory.Show
End Sub

Private Sub cmdBankTransfer_Click()
   If txtBC.text = "" Or txtClientList.text = "" Then
      If txtClientList.text = "" Then
         MsgBox "Please select a client account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtClientList
         Exit Sub
      Else
         MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtBC
         Exit Sub
      End If
   End If

   Load frmBankTransactions
'   frmDemands3.tabDmdRcpt.Tab = 2
'   frmDemands3.tabPayment.Tab = 2
   frmBankTransactions.Show
   frmBankTransactions.ZOrder 0
End Sub

Private Sub cmdBRP_Click()
   If txtBC.text = "" Or txtClientList.text = "" Then
      If txtClientList.text = "" Then
         MsgBox "Please select a client account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtClientList
         Exit Sub
      Else
         MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtBC
         Exit Sub
      End If
   End If

   Load frmBankTransactions
'   frmDemands3.tabDmdRcpt.Tab = 2
'   frmDemands3.tabPayment.Tab = 1
   frmBankTransactions.Show
   frmBankTransactions.ZOrder 0
End Sub

Private Sub cmdBTCancel_Click()
   cboClientName.text = ""
   Label13(21).Caption = ""
   cboANF.text = ""
   Label13(13).Caption = ""
   Label13(19).Caption = ""
   cboANT.text = ""
   Label13(20).Caption = ""
   Label13(19).Caption = ""
   Label13(17).Caption = ""
   txtBkTrDate.text = ""
   txtBkTrRef.text = ""
   cboFundBankTransf.text = ""
   txtBkTrDes.text = ""
   txtBkTrAmt.text = ""
   cboANF.Enabled = False
   cboANT.Enabled = False
   cboFundBankTransf.Enabled = False
End Sub

Private Sub cmdBTSave_Click()
'   If cboClientName.text = "" Then
'      MsgBox "Please select the Client first.", vbInformation + vbOKOnly, "Transfer"
'      cboClientName.SetFocus
'      Exit Sub
'   End If
'   If cboANF.text = "" Then
'      MsgBox "Please select the Account Number from first.", vbInformation + vbOKOnly, "Transfer"
'      If cboANF.Enabled Then cboANF.SetFocus
'      Exit Sub
'   End If
'   If cboANT.text = "" Then
'      MsgBox "Please select the Account Number to first.", vbInformation + vbOKOnly, "Transfer"
'      If cboANT.Enabled Then cboANT.SetFocus
'      Exit Sub
'   End If
'   If cboANT.text = cboANF.text Then
'      MsgBox "You need select another Account Number.", vbCritical + vbOKOnly, "Account Number Selection"
'      cboANT.SetFocus
'      Exit Sub
'   End If
'   If cboFundBankTransf.text = "" Then
'      MsgBox "Please select the Fund first.", vbInformation + vbOKOnly, "Transfer"
'      cboFundBankTransf.SetFocus
'      Exit Sub
'   End If
'   If txtBkTrAmt.text = "" Then
'      MsgBox "Please insert the Payment Value first.", vbInformation + vbOKOnly, "Transfer"
'      txtBkTrAmt.SetFocus
'      Exit Sub
'   End If
'
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL, szSQL1 As String
'   Dim valueF, valueT As String
'
'   On Error GoTo ErrorHandler
'   adoConn.Open getConnectionString
'
'   If valueF < 0 Then
'      MsgBox "Sorry, this Client hasn't a sufficient funds for this transfer!!.", vbCritical + vbOKOnly, "Insufficient Balance"
'      txtBkTrAmt.SetFocus
'      adoRst.Close
'      adoConn.Close
'      Set adoRst = Nothing
'      Set adoConn = Nothing
'      Exit Sub
'   Else
'      Dim tt As String
'      Dim foundID As String
'
'      foundID = cboFundBankTransf.BoundColumn
'
'      szSQL = "SELECT MY_ID " & _
'              "FROM tlbClientBanks " & _
'              "WHERE CLIENT_ID = '" & Label13(21).Caption & "' AND " & _
'                     "BANK_AC_NUM = '" & cboANF.text & "' AND BANK_SC = '" & Label13(13).Caption & "';"
''      Debug.Print szSQL
'      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'      tt = "DR"
'      szSQL = "INSERT INTO BankTransactions (ClientBankID, TranType, TranDate, Ref, Fund, Details, Amount) " & _
'              "VALUES (" & adoRst!MY_ID & ", '" & tt & "', " & _
'                       "'" & txtBkTrDate.text & "', '" & txtBkTrRef.text & "', " & _
'                       "" & foundID & ", '" & txtBkTrDes.text & "', " & txtBkTrAmt.text & ");"
'      adoConn.Execute szSQL
'      adoRst.Close
'
'      szSQL = "SELECT MY_ID " & _
'              "FROM tlbClientBanks  " & _
'              "WHERE CLIENT_ID = '" & Label13(21).Caption & "' AND BANK_AC_NUM = '" & cboANT.text & "' AND BANK_SC = '" & Label13(20).Caption & "'; "
'      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'      tt = "CR"
'      szSQL = "INSERT INTO BankTransactions (ClientBankID, TranType, TranDate, Ref, Fund, Details, Amount) " & _
'               "VALUES (" & adoRst!MY_ID & ", '" & tt & "', '" & txtBkTrDate.text & "', '" & txtBkTrRef.text & "', " & foundID & ", '" & txtBkTrDes.text & "', " & txtBkTrAmt.text & ");"
''      Debug.Print szSQL
'      adoConn.Execute szSQL
'   End If
'
'   MsgBox "Data has been saved successfully"
'   cboClientName.text = ""
'   Label13(21).Caption = ""
'   cboANF.text = ""
'   Label13(13).Caption = ""
'   Label13(19).Caption = ""
'   cboANT.text = ""
'   Label13(20).Caption = ""
'   Label13(19).Caption = ""
'   Label13(17).Caption = ""
'   txtBkTrDate.text = ""
'   txtBkTrRef.text = ""
'   cboFundBankTransf.text = ""
'   txtBkTrDes.text = ""
'   txtBkTrAmt.text = ""
'
'NoRes:
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
End Sub

Private Sub cmdCBHFilter_Click()
   Dim i As Integer

   If cmdCBHFilter.Caption = "<<< Clear" Then
      txtCBHDtFrm.text = ""
      txtCBHDtTo.text = ""
      txtCBHDtFrm.Locked = False
      txtCBHDtTo.Locked = False

      For i = 1 To flxCashBook.Rows - 1
         flxCashBook.RowHeight(i) = 240
      Next i
      'Resolved by BOSL
      'issue 523 cashbook total figure not refreshing
      'added by anol 19 Jan 2015
      CalDrCrCBHistory
      'End of modification
      
      cmdCBHFilter.Caption = "Filter >>>"
      FocusControl txtCBHDtFrm
   Else
      If txtCBHDtFrm.text = "" And txtCBHDtTo.text = "" Then
         For i = 1 To flxCashBook.Rows - 1
            flxCashBook.RowHeight(i) = 240
         Next i
      End If

      If txtCBHDtFrm.text = "" Then
         FocusControl txtCBHDtFrm
         Exit Sub
      End If
      If txtCBHDtTo.text = "" Then
         FocusControl txtCBHDtTo
         Exit Sub
      End If

      For i = 1 To flxCashBook.Rows - 1
         flxCashBook.RowHeight(i) = 240
      Next i

      For i = 1 To flxCashBook.Rows - 1
         If CDate(flxCashBook.TextMatrix(i, 1)) < CDate(txtCBHDtFrm.text) Or _
               CDate(flxCashBook.TextMatrix(i, 1)) > CDate(txtCBHDtTo.text) Then
            flxCashBook.RowHeight(i) = 0
         End If
      Next i

      cmdCBHFilter.Caption = "<<< Clear"
      txtCBHDtFrm.Locked = True
      txtCBHDtTo.Locked = True
      CalDrCrCBHistory
   End If
End Sub

Private Sub createTable(adoConn As ADODB.Connection)
    
     Dim adoRST As New ADODB.Recordset
     On Error GoTo CreateReportCashbookHistory
       
       adoRST.Open "SELECT * FROM ReportCashbookHistory;", adoConn, adOpenStatic, adLockReadOnly
       adoRST.Close
    
       GoTo alreadycreated
    
CreateReportCashbookHistory:
           adoConn.Execute _
              "CREATE TABLE ReportCashbookHistory " & _
                 "(" & _
                    "ReportingDate DateTime  NOT NULL, " & _
                    "SessionID     TEXT(100) NOT NULL, " & _
                    "ClientID      TEXT(10), " & _
                    "iRow      Number, " & _
                    "TDate DateTime, " & _
                    "No   TEXT(50) NOT NULL, " & _
                    "tTYpe      TEXT(100), " & _
                    "Account      TEXT(100), " & _
                    "Reference      TEXT(200), " & _
                    "Detail      TEXT(250), " & _
                    "Debit       CURRENCY, " & _
                    "Credit         CURRENCY, " & _
                    "Reconciled        TEXT(10), " & _
                    "StDate        TEXT(20), " & _
                    "PRIMARY KEY (ReportingDate, SessionID, iRow)" & _
                 ");"
        
alreadycreated:
End Sub
Private Sub cmdCbHReport_Click()
   If txtBC.text = "" Then
      MsgBox "Please select a Bank.", vbCritical + vbOKOnly, "Cashbook History"
      FocusControl txtBC
      Exit Sub
   End If
   'added by anol 08  Aug 2016
   If Trim(txtCBHDtFrm.text) = "" And cmdCBHFilter.Caption = "Filter >>>" Then
      txtCBHDtFrm.Tag = "01/01/2000"
   Else
      txtCBHDtFrm.Tag = txtCBHDtFrm.text
   End If

   If Trim(txtCBHDtTo.text) = "" And cmdCBHFilter.Caption = "Filter >>>" Then
        txtCBHDtTo.Tag = Date
   Else
        txtCBHDtTo.Tag = txtCBHDtTo.text
   End If
  
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsAdd As New ADODB.Recordset
    Call createTable(adoConn)
    Dim sessionID As String
    Dim reportingDate As String
    Dim i As Integer
    reportingDate = Format(Date, "dd mmmm yyyy")
    sessionID = GetTimeStamp
    adoConn.Execute _
    "DELETE FROM ReportCashbookHistory WHERE SessionID = '" & sessionID & "';"
    adoConn.Execute "DELETE FROM ReportCashbookHistory WHERE ReportingDate < #" & reportingDate & "# ;"
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
    rsAdd.Open " Select * from ReportCashbookHistory where 1=2", adoConn, adOpenKeyset, adLockBatchOptimistic
    With rsAdd
    For i = 1 To flxCashBook.Rows - 1
             If flxCashBook.RowHeight(i) <> 0 Then
                .AddNew
                !reportingDate = reportingDate
                !sessionID = sessionID
                !ClientID = frmCashbook.txtClientList.Tag
                !iRow = i
                !TDate = flxCashBook.TextMatrix(i, 1)
                !No = flxCashBook.TextMatrix(i, 2)
                !TType = flxCashBook.TextMatrix(i, 10)
                !Account = flxCashBook.TextMatrix(i, 3)
                !Reference = flxCashBook.TextMatrix(i, 4)
                !Detail = flxCashBook.TextMatrix(i, 5)
                !Debit = Val(flxCashBook.TextMatrix(i, 6))
                !Credit = Val(flxCashBook.TextMatrix(i, 7))
                !Reconciled = flxCashBook.TextMatrix(i, 8)
                !StDate = flxCashBook.TextMatrix(i, 9)
             End If
             .UpdateBatch
    Next
    End With
    adoConn.Close
   If txtClientList.text = "Consolidated" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookHistoryRpt_cons.rpt")
        
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
           Report.EnableParameterPrompting = False
           Report.DiscardSavedData
        
           'Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
           Report.ParameterFields(1).AddCurrentValue CDate(Format(txtCBHDtFrm.Tag, "dd mmmm yyyy"))
           Report.ParameterFields(2).AddCurrentValue CDate(Format(txtCBHDtTo.Tag, "dd mmmm yyyy"))
           Report.ParameterFields(3).AddCurrentValue frmCashbook.CalDrCrAcBalance(CDate(Format(txtCBHDtFrm.Tag, "dd mmmm yyyy")), CDate(Format(txtCBHDtTo.Tag, "dd mmmm yyyy")))
           'resolved by BOSL
           'issue 523 cashbook report was not working properly
           'added by anol 19 Jan 2015
           Report.ParameterFields(4).AddCurrentValue frmCashbook.SelectedConBankID
           Report.ParameterFields(5).AddCurrentValue sessionID
           Set rep = New frmReport
           Load rep
           rep.LoadReportViewer Report
   Else
           Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookHistoryRpt.rpt")
        
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
           Report.EnableParameterPrompting = False
           Report.DiscardSavedData
        
           Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
           Report.ParameterFields(2).AddCurrentValue CDate(Format(txtCBHDtFrm.Tag, "dd mmmm yyyy"))
           Report.ParameterFields(3).AddCurrentValue CDate(Format(txtCBHDtTo.Tag, "dd mmmm yyyy"))
           Report.ParameterFields(4).AddCurrentValue frmCashbook.CalDrCrAcBalance(CDate(Format(txtCBHDtFrm.Tag, "dd mmmm yyyy")), CDate(Format(txtCBHDtTo.Tag, "dd mmmm yyyy")))
           'resolved by BOSL
           'issue 523 cashbook report was not working properly
           'added by anol 19 Jan 2015
           Report.ParameterFields(5).AddCurrentValue frmCashbook.txtClientList.Tag
           Report.ParameterFields(6).AddCurrentValue sessionID
           Set rep = New frmReport
           Load rep
           rep.LoadReportViewer Report
   End If
   
'   frmMMain.fraCmdButton.Enabled = False
'   Me.Enabled = False
'   Load frmDtRange4CB
'   frmDtRange4CB.Show
'
'   Exit Sub
'
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'   Dim rep As frmReport
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookRptPay.rpt")
'
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue txtbc.Tag
'
'   Set rep = New frmReport
'   Load rep
'   rep.LoadReportViewer Report
'
'   Frame4(0).Visible = False
'   tabCashbook.Enabled = True
'   cboClientID.Locked = False
'   cboBC.Locked = False
End Sub

Private Sub cmdCbHRptCancel_Click()
   Frame4(0).Visible = False
   tabCashbook.Enabled = True
   txtClientList.Locked = False
   txtBC.Locked = False
End Sub

Private Sub cmdCbHRptOk_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   If optCbHRpt(0).Value Then          'All Receipt
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookReceipts.rpt")
   End If

   If optCbHRpt(1).Value Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookPayments.rpt")
   End If

   If optCbHRpt(2).Value Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookRptPay.rpt")
   End If

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtBC.Tag

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   
   Frame4(0).Visible = False
   tabCashbook.Enabled = True
   txtClientList.Locked = False
   txtBC.Locked = False
End Sub

Private Sub cmdClinetAddAtch_Click()
    ' memo attachments
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub

   AddNewAttachmentInCombo cmbFiles, "BankMemo", cmdBC.Tag 'cmdBC.Tag is ID

   ShowMsgInTaskBar "File has been saved successfully."
End Sub

Private Sub cmdCopy_Click()
   If tabCashbook.Tab < 2 Or tabCashbook.Tab > 3 Then Exit Sub

   If tabCashbook.Tab = 3 Then GoTo TAB3

TAB2:
   If flxCashBook.row < 1 Then
      MsgBox "Select a transaction from the grid.", vbInformation + vbOKOnly, "Copy/Copy Transaction"
      FocusControl flxCashBook
      Exit Sub
   End If
   
   'frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + cmdCopy.Top + Me.Top + 1150
   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + cmdCopy.Left + 80
   If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BR" Or _
         Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BP" Then
      frmPopUpMenu.CallingFrom "CB_BANK"
   End If
   If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 3) = "SRR" Then
      frmPopUpMenu.CallingFrom "CB_SRR"
   Else
      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "SR" Then frmPopUpMenu.CallingFrom "CB_SR"
      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "SA" Then frmPopUpMenu.CallingFrom "CB_SA"
   End If
   
   
   frmPopUpMenu.Show


TAB3:
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub

   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), cmdBC.Tag, "BankMemo" 'cmdBC.Tag is ID

   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdHistoryReport_Click()
   If txtBC.text = "" Then
  
      MsgBox "Please select Bank code", vbInformation + vbOKOnly, "Bank Reconciliation History Report"
      FocusControl cmdBC
     
      Exit Sub
   End If

   Me.Enabled = False
   Load frmBkRecHistReport
   frmBkRecHistReport.szBankID = cmdBC.Tag
   
   frmBkRecHistReport.Show
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdPayAllocate_Click()
   If cmbSPSupplier.text = "" Then Exit Sub

   Dim iRow As Integer

   If cmdPayAllocate.Caption = "All&ocation Only" Then
      Frame5(5).Visible = False
      Frame5(1).Left = Frame5(5).Left
      Frame5(1).Top = Frame5(5).Top
      Frame5(1).Visible = True
      cmdPayAllocate.Caption = "&Payment Only"
      Label3(5).Visible = True
      txtAllocatedDiff(1).Visible = True
      Label3(4).Visible = True

      cTotalAdjI = 0
      cTotalSI = 0

      For iRow = 1 To flxSPayment.Rows - 2
         If (flxSPayment.TextMatrix(iRow, 2) = "AdjI") Then
            cTotalAdjI = cTotalAdjI + CInt(flxSPayment.TextMatrix(iRow, 8))
         Else
            cTotalSI = cTotalSI + CInt(flxSPayment.TextMatrix(iRow, 8))
         End If
      Next iRow

      Frame5(1).Enabled = True
   Else
      Dim adoConn As New ADODB.Connection

      Frame5(5).Visible = True
      Frame5(1).Visible = False
      cmdPayAllocate.Caption = "All&ocation Only"
      Label3(5).Visible = False
      txtAllocatedDiff(1).Visible = False
      Label3(4).Visible = False
      lblAllocating(1).Visible = False

      adoConn.Open getConnectionString

      ConfigureFlxSCrPoA
      ConfigureFlxSPayment
      LoadFlxSPayment adoConn
      LoadFlxSCrPoA adoConn

      adoConn.Close
      Set adoConn = Nothing
   End If
End Sub

Private Sub cmdPayAllocateSave_Click()
   If MsgBox("Do you wish to save allocations?", vbQuestion + vbYesNo, "Allocation") = vbYes Then
      If SavingAllocationS Then
         ShowMsgInTaskBar "Allocations have been saved successfully."

         flxSPayment.Enabled = True
         flxSCrPoA.Enabled = True
         cmdPayAllocateSave.Enabled = False
         lblAllocating(1).Visible = False
         Frame5(5).Enabled = True                     'Payment - Saving
         Frame5(1).Enabled = False                    'Allocation - Saving
         txtAllocatedDiff(1).text = "0.00"
         ConfigureFlxSPayment
         cmdPayAutomatic.Enabled = True
         cmdPayAllocate_Click
      End If
   End If
End Sub

Private Sub cmdPayAutomatic_Click()
   Frame4(1).Left = tabCashbook.Left + tabPayRpt.Left + Frame5(1).Left + Frame5(1).Width - 60
   Frame4(1).Top = tabCashbook.Top + tabPayRpt.Top + Frame5(1).Top - Frame4(1).Height + Frame5(1).Height - 20
   Frame4(1).Visible = True
   FocusControl cmdAutoAllocSel

   tabCashbook.Enabled = False
   tabPayRpt.Enabled = False
End Sub

Private Sub cmdPaymentDiscard_Click()
   If MsgBox("Do you wish to discard all changes?", vbQuestion + vbYesNo, "Payment") = vbNo Then Exit Sub

   Dim iRow As Integer

   If bChangesMade Then
      For iRow = 1 To flxSPayment.Rows - 2
         If Val(flxSPayment.TextMatrix(iRow, 10)) > 0 Then
            flxSPayment.TextMatrix(iRow, 10) = "0.00"
         End If
      Next iRow
      txtPaymentEntered.text = "0.00"
      flxSCrPoA.Enabled = True
      txtSPaymentTotal.text = "0.00"
      bChangesMade = False
      cGridRptTotal = 0
      bTotalPayTyped = False
   End If
End Sub

Private Sub cmdReconcile_Click()
   If flxStatementReconcile.row > 0 Then
      flxStatementReconcile_DblClick
   End If
End Sub

Private Sub cmdReconPrint_Click()
On Error GoTo Err
   If txtClientList.text = "" Then
      MsgBox "Please select a client", vbInformation, "Warning"
      FocusControl txtClientList
      Exit Sub
   End If
   If txtBC.text = "" Then
      MsgBox "Please select a bank", vbInformation, "Warning"
      FocusControl cmdBC
      Exit Sub
   End If
   If IsDate(txtStatementDate.text) = False Then
        MsgBox "Please select a Statement Date", vbInformation, "Warning"
      FocusControl txtStatementDate
      Exit Sub
   End If
   If frmCashbook.txtClientList.text = "Consolidated" Then
           
              Dim reportApp As New CRAXDRT.Application
              Dim Report As CRAXDRT.Report
              Dim rep As frmReport
              If optReconciliation(0).Value Then              'Show unreconciled transactions only
                    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookUnReconSaved_Cons.rpt")
              Else
                     Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookReconUnreconSaved_Cons.rpt")
              End If
        
              Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
              Report.EnableParameterPrompting = False
              Report.DiscardSavedData
        
              Report.ParameterFields(1).AddCurrentValue txtBC.Tag 'Passing bank code
              Report.ParameterFields(2).AddCurrentValue txtAcBal.text 'passing Bank Balance
              Report.ParameterFields(3).AddCurrentValue Format(txtStOpenBal.text, "0.00") 'Opening statement Balance
              Report.ParameterFields(4).AddCurrentValue CDate(txtStatementDate.text) 'statement date but this have no use in actual report
              'Below one is Unreconciled Balance: Val("333") '
              Report.ParameterFields(5).AddCurrentValue CDbl(Val(txtAcBal.text) - Val(lblClosingBalance.Caption))  ' difference between sum of total of dr and cr (Name unreconbal in parameter)
              Report.ParameterFields(6).AddCurrentValue CDbl(txtStOpenBal.text) 'This is stclosingbal in report
              'added by anol  26 Jan 2015
              'issue 523
               Report.ParameterFields(7).AddCurrentValue (SelectedConBankID)
              Set rep = New frmReport
              Load rep
              rep.LoadReportViewer Report
              Exit Sub
           
   Else
            If optReconciliation(0).Value Then              'Show unreconciled transactions only
'                      Dim reportApp As New CRAXDRT.Application
'                      Dim Report As CRAXDRT.Report
'                      Dim rep As frmReport
'
                      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookUnReconSaved.rpt")
                
                      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
                
                      Report.EnableParameterPrompting = False
                      Report.DiscardSavedData
                
                      Report.ParameterFields(1).AddCurrentValue txtBC.Tag
                      Report.ParameterFields(2).AddCurrentValue txtAcBal.text
                      Report.ParameterFields(3).AddCurrentValue Format(txtStOpenBal.text, "0.00")
                      Report.ParameterFields(4).AddCurrentValue CDate(txtStatementDate.text)
                      'Below one is Unreconciled Balance: Val("333") '
                      Report.ParameterFields(5).AddCurrentValue CDbl(Val(Label1(46).Caption) - Val(Label1(47).Caption))
                      Report.ParameterFields(6).AddCurrentValue CDbl(txtStOpenBal.text)
                      'added by anol  26 Jan 2015
                      'issue 523
                       Report.ParameterFields(7).AddCurrentValue txtClientList.Tag
                      Set rep = New frmReport
                      Load rep
                      rep.LoadReportViewer Report
                      Exit Sub
                   End If
   End If
'  Show all transactions
   'frmMMain.fraCmdButton.Enabled = False
   Me.Enabled = False
   Load frmDtRange4BRec
   frmDtRange4BRec.Top = Me.Top + (Me.Height / 2) - (frmDtRange4BRec.Height / 2)
   frmDtRange4BRec.Left = Me.Left + (Me.Width / 2) - (frmDtRange4BRec.Width / 2)
   frmDtRange4BRec.Show
   Exit Sub
Err:
   MsgBox Err.description
End Sub

Private Sub cmdRefresh_Click()
   cboBC_Click
End Sub

Private Sub cmdRptAllocate_Click()
   If txtTenantID.text = "" Then Exit Sub

   Dim iRow As Integer

   If cmdRptAllocate.Caption = "All&ocation Only" Then
      Frame5(0).Visible = False
      Frame5(4).Left = Frame5(0).Left
      Frame5(4).Top = Frame5(0).Top
      Frame5(4).Visible = True
      cmdRptAllocate.Caption = "&Receipt Only"
      Label3(1).Visible = True
      txtAllocatedDiff(0).Visible = True
      Label3(2).Visible = True

      cTotalAdjI = 0
      cTotalSI = 0

      For iRow = 1 To flxTReceipt.Rows - 2
         If (flxTReceipt.TextMatrix(iRow, 2) = "AdjI") Then
            cTotalAdjI = cTotalAdjI + CInt(flxTReceipt.TextMatrix(iRow, 8))
         Else
            cTotalSI = cTotalSI + CInt(flxTReceipt.TextMatrix(iRow, 8))
         End If
      Next iRow

      Frame5(4).Enabled = True
   Else
      Dim adoConn As New ADODB.Connection

      Frame5(0).Visible = True
      Frame5(4).Visible = False
      cmdRptAllocate.Caption = "All&ocation Only"
      Label3(1).Visible = False
      txtAllocatedDiff(0).Visible = False
      Label3(2).Visible = False
      lblAllocating(0).Visible = False

      adoConn.Open getConnectionString

      ConfigureFlxTCrPoA
      ConfigureFlxTReceipt
      LoadFlxTReceipt adoConn
      LoadFlxSCrPoA adoConn

      adoConn.Close
      Set adoConn = Nothing
   End If
End Sub

Private Sub cmdRptAllocateSave_Click()
   If MsgBox("Do you wish to save allocations?", vbQuestion + vbYesNo, "Allocation") = vbYes Then
      If SavingAllocationR Then
         ShowMsgInTaskBar "Allocations have been saved successfully."

         flxTReceipt.Enabled = True
         flxTCrPoA.Enabled = True
         cmdRptAllocateSave.Enabled = False
         lblAllocating(0).Visible = False
         Frame5(0).Enabled = True                     'Receipt - Saving
         Frame5(4).Enabled = False                    'Allocation - Saving
         txtAllocatedDiff(0).text = "0.00"
         ConfigureFlxTReceipt
         cmdRptAutomatic.Enabled = True
         cmdRptAllocate_Click
      End If
   End If
End Sub

Private Function SavingAllocationS() As Boolean
   Dim iRow  As Integer, szSQL  As String
   Dim lRT_ID As Long
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset

   adoConn.Open getConnectionString

'     find NEXT transaction number of tlbPayment table
   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM tlbPayment"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   rstRst.Close

   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM PayTransactions;"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
   rstRst.Close

'  Update the credit Out-Standing amount by receipt amount
   If CCur(flxSCrPoA.TextMatrix(Label10(6).Caption, 8)) - _
      CCur(flxSCrPoA.TextMatrix(Label10(6).Caption, 9)) = 0 Then
      szSQL = "UPDATE tlbPayment " & _
              "SET OSAmount = " & CCur(flxSCrPoA.TextMatrix(Label10(6).Caption, 8)) - _
                                  CCur(flxSCrPoA.TextMatrix(Label10(6).Caption, 9)) & ", " & _
                  "PaymentView = False " & _
              "WHERE TransactionID = " & CLng(Label10(5).Caption) & ";"
   Else
      szSQL = "UPDATE tlbPayment " & _
              "SET OSAmount = " & flxSCrPoA.TextMatrix(Label10(6).Caption, 8) - _
                                  flxSCrPoA.TextMatrix(Label10(6).Caption, 9) & " " & _
              "WHERE TransactionID = " & CLng(Label10(5).Caption) & ";"
   End If
   adoConn.Execute szSQL

'  Update the Invoice out standing amount by receipt amount
   szSQL = "SELECT * FROM PayTransactions;"
   rstRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 15) = "A" Then
         If CCur(flxSPayment.TextMatrix(iRow, 9)) - _
             CCur(flxSPayment.TextMatrix(iRow, 10)) = 0 Then                'OutStanding Amount = 0
             szSQL = "UPDATE tlbPayment " & _
                    "SET tlbPayment.OSAmount = " & _
                        CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                         CCur(flxSPayment.TextMatrix(iRow, 10)) & ", " & _
                        "PaymentView = False " & _
                    "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(iRow, 19)) & ";"
         Else                                                               'OutStanding Amount > 0
            szSQL = "UPDATE tlbPayment " & _
                    "SET tlbPayment.OSAmount = " & _
                        CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                         CCur(flxSPayment.TextMatrix(iRow, 10)) & " " & _
                    "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(iRow, 19)) & ";"
         End If
         adoConn.Execute szSQL

'         ~CREDIT NOTE~ & ~RECEIPT ON ACCOUNT~
         If flxSCrPoA.TextMatrix(Label10(6).Caption, 13) = 7 Or _
            flxSCrPoA.TextMatrix(Label10(6).Caption, 13) = 9 Then
            With rstRst
               .AddNew
               !TranType = IIf(UCase(flxSCrPoA.TextMatrix(Label10(6).Caption, 1)) = "ADJC", "ADJAL", "AL")
               !TransactionID = lRT_ID
               !Alloc_Unalloc = 1
               !FromTran = CLng(Label10(5).Caption)
               !ToTran = CLng(flxSPayment.TextMatrix(iRow, 19))
               !AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
               !PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 10))
               !Discount = 0
               !IsSageUpdate = IIf(UCase(flxSCrPoA.TextMatrix(Label10(6).Caption, 1)) = "ADJC", False, True)
               !UpdateSage = False
               !BankCode = txtBC.Tag
               !nominalCode = rstRst!BankCode
               lRT_ID = lRT_ID + 1
               .Update
            End With
         End If
      End If
   Next iRow
   ConfigureFlxSCrPoA
   ConfigureFlxSPayment
   cmbSPSupplier.text = ""
   txtBC.text = ""
   txtPaymentEntered.text = "0.00"
   txtSPaymentTotal.text = "0.00"

   flxSCrPoA.Enabled = True
   ReDim baChangesMade(flxSPayment.Rows) As Boolean

   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   SavingAllocationS = True
End Function

Private Function SavingAllocationR() As Boolean
   Dim iRow  As Integer, lT_ID As Long
   Dim lRT_ID As Long, szSQL  As String
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset

   adoConn.Open getConnectionString

'     find NEXT transaction number of tlbReceipt table
   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM tlbReceipt"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
   rstRst.Close

   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM RptTransactions;"
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
   rstRst.Close

'  Update the credit Out-Standing amount by receipt amount
   If CCur(flxTCrPoA.TextMatrix(Label10(4).Caption, 8)) - _
      CCur(flxTCrPoA.TextMatrix(Label10(4).Caption, 9)) = 0 Then
      szSQL = "UPDATE tlbReceipt " & _
              "SET    OSAmount = " & CCur(flxTCrPoA.TextMatrix(Label10(4).Caption, 8)) - _
                                  CCur(flxTCrPoA.TextMatrix(Label10(4).Caption, 9)) & ", " & _
                  "ReceiptView = False " & _
              "WHERE TransactionID = " & CLng(Label10(2).Caption) & ";"
   Else
      szSQL = "UPDATE tlbReceipt " & _
              "SET    OSAmount = " & flxTCrPoA.TextMatrix(Label10(4).Caption, 8) - _
                                  flxTCrPoA.TextMatrix(Label10(4).Caption, 9) & " " & _
              "WHERE TransactionID = " & CLng(Label10(2).Caption) & ";"
   End If
   adoConn.Execute szSQL

'  Update the Invoice out standing amount by receipt amount
   szSQL = "SELECT * FROM RptTransactions;"
   rstRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxTReceipt.Rows - 1
      If flxTReceipt.TextMatrix(iRow, 15) = "A" Then
         If CCur(flxTReceipt.TextMatrix(iRow, 9)) - _
             CCur(flxTReceipt.TextMatrix(iRow, 10)) = 0 Then                'OutStanding Amount = 0
             szSQL = "UPDATE tlbReceipt " & _
                    "SET tlbReceipt.OSAmount = " & _
                        CCur(flxTReceipt.TextMatrix(iRow, 9)) - _
                         CCur(flxTReceipt.TextMatrix(iRow, 10)) & ", " & _
                        "ReceiptView = False " & _
                    "WHERE tlbReceipt.TransactionID = " & CLng(flxTReceipt.TextMatrix(iRow, 19)) & ";"
         Else                                                               'OutStanding Amount > 0
            szSQL = "UPDATE tlbReceipt " & _
                    "SET tlbReceipt.OSAmount = " & _
                        CCur(flxTReceipt.TextMatrix(iRow, 9)) - _
                         CCur(flxTReceipt.TextMatrix(iRow, 10)) & " " & _
                    "WHERE tlbReceipt.TransactionID = " & CLng(flxTReceipt.TextMatrix(iRow, 19)) & ";"
         End If
         adoConn.Execute szSQL

'         ~CREDIT NOTE~ & ~RECEIPT ON ACCOUNT~
         If flxTCrPoA.TextMatrix(Label10(4).Caption, 13) = 2 Or _
            flxTCrPoA.TextMatrix(Label10(4).Caption, 13) = 4 Then
            With rstRst
               .AddNew
               !TranType = IIf(UCase(flxTCrPoA.TextMatrix(Label10(4).Caption, 1)) = "ADJC", "ADJAL", "AL")
               !TransactionID = lRT_ID
               !Alloc_Unalloc = 1
               !FromTran = CLng(Label10(2).Caption)
               !ToTran = CLng(flxTReceipt.TextMatrix(iRow, 19))
               !AllocDate = Format(txtTRDate.text, "dd mmmm yyyy")
               !receiptAmount = CCur(flxTReceipt.TextMatrix(iRow, 10))
               !Discount = 0
               !IsSageUpdate = IIf(UCase(flxTCrPoA.TextMatrix(Label10(4).Caption, 1)) = "ADJC", False, True)
               !UpdateSage = False
               !BankCode = txtBC.Tag
               !nominalCode = rstRst!BankCode
               lRT_ID = lRT_ID + 1
               .Update
            End With
         End If
      End If
   Next iRow
   ConfigureFlxTCrPoA
   ConfigureFlxTReceipt
   txtTenantID.text = ""
   txtBC.text = ""
   txtAcBal.text = ""
   txtReceiptEntered.text = "0.00"
   txtTReceiptTotal.text = "0.00"

   flxTCrPoA.Enabled = True
   ReDim baChangesMade(flxTReceipt.Rows) As Boolean

   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   SavingAllocationR = True
End Function

Private Sub AutomaticAllocationTR()
   Dim iRow As Integer, iInd As Integer, j As Integer, valx As Integer, OSAmount As Double, ProcessAmount As Double, iGridRow As Integer
   Dim dSortIndex() As Integer
   Dim bTrue As Boolean, iCount As Integer, iActiveRow As Integer

   If (flxTCrPoA.Rows = 0) Then
      Exit Sub
   End If

   If (optOIF) Then
      ReDim dSortIndex(flxTReceipt.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxTReceipt.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxTReceipt.TextMatrix(dSortIndex(iInd), 5)) > (CDate(flxTReceipt.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd
         If Not bTrue Then
            dSortIndex(iCount) = iGridRow
         End If
         iCount = iCount + 1
      Next iGridRow
   Else
      ReDim dSortIndex(flxTReceipt.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxTReceipt.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxTReceipt.TextMatrix(dSortIndex(iInd), 5)) < (CDate(flxTReceipt.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd
         If Not bTrue Then
            dSortIndex(iCount) = iGridRow
         End If
         iCount = iCount + 1
      Next iGridRow
   End If

   iActiveRow = IIf(flxTCrPoA.row = 0, 1, flxTCrPoA.row)

   Label10(4).Caption = iActiveRow
   Label10(2).Caption = flxTCrPoA.TextMatrix(iActiveRow, 0)

   OSAmount = CDbl(flxTCrPoA.TextMatrix(iActiveRow, 8))
   ProcessAmount = IIf(flxTCrPoA.TextMatrix(iActiveRow, 9) = "0.00", CDbl(flxTCrPoA.TextMatrix(iActiveRow, 8)), CDbl(flxTCrPoA.TextMatrix(iActiveRow, 9)))

   If (flxTCrPoA.TextMatrix(iActiveRow, 3) = "ADJC") Then

      If ProcessAmount > cTotalAdjI Then
         ProcessAmount = cTotalAdjI
         flxTCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxTCrPoA.TextMatrix(iActiveRow, 9) = Format(ProcessAmount, "0.00")
      End If

      For j = 0 To flxTReceipt.Rows - 3
         If (flxTReceipt.TextMatrix(dSortIndex(j), 2) = "ADJI") Then
            If (ProcessAmount > flxTReceipt.TextMatrix(dSortIndex(j), 9)) Then
               flxTReceipt.TextMatrix(dSortIndex(j), 10) = flxTReceipt.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxTReceipt.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               ProcessAmount = ProcessAmount - flxTReceipt.TextMatrix(dSortIndex(j), 9)
               flxTReceipt.TextMatrix(dSortIndex(j), 15) = "A"
               flxTReceipt.TextMatrix(dSortIndex(j), 16) = Label10(2).Caption
            Else
               flxTReceipt.TextMatrix(dSortIndex(j), 10) = Format(ProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxTReceipt.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxTReceipt.TextMatrix(dSortIndex(j), 15) = "A"
               flxTReceipt.TextMatrix(dSortIndex(j), 16) = Label10(2).Caption
               Exit For
            End If
         End If
      Next j
   Else
      If ProcessAmount > cTotalSI Then
         ProcessAmount = cTotalAdjI
         flxTCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxTCrPoA.TextMatrix(iActiveRow, 9) = Format(ProcessAmount, "0.00")
      End If

      For j = 0 To flxTReceipt.Rows - 3
         If (Not flxTReceipt.TextMatrix(dSortIndex(j), 2) = "AdjI") Then
            If (ProcessAmount > flxTReceipt.TextMatrix(dSortIndex(j), 9)) Then
               flxTReceipt.TextMatrix(dSortIndex(j), 10) = flxTReceipt.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxTReceipt.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               ProcessAmount = ProcessAmount - flxTReceipt.TextMatrix(dSortIndex(j), 9)
               flxTReceipt.TextMatrix(dSortIndex(j), 15) = "A"
               flxTReceipt.TextMatrix(dSortIndex(j), 16) = Label10(2).Caption
            Else
               flxTReceipt.TextMatrix(dSortIndex(j), 10) = Format(ProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxTReceipt.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxTReceipt.TextMatrix(dSortIndex(j), 15) = "A"
               flxTReceipt.TextMatrix(dSortIndex(j), 16) = Label10(2).Caption
               Exit For
            End If
         End If
       Next j
   End If

   Frame4(1).Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True

   cmdRptAllocateSave.Enabled = True
   FocusControl cmdRptAllocateSave

   cmdRptAutomatic.Enabled = False
End Sub

Private Sub AutomaticAllocationSP()
   Dim iRow As Integer, iInd As Integer, j As Integer, valx As Integer, dOSAmount As Double
   Dim dSortIndex() As Integer, dProcessAmount As Double, iGridRow As Integer
   Dim bTrue As Boolean, iCount As Integer, iActiveRow As Integer

   If (flxSCrPoA.Rows = 0) Then Exit Sub

   If (optOIF) Then
      ReDim dSortIndex(flxSCrPoA.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSCrPoA.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxSCrPoA.TextMatrix(dSortIndex(iInd), 5)) > (CDate(flxSCrPoA.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd

         If Not bTrue Then dSortIndex(iCount) = iGridRow

         iCount = iCount + 1
      Next iGridRow
   Else
      ReDim dSortIndex(flxSCrPoA.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSCrPoA.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxSCrPoA.TextMatrix(dSortIndex(iInd), 5)) < (CDate(flxSCrPoA.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd
         If Not bTrue Then
            dSortIndex(iCount) = iGridRow
         End If
         iCount = iCount + 1
      Next iGridRow
   End If

   iActiveRow = IIf(flxSCrPoA.row = 0, 1, flxSCrPoA.row)

   Label10(6).Caption = iActiveRow
   Label10(5).Caption = flxSCrPoA.TextMatrix(iActiveRow, 0)

   dOSAmount = CDbl(flxSCrPoA.TextMatrix(iActiveRow, 8))
   dProcessAmount = IIf(flxSCrPoA.TextMatrix(iActiveRow, 9) = "0.00", CDbl(flxSCrPoA.TextMatrix(iActiveRow, 8)), CDbl(flxSCrPoA.TextMatrix(iActiveRow, 9)))

   If (flxSCrPoA.TextMatrix(iActiveRow, 3) = "ADJC") Then
      If dProcessAmount > cTotalAdjI Then
         dProcessAmount = cTotalAdjI
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(dProcessAmount, "0.00")
      End If

      For j = 0 To flxSCrPoA.Rows - 3
         If (flxSCrPoA.TextMatrix(dSortIndex(j), 2) = "ADJI") Then
            If (dProcessAmount > flxSCrPoA.TextMatrix(dSortIndex(j), 9)) Then
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               dProcessAmount = dProcessAmount - flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
            Else
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               Exit For
            End If
         End If
      Next j
   Else
      If dProcessAmount > cTotalSI Then
         dProcessAmount = cTotalAdjI
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(dProcessAmount, "0.00")
      End If

      For j = 0 To flxSCrPoA.Rows - 3
         If (Not flxSCrPoA.TextMatrix(dSortIndex(j), 2) = "AdjI") Then
            If (dProcessAmount > flxSCrPoA.TextMatrix(dSortIndex(j), 9)) Then
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               dProcessAmount = dProcessAmount - flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
            Else
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               Exit For
            End If
         End If
       Next j
   End If

   Frame4(1).Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True

   cmdPayAllocateSave.Enabled = True
   FocusControl cmdPayAllocateSave

   cmdPayAutomatic.Enabled = False
End Sub

Private Sub cmdAutoAllocSel_Click()
   If tabPayRpt.Tab = 0 Then AutomaticAllocationTR

   If tabPayRpt.Tab = 1 Then AutomaticAllocationSP
End Sub

Private Sub cmdAutoAllocSelCancel_Click()
   Frame4(1).Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True
   FocusControl cmdRptAutomatic
End Sub

Private Sub cmdRptAutomatic_Click()
   Frame4(1).Left = tabCashbook.Left + tabPayRpt.Left + Frame5(4).Left + Frame5(4).Width - 60
   Frame4(1).Top = tabCashbook.Top + tabPayRpt.Top + Frame5(4).Top - Frame4(1).Height + Frame5(4).Height - 20
   Frame4(1).Visible = True
   FocusControl cmdAutoAllocSel

   tabCashbook.Enabled = False
   tabPayRpt.Enabled = False
End Sub

Private Sub cmdBankOK_Click()
   If optReceipt Then
      BANK_TYPE = "Bank Receipt"
      Label3(3).Visible = True
      Label3(0).Visible = False
   Else
      BANK_TYPE = "Bank Payment"
      Label3(0).Visible = True
      Label3(3).Visible = False
   End If

   txtDateBk(0).text = Format(Date, "dd/mm/yyyy")

   HandleTextBoxesBk True, True

   cmdNewBk(0).Enabled = False

   fraBank.Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True

   FocusControl cboBRPClient
   bChangesMade = True
End Sub

Private Sub HandleTextBoxesBk(bClear As Boolean, bEnable As Boolean)
   If bClear Then
      txtBkAc(0).text = ""
      txtNCBk(0).text = ""
      txtDeptBk(0).text = ""
      txtDetailsBk(0).text = ""
      txtNetBk(0).text = ""
      txtVatBk(0).text = ""
      txtReference(0).text = ""
      txtBkAcName(0).text = ""
      txtNCNameBk(0).text = ""
      txtDeptBkName(0).text = ""
      txtTotalBk(0).text = "0.00"
   End If

   txtBkAc(0).Enabled = bEnable
   txtDateBk(0).Enabled = bEnable
   txtNCBk(0).Enabled = bEnable
   txtDeptBk(0).Enabled = bEnable
   txtDetailsBk(0).Enabled = bEnable
   txtNetBk(0).Enabled = bEnable
   txtVatBk(0).Enabled = bEnable
   cboBRPClient.Enabled = bEnable

   cmdBkList(0).Enabled = bEnable
   cmdNCBk(0).Enabled = bEnable
   cmdDeptBk(0).Enabled = bEnable
   cmdTaxListBk(0).Enabled = bEnable

   cmdUpdateBk(0).Enabled = bEnable
End Sub

Private Sub cmdBkList_Click(Index As Integer)
   LoadBankAccount

   picLeaseList.Left = txtBkAc(0).Left + fraBkInput(0).Left + tabPayRpt.Left + tabCashbook.Left
   picLeaseList.Top = txtBkAc(0).Top + txtBkAc(0).Height + _
                                    fraBkInput(0).Top + tabPayRpt.Top + tabCashbook.Top
   picLeaseList.Width = 3600
   cmdGridUnitLookup(0).Left = picLeaseList.Width - cmdGridUnitLookup(0).Width
   Shape4(6).Width = picLeaseList.Width - cmdGridUnitLookup(0).Width - 50
   flxLeaseList.Width = 3400
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0
   FocusControl flxLeaseList
   sTextBox = "Bank"
End Sub

Private Sub LoadBankAccount()
   flxLeaseList.Clear
   flxLeaseList.Rows = 2
   flxLeaseList.Cols = 4
   flxLeaseList.ColWidth(1) = 800
   flxLeaseList.ColWidth(2) = 2500

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.

   flxLeaseList.ColWidth(0) = 0
   flxLeaseList.ColWidth(3) = 0
   Label20(0).Width = 700
   Label20(0).Left = 50
   Label20(1).Width = 2600
   Label20(1).Left = Label20(0).Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchID.Width = 700
   txtTenantSearchID.Left = 40
   
   txtTenantSearchName.Width = 2400
   txtTenantSearchName.Left = txtTenantSearchID.Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchUnitName.Visible = False
   
         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   Label20(0).Caption = "Bank Code"
   Label20(1).Caption = "Bank Name"
   Label20(2).Visible = False
   
   '~~~ End of config
   ' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup your Client Bank Accounts."
   Else
      rRow = 1
      While Not adoRST.EOF
         flxLeaseList.TextMatrix(rRow, 1) = adoRST.Fields.Item("BNC").Value
         flxLeaseList.TextMatrix(rRow, 2) = adoRST.Fields.Item("BNN").Value
         flxLeaseList.AddItem ""
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
   End If

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Load Bank Account in Demand"

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdCancelBk_Click(Index As Integer)
   Dim iRow As Integer, iRemRec As Integer

   On Error GoTo ErrorHandler

   If Not cmdEditBk(0).Enabled Then
      If MsgBox("Do you want to cancel Edit?", vbQuestion + vbYesNo, "Edit Record") = vbNo Then Exit Sub
      HandleTextBoxesBk True, False
      iSelected = 0
      cmdEditBk(0).Enabled = True
   End If

   If Not cmdNewBk(0).Enabled Then
      If MsgBox("Do you want to cancel the new records?", vbQuestion + vbYesNo, "Add Record") = vbNo Then Exit Sub
      If cmdUpdateBk(0).Enabled Then
         HandleTextBoxesBk True, False
         cmdNewBk(0).Enabled = True
         bChangesMade = False
         flxBankPay(0).Enabled = True
         flxBankPay(0).row = 0
         Label3(0).Visible = False
         Label3(3).Visible = False
         Exit Sub
      End If

      iRemRec = 0
      iRow = 1
      If flxBankPay(0).Rows > 2 Then
         While iRow <= flxBankPay(0).Rows - 1
            If flxBankPay(0).TextMatrix(iRow, 14) = "" Then
               flxBankPay(0).RemoveItem (iRow)
               iRow = iRow - 1
            End If
            iRow = iRow + 1
         Wend
      Else
         For iRow = 0 To flxBankPay(0).Cols - 1
            flxBankPay(0).TextMatrix(1, iRow) = ""
         Next iRow
      End If

      HandleTextBoxesBk True, False
      cmdNewBk(0).Enabled = True
   End If

   bChangesMade = False
   flxBankPay(0).Enabled = True
   flxBankPay(0).row = 0

ErrorHandler:
   
End Sub

Private Sub cmdClose_Click()
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub cmdCloseBk_Click(Index As Integer)
   If Not cmdNewBk(0).Enabled Or Not cmdEditBk(0).Enabled Or bChangesMade Then
      If MsgBox("You want to close this window? Your data may be lost.", vbInformation + vbYesNo, "Close this window") = vbNo Then Exit Sub
   End If
   Label3(0).Visible = False
   Label3(3).Visible = False
   Unload Me
End Sub

Private Sub cmdContactCancel_Click()
   cmdContactSave.Enabled = False
   cmdContactCancel.Enabled = False
   cmdContactEdit.Enabled = True
   UnLockedContactFields False
End Sub

Private Sub cmdContactEdit_Click()
   If txtBC.text = "" Or txtClientList.text = "" Then Exit Sub
   UnLockedContactFields True
   cmdContactSave.Enabled = True
   cmdContactCancel.Enabled = True
   cmdContactEdit.Enabled = False
End Sub

Private Sub UnLockedContactFields(bUnLocked As Boolean)
   txtContact.Locked = Not bUnLocked
   txtTel.Locked = Not bUnLocked
   txtFax.Locked = Not bUnLocked
   txtMobile.Locked = Not bUnLocked
   txtEmail.Locked = Not bUnLocked
   txtWebsite.Locked = Not bUnLocked
End Sub

Private Sub cmdContactSave_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   On Error GoTo ErrorExc

   adoConn.Open getConnectionString

   szSQL = "UPDATE tlbBank AS B " & _
           "SET B.Contact = '" & txtContact.text & "', " & _
               "B.Tel = '" & txtTel.text & "', " & _
               "B.Fax = '" & txtFax.text & "', " & _
               "B.Mobile = '" & txtMobile.text & "', " & _
               "B.eMail = '" & txtEmail.text & "', " & _
               "B.Website = '" & txtWebsite.text & "' " & _
           "WHERE B.BANK_ID = '" & adoBank.Recordset.Fields.Item("BANK_ID").Value & "';"

   adoConn.Execute szSQL

   UnLockedContactFields False
   cmdContactSave.Enabled = False
   cmdContactCancel.Enabled = False
   cmdContactEdit.Enabled = True
   MsgBox "Contact details have been updated successfully."
   Exit Sub

ErrorExc:
   UnLockedContactFields False
   MsgBox "Contact details have not been updated."
End Sub

Private Sub cmdDeptBk_Click(Index As Integer)
   MousePointer = vbHourglass
   LoadDeptBk

   picLeaseList.Left = txtDeptBk(0).Left + fraBkInput(0).Left + tabPayRpt.Top + tabCashbook.Top
   picLeaseList.Top = txtDeptBk(0).Top + txtDeptBk(0).Height + fraBkInput(0).Top + tabPayRpt.Top + tabCashbook.Top
   
   picLeaseList.Width = 3600
   cmdGridUnitLookup(0).Left = picLeaseList.Width - cmdGridUnitLookup(0).Width
   Shape4(6).Width = picLeaseList.Width - cmdGridUnitLookup(0).Width - 50
   flxLeaseList.Width = 3400
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0
   FocusControl flxLeaseList
   sTextBox = "Dept"
   MousePointer = vbDefault
End Sub

Private Sub LoadDeptBk()
   flxLeaseList.Clear
   flxLeaseList.Rows = 2
   flxLeaseList.Cols = 4
   flxLeaseList.ColWidth(1) = 800
   flxLeaseList.ColWidth(2) = 2500
   flxLeaseList.ColAlignment = vbLeftJustify
   
    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
    
   flxLeaseList.ColWidth(0) = 0
   flxLeaseList.ColWidth(3) = 0
   Label20(0).Width = 1400
   Label20(0).Left = 50
   Label20(1).Width = 2600
   Label20(1).Left = Label20(0).Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchID.Width = 700
   txtTenantSearchID.Left = 40
   
   txtTenantSearchName.Width = 2400
   txtTenantSearchName.Left = txtTenantSearchID.Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchUnitName.Visible = False
   
         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   Label20(0).Caption = "Fund"
   Label20(1).Caption = "Name"
   Label20(2).Visible = False
   
   '~~~ End of config

   ' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   flxLeaseList.Clear
   flxLeaseList.TextMatrix(0, 0) = "Dept. ID"
   flxLeaseList.TextMatrix(0, 1) = "Department Name"

   Dim rRow As Integer
   rRow = 1
   While Not adoRST.EOF
      flxLeaseList.TextMatrix(rRow, 1) = adoRST.Fields.Item("FundID").Value
      flxLeaseList.TextMatrix(rRow, 2) = adoRST.Fields.Item("FundName").Value
      flxLeaseList.AddItem ""
      rRow = rRow + 1
      adoRST.MoveNext
   Wend

   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdEditBk_Click(Index As Integer)
   If cmdNewBk(0).Enabled = False Then Exit Sub

   If iSelected = 0 Then
      MsgBox "Select at least 1 row.", vbInformation + vbOKOnly, "Edit Record"
      Exit Sub
   End If

   HandleTextBoxesBk True, True

   txtBkAc(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 0)
   txtDateBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 2)
   cboBRPClient.Value = flxBankPay(0).TextMatrix(flxBankPay(0).row, 5)
   txtNCBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 6)
   txtDeptBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 8)
   txtReference(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 7)
   txtDetailsBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 9)
   txtNetBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 10)
   cmdTaxListBk(0).Caption = flxBankPay(0).TextMatrix(flxBankPay(0).row, 3)
   txtVatBk(0).text = flxBankPay(0).TextMatrix(flxBankPay(0).row, 11)
   txtTotalBk(0).text = Format(flxBankPay(0).TextMatrix(flxBankPay(0).row, 12), "0.00")

   flxBankPay(0).TextMatrix(flxBankPay(0).row, 15) = "1"

   iCurEditRow = flxBankPay(0).row
   cmdEditBk(0).Enabled = False
   flxBankPay(0).Enabled = False
   bChangesMade = True
End Sub

Private Sub cmdGridUnitLookup_Click(Index As Integer)
   picLeaseList.Visible = False
   tabCashbook.Enabled = True
   tabPayRpt.Enabled = True
End Sub

Private Sub cmdNCBk_Click(Index As Integer)
   LoadNominalCodeBk

   picLeaseList.Left = txtNCBk(0).Left + fraBkInput(0).Left + tabPayRpt.Left + tabCashbook.Left
   picLeaseList.Top = txtNCBk(0).Top + txtNCBk(0).Height + fraBkInput(0).Top + tabPayRpt.Top + tabCashbook.Top
   picLeaseList.Width = 3600
   cmdGridUnitLookup(0).Left = picLeaseList.Width - cmdGridUnitLookup(0).Width
   Shape4(6).Width = picLeaseList.Width - cmdGridUnitLookup(0).Width - 50
   flxLeaseList.Width = 3400
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0
   FocusControl flxLeaseList
   sTextBox = "NC"
End Sub

Private Sub LoadNominalCodeBk()
   flxLeaseList.Clear
   flxLeaseList.Rows = 2
   flxLeaseList.Cols = 4
   flxLeaseList.ColWidth(1) = 800
   flxLeaseList.ColWidth(2) = 2500
   flxLeaseList.ColAlignment = vbLeftJustify
   
    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
    
   flxLeaseList.ColWidth(0) = 0
   flxLeaseList.ColWidth(3) = 0
   Label20(0).Width = 1400
   Label20(0).Left = 50
   Label20(1).Width = 2600
   Label20(1).Left = Label20(0).Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchID.Width = 700
   txtTenantSearchID.Left = 40
   
   txtTenantSearchName.Width = 2400
   txtTenantSearchName.Left = txtTenantSearchID.Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchUnitName.Visible = False
   
         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   Label20(0).Caption = "Code"
   Label20(1).Caption = "Name"
   Label20(2).Visible = False
   
   '~~~ End of config
   

' Error Handler
  On Error GoTo Error_Handler

   Dim adoConn As New ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT NominalLedger.* FROM NominalLedger ORDER BY CODE;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim i As Integer
   i = 1
   While Not adoRST.EOF
      flxLeaseList.TextMatrix(i, 1) = adoRST.Fields.Item("Code").Value
      flxLeaseList.TextMatrix(i, 2) = adoRST.Fields.Item("Name").Value
      flxLeaseList.AddItem ""
      i = i + 1
      adoRST.MoveNext
   Wend
   
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
   
   flxLeaseList.Sort = 1
   Exit Sub

' Error Handling Code
Error_Handler:

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdNewBk_Click(Index As Integer)
   If cmdUpdateBk(0).Enabled Then
      FocusControl cmdUpdateBk(0)
      Exit Sub
   End If

   fraBank.Left = tabCashbook.Left + tabPayRpt.Left + Frame5(2).Left + cmdNewBk(0).Left - 60
   fraBank.Top = tabCashbook.Top + tabPayRpt.Top + Frame5(2).Top - fraBank.Height - 20
   fraBank.Visible = True
   FocusControl cmdBankOK

   tabCashbook.Enabled = False
   tabPayRpt.Enabled = False
End Sub

Private Sub InstantUnLock()
    Dim adoConn As New ADODB.Connection
    Dim rsLockDialog As New ADODB.Recordset
    Dim selcol As Integer
    Dim selRow As Integer
    Dim strSQL As String
    Dim colTransactionIDHerePayment As String
    Dim colTransactionIDHereReceipt As String
    Dim colTransactionIDHereBankPayment As String
    Dim i As Integer
    selcol = flxStatementReconcile.col
    selRow = flxStatementReconcile.row
'     colTransactionIDOtherReceipt = ""
'   colTransactionIDOtherPayment = ""
'   colTransactionIDOtherBankReceipt = ""
   
   
      'colTransactionIDOther varibale contains the transaction ID that is locked by other screen
     adoConn.Open getConnectionString
       If Len(colTransactionIDOtherPayment) > 0 Then ' This procedure is only for unlock the record on each cell browsing written by anol 20190412
                 strSQL = "Select DateTimeStamp ,UserSessionID,transactionID " & _
                        "from tlbPayment as Pt  where  (UserSessionID='' or isnull(UserSessionID='')) AND TransactionID in (" & colTransactionIDOtherPayment & ")"
                 rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
                 
                 While Not rsLockDialog.EOF
                         flxStatementReconcile.col = 0
                         For i = 1 To flxStatementReconcile.Rows - 1
                             If flxStatementReconcile.TextMatrix(i, btRecColNo) = rsLockDialog("transactionID").Value Then
                                   flxStatementReconcile.row = i
                                   flxStatementReconcile.CellBackColor = vbWhite
                                   'now you need to lock it for this screen
                                    colTransactionIDHerePayment = colTransactionIDHerePayment & flxStatementReconcile.TextMatrix(i, btRecColNo) & ","
                                    flxStatementReconcile.TextMatrix(i, 13) = "" 'we are not loading sessionID in this column for current screen lock
                                    flxStatementReconcile.TextMatrix(i, 14) = ""
                                    flxStatementReconcile.TextMatrix(i, 15) = ""
                                    flxStatementReconcile.TextMatrix(i, 16) = ""
                                    flxStatementReconcile.TextMatrix(i, 17) = ""
                             End If
                          Next i
                       rsLockDialog.MoveNext
                 Wend
                
                
                 If Len(colTransactionIDHerePayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
                     colTransactionIDHerePayment = Left(colTransactionIDHerePayment, Len(colTransactionIDHerePayment) - 1)
                 End If
                 If Len(colTransactionIDHerePayment) > 0 Then
                     'again locking those records for current screen
                     adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                    SystemUser & "',MachineName='" & WS_Name & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in  (" & colTransactionIDHerePayment & ")"
                                    haveYouLockedAnyReccord = True
                 End If
                 rsLockDialog.Close
                 Set rsLockDialog = Nothing
       End If
       If Len(colTransactionIDOtherReceipt) > 0 Then
                 strSQL = "Select DateTimeStamp ,UserSessionID,transactionID " & _
                        "from tlbReceipt as Pt  where  (UserSessionID='' or isnull(UserSessionID='')) AND TransactionID in (" & colTransactionIDOtherReceipt & ")"
                 rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
                 
                 While Not rsLockDialog.EOF
                         flxStatementReconcile.col = 0
                         For i = 1 To flxStatementReconcile.Rows - 1
                             If flxStatementReconcile.TextMatrix(i, btRecColNo) = rsLockDialog("transactionID").Value Then
                                   flxStatementReconcile.row = i
                                   flxStatementReconcile.CellBackColor = vbWhite
                                   'now you need to lock it for this screen
                                    colTransactionIDHerePayment = colTransactionIDHerePayment & flxStatementReconcile.TextMatrix(i, btRecColNo) & ","
                                    flxStatementReconcile.TextMatrix(i, 13) = "" 'we are not loading sessionID in this column for current screen lock
                                    flxStatementReconcile.TextMatrix(i, 14) = ""
                                    flxStatementReconcile.TextMatrix(i, 15) = ""
                                    flxStatementReconcile.TextMatrix(i, 16) = ""
                                    flxStatementReconcile.TextMatrix(i, 17) = ""
                             End If
                          Next i
                       rsLockDialog.MoveNext
                 Wend
                
                
                 If Len(colTransactionIDHereReceipt) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
                     colTransactionIDHereReceipt = Left(colTransactionIDHereReceipt, Len(colTransactionIDHereReceipt) - 1)
                 End If
                 If Len(colTransactionIDHereReceipt) > 0 Then
                     'again locking those records for current screen
                     adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                    SystemUser & "',MachineName='" & WS_Name & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in  (" & colTransactionIDHereReceipt & ")"
                                    haveYouLockedAnyReccord = True
                 End If
                 rsLockDialog.Close
                 Set rsLockDialog = Nothing
       End If
       If Len(colTransactionIDOtherBankReceipt) > 0 Then
                 strSQL = "Select DateTimeStamp ,UserSessionID,MY_ID " & _
                        "from tlbBankPayment as Pt  where  (UserSessionID='' or isnull(UserSessionID='')) AND MY_ID in (" & colTransactionIDOtherBankReceipt & ")"
                 rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
                 
                 While Not rsLockDialog.EOF
                         flxStatementReconcile.col = 0
                         For i = 1 To flxStatementReconcile.Rows - 1
                             If flxStatementReconcile.TextMatrix(i, btRecColNo) = rsLockDialog("MY_ID").Value Then
                                   flxStatementReconcile.row = i
                                   flxStatementReconcile.CellBackColor = vbWhite
                                   'now you need to lock it for this screen
                                    colTransactionIDHereBankPayment = colTransactionIDHereBankPayment & "'" & flxStatementReconcile.TextMatrix(i, btRecColNo) & "',"
                                    flxStatementReconcile.TextMatrix(i, 13) = "" 'we are not loading sessionID in this column for current screen lock
                                    flxStatementReconcile.TextMatrix(i, 14) = ""
                                    flxStatementReconcile.TextMatrix(i, 15) = ""
                                    flxStatementReconcile.TextMatrix(i, 16) = ""
                                    flxStatementReconcile.TextMatrix(i, 17) = ""
                             End If
                          Next i
                       rsLockDialog.MoveNext
                 Wend
                
                
                 If Len(colTransactionIDHereBankPayment) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
                     colTransactionIDHereBankPayment = Left(colTransactionIDHereBankPayment, Len(colTransactionIDHereBankPayment) - 1)
                 End If
                 If Len(colTransactionIDHereBankPayment) > 0 Then
                     'again locking those records for current screen
                     adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                                    SystemUser & "',MachineName='" & WS_Name & "'," & _
                                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID in  (" & colTransactionIDHereBankPayment & ")"
                                    haveYouLockedAnyReccord = True
                 End If
                 rsLockDialog.Close
                 Set rsLockDialog = Nothing
       End If
        adoConn.Close
        Set adoConn = Nothing
        
  
    flxStatementReconcile.col = selcol
    flxStatementReconcile.row = selRow
End Sub
Private Sub cmdReconAll_Click()
   Dim i As Integer
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   adoConn.Open getConnectionString
   Dim adoRST As New ADODB.Recordset
   If MsgBox("Do you want to reset all saved entries?", vbYesNo + vbInformation, "Bank concilation") = vbNo Then Exit Sub
   'below line moved to the top by anol
   lblClosingBalance.Caption = Format(txtStOpenBal.text, "0.00")
    szSQL = "SELECT ReconDate, BankCode " & _
                       "FROM tlbBankReconcilation " & _
                       "WHERE BankCode = '" & txtAccountName.text & "' AND clientID='" & txtClientList.Tag & "' Order by ReconDate Desc;"
            'Debug.Print szSQL
               adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            
               If Not adoRST.EOF Then
                    txtStatementDate.text = Format(adoRST.Fields("ReconDate").Value, "dd/MM/yyyy")
               Else
                    txtStatementDate.text = ""
               End If
            
               adoRST.Close
               Set adoRST = Nothing
               
   ResetAllSavedBankRecon adoConn
   ConfigFlxStatementReconcile
   
    adoConn.Execute "Update tlbPayment Set  ReconNow =null,Reconciled=null where Right(ReconNow,5)='Saved'"
    adoConn.Execute "Update tlbReceipt Set  ReconNow =null,Reconciled=null where Right(ReconNow,5)='Saved'"
    adoConn.Execute "Update tlbBankPayment Set  ReconNow =null,Reconciled=null where Right(ReconNow,5)='Saved'"

   
            
            
            
   'release all locks before loading the grid
    adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
    adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
    adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
            "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
            
            
            
            
   LoadFlxStatementReconcile adoConn
  
               
   optReconciliation(0).Value = True
   'lblClosingBalance.Caption = Format(txtStOpenBal.text, "0.00")
   txtProjClosingBal.text = ""
   'issue 523
   'Modified by anol 21 Jan 2015
   txtStOpenBal.Locked = False
   For i = 1 To flxCashBook.Rows - 1
      flxCashBook.row = i
      If flxCashBook.TextMatrix(i, 8) = "YES" Then
            txtStOpenBal.Locked = True
      End If
   Next i
  If flxCashBook.Rows = 2 Then
      If flxCashBook.TextMatrix(1, 8) = "YES" Then
            txtStOpenBal.Locked = True
      End If
  End If
  If flxCashBook.Rows = 1 Then
         txtStOpenBal.Locked = True
  End If
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub ResetAllSavedBankRecon(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim rstBank    As New ADODB.Recordset

   curOpeningBal = 0

   szSQL = "SELECT SUM(Reconciled) " & _
           "FROM tlbReceipt " & _
           "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                 "BankCode = '" & txtAccountName.text & "';"
   rstBank.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not rstBank.EOF Then curOpeningBal = curOpeningBal + _
         IIf(IsNull(rstBank.Fields.Item(0).Value), 0, rstBank.Fields.Item(0).Value)
   rstBank.Close

   szSQL = "SELECT SUM(Reconciled) " & _
           "FROM tlbPayment " & _
           "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                 "BankCode = '" & txtAccountName.text & "';"
   rstBank.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not rstBank.EOF Then curOpeningBal = curOpeningBal + _
         IIf(IsNull(rstBank.Fields.Item(0).Value), 0, rstBank.Fields.Item(0).Value)
   rstBank.Close

   szSQL = "SELECT SUM(Reconciled) " & _
           "FROM tlbBankPayment " & _
           "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                 "BANK_AC = '" & txtAccountName.text & "';"
   rstBank.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not rstBank.EOF Then curOpeningBal = curOpeningBal + _
         IIf(IsNull(rstBank.Fields.Item(0).Value), 0, rstBank.Fields.Item(0).Value)

   rstBank.Close
   Set rstBank = Nothing
   
   'issue 523
   'Modified by anol 20 Jan 2015
   'UpdateBankClosingBalance adoConn, Val(txtStOpenBal.text) - curOpeningBal, txtbc.Tag
'      adoConn.Execute "UPDATE tlbClientBanks " & _
'                   "SET ClosingBal = " & Val(txtStOpenBal.text) - curOpeningBal & " " & _
'                   "WHERE NominalCode = '" & txtBC.Tag & "' and Client_ID= '" & txtClientList.Tag & "';"
'modified by anol 02 08 2016
    adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET ClosingBal = " & Val(lblClosingBalance.Caption) & ",spare2='" & txtStatementDate.text & "' " & _
                   "WHERE NominalCode = '" & txtBC.Tag & "' and Client_ID= '" & txtClientList.Tag & "';"

   adoConn.Execute "UPDATE tlbReceipt " & _
                   "SET Reconciled = NULL, " & _
                       "ReconNow = NULL " & _
                   "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                       "BankCode = '" & txtAccountName.text & "';"

   adoConn.Execute "UPDATE tlbPayment " & _
                   "SET Reconciled = NULL, " & _
                       "ReconNow = NULL " & _
                   "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                       "BankCode = '" & txtAccountName.text & "';"

   adoConn.Execute "UPDATE tlbBankPayment " & _
                   "SET Reconciled = NULL, " & _
                       "ReconNow = NULL " & _
                   "WHERE RIGHT(ReconNow, 5) = 'Saved' AND " & _
                       "BANK_AC = '" & txtAccountName.text & "';"

   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET PCB = NULL " & _
                   "WHERE NominalCode = '" & txtAccountName.text & "';"
End Sub

Private Sub flxStatementReconcile_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 70 And Shift = 3 Then     '>> fixing data at run time
      If MsgBox("Do you want to fix part reconciled data without changing bank account balance?", vbQuestion + vbYesNo, "Fixing") = vbNo Then Exit Sub

      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString

      adoConn.Execute _
         "UPDATE tlbReceipt " & _
         "SET Reconciled = Amount, " & _
             "ReconNow = Left(ReconNow, 11) + 'Full' " & _
         "WHERE Type = 3 AND Details = 'BATCH RECEIPT ' AND " & _
               "BankCode = '" & txtBC.Tag & "' AND " & _
               "RIGHT(ReconNow, 4) = 'Part';"

      adoConn.Close
      Set adoConn = Nothing

      MsgBox "Reload the form please.", vbInformation + vbOKOnly, "Fixed"
   End If
End Sub

Public Sub SavedPreBankRecTrans(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, i As Integer, sFullPartSaved As Single

   SpreadAmount 3

'        ReconNow -> StatementDate#Full/Part/Saved
   With flxStatementReconcile
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, btRecColNo) <> "" And Val(.TextMatrix(iRow, 8)) <> 0 And _
               .TextMatrix(iRow, 0) <> "-" And .TextMatrix(iRow, 11) = "M" Then
            If .TextMatrix(iRow, 3) = "Sales Receipt" Or _
                  .TextMatrix(iRow, 3) = "Sales Receipt on Account" Or _
                  Left(.TextMatrix(iRow, 2), 3) = "SRR" Then
               szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = " & Val(.TextMatrix(iRow, 8)) & ", " & _
                           "ReconNow = '" & txtStatementDate.text & "#" & .TextMatrix(iRow, 9) & "' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            If .TextMatrix(iRow, 3) = "Bank Receipt" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = " & Val(.TextMatrix(iRow, 8)) & ", " & _
                           "ReconNow = '" & txtStatementDate.text & "#" & .TextMatrix(iRow, 9) & "' " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If

            If .TextMatrix(iRow, 3) = "Bank Payment" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = " & Val(.TextMatrix(iRow, 8)) & ", " & _
                           "ReconNow = '" & txtStatementDate.text & "#" & .TextMatrix(iRow, 9) & "' " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If
            If .TextMatrix(iRow, 3) = "Purchase Payment" Or _
                  .TextMatrix(iRow, 3) = "Purchase Payment on Account" Or _
                  Left(.TextMatrix(iRow, 2), 3) = "PPR" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = " & Val(.TextMatrix(iRow, 8)) & ", " & _
                           "ReconNow = '" & txtStatementDate.text & "#" & .TextMatrix(iRow, 9) & "' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            
            Debug.Print iRow
            Debug.Print "Below" & szSQL
            adoConn.Execute szSQL
         End If

         If .TextMatrix(iRow, btRecColNo) <> "" And Val(.TextMatrix(iRow, 8)) = 0 And _
            .TextMatrix(iRow, 0) <> "-" And .TextMatrix(iRow, 11) = "M" Then
            If .TextMatrix(iRow, 3) = "Sales Receipt" Or _
                  .TextMatrix(iRow, 3) = "Sales Receipt on Account" Or _
                  Left(.TextMatrix(iRow, 2), 3) = "SRR" Then
               szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            If .TextMatrix(iRow, 3) = "Bank Receipt" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If

            If .TextMatrix(iRow, 3) = "Bank Payment" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If
            If .TextMatrix(iRow, 3) = "Purchase Payment" Or _
                  .TextMatrix(iRow, 3) = "Purchase Payment on Account" Or _
                  Left(.TextMatrix(iRow, 2), 3) = "PPR" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            Debug.Print iRow
            Debug.Print "Below" & szSQL
            adoConn.Execute szSQL
         End If
      Next iRow
   End With

   'On Error Resume Next
'issue 523
'modified by anol 19 Jan 2015
   szSQL = "UPDATE tlbClientBanks " & _
           "SET PCB = " & Val(txtProjClosingBal.text) & ", " & _
               "spare2 = '" & Format(txtStatementDate.text, "dd mmmm yyyy") & "',ClosingBal = " & Val(lblClosingBalance.Caption) & "  " & _
           "WHERE NominalCode = '" & txtAccountName.text & "' and Client_ID= '" & txtClientList.Tag & "';"
   adoConn.Execute szSQL

'   UpdateBankClosingBalance adoConn, Val(lblClosingBalance.Caption), txtbc.Tag
'issue 523
'modified by anol 19 Jan 2015
'rem by anol 02 08 2016
'   adoconn.Execute "UPDATE tlbClientBanks " & _
'                   "SET ClosingBal = " & Val(lblClosingBalance.Caption) & " " & _
'                   "WHERE NominalCode = '" & txtAccountName.text & "' and Client_ID= '" & txtClientList.Tag & "';"

End Sub

'Private Sub UpdateBankReconSplit_(adoConn As ADODB.Connection)
'   Dim rstSrc       As New ADODB.Recordset
'   Dim rstDST       As New ADODB.Recordset
'   Dim szaTemp()    As String
'
'   rstDST.Open "SELECT * FROM tlbBankReconcilation;", adoConn, adOpenDynamic, adLockOptimistic
'
''   With rstSrc
''      .Open "SELECT * " & _
''            "FROM tlbPayment AS P " & _
''            "WHERE ReconNow <>'' AND P.Type IN (8, 9, 24) AND CSTR(P.TransactionID) NOT IN (" & _
''                 "SELECT RefID FROM tlbBankReconcilation WHERE TransactionType IN (8, 9, 24))" & _
''                  ";", adoConn, adOpenStatic, adLockReadOnly
'''szHeader$ = "+|<Date|<TranID|<Type|<Account|<Reference|>ReceiptValue|>PaymentValue" & _
'''            "|>Statement|<Reconciliation|ID|Flag|Statement date"
''
''      While Not .EOF
''         rstDST.AddNew
''         rstDST.Fields.Item("MY_ID").Value = UniqueID()
''         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("Type").Value
''         rstDST.Fields.Item("RefID").Value = .Fields.Item("TransactionID").Value
''         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("SageAccountNumber").Value
''         rstDST.Fields.Item("UnitID").Value = .Fields.Item("UnitID").Value
''         rstDST.Fields.Item("TDate").Value = .Fields.Item("PDate").Value
''         rstDST.Fields.Item("DDate").Value = .Fields.Item("DDate").Value
''         rstDST.Fields.Item("TRef").Value = .Fields.Item("Ref").Value
''         rstDST.Fields.Item("Details").Value = .Fields.Item("Details").Value
''         rstDST.Fields.Item("Amount").Value = .Fields.Item("Amount").Value
''         rstDST.Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
''         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value
''
''         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
''         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
''         rstDST.Fields.Item("ReconType").Value = szaTemp(1)
''
''         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BankCode").Value
''         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NominalCode").Value
''         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("ExtRef").Value
''         rstDST.Fields.Item("TranMth").Value = .Fields.Item("PayAmtType").Value
''         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("SlNumber").Value
''         rstDST.Fields.Item("FundID").Value = .Fields.Item("FundID").Value
''         rstDST.Fields.Item("Recoverable").Value = .Fields.Item("Recoverable").Value
''         rstDST.Update
''
''         .MoveNext
''      Wend
''
''      .Close
''   End With
'   With rstSrc
'      .Open "SELECT * " & _
'            "FROM tlbReceipt AS P " & _
'            "WHERE ReconNow <>'' AND P.Type IN (3, 4, 23) AND CSTR(P.TransactionID) NOT IN (" & _
'                 "SELECT RefID FROM tlbBankReconcilation WHERE TransactionType IN (3, 4, 23))" & _
'                  ";", adoConn, adOpenStatic, adLockReadOnly
'
'      While Not .EOF
'         rstDST.AddNew
'         rstDST.Fields.Item("MY_ID").Value = UniqueID()
'         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("Type").Value
'         rstDST.Fields.Item("RefID").Value = .Fields.Item("TransactionID").Value
'         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("SageAccountNumber").Value
'         rstDST.Fields.Item("UnitID").Value = .Fields.Item("UnitID").Value
'         rstDST.Fields.Item("TDate").Value = .Fields.Item("RDate").Value
'         rstDST.Fields.Item("DDate").Value = .Fields.Item("DDate").Value
'         rstDST.Fields.Item("TRef").Value = .Fields.Item("Ref").Value
'         rstDST.Fields.Item("Details").Value = .Fields.Item("Details").Value
'         rstDST.Fields.Item("Amount").Value = .Fields.Item("Amount").Value
'         rstDST.Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
'         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value
'
'         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
'         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
'         rstDST.Fields.Item("ReconType").Value = szaTemp(1)
'
'         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BankCode").Value
'         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NominalCode").Value
'         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("ExtRef").Value
'         rstDST.Fields.Item("TranMth").Value = .Fields.Item("RptAmtType").Value
'         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("SlNumber").Value
'         rstDST.Fields.Item("FundID").Value = .Fields.Item("FundID").Value
'         'added by anol 20 jan 2015
'         'issue 523
'          'rstDST.Fields.Item("ClientID").Value =  txtClientList.tag
'         rstDST.Update
'
'         .MoveNext
'      Wend
'
'      .Close
'   End With
'   With rstSrc
'      .Open "SELECT * " & _
'            "FROM tlbBankPayment AS P " & _
'            "WHERE ReconNow <>'' AND CSTR(P.MY_ID) NOT IN (" & _
'                 "SELECT RefID FROM tlbBankReconcilation WHERE TransactionType IN (11, 12))" & _
'                  ";", adoConn, adOpenStatic, adLockReadOnly
'
'      While Not .EOF
'         rstDST.AddNew
'         rstDST.Fields.Item("MY_ID").Value = UniqueID()
'         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("TransactionType").Value
'         rstDST.Fields.Item("RefID").Value = .Fields.Item("MY_ID").Value
'         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("BANK_AC").Value
'         rstDST.Fields.Item("UnitID").Value = .Fields.Item("PropertyID").Value
'         rstDST.Fields.Item("TDate").Value = .Fields.Item("TRAN_DATE").Value
'         rstDST.Fields.Item("DDate").Value = .Fields.Item("TRAN_DATE").Value
'         rstDST.Fields.Item("TRef").Value = .Fields.Item("PROJ_REF").Value
'         rstDST.Fields.Item("Details").Value = .Fields.Item("DESCRIPTION").Value
'         rstDST.Fields.Item("Amount").Value = .Fields.Item("NET_AMOUNT").Value + .Fields.Item("VAT").Value
'         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value
'
'         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
'         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
'         rstDST.Fields.Item("ReconType").Value = szaTemp(1)
'
'         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BANK_AC").Value
'         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NOMINAL_CODE").Value
'         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("PROJ_REF").Value
'         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("TRAN_ID").Value
'         rstDST.Fields.Item("FundID").Value = .Fields.Item("DEPT_ID").Value
'         rstDST.Fields.Item("ClientID").Value = .Fields.Item("ClientID").Value
'         rstDST.Update
'
'         .MoveNext
'      Wend
'
'      .Close
'   End With
'
'   rstDST.Close
'   Set rstDST = Nothing
'   Set rstSrc = Nothing
'End Sub

Private Function CheckTranDate_StDate() As Boolean
   CheckTranDate_StDate = True

   Dim iRow As Integer
   On Error Resume Next

   With flxStatementReconcile
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 11) = "M" Or .TextMatrix(iRow, 9) = "Saved" Then
            If CDate(txtStatementDate.text) < CDate(.TextMatrix(iRow, 1)) Then
               CheckTranDate_StDate = False
               Exit Function
            End If
         End If
      Next iRow
   End With
End Function

Private Sub cmdReconSave_Click()
   If txtStOpenBal.text = "" Then
      MsgBox "Please input the statement opening balance.", vbOKOnly, "Warning"
      FocusControl txtStOpenBal
      Exit Sub
   End If
    
   If Trim(txtProjClosingBal.text) = "" Then
        MsgBox "Please enter a " & Label1(24).Caption & ".", vbOKOnly, "Warning"
        FocusControl txtProjClosingBal
        Exit Sub
   End If
   If Len(txtBC.text) = 0 Then
        MsgBox "Please input the bank Code.", vbOKOnly, "Warning"
        FocusControl cmdBC
        Exit Sub
   End If
   If txtStatementDate.text = "" Then
      MsgBox "Please input the statement date.", vbOKOnly, "Warning"
      FocusControl txtStatementDate
      Exit Sub
   End If
   
   
   
   Dim adoConn As New ADODB.Connection
   Dim bBankRec As Boolean
   adoConn.Open getConnectionString
   
   'validation placed on 02 08 2016 by anol
   If Not CheckTranDate_StDate Then
      MsgBox "You cannot reconcile transactions with a date after the bank statement date", vbOKOnly, "Warning"
      Exit Sub
   End If

   If txtStatementDate.text <> "" Then
      If Not IsBankStDtValid(adoConn) Then
         MsgBox " A Bank Reconciliation has been completed at this date. Please select another date.", vbCritical + vbOKOnly, "Invalid Bank Reconciliation Date."
         FocusControl txtStatementDate
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If
   End If
  'End of validation
   
   
'  Saving the entries but it will not book as reconciled
   If Val(txtProjClosingBal.text) <> Val(lblClosingBalance.Caption) Then
      If MsgBox("Do you wish to save this reconciliation and return to it later?", vbQuestion + vbYesNo, _
                "Cashbook") = vbNo Then
         FocusControl txtProjClosingBal
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      Else
         adoConn.BeginTrans
         SavedPreBankRecTrans adoConn
         adoConn.CommitTrans

         MsgBox "Transactions have been saved.", vbOKOnly, "Saved"
           'added by anol 26/ 04/ 2016
         Label1(23).Caption = "Statement Opening Balance:"
         Label1(24).Caption = "Projected Closing Balance:"
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If
   End If

   If MsgBox("Your bank reconciliation agrees. Do you wish to complete it?", vbQuestion + vbYesNo, _
             "Bank Reconciliation") = vbNo Then
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   'validation is moving at the top of this procedure anol 02 08 2016

'  System will book bank reconciliation
   adoConn.BeginTrans

   bBankRec = SaveBankReconciliation(adoConn)

   If Not bBankRec Then
   'rem out by anol 27 07 2016
'            Dim sChoice As Single
'            sChoice = MsgBox("There are some reconciled transactions found dated after their reconciled statement date." + Chr(13) + _
'                          "Please contact PCM Support to correct these transactions before proceeding further." + Chr(13) + _
'                          "Click OK to print a list of these transactions.", vbCritical + vbOKCancel, _
'                          "Incorrect Data")
'            If sChoice = vbOK Then
'               ShowReport App.Path & szReportPath & "\TranDt_BankRecDate.rpt"
'            End If
      adoConn.RollbackTrans
      adoConn.Close
      'MsgBox "An error occured.Bank Reconciliation has not been saved.", vbInformation, "warning"
      Set adoConn = Nothing
      Exit Sub
   End If

   adoConn.CommitTrans

   ConfigFlxStatementReconcile
   If txtClientList.text = "Consolidated" Then
        LoadFlxStatementReconcileConsolidated adoConn, SelectedConBankID
   Else
        LoadFlxStatementReconcile adoConn
   End If
   'loadflxCashBook
'   LoadLstBankStDates adoConn

'  Refresh the cash book history
   LoadFlxCashBook adoConn
'-------------------------------
   adoConn.Close
   Set adoConn = Nothing
    
   txtProjClosingBal.text = ""
   optReconciliation(0).Value = True

   If bBankRec Then
        ShowMsgInTaskBar "Bank Reconciliation has been saved."
        'added by anol 26/ 04/ 2016
        Label1(23).Caption = "Statement Opening Balance:"
        Label1(24).Caption = "Projected Closing Balance:"
  End If
    
   FocusControl txtStatementDate
   txtStatementDate.SelLength = Len(txtStatementDate.text)
   'added by anol 02 08 2016
   bChangesMade = False
   'lblClosingBalance.Caption = "0.00"
End Sub

Private Function IsBankStDtValid(adoConn As ADODB.Connection) As Boolean
   Dim adoRST  As New ADODB.Recordset
   Dim szSQL   As String
   Dim r       As Integer
'issue 523
'Modified by anol 20 Jan 2015
    If txtClientList.text = "Consolidated" Then
        szSQL = "SELECT ReconDate, B.BankCode " & _
           "FROM tlbBankReconcilation B,ConsolidatedBankList CB " & _
           "WHERE CB.BankCode=B.clientID and CB.conBankID=" & SelectedConBankID & " AND " & _
               "ReconDate >= #" & Format(txtStatementDate.text, "dd mmmm yyyy") & "# " & _
           "GROUP BY ReconDate, B.BankCode;"
    Else
    
        szSQL = "SELECT ReconDate, BankCode " & _
           "FROM tlbBankReconcilation " & _
           "WHERE BankCode = '" & txtAccountName.text & "' AND clientID='" & txtClientList.Tag & "'AND " & _
               "ReconDate >= #" & Format(txtStatementDate.text, "dd mmmm yyyy") & "# " & _
           "GROUP BY ReconDate, BankCode;"
    End If
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   IsBankStDtValid = adoRST.EOF

   adoRST.Close
   Set adoRST = Nothing
End Function

'
'Private Sub LoadLstBankStDates(adoConn As ADODB.Connection)
'   Dim adoRST  As New ADODB.Recordset
'   Dim szSQL   As String
'   Dim r       As Integer
'
'   szSQL = "SELECT ReconDate, BankCode " & _
'           "FROM tlbBankReconcilation " & _
'           "WHERE BankCode = '" & txtAccountName.text & "' " & _
'           "GROUP BY ReconDate, BankCode " & _
'           "ORDER BY ReconDate DESC;"
''Debug.Print szSQL
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   lstBankStDates.Clear
'
'   While Not adoRST.EOF
'      If Not IsNull(adoRST.Fields.Item(0).Value) Then
'         lstBankStDates.AddItem adoRST.Fields.Item(0).Value & "#" & adoRST.Fields.Item(1).Value
'      End If
'
'      adoRST.MoveNext
'   Wend
'
'   adoRST.Close
'   Set adoRST = Nothing
'End Sub
'
'Private Function MatchedBankStDt(adoConn As ADODB.Connection) As Boolean
'   Dim szSQL As String
'   Dim adoRST As New ADODB.Recordset
'
'   MatchedBankStDt = False
'
'   szSQL = "SELECT DISTINCT LEFT(ReconNow,10) AS ReconDt " & _
'           "FROM tlbReceipt " & _
'           "WHERE ReconNow <> '' AND RIGHT(ReconNow, 4) = 'Full' AND " & _
'                  "BankCode = '" & txtAccountName.text & "' " & _
'           "UNION " & _
'           "SELECT DISTINCT LEFT(ReconNow,10) AS ReconDt " & _
'           "FROM tlbBankPayment " & _
'           "WHERE  ReconNow <> '' AND RIGHT(ReconNow, 4) = 'Full' AND " & _
'                  "BANK_AC = '" & txtAccountName.text & "' " & _
'           "UNION " & _
'           "SELECT DISTINCT LEFT(ReconNow,10) AS ReconDt " & _
'           "FROM tlbPayment " & _
'           "WHERE ReconNow <> '' AND RIGHT(ReconNow, 4) = 'Full' AND " & _
'                  "BankCode = '" & txtAccountName.text & "';"
''Debug.Print szSQL
'
'   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'   While Not adoRST.EOF
'      If adoRST.Fields.Item("ReconDt").Value = txtStatementDate.text Then
'         MatchedBankStDt = True
'         adoRST.Close
'         Set adoRST = Nothing
'         Exit Function
'      End If
'
'      adoRST.MoveNext
'   Wend
'
'   adoRST.Close
'   Set adoRST = Nothing
'End Function

Private Function SaveBankReconciliation(adoConn As ADODB.Connection) As Boolean
   Dim szSQL As String, iRow As Integer, i As Integer
   Dim sFullPartSaved As Single
   Dim rstSrc       As New ADODB.Recordset
   Dim rstDST       As New ADODB.Recordset
   Dim szaTemp()    As String

   On Error GoTo Err:
'
'   If MatchedBankStDt(adoConn) Then
'      If MsgBox("This statement date has already been reconciled." & Chr(13) & _
'                "Do you wish to enter a new statement date?", vbYesNo, "Bank Reconciliation") = vbYes Then
'         txtStatementDate.SetFocus
'         txtStatementDate.SelLength = Len(txtStatementDate.text)
'         SaveBankReconciliation = False
'         Exit Function
'      End If
'   End If

'  The following procedure will spread the header amount among the child rows.
   SpreadAmount sFullPartSaved

   ResetAllSavedBankRecon adoConn
   rstDST.Open "SELECT * FROM tlbBankReconcilation;", adoConn, adOpenDynamic, adLockOptimistic

'  Save in DataBase
'  ReconNow -> StatementDate#Full/Part/Saved
   With flxStatementReconcile
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, btRecColNo) <> "" And Val(.TextMatrix(iRow, 8)) <> 0 And _
               .TextMatrix(iRow, 0) <> "+" And .TextMatrix(iRow, 0) <> ">" And .TextMatrix(iRow, 11) = "M" Then
            If .TextMatrix(iRow, 3) = "Sales Receipt" Or .TextMatrix(iRow, 3) = "Sales Receipt on Account" Then
               
'               szSQL = "UPDATE tlbReceipt AS T " & _
'                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
'                              Val(.TextMatrix(iRow, 8)) & ", " & _
'                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
'                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
             'resolved by BOSL
             'issue 486
             'Modified by anol 22 Oct 2014
             szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = " & Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
               adoConn.Execute szSQL
                        'issue 486 modified below line by anol 08 Jan 2015
                        If Val(.RowHeight(iRow)) > 0 Then
                           adoConn.Execute szSQL
                        End If
               rstSrc.Open "SELECT P.Type AS TransactionType, P.TransactionID AS RefID, " & _
                              "P.SageAccountNumber AS AccountNum, P.UnitID, P.RDate AS TDate, " & _
                              "P.DDate, P.Ref AS TRef, P.Details, P.ReconNow, P.BankCode, " & _
                              "P.NominalCode, P.ExtRef, P.RptAmtType AS TranMth, P.SlNumber, " & _
                              "P.FundID, '' AS Recoverable " & _
                           "FROM tlbReceipt AS P " & _
                           "WHERE P.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";", adoConn, adOpenStatic, adLockReadOnly
            End If
            If Left(.TextMatrix(iRow, 2), 3) = "SRR" Then
               szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
                              Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
               adoConn.Execute szSQL

               rstSrc.Open "SELECT P.Type AS TransactionType, P.TransactionID AS RefID, " & _
                              "P.SageAccountNumber AS AccountNum, P.UnitID, P.RDate AS TDate, " & _
                              "P.DDate, P.Ref AS TRef, P.Details, P.ReconNow, P.BankCode, " & _
                              "P.NominalCode, P.ExtRef, P.RptAmtType AS TranMth, P.SlNumber, " & _
                              "P.FundID, '' AS Recoverable " & _
                           "FROM tlbReceipt AS P " & _
                           "WHERE P.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";", adoConn, adOpenStatic, adLockReadOnly
            End If
            If .TextMatrix(iRow, 3) = "Bank Receipt" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
                              Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
               adoConn.Execute szSQL

               rstSrc.Open "SELECT P.TransactionType, P.MY_ID AS RefID, " & _
                              "P.BANK_AC AS AccountNum, P.UNIT_ID AS UnitID, P.TRAN_DATE AS TDate, " & _
                              "P.TRAN_DATE AS DDate, P.TRANS AS TRef, P.DESCRIPTION AS Details, " & _
                              "P.ReconNow, P.BANK_AC AS BankCode, P.NOMINAL_CODE AS NominalCode, " & _
                              "P.PROJ_REF AS ExtRef, '' AS TranMth, P.TRAN_ID AS SlNumber, " & _
                              "P.DEPT_ID AS FundID, '' AS Recoverable " & _
                           "FROM tlbBankPayment AS P " & _
                           "WHERE P.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';", adoConn, adOpenStatic, adLockReadOnly
            End If
            If .TextMatrix(iRow, 3) = "Bank Payment" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
                              Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
               adoConn.Execute szSQL

               rstSrc.Open "SELECT P.TransactionType, P.MY_ID AS RefID, " & _
                              "P.BANK_AC AS AccountNum, P.UNIT_ID AS UnitID, P.TRAN_DATE AS TDate, " & _
                              "P.TRAN_DATE AS DDate, P.TRANS AS TRef, P.DESCRIPTION AS Details, " & _
                              "P.ReconNow, P.BANK_AC AS BankCode, P.NOMINAL_CODE AS NominalCode, " & _
                              "P.PROJ_REF AS ExtRef, '' AS TranMth, P.TRAN_ID AS SlNumber, " & _
                              "P.DEPT_ID AS FundID, '' AS Recoverable " & _
                           "FROM tlbBankPayment AS P " & _
                           "WHERE P.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';", adoConn, adOpenStatic, adLockReadOnly
            End If
            If .TextMatrix(iRow, 3) = "Purchase Payment" Or _
                  .TextMatrix(iRow, 3) = "Purchase Payment on Account" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
                              Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
               adoConn.Execute szSQL

               rstSrc.Open "SELECT P.Type AS TransactionType, P.TransactionID AS RefID, " & _
                              "P.SageAccountNumber AS AccountNum, P.UnitID, P.PDate AS TDate, " & _
                              "P.DDate, P.Ref AS TRef, P.Details, P.ReconNow, P.BankCode, " & _
                              "P.NominalCode, P.ExtRef, P.PayAmtType AS TranMth, P.SlNumber, " & _
                              "P.FundID, P.Recoverable " & _
                           "FROM tlbPayment AS P " & _
                           "WHERE P.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";", adoConn, adOpenStatic, adLockReadOnly

            End If
            If Left(.TextMatrix(iRow, 2), 3) = "PPR" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = IIF(ISNULL(T.Reconciled), 0, VAL(T.Reconciled)) + " & _
                              Val(.TextMatrix(iRow, 8)) & ", DateTimeStamp = '',Module = '',UserSessionID = '',WindowsUserName = '',MachineName = '',PrestigeUserName = '',ServerIPaddress = '', " & _
                           "ReconNow = '" & txtStatementDate.text & "#Full' " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"

               adoConn.Execute szSQL

               rstSrc.Open "SELECT P.Type AS TransactionType, P.TransactionID AS RefID, " & _
                              "P.SageAccountNumber AS AccountNum, P.UnitID, P.PDate AS TDate, " & _
                              "P.DDate, P.Ref AS TRef, P.Details, P.ReconNow, P.BankCode, " & _
                              "P.NominalCode, P.ExtRef, P.PayAmtType AS TranMth, P.SlNumber, " & _
                              "P.FundID, P.Recoverable " & _
                           "FROM  tlbPayment AS P " & _
                           "WHERE P.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";", adoConn, adOpenStatic, adLockReadOnly
            End If

'szHeader$ = "+|<Date|<TranID|<Type|<Account|<Reference|>ReceiptValue|>PaymentValue" & _
'            "|>Statement|<Reconciliation|ID|Flag|Statement date"
            rstDST.AddNew
            rstDST.Fields.Item("MY_ID").Value = UniqueID()
            rstDST.Fields.Item("TransactionType").Value = rstSrc.Fields.Item("TransactionType").Value
            rstDST.Fields.Item("RefID").Value = rstSrc.Fields.Item("RefID").Value
            rstDST.Fields.Item("AccountNum").Value = rstSrc.Fields.Item("AccountNum").Value
            rstDST.Fields.Item("UnitID").Value = rstSrc.Fields.Item("UnitID").Value
            rstDST.Fields.Item("TDate").Value = rstSrc.Fields.Item("TDate").Value
            rstDST.Fields.Item("DDate").Value = rstSrc.Fields.Item("DDate").Value
            rstDST.Fields.Item("TRef").Value = rstSrc.Fields.Item("TRef").Value
            rstDST.Fields.Item("Details").Value = rstSrc.Fields.Item("Details").Value
            If .TextMatrix(iRow, 6) = "" Then
               rstDST.Fields.Item("Amount").Value = .TextMatrix(iRow, 7)
            Else
               rstDST.Fields.Item("Amount").Value = .TextMatrix(iRow, 6)
            End If
            rstDST.Fields.Item("OSAmount").Value = 0
            rstDST.Fields.Item("ReconAmount").Value = .TextMatrix(iRow, 8)

            szaTemp = Split(rstSrc.Fields.Item("ReconNow").Value, "#")
            rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
            rstDST.Fields.Item("ReconType").Value = szaTemp(1)

            rstDST.Fields.Item("BankCode").Value = rstSrc.Fields.Item("BankCode").Value
            rstDST.Fields.Item("NominalCode").Value = rstSrc.Fields.Item("NominalCode").Value
            rstDST.Fields.Item("ExtRef").Value = rstSrc.Fields.Item("ExtRef").Value
            rstDST.Fields.Item("TranMth").Value = rstSrc.Fields.Item("TranMth").Value
            rstDST.Fields.Item("SlNumber").Value = rstSrc.Fields.Item("SlNumber").Value
            rstDST.Fields.Item("FundID").Value = rstSrc.Fields.Item("FundID").Value
            If IsNull(rstSrc.Fields.Item("Recoverable").Value) Or rstSrc.Fields.Item("Recoverable").Value = "" Then
               rstDST.Fields.Item("Recoverable").Value = 0
            Else
               rstDST.Fields.Item("Recoverable").Value = Val(rstSrc.Fields.Item("Recoverable").Value)
            End If
            'Issue 523
            'added by anol 20 Jan 2015
            rstDST.Fields.Item("ClientID").Value = txtClientList.Tag
            'end of modification
            rstDST.Update
            rstSrc.Close
'Issue 523
'Modified by anol 20 Jan 2015
           ' UpdateBankClosingBalance adoConn, Val(txtProjClosingBal.text), txtbc.Tag
                If txtClientList.text = "Consolidated" Then
                     adoConn.Execute "UPDATE ConsolidatedBankList " & _
                    "SET ClosingBal = " & Val(txtProjClosingBal.text) & " " & _
                    "WHERE conBankID = " & SelectedConBankID & ";"
                    
                    adoConn.Execute "UPDATE tlbClientBanks " & _
                    "SET conBankReadOnly = 1 " & _
                    "WHERE ConsolidatedBankID = " & SelectedConBankID & ";"
                Else
                    adoConn.Execute "UPDATE tlbClientBanks " & _
                    "SET ClosingBal = " & Val(txtProjClosingBal.text) & " " & _
                    "WHERE NominalCode = '" & txtBC.Tag & "' and Client_ID= '" & txtClientList.Tag & "';"
                End If
         End If

         If .TextMatrix(iRow, btRecColNo) <> "" And Val(.TextMatrix(iRow, 8)) = 0 And _
            .TextMatrix(iRow, 0) <> "+" And .TextMatrix(iRow, 0) <> ">" And .TextMatrix(iRow, 11) = "M" Then
            If .TextMatrix(iRow, 3) = "Sales Receipt" Or .TextMatrix(iRow, 3) = "Sales Receipt on Account" Then
               szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            If .TextMatrix(iRow, 3) = "Sales Receipt Refund" Then
               szSQL = "UPDATE tlbReceipt AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            If .TextMatrix(iRow, 3) = "Bank Receipt" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If

            If .TextMatrix(iRow, 3) = "Bank Payment" Then
               szSQL = "UPDATE tlbBankPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.MY_ID = '" & .TextMatrix(iRow, btRecColNo) & "';"
            End If
            If .TextMatrix(iRow, 3) = "Purchase Payment" Or _
                  .TextMatrix(iRow, 3) = "Purchase Payment on Account" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If
            If .TextMatrix(iRow, 3) = "Purchase Payment Refund" Then
               szSQL = "UPDATE tlbPayment AS T " & _
                       "SET T.Reconciled = NULL, " & _
                           "ReconNow = NULL " & _
                       "WHERE T.TransactionID = " & .TextMatrix(iRow, btRecColNo) & ";"
            End If

            adoConn.Execute szSQL
         End If
      Next iRow
   End With
   'below line is commnted by anol due to syntax error n 21 Jan 2015
   
   'rstDST.Close

''  Bank reconciliation splits will be saved now
'   UpdateBankReconSplit adoConn
'issue 523
'modified by anol 19 Jan 2015
    If txtClientList.text = "Consolidated" Then
         szSQL = "UPDATE ConsolidatedBankList  " & _
                    "SET  " & _
                        "StatementDate = '" & Format(txtStatementDate.text, "dd mmmm yyyy") & "',  " & _
                        "SOB = " & Val(txtProjClosingBal.text) & " " & _
                    "WHERE conBankID = " & SelectedConBankID & ";"
            adoConn.Execute szSQL
    Else
            szSQL = "UPDATE tlbClientBanks " & _
                    "SET PCB = 0, " & _
                        "spare2 = '" & Format(txtStatementDate.text, "dd mmmm yyyy") & "',  " & _
                        "SOB = " & Val(txtProjClosingBal.text) & " " & _
                    "WHERE NominalCode = '" & txtAccountName.text & "' and Client_ID= '" & txtClientList.Tag & "';"
            adoConn.Execute szSQL
    End If
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT * FROM tlbBankReconClosingBal;"

   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With adoRST
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = txtAccountName.Tag
      .Fields.Item("BankCode").Value = txtAccountName.text
      .Fields.Item("StatementDate").Value = Format(txtStatementDate.text, "dd mmmm yyyy")
      .Fields.Item("AccBal").Value = CCur(txtAcBal.text)
      .Fields.Item("StOpenBal").Value = CCur(txtStOpenBal.text)
      .Fields.Item("ProjClBal").Value = CCur(txtProjClosingBal.text)
      .Update
      .Close
   End With
   Set adoRST = Nothing
'   'rollback implementation 18 July 2016
'   szSQL = "SELECT * FROM  tlbBankReconcilation WHERE TDate > ReconDate;"
'   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'   If adoRst.EOF Then
'        SaveBankReconciliation = True
'   Else
'        SaveBankReconciliation = False
'   End If
   SaveBankReconciliation = True
   If checksumRollbankBankrecon(adoConn, LastBankReconDate(adoConn)) = False Or checksumRollbankBankrecon(adoConn, txtStatementDate.text) = False Then
         SaveBankReconciliation = False
   End If
   Exit Function
Err:
   SaveBankReconciliation = False
End Function
Private Function LastBankReconDate(adoConn As ADODB.Connection) As String
    On Error GoTo Err
    Dim adoRST As New ADODB.Recordset
    Dim szSQL As String
    Dim temp As String
    szSQL = "Select MAX(StatementDate) as DT FROM tlbBankReconClosingBal B where StatementDate<>#" & txtStatementDate.text & "# AND ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & "'"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        temp = IIf(IsNull(adoRST("DT").Value), "1 Mar 1950", adoRST("DT").Value)
    End If
    adoRST.Close
    If IsDate(temp) Then
        LastBankReconDate = temp
    Else
        LastBankReconDate = "1 Mar 1950"
    End If
    Exit Function
Err:
    LastBankReconDate = "1 Mar 1950"
End Function
Private Function checksumRollbankBankrecon(adoConn As ADODB.Connection, szStatementDate As String) As Boolean  ' true means check passed and false means failed
'this procedure is testing before rollback , selected period is consistent then it will start rollback.
'he first recondate of tlbBankReconClosingBal openning balance is not correct even acc balance is not correct
'written by anol 2019 05 28
    Dim adoRST As New ADODB.Recordset
    Dim dbamount As Double
    Dim dbamountCompr As Double
    Dim szSQL As String
    'szStatementDate = "30/04/2019"
    szSQL = "Select SUM(ProjClBal-StOpenBal) as amt FROM tlbBankReconClosingBal B where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
            "' AND StatementDate=#" & Format(szStatementDate, "dd MMM yyyy") & "#"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        dbamountCompr = IIf(IsNull(adoRST("amt").Value), "0", adoRST("amt").Value)
    End If
    If dbamountCompr = 0 Then
        checksumRollbankBankrecon = True
        Exit Function
    End If
    adoRST.Close
    szSQL = "Select Sum(Reconciled) as amt From tlbReceipt R,Units,Property P where " & _
                    "R.UnitID=Units.UnitNumber AND Units.PropertyID=P.PropertyID and P.ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
                    "' AND  CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))=#" & Format(szStatementDate, "dd MMM yyyy") & "# AND Reconnow is not NULL"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        dbamount = IIf(IsNull(adoRST("amt").Value), "0", adoRST("amt").Value)
    End If
    adoRST.Close
    szSQL = "Select  Sum(Reconciled) as amt From tlbPayment P where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
                    "' AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))=#" & Format(szStatementDate, "dd MMM yyyy") & "#"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        dbamount = dbamount + IIf(IsNull(adoRST("amt").Value), "0", adoRST("amt").Value)
    End If
    adoRST.Close
    szSQL = "Select  Sum(Reconciled) as amt From  tlbBankPayment B where ClientID='" & txtClientList.Tag & "' AND Bank_AC='" & txtBC.Tag & _
                    "' AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))=#" & Format(szStatementDate, "dd MMM yyyy") & "#"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        dbamount = dbamount + IIf(IsNull(adoRST("amt").Value), "0", adoRST("amt").Value)
    End If
    adoRST.Close
    Set adoRST = Nothing
    If dbamountCompr - dbamount = 0 Then
        checksumRollbankBankrecon = True
    Else
        checksumRollbankBankrecon = False
        MsgBox "Could not Save Bank reconciliation.", vbInformation, "This statement cannot be reconciled"
    End If
End Function
Private Sub SpreadAmount(sFPS As Single)
   Dim iRow As Integer, i As Integer

   With flxStatementReconcile
      If sFPS = 3 Then
         'optReconciliation_Click 0
        '18 Feb 2015
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 9) = "Full" Or .TextMatrix(iRow, 0) = "-" Then _
                  .RowHeight(iRow) = 0
            If .TextMatrix(iRow, 0) = ">" Then .TextMatrix(iRow, 0) = "+"
         Next iRow
      'End of modification
         For iRow = 1 To .Rows - 1
            If .RowHeight(iRow) > 0 And Val(.TextMatrix(iRow, 8)) <> 0 Then
               .TextMatrix(iRow, 9) = "Saved"
            End If
         Next iRow
      Else
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) = "+" Or .TextMatrix(iRow, 0) = ">" Then
               i = iRow + 1
               If i < .Rows Then
                  If .TextMatrix(iRow, 8) = .TextMatrix(iRow, 6) And i < .Rows Then   'Fully booked
                     While .TextMatrix(i, 0) <> "+" And _
                           .TextMatrix(i, 0) <> ">" And _
                           .TextMatrix(i, 0) <> ""

                        .TextMatrix(i, 8) = .TextMatrix(i, 6)
                        .TextMatrix(i, 11) = "M"

                        i = i + 1
                     Wend
                  Else                                                                 'Part booked
                     While i < .Rows
                        If .TextMatrix(i, 0) <> "+" And _
                              .TextMatrix(i, 0) <> ">" And _
                              .TextMatrix(i, 0) <> "" Then
                           .TextMatrix(i, 8) = .TextMatrix(iRow, 8)
                           .TextMatrix(i, 11) = "M"
                        End If
                        i = i + 1
                     Wend
                  End If
               End If
            End If
         Next iRow
      End If
   End With
End Sub

Private Sub cmdSaveBk_Click(Index As Integer)
   If flxBankPay(0).TextMatrix(1, 0) = "" Then
      MsgBox "No data to save!", vbInformation + vbOKOnly, "Bank Payment & Receipt"
      Exit Sub
   End If

   If cmdUpdateBk(0).Enabled Then
      MsgBox "Please update data first.", vbInformation + vbOKOnly, "Saving Data"
      Exit Sub
   End If
   If MsgBox("Are you sure to save?", vbYesNo + vbQuestion, "Saving Data") = vbNo Then Exit Sub

   Dim szSQL As String
   Dim iRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim Rst1 As New ADODB.Recordset

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM tlbBankPayment"
   Rst1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

'Add New Records
   For iRow = 1 To flxBankPay(0).Rows - 1
      If flxBankPay(0).TextMatrix(iRow, 14) = "" Then
         Rst1.AddNew
         Rst1!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
         flxBankPay(0).TextMatrix(iRow, 14) = Rst1!My_ID
         Rst1!TRAN_ID = SlNumber(flxBankPay(0).TextMatrix(iRow, 1), "tlbBankPayment", adoConn)
         Rst1!BANK_AC = flxBankPay(0).TextMatrix(iRow, 0)
         Rst1!TRANS = flxBankPay(0).TextMatrix(iRow, 1)
         Rst1!TRAN_DATE = Format(flxBankPay(0).TextMatrix(iRow, 2), "DD MMMM YYYY")
         Rst1!UNIT_ID = flxBankPay(0).TextMatrix(iRow, 5)
         Rst1!Nominal_code = flxBankPay(0).TextMatrix(iRow, 6)
         Rst1!PROJ_REF = flxBankPay(0).TextMatrix(iRow, 7)              'Reference
         Rst1!DEPT_ID = flxBankPay(0).TextMatrix(iRow, 8)               'Fund
         Rst1!description = flxBankPay(0).TextMatrix(iRow, 9)
         Rst1!NET_AMOUNT = IIf(flxBankPay(0).TextMatrix(iRow, 10) = "", 0, CCur(flxBankPay(0).TextMatrix(iRow, 10)))
         Rst1!vat = IIf(IsNull(flxBankPay(0).TextMatrix(iRow, 11)), 0, Format(CCur(flxBankPay(0).TextMatrix(iRow, 11)), "0.00"))
         Rst1!TAX_CODE = flxBankPay(0).TextMatrix(iRow, 3)
         Rst1!TransactionType = IIf(flxBankPay(0).TextMatrix(iRow, 1) = "BR", 12, 11)     'Bank Receipt = sdoBR 12, Bank Payment = sdoBP 11
         Rst1.Update
      End If
   Next iRow
   Rst1.Close

   ShowMsgInTaskBar "Data has been saved successfully"
   flxBankPay(0).Clear
   flxBankPay(0).Rows = 2
   HandleTextBoxesBk True, False
   cmdNewBk(0).Enabled = True
   Label3(0).Visible = False
   Label3(3).Visible = False

   adoConn.Close

   Set Rst1 = Nothing
   Set adoConn = Nothing

   bChangesMade = False
End Sub
'
'Private Sub cmdSetAmtType_Click()
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   LoadRptAmtType "RECEIPT AMOUNT TYPE", adoConn, cmbRptAmtType
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub SupTotalPay()
   txtTReceiptTotal.text = "0.00"
   Dim iRow As Integer
   For iRow = 1 To flxTReceipt.Rows - 1
      If flxTReceipt.TextMatrix(iRow, 1) = "PI" Then
         txtTReceiptTotal.text = Format(CDbl(txtTReceiptTotal.text) + _
                                 CDbl(IIf(flxTReceipt.TextMatrix(iRow, 9) = _
                                 "", 0, flxTReceipt.TextMatrix(iRow, 9))), "0.00")
      Else
         txtTReceiptTotal.text = Format(CDbl(txtTReceiptTotal.text) - _
                                 CDbl(flxTReceipt.TextMatrix(iRow, 9)), "0.00")
      End If
   Next iRow
End Sub
'
'Private Sub cmdSPAmtType_Click()
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   LoadRptAmtType "RECEIPT AMOUNT TYPE", adoConn, cmbSPAmtType
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub cmdSPayAll_Click()
   Dim iRow As Integer, cDiff As Currency

   If bTotalPayTyped Then cDiff = Val(txtSPaymentTotal.text)

   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 2) <> "ADJI" And flxSPayment.TextMatrix(iRow, 9) <> "" Then
         If bTotalPayTyped Then
            If cDiff > Val(flxSPayment.TextMatrix(iRow, 9)) Then
               flxSPayment.TextMatrix(iRow, 10) = flxSPayment.TextMatrix(iRow, 9)
               cDiff = cDiff - Val(flxSPayment.TextMatrix(iRow, 9))
            Else
               flxSPayment.TextMatrix(iRow, 10) = Format(cDiff, "0.00")
               cDiff = 0
               baChangesMade(iRow) = IIf(Val(flxSPayment.TextMatrix(iRow, 10)) > 0, True, False)
               Exit For
            End If
         Else
            flxSPayment.TextMatrix(iRow, 10) = flxSPayment.TextMatrix(iRow, 9)
            txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + Val(flxSPayment.TextMatrix(iRow, 9)), "0.00")
         End If
         baChangesMade(iRow) = IIf(Val(flxSPayment.TextMatrix(iRow, 10)) > 0, True, False)
      End If
   Next iRow

   cGridSPTotal = TotalPaymentEntered
   txtPaymentEntered.text = Format(cGridSPTotal, "0.00")
End Sub

Private Sub cmdSPClose_Click()
   Unload Me
End Sub

Private Sub cmdSPFull_Click()
   If flxSPayment.row = 0 Then Exit Sub
   If flxSPayment.TextMatrix(flxSPayment.row, 2) = "ADJI" Then Exit Sub

   On Error GoTo ErrorHandler

   flxSPayment.col = 9
   If flxSPayment.row > 0 And flxSPayment.row <= flxSPayment.Rows - 1 Then
      If Val(flxSPayment.TextMatrix(flxSPayment.row, 10)) > 0 And Not bTotalPayTyped Then
         txtSPaymentTotal.text = Val(txtSPaymentTotal.text) - Val(flxSPayment.TextMatrix(flxSPayment.row, 10))
      End If

      If bTotalPayTyped Then                'Payment amount has put in the "Total Payment Amt"
         If Val(txtDiffPay.text) > Val(flxSPayment.TextMatrix(flxSPayment.row, 9)) Then
            flxSPayment.TextMatrix(flxSPayment.row, 10) = flxSPayment.TextMatrix(flxSPayment.row, 9)
         Else
            flxSPayment.TextMatrix(flxSPayment.row, 10) = IIf(baChangesMade(flxSPayment.row), flxSPayment.TextMatrix(flxSPayment.row, 10), txtDiffPay.text)
         End If
      Else
         flxSPayment.TextMatrix(flxSPayment.row, 10) = flxSPayment.TextMatrix(flxSPayment.row, 9)
         txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + CCur(flxSPayment.TextMatrix(flxSPayment.row, 10)), "0.00")
      End If

      cGridSPTotal = TotalPaymentEntered
      baChangesMade(flxSPayment.row) = IIf(Val(flxSPayment.TextMatrix(flxSPayment.row, 10)) > 0, True, False)
      txtPaymentEntered.text = Format(cGridSPTotal, "0.00")

      flxSPayment.row = flxSPayment.row + 1
      flxSPayment_Click
   End If
   Exit Sub

ErrorHandler:
   Debug.Print "Reached the end of the records"
End Sub

Private Sub cmdSupplierPayment_Click()
   If txtBC.text = "" Or txtClientList.text = "" Then
      If txtClientList.text = "" Then
         MsgBox "Please select a client account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtClientList
         Exit Sub
      Else
         MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl cmdBC
         Exit Sub
      End If
   End If

   Load frmPurchaseExpense
   frmPurchaseExpense.Show
   frmPurchaseExpense.tabPurExp.Tab = 1
   frmPurchaseExpense.tabPayment.Tab = 0
   frmPurchaseExpense.ZOrder 0
End Sub

Private Sub cmdTenantReceipt_Click()
   If txtBC.text = "" Or txtClientList.text = "" Then
      If txtClientList.text = "" Then
         MsgBox "Please select a client account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl txtClientList
         Exit Sub
      Else
         MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Cashbook"
         FocusControl cmdBC
         Exit Sub
      End If
   End If

   Load frmDemands3
   frmDemands3.Show
   frmDemands3.tabDmdRcpt.Tab = 2
   frmDemands3.tabPayment.Tab = 0
   frmDemands3.ZOrder 0
End Sub

Private Function TotalPaymentEntered() As Currency
   Dim i As Integer

   For i = 1 To flxSPayment.Rows - 1
      TotalPaymentEntered = TotalPaymentEntered + IIf(flxSPayment.TextMatrix(i, 10) = "", 0, Val(flxSPayment.TextMatrix(i, 10)))
   Next i
End Function

Private Function TotalReceiptEntered() As Currency
   Dim i As Integer

   For i = 1 To flxTReceipt.Rows - 1
      TotalReceiptEntered = TotalReceiptEntered + IIf(flxTReceipt.TextMatrix(i, 10) = "", 0, Val(flxTReceipt.TextMatrix(i, 10)))
   Next i
End Function

Private Sub cmdTRClose_Click()
   Unload Me
End Sub

Private Sub cmdTaxListBk_Click(Index As Integer)
   LoadVATBk

   picLeaseList.Left = txtVatBk(0).Left - 400 + tabPayRpt.Left + tabCashbook.Left
   picLeaseList.Top = txtVatBk(0).Top + txtVatBk(0).Height + tabPayRpt.Top + tabCashbook.Top
   picLeaseList.Width = 2300
   cmdGridUnitLookup(0).Left = picLeaseList.Width - cmdGridUnitLookup(0).Width
   Shape4(6).Width = picLeaseList.Width - cmdGridUnitLookup(0).Width - 50
   flxLeaseList.Width = 2000
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0
   FocusControl flxLeaseList
   sTextBox = "VAT"
End Sub

Private Sub LoadVATBk()
   flxLeaseList.Clear
   flxLeaseList.Cols = 4
   flxLeaseList.ColWidth(1) = 800
   flxLeaseList.ColWidth(2) = 800
   
   flxLeaseList.TextMatrix(0, 1) = "CODE"
   flxLeaseList.TextMatrix(0, 2) = "RATE"
   
      '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
    
   flxLeaseList.ColWidth(0) = 0
   flxLeaseList.ColWidth(3) = 0
   Label20(0).Width = 600
   Label20(0).Left = 50
   Label20(1).Width = 600
   Label20(1).Left = Label20(0).Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchID.Width = 600
   txtTenantSearchID.Left = 40
   
   txtTenantSearchName.Width = 600
   txtTenantSearchName.Left = txtTenantSearchID.Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchUnitName.Visible = False
   
         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   Label20(0).Caption = "Code"
   Label20(1).Caption = "Rate"
   Label20(2).Visible = False
   
   '~~~ End of config
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
   Conn2.Open getConnectionString

   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxLeaseList.Clear

      rstRec.MoveFirst
      flxLeaseList.ColAlignment(1) = vbRightJustify

      rRow = 1
      While Not rstRec.EOF
         flxLeaseList.TextMatrix(rRow, 1) = rstRec!VAT_CODE
         flxLeaseList.TextMatrix(rRow, 2) = rstRec!VAT_RATE
         rstRec.MoveNext
         If Not rstRec.EOF Then flxLeaseList.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close
   
   Set rstRec = Nothing
   Set Conn2 = Nothing
End Sub

Private Sub cmdTenantLookup_Click()
   If txtBC.text = "" Then
      MsgBox "Please select the client's bank account", vbCritical + vbOKOnly, "Receipt - Bank Account missing"
      FocusControl cmdBC
      Exit Sub
   End If

   Me.MousePointer = vbHourglass

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   sTextBox = "Lease"

   ConfigureFlxLeaseList
   If txtClientList.Tag = "ALL" And cboRptPropertyList.Column(0) = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag <> "ALL" And cboRptPropertyList.Column(0) = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & txtClientList.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag = "ALL" And cboRptPropertyList.Column(0) <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = '" & cboRptPropertyList.Column(0) & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag <> "ALL" And cboRptPropertyList.Column(0) <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & txtClientList.Tag & "' AND " & _
               "Units.PropertyID = '" & cboRptPropertyList.Column(0) & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   PopulateTenantLookup adoConn, szSQL

   adoConn.Close
   Set adoConn = Nothing

   txtTenantSearchID.text = ""
   txtTenantSearchName.text = ""
   txtTenantSearchUnitName.text = ""
   picLeaseList.Top = tabCashbook.Top + tabPayRpt.Top + Frame8(0).Top + txtTenantID.Top + txtTenantID.Height + 5
   picLeaseList.Left = tabCashbook.Left + tabPayRpt.Left + Frame8(0).Left + txtTenantID.Left + 5
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0

   Me.MousePointer = vbDefault
End Sub

Private Sub ConfigureFlxLeaseList()
   Dim szHeader As String
   picLeaseList.Width = 6375
   flxLeaseList.Width = 6255
   Label20(0).Caption = "Tenant Id"
   Label20(1).Caption = "Tenant Name"
   Label20(0).Width = 690
   Label20(0).Left = 120
   Label20(1).Width = 930
   Label20(1).Left = 1560
   Label20(2).Visible = True
   txtTenantSearchID.Left = 120
   txtTenantSearchName.Left = 1560
   txtTenantSearchUnitName.Left = 4080
   txtTenantSearchUnitName.Visible = True
   txtTenantSearchID.Width = 1335
   txtTenantSearchName.Width = 2415
   txtTenantSearchUnitName.Width = 1935
   Shape4(6).Width = 6015
   cmdGridUnitLookup(0).Left = 6080
   
   flxLeaseList.Clear
   flxLeaseList.Cols = 4
   flxLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Tenant ID|<Tenant Name|<Unit Name"
   flxLeaseList.FormatString = szHeader$
   flxLeaseList.ColWidth(0) = Label20(0).Left - flxLeaseList.Left   '240        Solid column
   flxLeaseList.ColWidth(1) = Label20(1).Left - Label20(0).Left - 20  '1400       'Tenant ID
   flxLeaseList.ColWidth(2) = Label20(2).Left - Label20(1).Left - 20         'Tenant Name
   flxLeaseList.ColWidth(3) = flxLeaseList.Left + flxLeaseList.Width - Label20(2).Left - 300 'Unit Name
   flxLeaseList.Rows = 2
End Sub

Private Sub LoadUnit()
End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtNote.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub cmdUnitMemoSave_Click()
   Dim conMemo As New ADODB.Connection
   Dim rstMemo_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   On Error GoTo Exception
   
   conMemo.Open getConnectionString

   sSQLQuery_ = "SELECT BankMemo " & _
                "FROM tlbClientBanks " & _
                "WHERE MY_ID = " & cmdBC.Tag & ";"
'Debug.Print sSQLQuery_
   rstMemo_.Open sSQLQuery_, conMemo, adOpenDynamic, adLockPessimistic
   
   If txtNote.text = "" Then
       rstMemo_.Fields.Item("BankMemo").Value = "<No memo saved>"
   Else
       rstMemo_.Fields.Item("BankMemo").Value = txtNote.text
   End If
   
   rstMemo_.Update

   rstMemo_.Close
   conMemo.Close
   Set rstMemo_ = Nothing
   Set conMemo = Nothing
   
   ShowMsgInTaskBar "Memo has been saved successfully."
   MemoButtonEnable False
   Exit Sub

Exception:
   
   MsgBox Err.Number & " - " & Err.description, vbOKOnly, "Error"
   rstMemo_.Close
   conMemo.Close
   Set rstMemo_ = Nothing
   Set conMemo = Nothing
End Sub

Private Sub cmdUpdateBk_Click(Index As Integer)
   If txtBkAc(0).text = "" Then
      MsgBox "Please select a Bank code.", vbInformation, "Bank"
      FocusControl txtBkAc(0)
      Exit Sub
   End If
   If txtNetBk(0).text = "" Then
      MsgBox "Please enter the amount.", vbInformation, "Bank"
      FocusControl txtNetBk(0)
      Exit Sub
   End If

   Dim sSql As String
   Dim adoConn As ADODB.Connection
   Dim adoRSTUnit As ADODB.Recordset
   Dim ClientName As String, PropertyName As String

   If (txtNetBk(0).text = "") Then
    MsgBox "Net Amound can not be empty", vbOKOnly, "Mandatory Data Mission"
    Exit Sub
   End If

   Set adoConn = New ADODB.Connection
   Set adoRSTUnit = New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString
   Dim iGrid As Integer
   If MsgBox("Do you want to update data?", vbYesNo + vbQuestion, "Update Data") = vbNo Then Exit Sub

   iGrid = 0

   If cmdEditBk(0).Enabled Then         'Not in Edit mode. New record adding
      If Not (flxBankPay(iGrid).Rows = 2 And flxBankPay(iGrid).TextMatrix(1, 0) = "") Then
         flxBankPay(iGrid).AddItem ""
      End If
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 0) = txtBkAc(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 1) = IIf(BANK_TYPE = "Bank Receipt", "BR", "BP")
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 2) = txtDateBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 3) = cmdTaxListBk(0).Caption
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 5) = cboBRPClient.Value
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 6) = txtNCBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 7) = txtReference(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 8) = txtDeptBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 9) = txtDetailsBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 10) = txtNetBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 11) = txtVatBk(0).text
      flxBankPay(iGrid).TextMatrix(flxBankPay(iGrid).Rows - 1, 12) = txtTotalBk(0).text

      HandleTextBoxesBk True, True
      cmdUpdateBk(0).Enabled = False
      cmdNewBk(0).Enabled = True
      FocusControl cmdNewBk(0)
   Else
      flxBankPay(0).TextMatrix(iCurEditRow, 0) = txtBkAc(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 2) = txtDateBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 3) = cmdTaxListBk(0).Caption
      flxBankPay(0).TextMatrix(iCurEditRow, 5) = cboBRPClient.Value
      flxBankPay(0).TextMatrix(iCurEditRow, 6) = txtNCBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 7) = txtReference(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 8) = txtDeptBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 9) = txtDetailsBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 10) = txtNetBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 11) = txtVatBk(0).text
      flxBankPay(0).TextMatrix(iCurEditRow, 12) = txtTotalBk(0).text

      HandleTextBoxesBk True, False
      cmdUpdateBk(0).Enabled = False
      cmdEditBk(0).Enabled = True
      cmdNewBk(0).Enabled = True
      FocusControl cmdNewBk(0)
   End If

   flxBankPay(0).row = 0      'reset the row selection to 0
   flxBankPay(0).Enabled = True
End Sub

Private Sub flxBankPay_DblClick(Index As Integer)
   If cmdEditBk(0).Enabled = False Then Exit Sub     'THE GRID IN THE EDIT MODE
End Sub

Private Sub flxBankPay_RowColChange(Index As Integer)
   If cmdEditBk(0).Enabled = False Then Exit Sub     'THE GRID IN THE EDIT MODE

   iSelected = 1

   If flxBankPay(Index).row <> flxBankPay(Index).Rows - 1 Then Exit Sub
End Sub

Private Sub flxSCrPoA_Click()
   If flxSCrPoA.RowSel = 0 Then Exit Sub
   If flxSCrPoA.TextMatrix(1, 0) = "" Then Exit Sub
   If cmdPayAllocate.Caption = "All&ocation Only" Then Exit Sub

   iCrPoARowSel = IIf(flxSCrPoA.TextMatrix(flxSCrPoA.RowSel, 8) > 0, flxSCrPoA.RowSel, 0)

   Dim i As Integer, iFlxCrPoACol As Integer

   iFlxCrPoACol = 9
   flxSCrPoA.col = iFlxCrPoACol

   If flxSCrPoA.TextMatrix(flxSCrPoA.row, 2) = "" Then Exit Sub

   txtCrPayment.BackColor = vbWhite
   txtCrPayment.Top = flxSCrPoA.CellTop + flxSCrPoA.Top
   txtCrPayment.Left = flxSCrPoA.CellLeft + flxSCrPoA.Left
   txtCrPayment.Width = flxSCrPoA.CellWidth
   txtCrPayment.Height = flxSCrPoA.RowHeight(flxSCrPoA.row) - 15
   txtCrPayment.text = flxSCrPoA.TextMatrix(flxSCrPoA.row, iFlxCrPoACol)
   txtCrPayment.Visible = True

   FocusControl txtCrPayment
   Label10(6).Caption = flxSCrPoA.row
End Sub

Private Sub flxSPayment_Click()
   Dim i As Integer, iFlxSPayCol As Integer

   If flxSPayment.TextMatrix(flxSPayment.row, 2) = "" Then Exit Sub

   iFlxSPayCol = 10
   flxSPayment.col = iFlxSPayCol

   szUndoText = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)

   If Not lblAllocating(1).Visible And flxSPayment.TextMatrix(flxSPayment.row, 2) <> "ADJI" Then
      txtSPayment.Top = flxSPayment.CellTop + flxSPayment.Top
      txtSPayment.Left = flxSPayment.CellLeft + flxSPayment.Left
      txtSPayment.Width = flxSPayment.ColWidth(iFlxSPayCol)
      txtSPayment.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtSPayment.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
      txtSPayment.Visible = True
      FocusControl txtSPayment
   End If
'  ALLOCATION - Place the txtCrReceipt text box in the grid to allocate agaist invoice
   If lblAllocating(1).Visible And Val(flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)) = 0 And Val(txtAllocatedDiff(1).text) > 0 Then
      If (InStr(lblAllocating(1).Caption, "ADJ") > 0 And InStr(flxSPayment.TextMatrix(flxSPayment.row, 2), "ADJ") > 0) Or _
         (InStr(lblAllocating(1).Caption, "ADJ") = 0 And InStr(flxSPayment.TextMatrix(flxSPayment.row, 2), "ADJ") = 0) Then
         txtCrReceipt.Top = flxSPayment.CellTop + flxSPayment.Top
         txtCrReceipt.Left = flxSPayment.CellLeft + flxSPayment.Left
         txtCrReceipt.Width = flxSPayment.ColWidth(iFlxSPayCol)
         txtCrReceipt.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
         txtCrReceipt.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
         txtCrReceipt.Visible = True
         FocusControl txtCrReceipt
         txtCrReceipt.BackColor = RGB(233, 232, 155)
         Label10(3).Caption = flxSPayment.row
      Else
         If InStr(lblAllocating(1).Caption, "ADJ") > 0 Then
            MsgBox "               Please select an Adjustment Invoice (ADJI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         Else
            MsgBox "                    Please select a Sales Invoice (SI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         End If
      End If
   End If
End Sub

Private Sub flxStatementReconcile_Click()
'    Dim selcol As Integer
'    'added by anol for locking issue 749 will not be editable on double click
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    selcol = flxStatementReconcile.col
'    flxStatementReconcile.col = 0
'    Dim rsLockCheck As New ADODB.Recordset
'    If flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbbankPayment " Then
'        rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbPayment where transactionID='" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) & _
'                "'", adoconn, adOpenStatic, adLockReadOnly
'    Else
'         rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbPayment where transactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) & _
'                "", adoconn, adOpenStatic, adLockReadOnly
'    End If
'    If Not rsLockCheck.EOF Then
'        If rsLockCheck("UserSessionID").Value <> "" And rsLockCheck("UserSessionID").Value <> UserSessionID Then 'you are not showing warning when it is locked by you (rsLockCheck("UserSessionID").Value <> UserSessionID)
'            flxStatementReconcile.CellBackColor = vbRed
'            MsgBox "The selected invoice is currently locked by '" & IIf(IsNull(rsLockCheck("WindowsUserName").Value), "", rsLockCheck("WindowsUserName").Value) & _
'                    "' on '" & IIf(IsNull(rsLockCheck("MachineName").Value), "", rsLockCheck("MachineName").Value) & "' in the '" & IIf(IsNull(rsLockCheck("Module").Value), "", rsLockCheck("Module").Value) & _
'                    "'" & vbCrLf & " screen for the Client '" & IIf(IsNull(rsLockCheck("ClientID").Value), "", rsLockCheck("ClientID").Value) & _
'                    "' and cannot be reconciled. Please wait until it is released.", vbInformation, "Warning"
'            Exit Sub
'        Else
'            flxStatementReconcile.CellBackColor = vbWhite
'            'you need to now lock it for your screen because other person has released it
'             If flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbbankPayment" Then
'                    adoconn.Execute "Update tlbbankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
'                    "',MachineName='" & WS_Name & "'," & _
'                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID='" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & "'"
'             ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbPayment" Then
'                     adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
'                    "',MachineName='" & WS_Name & "'," & _
'                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & ""
'             ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbReceipt" Then
'                   adoconn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
'                    "',MachineName='" & WS_Name & "'," & _
'                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & ""
'             End If
'
'        End If
'    End If
'    rsLockCheck.Close
'    Set rsLockCheck = Nothing
'    adoconn.Close
'    Set adoconn = Nothing
''   flxStatementReconcile.col = 0
''   If flxStatementReconcile.CellBackColor = vbRed Then
''        MsgBox "The selected invoice is currently locked by '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 14) & _
''                "' on '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 15) & "' in the '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 16) & "'" & vbCrLf & "" & _
''                        "screen for the Client '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 17) & _
''                        "' and cannot be reconciled. Please wait until it is released.", vbInformation, "Warning"
''        Exit Sub
''   End If
'   flxStatementReconcile.col = selcol
   
   
   Dim iRow As Integer

   iRow = flxStatementReconcile.row

   If iRow = flxStatementReconcile.Rows - 1 Then Exit Sub

   If flxStatementReconcile.col = 0 Then
      If flxStatementReconcile.TextMatrix(iRow, 0) = "+" And flxStatementReconcile.RowHeight(iRow + 1) = 0 Then
         flxStatementReconcile.TextMatrix(iRow, 0) = ">"
         For iRow = iRow + 1 To flxStatementReconcile.Rows - 1
            If flxStatementReconcile.TextMatrix(iRow, 0) = "+" Or flxStatementReconcile.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxStatementReconcile.TextMatrix(iRow, 0) = "-" Then flxStatementReconcile.RowHeight(iRow) = 240
         Next iRow
         Exit Sub
      End If
      If flxStatementReconcile.TextMatrix(iRow, 0) = ">" And flxStatementReconcile.RowHeight(iRow + 1) = 240 Then
         flxStatementReconcile.TextMatrix(iRow, 0) = "+"
         For iRow = iRow + 1 To flxStatementReconcile.Rows - 1
            If flxStatementReconcile.TextMatrix(iRow, 0) = "+" Or flxStatementReconcile.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxStatementReconcile.TextMatrix(iRow, 0) = "-" Then flxStatementReconcile.RowHeight(iRow) = 0
         Next iRow
         Exit Sub
      End If
   End If
End Sub

Private Sub flxStatementReconcile_DblClick()
    Dim InvoiceType As String
    Dim selcol As Integer
    'added by anol for locking issue 749 will not be editable on double click
    Dim adoConn As New ADODB.Connection
    If txtStatementDate.Visible = False Then
        'This means you are in consolidated bank section and your bank recon column need to be freezed.
        Exit Sub
    End If
    If flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) = "" Then Exit Sub 'that means no transaction ID in the row
    adoConn.Open getConnectionString
    selcol = flxStatementReconcile.col
    flxStatementReconcile.col = 0
    Dim rsLockCheck As New ADODB.Recordset
    If flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbBankPayment" Or _
                flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Bank Receipt" Or _
                    flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Bank Payment" Then
        rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbBankPayment where MY_ID='" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) & _
                "'", adoConn, adOpenStatic, adLockReadOnly
    ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbReceipt" Or _
                flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Sales Receipt on Account" Or _
                flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Sales Receipt Refund" Or _
                flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Sales Receipt" Then
         rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbReceipt where transactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) & _
                "", adoConn, adOpenStatic, adLockReadOnly
    ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbPayment" Or _
                flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Purchase Payment On Account" Or _
                    flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Purchase Payment Refund" Or _
                        flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) = "Purchase Payment" Then
         rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbPayment where transactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, btRecColNo) & _
                "", adoConn, adOpenStatic, adLockReadOnly
    End If
    If Not rsLockCheck.EOF Then
        If rsLockCheck("UserSessionID").Value <> "" And rsLockCheck("UserSessionID").Value <> UserSessionID Then 'you are not showing warning when it is locked by you (rsLockCheck("UserSessionID").Value <> UserSessionID)
            flxStatementReconcile.CellBackColor = vbRed

            MsgBox "The selected " & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 3) & " is currently locked by '" & IIf(IsNull(rsLockCheck("WindowsUserName").Value), "", rsLockCheck("WindowsUserName").Value) & _
                    "' on '" & IIf(IsNull(rsLockCheck("MachineName").Value), "", rsLockCheck("MachineName").Value) & "' in the '" & IIf(IsNull(rsLockCheck("Module").Value), "", rsLockCheck("Module").Value) & _
                    "'" & vbCrLf & " screen for the Client '" & IIf(IsNull(rsLockCheck("ClientID").Value), "", rsLockCheck("ClientID").Value) & _
                    "' and cannot be reconciled. Please wait until it is released.", vbInformation, "Warning"
            Exit Sub
        Else
            flxStatementReconcile.CellBackColor = vbWhite
            'you need to now lock it for your screen because other person has released it
             If flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbbankPayment" Then
                    adoConn.Execute "Update tlbbankPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
                    "',MachineName='" & WS_Name & "'," & _
                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID='" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & "'"
                    haveYouLockedAnyReccord = True
             ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbPayment" Then
                     adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
                    "',MachineName='" & WS_Name & "'," & _
                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & ""
                    haveYouLockedAnyReccord = True
             ElseIf flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 18) = "tlbReceipt" Then
                   adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='" & Now & "',Module='Cashbook',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & _
                    "',MachineName='" & WS_Name & "'," & _
                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID=" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 10) & ""
                    haveYouLockedAnyReccord = True
             End If
             
        End If
    End If
    rsLockCheck.Close
    Set rsLockCheck = Nothing
    adoConn.Close
    Set adoConn = Nothing
'   flxStatementReconcile.col = 0
'   If flxStatementReconcile.CellBackColor = vbRed Then
'        MsgBox "The selected invoice is currently locked by '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 14) & _
'                "' on '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 15) & "' in the '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 16) & "'" & vbCrLf & "" & _
'                        "screen for the Client '" & flxStatementReconcile.TextMatrix(flxStatementReconcile.row, 17) & _
'                        "' and cannot be reconciled. Please wait until it is released.", vbInformation, "Warning"
'        Exit Sub
'   End If
   flxStatementReconcile.col = selcol
   
   
   
       With flxStatementReconcile
       'Below line has been added by anol to stop Reconciled Balalnce to be reconcile again
        If .TextMatrix(.row, 9) <> "Full" Then
          If .TextMatrix(.row, 1) <> "" And .TextMatrix(.row, 0) <> "-" Then
             If Val(.TextMatrix(.row, 8)) = 0 Then
                .TextMatrix(.row, 8) = Format(IIf(.TextMatrix(.row, 6) <> "", .TextMatrix(.row, 6), Val(.TextMatrix(.row, 7)) * (-1)), "0.00")
                                             
    
                .TextMatrix(.row, 11) = "M"            'Modified
    
                Label1(27).Caption = Format(UnclearedBalance, "0.00")
                lblClosingBalance.Caption = Format(Val(lblClosingBalance.Caption) + Val(.TextMatrix(.row, 8)), "0.00")
                Label1(50).Caption = Format(Val(Label1(50).Caption) + Val(.TextMatrix(.row, 8)), "0.00")
             Else
                If .TextMatrix(.row, 9) = "" Then .TextMatrix(.row, 11) = ""
    
                Label1(27).Caption = Format(UnclearedBalance, "0.00")
                lblClosingBalance.Caption = Format(Val(lblClosingBalance.Caption) - Val(.TextMatrix(.row, 8)), "0.00")
                Label1(50).Caption = Format(Val(Label1(50).Caption) - Val(.TextMatrix(.row, 8)), "0.00")
                .TextMatrix(.row, 8) = "0.00"
             End If
            End If
          End If
       End With
End Sub

Private Sub flxStatementReconcile_KeyPress(KeyAscii As Integer)
   With flxStatementReconcile
      If .TextMatrix(.row, 1) <> "" And KeyAscii = 13 Then
         flxStatementReconcile_DblClick
      End If
   End With
End Sub

Private Sub cmdReconcileAll_Click()
   Dim i As Integer, cReconTotal As Currency

   cReconTotal = 0

   With flxStatementReconcile
      For i = 1 To .Rows - 1
         If .RowHeight(i) > 0 Then
            .row = i
            If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 0) <> "-" And Val(.TextMatrix(i, 8)) = 0 Then
               .TextMatrix(i, 8) = Format(IIf(.TextMatrix(i, 6) <> "", _
                                              .TextMatrix(i, 6), _
                                          Val(.TextMatrix(i, 7)) * (-1)), "0.00")
               .TextMatrix(i, 11) = "M"                                                'Modified
   
               Label1(27).Caption = Format(Val(Label1(27).Caption) + _
                                    (Val(.TextMatrix(i, 8)) * (-1)), "0.00")
               lblClosingBalance.Caption = Format(Val(lblClosingBalance.Caption) - _
                                    (Val(.TextMatrix(i, 8)) * (-1)), "0.00")
               Label1(50).Caption = Format(Val(Label1(50).Caption) + _
                                    Val(.TextMatrix(i, 8)), "0.00")
            End If
         End If
      Next i
   End With
End Sub

Private Sub Sum_RptPayVal()
   Dim iRow As Integer
   On Error Resume Next

   Label1(46).Caption = "0.00"
   
'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'Used 1 iteration to calculate summary of payment, receipt and statement value instead of using
'three different iterations which is redundant.
'Also checks for the transactions whose statement value does not match with the transaction value as
'this leads to mismatch of the balance of the statement of the account against the calculated
'reconciled summary figure showing at the bottom.
'
'   If optReconciliation(0).Value Then                             'Receipt Value
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" And flxStatementReconcile.RowHeight(iRow) > 0 Then
'            Label1(46).Caption = Val(Label1(46).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 6))
'            Label1(46).Caption = Format(Label1(46).Caption, "0.00")
'         End If
'      Next iRow
'
'      Label1(47).Caption = "0.00"                                 'Payment Value
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" And flxStatementReconcile.RowHeight(iRow) > 0 Then
'            Label1(47).Caption = Val(Label1(47).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 7))
'            Label1(47).Caption = Format(Label1(47).Caption, "0.00")
'         End If
'      Next iRow
'
''szHeader$ = "+|<Date|<TranID|<Type|<Account|<Reference|>ReceiptValue|>PaymentValue|>Statement|<Reconciliation| |Flag"
''                 1       2      3     4          5            6              7         8             9       10  11
'      Label1(50).Caption = "0.00"
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" And flxStatementReconcile.RowHeight(iRow) > 0 Then
'            Label1(50).Caption = Val(Label1(50).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 8))
'         End If
'      Next iRow
'   End If
'
'   If optReconciliation(1).Value Then
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" Then
'            Label1(46).Caption = Val(Label1(46).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 6))
'            Label1(46).Caption = Format(Label1(46).Caption, "0.00")
'         End If
'      Next iRow
'
'      Label1(47).Caption = "0.00"
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" Then
'            Label1(47).Caption = Val(Label1(47).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 7))
'            Label1(47).Caption = Format(Label1(47).Caption, "0.00")
'         End If
'      Next iRow
'
'      Label1(50).Caption = "0.00"
'      For iRow = 1 To flxStatementReconcile.Rows - 1
'         If Trim(flxStatementReconcile.TextMatrix(iRow, 0)) <> "-" Then
'            Label1(50).Caption = Val(Label1(50).Caption) + CCur(flxStatementReconcile.TextMatrix(iRow, 8))
'            Debug.Print CCur(flxStatementReconcile.TextMatrix(iRow, 8))
'         End If
'      Next iRow
'   End If

   Dim receipt, currentReceipt, payment, currentPayment, statement, currentStatement As Double
   receipt = 0
   payment = 0
   statement = 0
   
   Dim ignoreWarning As Boolean
   ignoreWarning = False
   
   For iRow = 1 To flxStatementReconcile.Rows - 1
        If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" Then
           
           If (flxStatementReconcile.TextMatrix(iRow, 6) <> "") Then
                receipt = receipt + CCur(flxStatementReconcile.TextMatrix(iRow, 6))
                currentReceipt = CDbl(flxStatementReconcile.TextMatrix(iRow, 6))
           Else
                currentReceipt = 0
           End If

           If (flxStatementReconcile.TextMatrix(iRow, 7) <> "") Then
                payment = payment + CCur(flxStatementReconcile.TextMatrix(iRow, 7))
                currentPayment = CDbl(flxStatementReconcile.TextMatrix(iRow, 7))
           Else
                currentPayment = 0
           End If
           
           If (flxStatementReconcile.TextMatrix(iRow, 8) <> "") Then
                If flxStatementReconcile.TextMatrix(iRow, 9) = "Full" Then
                    statement = statement + CCur(flxStatementReconcile.TextMatrix(iRow, 8))
                End If
                currentStatement = CDbl(flxStatementReconcile.TextMatrix(iRow, 8))
           Else
                currentStatement = 0
           End If
           
           
           If (flxStatementReconcile.TextMatrix(iRow, 9) = "Full") And (currentReceipt - currentPayment) <> currentStatement Then
           
                If Not ignoreWarning Then
                    Dim message As String
                    message = "The statement value of the reconciled transaction with Reference: " & flxStatementReconcile.TextMatrix(iRow, 2) & _
                    " Dated: " & flxStatementReconcile.TextMatrix(iRow, 1) & " of the A/C: " & _
                    flxStatementReconcile.TextMatrix(iRow, 4) & " does not match with the transaction figure. " & _
                    vbNewLine & vbNewLine & _
                    "Click YES to confirm and ignore the warning."
                    
                    message = "WARNING: " & vbNewLine & message
                
                    flxStatementReconcile.row = iRow
                    flxStatementReconcile.RowSel = iRow
                    Dim r As Integer
                    
                    For r = 1 To flxStatementReconcile.Cols - 1
                         flxStatementReconcile.col = r
                         flxStatementReconcile.CellBackColor = RGB(255, 159, 159)
                    Next r
                    
                    If MsgBox(message, vbYesNo, "Incorrect Statement Value") = vbYes Then
                        ignoreWarning = True
                    End If
                End If
           End If
        End If
   Next iRow
   
   Label1(46).Caption = Format(receipt, "0.00")
   Label1(47).Caption = Format(payment, "0.00")
   Label1(50).Caption = Format(statement, "0.00")
'END OF MODIFICATION

End Sub

Private Function UnclearedBalance() As Currency
   Dim iRow As Integer
   On Error Resume Next

   UnclearedBalance = 0
   For iRow = 1 To flxStatementReconcile.Rows - 1
      If flxStatementReconcile.TextMatrix(iRow, 0) <> "-" And flxStatementReconcile.RowHeight(iRow) > 0 Then
         UnclearedBalance = UnclearedBalance + CCur(flxStatementReconcile.TextMatrix(iRow, 6))
         UnclearedBalance = UnclearedBalance - CCur(flxStatementReconcile.TextMatrix(iRow, 7))
         If flxStatementReconcile.TextMatrix(iRow, 9) <> "Full" And flxStatementReconcile.RowHeight(iRow) > 0 Then _
            UnclearedBalance = UnclearedBalance - CCur(flxStatementReconcile.TextMatrix(iRow, 8))
      End If
   Next iRow
End Function

Private Sub flxStatementReconcile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 2 Then
      'MsgBox flxStatementReconcile.row
   End If
End Sub

Private Sub flxStatementReconcile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxStatementReconcile.ToolTipText = flxStatementReconcile.TextMatrix(flxStatementReconcile.MouseRow, flxStatementReconcile.MouseCol)
    If flxStatementReconcile.MouseCol = 1 Then
        flxStatementReconcile.ToolTipText = flxStatementReconcile.TextMatrix(flxStatementReconcile.MouseRow, 17)
    ElseIf flxStatementReconcile.MouseCol = 2 Then
         flxStatementReconcile.ToolTipText = flxStatementReconcile.TextMatrix(flxStatementReconcile.MouseRow, 20)
    End If
    
End Sub

Private Sub flxTCrPoA_Click()
   If flxTCrPoA.RowSel = 0 Then Exit Sub
   If flxTCrPoA.TextMatrix(1, 0) = "" Then Exit Sub
   If cmdRptAllocate.Caption = "All&ocation Only" Then Exit Sub

   iCrPoARowSel = IIf(flxTCrPoA.TextMatrix(flxTCrPoA.RowSel, 8) > 0, flxTCrPoA.RowSel, 0)

   Dim i As Integer, iFlxCrPoACol As Integer

   iFlxCrPoACol = 9
   flxTCrPoA.col = iFlxCrPoACol

   If flxTCrPoA.TextMatrix(flxTCrPoA.row, 2) = "" Then Exit Sub

   txtCrReceipt.BackColor = vbWhite
   txtCrReceipt.Top = flxTCrPoA.CellTop + flxTCrPoA.Top
   txtCrReceipt.Left = flxTCrPoA.CellLeft + flxTCrPoA.Left
   txtCrReceipt.Width = flxTCrPoA.CellWidth
   txtCrReceipt.Height = flxTCrPoA.RowHeight(flxTCrPoA.row) - 15
   txtCrReceipt.text = flxTCrPoA.TextMatrix(flxTCrPoA.row, iFlxCrPoACol)
'   flxTCrPoA.Enabled = False
   txtCrReceipt.Visible = True

   FocusControl txtCrReceipt
   Label10(4).Caption = flxTCrPoA.row
End Sub

Private Sub flxLeaseList_Click()
   If sTextBox = "Lease" Then
      txtTenantID.text = flxLeaseList.TextMatrix(flxLeaseList.row, 1) & " \ " & flxLeaseList.TextMatrix(flxLeaseList.row, 2)

      Dim adoConn As New ADODB.Connection

      ConfigureFlxTReceipt
      ConfigureFlxTCrPoA

      adoConn.Open getConnectionString

      LoadFlxTReceipt adoConn
      LoadFlxTCrPoA adoConn

      adoConn.Close
      Set adoConn = Nothing

      ReDim baChangesMade(flxTReceipt.Rows) As Boolean

      Frame5(0).Enabled = True
   End If

   If sTextBox = "Bank" Then
      txtBkAc(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
      cmdTaxListBk(0).Caption = "T1"
      nTaxCode = TaxRate(1)

      FocusControl txtReference(0)
   End If

   If sTextBox = "NC" Then
      txtNCBk(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
      FocusControl cmdDeptBk(0)
   End If

   If sTextBox = "Dept" Then
      txtDeptBk(0).text = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
      FocusControl txtDetailsBk(0)
   End If

   If sTextBox = "VAT" Then
      nTaxCode = flxLeaseList.TextMatrix(flxLeaseList.row, 2)
      cmdTaxListBk(0).Caption = flxLeaseList.TextMatrix(flxLeaseList.row, 1)
      txtVatBk(0).text = Format(IIf(txtNetBk(0).text = "", 0, Val(txtNetBk(0).text)) * _
                     (nTaxCode / 100), "0.00")
      FocusControl cmdUpdateBk(0)
   End If

   If sTextBox = "receipt" Then
      txtTenantID.text = flxLeaseList.TextMatrix(flxLeaseList.row, 1) & " \ " & flxLeaseList.TextMatrix(flxLeaseList.row, 2)
   End If

   If sTextBox = "Supplier" Then
      tabCashbook.Enabled = True
      tabPayRpt.Enabled = True
   End If

   flxLeaseList.Clear
   flxLeaseList.Cols = 2
   flxLeaseList.Rows = 2
   picLeaseList.Visible = False
End Sub

Private Sub LoadFlxTCrPoA(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim SQLStr1 As String, szaTenant() As String
   Dim iRow As Integer, szDataPath As String

   szaTenant = Split(txtTenantID.text, " \ ")

'   Get the details for the demand type selected
   SQLStr1 = "SELECT RPT.TransactionID, RPT.SageAccountNumber, " & _
                  "RPT.UnitID, RPT.RDate AS Dt, RPT.Ref, RPT.Details, " & _
                  "RPT.Amount, RPT.OSAmount as OS, RPT.DemandRef as DR, " & _
                  "RPT.AdjTag, tlbTransactionTypes.DESCRIPTION, " & _
                  "tlbTransactionTypes.TYPE_ID, RPT.Type as TT "
   SQLStr1 = SQLStr1 + _
             "FROM tlbReceipt AS RPT, tlbTransactionTypes " & _
             "WHERE RPT.SageAccountNumber = '" & szaTenant(0) & "' And " & _
                   "RPT.ReceiptView = True And RPT.Type = tlbTransactionTypes.TYPE_ID And " & _
                   "(tlbTransactionTypes.TYPE_ID = 2 OR tlbTransactionTypes.TYPE_ID = 4) " & _
             "Order By TransactionID;"
'Debug.Print SQLStr1

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRST.EOF
      flxTCrPoA.TextMatrix(iRow, 0) = adoRST!TransactionID
      If InStr(adoRST!description, "Credit") > 0 Then
         flxTCrPoA.TextMatrix(iRow, 1) = IIf(adoRST!AdjTag = "Y", "ADJC", adoRST!description)
      Else
         flxTCrPoA.TextMatrix(iRow, 1) = adoRST!description
      End If
      flxTCrPoA.TextMatrix(iRow, 2) = adoRST!SageAccountNumber
      flxTCrPoA.TextMatrix(iRow, 3) = IIf(IsNull(adoRST!unitid), "", adoRST!unitid)
      flxTCrPoA.TextMatrix(iRow, 4) = Format(adoRST!dt, "dd/mm/yyyy")
      flxTCrPoA.TextMatrix(iRow, 5) = IIf(IsNull(adoRST!ref), "", adoRST!ref)
      flxTCrPoA.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
      flxTCrPoA.TextMatrix(iRow, 7) = Format(adoRST!amount, "0.00")
      flxTCrPoA.TextMatrix(iRow, 8) = Format(adoRST!OS, "0.00")
      flxTCrPoA.TextMatrix(iRow, 9) = "0.00"
      flxTCrPoA.TextMatrix(iRow, 11) = IIf(IsNull(adoRST!DR), "", adoRST!DR)
      flxTCrPoA.TextMatrix(iRow, 12) = IIf(Val(flxTCrPoA.TextMatrix(iRow, 12)) = -1, "P/ADJ/CR", Format(flxTCrPoA.TextMatrix(iRow, 12), "0.00"))
      flxTCrPoA.TextMatrix(iRow, 13) = adoRST!TYPE_ID

      adoRST.MoveNext
      If Not adoRST.EOF Then flxTCrPoA.AddItem ""
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub LoadFlxSCrPoA(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim SQLStr1 As String
   Dim iRow As Integer, szDataPath As String

'   Get the details for the demand type selected
   SQLStr1 = "SELECT PYT.TransactionID, PYT.SageAccountNumber, " & _
                  "PYT.UnitID, PYT.PDate AS Dt, PYT.Ref, PYT.Details, " & _
                  "PYT.Amount, PYT.OSAmount as OS, PYT.PI as DR, " & _
                  "PYT.AdjTag, tlbTransactionTypes.DESCRIPTION, " & _
                  "tlbTransactionTypes.TYPE_ID, PYT.Type as TT "
   SQLStr1 = SQLStr1 + _
             "FROM tlbPayment AS PYT, tlbTransactionTypes " & _
             "WHERE PYT.SageAccountNumber = '" & cmbSPSupplier.Column(0) & "' And " & _
                   "PYT.PaymentView = True And PYT.Type = tlbTransactionTypes.TYPE_ID And " & _
                   "(tlbTransactionTypes.TYPE_ID = 7 OR tlbTransactionTypes.TYPE_ID = 9) " & _
             "Order By TransactionID;"
'Debug.Print SQLStr1
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRST.EOF
      flxSCrPoA.TextMatrix(iRow, 0) = adoRST!TransactionID
      If InStr(adoRST!description, "Credit") > 0 Then
         flxSCrPoA.TextMatrix(iRow, 1) = IIf(adoRST!AdjTag = "Y", "ADJC", adoRST!description)
      Else
         flxSCrPoA.TextMatrix(iRow, 1) = adoRST!description
      End If
      flxSCrPoA.TextMatrix(iRow, 2) = adoRST!SageAccountNumber
      flxSCrPoA.TextMatrix(iRow, 3) = IIf(IsNull(adoRST!unitid), "", adoRST!unitid)
      flxSCrPoA.TextMatrix(iRow, 4) = Format(adoRST!dt, "dd/mm/yyyy")
      flxSCrPoA.TextMatrix(iRow, 5) = IIf(IsNull(adoRST!ref), "", adoRST!ref)
      flxSCrPoA.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
      flxSCrPoA.TextMatrix(iRow, 7) = Format(adoRST!amount, "0.00")
      flxSCrPoA.TextMatrix(iRow, 8) = Format(adoRST!OS, "0.00")
      flxSCrPoA.TextMatrix(iRow, 9) = "0.00"
      flxSCrPoA.TextMatrix(iRow, 11) = IIf(IsNull(adoRST!DR), "", adoRST!DR)
      flxSCrPoA.TextMatrix(iRow, 12) = IIf(Val(flxSCrPoA.TextMatrix(iRow, 12)) = -1, "P/ADJ/CR", Format(flxSCrPoA.TextMatrix(iRow, 12), "0.00"))
      flxSCrPoA.TextMatrix(iRow, 13) = adoRST!TYPE_ID

      adoRST.MoveNext
      If Not adoRST.EOF Then flxSCrPoA.AddItem ""
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub LoadFlxTReceipt(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset, rdoSplits As New ADODB.Recordset
   Dim szSQL As String, szaTenant() As String, iRow As Integer, szDataPath As String, iSpRow As Integer

   szaTenant = Split(txtTenantID.text, " \ ")

'   INCLUDED NUMBER OF SPLITS
   szSQL = "SELECT tlbReceipt.TransactionID, tlbReceipt.DemandRef, tlbReceipt.AdjTag, tlbReceipt.SageAccountNumber, " & _
                  "tlbReceipt.UnitID, tlbReceipt.DDate, tlbReceipt.Ref, tlbReceipt.Details, tlbReceipt.Amount, " & _
                  "tlbReceipt.OSAmount, tlbReceipt.Type, " & _
                  "                           TT.DESCRIPTION, Count([DSR].[DSR]) AS TOTAL_SPLIT " & _
             "FROM (tlbReceipt INNER JOIN tlbTransactionTypes AS TT ON tlbReceipt.Type = TT.TYPE_ID) INNER JOIN " & _
                  "DemandSplitRecords AS DSR ON tlbReceipt.DemandRef = DSR.DemandID " & _
             "WHERE tlbReceipt.SageAccountNumber = '" & szaTenant(0) & "' And " & _
                   "tlbReceipt.ReceiptView=True AND TT.TYPE_ID=1 " & _
             "GROUP BY tlbReceipt.TransactionID, tlbReceipt.AdjTag, tlbReceipt.SageAccountNumber, tlbReceipt.UnitID, " & _
                  "tlbReceipt.DDate, tlbReceipt.Ref, tlbReceipt.Details, tlbReceipt.Amount, " & _
                  "tlbReceipt.OSAmount, tlbReceipt.DemandRef, tlbReceipt.Type, " & _
                  "                           TT.DESCRIPTION " & _
             "ORDER BY TransactionID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRST.EOF
      If adoRST!TOTAL_SPLIT > 1 Then
         flxTReceipt.TextMatrix(iRow, 0) = "+"
'  Add all splits below the header data
         szSQL = "SELECT DSR.SplitID, RIGHT(TT.CONSTANT, 2) AS Type, DR.SageAccountNumber as Tenant, " & _
                     "DR.UnitNumber, DSR.DueDate, DSR.SageRef, DSR.Description, DSR.TotalAmount, " & _
                     "DSR.DemandID " & _
                 "FROM DemandRecords as DR, DemandSplitRecords as DSR, tlbReceipt as RPT, tlbTransactionTypes as TT " & _
                 "WHERE DSR.DemandID = RPT.DemandRef AND RPT.TransactionID = " & adoRST!TransactionID & " AND " & _
                     "DSR.DemandID = DR.DemandID AND DR.TransactionType = TT.TYPE_ID " & _
                 "ORDER BY DSR.SplitID;"
         rdoSplits.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Debug.Print szSQL
         For iSpRow = 1 To adoRST!TOTAL_SPLIT
            flxTReceipt.AddItem ""
            flxTReceipt.TextMatrix(iRow + iSpRow, 0) = "-"
            flxTReceipt.TextMatrix(iRow + iSpRow, 1) = rdoSplits!splitID
            flxTReceipt.TextMatrix(iRow + iSpRow, 3) = rdoSplits!Tenant
            flxTReceipt.TextMatrix(iRow + iSpRow, 4) = rdoSplits!UnitNumber
            flxTReceipt.TextMatrix(iRow + iSpRow, 5) = rdoSplits!DueDate
            flxTReceipt.TextMatrix(iRow + iSpRow, 7) = rdoSplits!description
            flxTReceipt.TextMatrix(iRow + iSpRow, 8) = rdoSplits!TotalAmount
            flxTReceipt.TextMatrix(iRow + iSpRow, 11) = 0                     'Discount
            flxTReceipt.TextMatrix(iRow + iSpRow, 12) = rdoSplits!SageRef
            flxTReceipt.TextMatrix(iRow + iSpRow, 14) = rdoSplits!Type
            flxTReceipt.RowHeight(iRow + iSpRow) = 0
            rdoSplits.MoveNext
         Next iSpRow

         rdoSplits.Close
         Set rdoSplits = Nothing
      End If
      flxTReceipt.TextMatrix(iRow, 19) = adoRST!TransactionID
      flxTReceipt.TextMatrix(iRow, 1) = adoRST!DemandRef
      If InStr(adoRST!description, "Invoice") > 0 Then
         flxTReceipt.TextMatrix(iRow, 2) = IIf(adoRST!AdjTag = "Y", "ADJI", adoRST!description)
      Else
         flxTReceipt.TextMatrix(iRow, 2) = adoRST!description
      End If

      flxTReceipt.TextMatrix(iRow, 3) = adoRST!SageAccountNumber
      flxTReceipt.TextMatrix(iRow, 4) = adoRST!unitid
      flxTReceipt.TextMatrix(iRow, 5) = IIf(Not IsNull(adoRST!dDate), Format(adoRST!dDate, "dd/mm/yyyy"), "")
      flxTReceipt.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!ref), "", adoRST!ref)
      flxTReceipt.TextMatrix(iRow, 7) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
      flxTReceipt.TextMatrix(iRow, 8) = Format(adoRST!amount, "0.00")
      flxTReceipt.TextMatrix(iRow, 9) = Format(adoRST!OSAmount, "0.00")
      flxTReceipt.TextMatrix(iRow, 10) = "0.00"
      flxTReceipt.TextMatrix(iRow, 12) = IIf(IsNull(adoRST!DemandRef), "", adoRST!DemandRef)
      flxTReceipt.TextMatrix(iRow, 14) = adoRST!Type

      adoRST.MoveNext
      If Not adoRST.EOF Then flxTReceipt.AddItem ""
      If flxTReceipt.TextMatrix(iRow, 0) = "+" Then iRow = iRow + iSpRow - 1
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub LoadFlxSPayment(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset, rdoSplits As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, szDataPath As String, iSpRow As Integer

'   INCLUDED NUMBER OF SPLITS
   szSQL = "SELECT Pt.TransactionID, Pt.PI, Pt.AdjTag, Pt.SageAccountNumber, " & _
                  "Pt.UnitID, Pt.DDate, Pt.Ref, Pt.Details, Pt.Amount, " & _
                  "Pt.OSAmount, Pt.PI, Pt.Type, TT.DESCRIPTION " & _
             "FROM tlbPayment AS Pt INNER JOIN tlbTransactionTypes AS TT ON Pt.Type = TT.TYPE_ID " & _
             "WHERE Pt.SageAccountNumber = '" & cmbSPSupplier.Column(0) & "' And " & _
                   "Pt.PaymentView=True AND TT.TYPE_ID=6 " & _
             "GROUP BY Pt.TransactionID, Pt.AdjTag, Pt.SageAccountNumber, Pt.UnitID, " & _
                  "Pt.DDate, Pt.Ref, Pt.Details, Pt.Amount, " & _
                  "Pt.OSAmount, Pt.PI, Pt.Type, TT.DESCRIPTION " & _
             "ORDER BY TransactionID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRST.EOF
      flxSPayment.TextMatrix(iRow, 19) = adoRST!TransactionID
      flxSPayment.TextMatrix(iRow, 1) = adoRST!TransactionID
      If InStr(adoRST!description, "Invoice") > 0 Then
         flxSPayment.TextMatrix(iRow, 2) = IIf(adoRST!AdjTag = "Y", "ADJI", adoRST!description)
      Else
         flxSPayment.TextMatrix(iRow, 2) = adoRST!description
      End If

      flxSPayment.TextMatrix(iRow, 3) = adoRST!SageAccountNumber
      flxSPayment.TextMatrix(iRow, 4) = adoRST!unitid
      flxSPayment.TextMatrix(iRow, 5) = IIf(Not IsNull(adoRST!dDate), Format(adoRST!dDate, "dd/mm/yyyy"), "")
      flxSPayment.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!ref), "", adoRST!ref)
      flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
      flxSPayment.TextMatrix(iRow, 8) = Format(adoRST!amount, "0.00")
      flxSPayment.TextMatrix(iRow, 9) = Format(adoRST!OSAmount, "0.00")
      flxSPayment.TextMatrix(iRow, 10) = "0.00"
      flxSPayment.TextMatrix(iRow, 12) = IIf(IsNull(adoRST!Pi), "", adoRST!Pi)
      flxSPayment.TextMatrix(iRow, 14) = adoRST!Type

      adoRST.MoveNext
      If Not adoRST.EOF Then flxSPayment.AddItem ""
      If flxSPayment.TextMatrix(iRow, 0) = "+" Then iRow = iRow + iSpRow - 1
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigureFlxTCrPoA()
   Dim szHeader As String

   flxTCrPoA.Clear
   flxTCrPoA.Cols = 14
   flxTCrPoA.Rows = 2

   szHeader$ = "<No.|<Type|<Tenant A/C|<Unit ID|<Date" & _
               "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
               "|>Receipt £|>Discount|<DemandID|>SAGE O/S £"
   flxTCrPoA.FormatString = szHeader$

   flxTCrPoA.ColWidth(0) = 700    'Transaction ID
   flxTCrPoA.ColWidth(1) = 1700   'Type
   flxTCrPoA.ColWidth(2) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
   flxTCrPoA.ColWidth(3) = 1000   'Unit ID
   flxTCrPoA.ColWidth(4) = 1000   'Date
   flxTCrPoA.ColWidth(5) = 1400   'Ref
   flxTCrPoA.ColWidth(6) = 2560   'Details
   flxTCrPoA.ColWidth(7) = 1100   'Amount
   flxTCrPoA.ColWidth(8) = 1100   'O/S Amount
   flxTCrPoA.ColWidth(9) = 1100   'Receipt
   flxTCrPoA.ColWidth(10) = 0     'Discount
   flxTCrPoA.ColWidth(11) = 0     'DemandID
   flxTCrPoA.ColWidth(12) = 0     'SAGE O/S £
   flxTCrPoA.ColWidth(13) = 0     'Type ID    ID of col 1

   flxTCrPoA.RowHeightMin = 285

   flxTCrPoA.row = 0
   flxTCrPoA.col = 0
End Sub

Private Sub ConfigureFlxSCrPoA()
   Dim szHeader As String

   flxSCrPoA.Clear
   flxSCrPoA.Cols = 14
   flxSCrPoA.Rows = 2

   szHeader$ = "<No.|<Type|<Tenant A/C|<Unit ID|<Date" & _
               "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
               "|>Receipt £|>Discount|<DemandID|>SAGE O/S £"
   flxSCrPoA.FormatString = szHeader$

   flxSCrPoA.ColWidth(0) = 700    'Transaction ID
   flxSCrPoA.ColWidth(1) = 1700   'Type
   flxSCrPoA.ColWidth(2) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
   flxSCrPoA.ColWidth(3) = 1000   'Unit ID
   flxSCrPoA.ColWidth(4) = 1000   'Date
   flxSCrPoA.ColWidth(5) = 1400   'Ref
   flxSCrPoA.ColWidth(6) = 2560   'Details
   flxSCrPoA.ColWidth(7) = 1100   'Amount
   flxSCrPoA.ColWidth(8) = 1100   'O/S Amount
   flxSCrPoA.ColWidth(9) = 1100   'Receipt
   flxSCrPoA.ColWidth(10) = 0     'Discount
   flxSCrPoA.ColWidth(11) = 0     'DemandID
   flxSCrPoA.ColWidth(12) = 0     'SAGE O/S £
   flxSCrPoA.ColWidth(13) = 0     'Type ID    ID of col 1

   flxSCrPoA.RowHeightMin = 285

   flxSCrPoA.row = 0
   flxSCrPoA.col = 0
End Sub

Private Sub ConfigureFlxTReceipt()
   Dim szHeader As String

   flxTReceipt.Clear
   flxTReceipt.Cols = 20
   flxTReceipt.Rows = 3
   flxSCrPoA.RowHeightMin = 285
   
   szHeader$ = "|<No.|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
               "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
               "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo."
   flxTReceipt.FormatString = szHeader$

   flxTReceipt.ColWidth(0) = 200    'Sign
   flxTReceipt.ColAlignment(0) = vbCenter
   flxTReceipt.ColWidth(1) = 620    'No
   flxTReceipt.ColWidth(2) = 1140   'Type
   flxTReceipt.ColWidth(3) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
   flxTReceipt.ColWidth(4) = 1000   'Unit ID
   flxTReceipt.ColWidth(5) = 1000   'Date
   flxTReceipt.ColWidth(6) = 1360   'Ref
   flxTReceipt.ColWidth(7) = 3000   'Details
   flxTReceipt.ColWidth(8) = 1100   'Amount
   flxTReceipt.ColWidth(9) = 1100   'O/S Amount
   flxTReceipt.ColWidth(10) = 1100   'Receipt
   flxTReceipt.ColWidth(11) = 0     'Discount
   flxTReceipt.ColWidth(12) = 0     'DemandID
   flxTReceipt.ColWidth(13) = 0     'SAGE O/S £
   flxTReceipt.ColWidth(14) = 0     'Transaction Type - linked with column 1 Type
   flxTReceipt.ColWidth(15) = 0     'R/A; R -> receipt, A -> allocation
   flxTReceipt.ColWidth(16) = 0     'allocation ref
   flxTReceipt.ColWidth(17) = 0     'allocation amount
   flxTReceipt.ColWidth(18) = 0     'Sage Department
   flxTReceipt.ColWidth(19) = 0     'Receipt No

   flxTReceipt.RowHeight(0) = 0
   
'   Label19(10).Left = 200
'   Label19(11).Left = 1000
'   Label19(12).Left = 1500
'   Label19(13).Left = 2000
'   Label19(14).Left = 2500
'   Label19(15).Left = 3000
'   Label19(17).Left = 3500
'   Label19(18).Left = 4000
'   Label19(19).Left = 4500
End Sub

Private Sub ConfigureFlxSPayment()
   Dim szHeader As String

   flxSPayment.Clear
   flxSPayment.Cols = 20
   flxSPayment.Rows = 3

   szHeader$ = "|<No.|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
               "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
               "|>Receipt £|>Discount|<DemandID|>SAGE O/S £|<RptNo."
   flxSPayment.FormatString = szHeader$

   flxSPayment.ColWidth(0) = 200    'Sign
   flxSPayment.ColAlignment(0) = vbCenter
   flxSPayment.ColWidth(1) = 620    'No
   flxSPayment.ColWidth(2) = 1140   'Type
   flxSPayment.ColWidth(3) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
   flxSPayment.ColWidth(4) = 1000   'Unit ID
   flxSPayment.ColWidth(5) = 1000   'Date
   flxSPayment.ColWidth(6) = 1360   'Ref
   flxSPayment.ColWidth(7) = 3000   'Details
   flxSPayment.ColWidth(8) = 1100   'Amount
   flxSPayment.ColWidth(9) = 1100   'O/S Amount
   flxSPayment.ColWidth(10) = 1100  'Receipt
   flxSPayment.ColWidth(11) = 0     'Discount
   flxSPayment.ColWidth(12) = 0     'DemandID
   flxSPayment.ColWidth(13) = 0     'SAGE O/S £
   flxSPayment.ColWidth(14) = 0     'Transaction Type - linked with column 1 Type
   flxSPayment.ColWidth(15) = 0     'R/A; R -> receipt, A -> allocation
   flxSPayment.ColWidth(16) = 0     'allocation ref
   flxSPayment.ColWidth(17) = 0     'allocation amount
   flxSPayment.ColWidth(18) = 0     'Sage Department
   flxSPayment.ColWidth(19) = 0     'Receipt No
End Sub

Private Sub flxTReceipt_Click()
   Dim i As Integer, iFlxTRptCol As Integer
   Dim iCurRowHeight As Integer

   If flxTReceipt.TextMatrix(flxTReceipt.row, 2) = "" Then Exit Sub

   If flxTReceipt.col = 0 And flxTReceipt.TextMatrix(flxTReceipt.row, 0) = "+" Then          'Expanding the grid
      flxTReceipt.TextMatrix(flxTReceipt.row, 0) = ">"
      iCurRowHeight = flxTReceipt.RowHeight(flxTReceipt.row)
      i = 1

      While flxTReceipt.TextMatrix(flxTReceipt.row + i, 0) = "-"
         flxTReceipt.RowHeight(flxTReceipt.row + i) = iCurRowHeight
         i = i + 1
      Wend
      Exit Sub
   End If
   If flxTReceipt.col = 0 And flxTReceipt.TextMatrix(flxTReceipt.row, 0) = ">" Then          'Squeezing the grid
      flxTReceipt.TextMatrix(flxTReceipt.row, 0) = "+"
      i = 1
      While flxTReceipt.TextMatrix(flxTReceipt.row + i, 0) = "-"
         flxTReceipt.RowHeight(flxTReceipt.row + i) = 0
         i = i + 1
      Wend
      Exit Sub
   End If

   iFlxTRptCol = 10
   flxTReceipt.col = iFlxTRptCol

   szUndoText = flxTReceipt.TextMatrix(flxTReceipt.row, iFlxTRptCol)

   If Not lblAllocating(0).Visible And flxTReceipt.TextMatrix(flxTReceipt.row, 2) <> "ADJI" Then
      txtTReceipt.Top = flxTReceipt.CellTop + flxTReceipt.Top
      txtTReceipt.Left = flxTReceipt.CellLeft + flxTReceipt.Left
      txtTReceipt.Width = flxTReceipt.ColWidth(iFlxTRptCol)
      txtTReceipt.Height = flxTReceipt.RowHeight(flxTReceipt.row) - 15
      txtTReceipt.text = flxTReceipt.TextMatrix(flxTReceipt.row, iFlxTRptCol)
      txtTReceipt.Visible = True
      FocusControl txtTReceipt
   End If
'  ALLOCATION - Place the txtCrReceipt text box in the grid to allocate agaist invoice
   If lblAllocating(0).Visible And Val(flxTReceipt.TextMatrix(flxTReceipt.row, iFlxTRptCol)) = 0 And Val(txtAllocatedDiff(0).text) > 0 Then
      If (InStr(lblAllocating(0).Caption, "ADJ") > 0 And InStr(flxTReceipt.TextMatrix(flxTReceipt.row, 2), "ADJ") > 0) Or _
         (InStr(lblAllocating(0).Caption, "ADJ") = 0 And InStr(flxTReceipt.TextMatrix(flxTReceipt.row, 2), "ADJ") = 0) Then
         txtCrReceipt.Top = flxTReceipt.CellTop + flxTReceipt.Top
         txtCrReceipt.Left = flxTReceipt.CellLeft + flxTReceipt.Left
         txtCrReceipt.Width = flxTReceipt.ColWidth(iFlxTRptCol)
         txtCrReceipt.Height = flxTReceipt.RowHeight(flxTReceipt.row) - 15
         txtCrReceipt.text = flxTReceipt.TextMatrix(flxTReceipt.row, iFlxTRptCol)
'         flxTReceipt.Enabled = False
         txtCrReceipt.Visible = True
         FocusControl txtCrReceipt
         txtCrReceipt.BackColor = RGB(233, 232, 155)
         Label10(3).Caption = flxTReceipt.row
      Else
         If InStr(lblAllocating(0).Caption, "ADJ") > 0 Then
            MsgBox "               Please select an Adjustment Invoice (ADJI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         Else
            MsgBox "                    Please select a Sales Invoice (SI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         End If
      End If
   End If
End Sub

Private Sub flxTReceipt_Scroll()
   If txtTReceipt.Visible Then      'The grid is in edting mode
      txtTReceipt.text = szUndoText
      flxTReceipt.Enabled = True
      txtTReceipt.Visible = False
   End If
   If txtCrReceipt.Visible Then
      txtCrReceipt.text = szUndoText
      flxTReceipt.Enabled = True
      txtCrReceipt.Visible = False
   End If
End Sub

Private Sub Form_Activate()
   bTotalRptTyped = False
   cGridRptTotal = 0
   cGridSPTotal = 0
   iCrPoARowSel = 0
   If IsLoadedAndVisible("frmCashbook") Then
        FocusControl cmdClientList
   End If

   If UCase(User) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
'      cboClientID.ListIndex = 2
'      cboBC.ListIndex = 0
'      tabCashbook.Tab = 3

      Exit Sub
   End If
End Sub
Private Sub fixLegacyDateformatReconow(adoConn As ADODB.Connection)
    'some legacy date format like '30 november 2011#FULL' was found in the reconnow field which is a text field that has been remformatted like '30/11/2011#FULL' with this procedure
    'while converting this text field to datetime they were not fitting function like LEFT(x,10) now they shall fit
    'written by anol 2019-12-02
    On Error GoTo Err
    Dim rsReceipt As New ADODB.Recordset
    Dim part
   rsReceipt.Open "Select  * from tlbReceipt R where len(R.reconnow)>15", adoConn, adOpenKeyset, adLockReadOnly
   While Not rsReceipt.EOF
            part = Split(rsReceipt("reconnow").Value, "#")
            If UBound(part) > 0 Then
               If Not IsDate(part(0)) Then
                    rsReceipt.Close
                    Set rsReceipt = Nothing
                    Exit Sub
               End If
               adoConn.Execute "Update tlbReceipt Set reconnow ='" & Format(part(0), "dd/mm/yyyy") & "#" & part(1) & "'  where TransactionID =" & rsReceipt("TransactionID").Value & ""
               'Debug.Print rsReceipt.RecordCount
            End If
        rsReceipt.MoveNext
   Wend
   rsReceipt.Close
   Set rsReceipt = Nothing
   Exit Sub
Err:
End Sub
Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection, adoRST As New ADODB.Recordset
   UserSessionID = GetTimeStamp
   Label2.Caption = ""
'   connect to database
   adoConn.Open getConnectionString

   Me.Height = 8745
   Me.Width = 13080
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
'   fraConinfo.BackColor = MODULEBACKCOLOR
   tabCashbook.BackColor = Me.BackColor

   iSelected = 0
   nTaxCode = GetVATRate(1, adoConn)
   yWorkingTabCashBook = 255
   yWorkingTabPayRpt = 255

   tabCashbook.Tab = 0
   tabPayRpt.Tab = 0

   PrepareListBankTransf adoConn

   ConfigureFlxCashBook

   bChangesMade = False

   'LoadClients adoConn
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        txtClientList.Tag = adoRST.Fields("CLIENTID").Value
        txtClientList.text = adoRST.Fields("CLIENTNAME").Value
   End If
'This fix have been remmed by anol 2020-08-11
'   Debug.Print time & "Legacy"
'   Call fixLegacyDateformatReconow(adoConn)
'   Debug.Print time & "Legacy1"
   adoConn.Close
   Set adoConn = Nothing
   bLoad = True
   'cboClientID.ListIndex = 0
  ' If cboBC.ListCount > 0 Then cboBC.ListIndex = 0
   
    'If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
        Call WheelHook(Me.hWnd)
  ' End If
End Sub

Private Sub ConfigureFlxCashBook()
   Dim i As Integer, iCol As Integer
   Dim szHeader As String

   flxCashBook.Cols = 10
   flxCashBook.RowHeight(0) = 0
   i = 35
'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
'                    ^           ^         ^           ^       ^      ^          ^           ^

   szHeader$ = "|<Date|<Trans|<Account|<Ref|<Details" & _
               "|>Debit|>Credit|>Reconciled|<Dt"
   flxCashBook.FormatString = szHeader$

   flxCashBook.ColWidth(0) = Label1(i).Left - flxCashBook.Left
   For iCol = 1 To flxCashBook.Cols - 2
      i = i + 1
      flxCashBook.ColWidth(iCol) = Label1(i).Left - Label1(i - 1).Left
   Next iCol
   flxCashBook.ColWidth(iCol) = flxCashBook.Left + flxCashBook.Width - Label1(i).Left - 320

   Label1(48).Left = Label1(40).Left
   Label1(48).Width = flxCashBook.ColWidth(6)

   Label1(49).Left = Label1(41).Left
   Label1(49).Width = flxCashBook.ColWidth(7)
    flxCashBook.Cols = 11
    flxCashBook.ColWidth(10) = 0
End Sub

Private Sub PrepareListBankTransf(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String, i As Integer, j As Integer
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT BANK TRANSFER COMBO ******************************************
   szSQL = "SELECT DISTINCT tlbClientBanks.CLIENT_ID, Client.ClientName " & _
           "FROM tlbClientBanks, Client " & _
           "WHERE tlbClientBanks.CLIENT_ID = Client.ClientID " & _
           "ORDER BY CLIENT_ID;"

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes
   While Not adoRST.EOF
      cboClientName.AddItem adoRST.Fields.Item("ClientName").Value
      adoRST.MoveNext
   Wend
   adoRST.Close
'*************************************** FUND ******************************************
   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund " & _
           "ORDER BY FundID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Debug.Print szSQL
   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count
   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   
   cboFundBankTransf.Column() = Data()
   
NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   
   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

NoRes:
   adoRST.Close
   Set adoRST = Nothing

   ConfigureFlxLeaseList

   szSQL = "SELECT Tenants.SageAccountNumber, Name, UnitNumber " & _
          "From Tenants, LeaseDetails " & _
          "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber " & _
         "ORDER BY Tenants.SageAccountNumber;"

   PopulateTenantLookup adoConn, szSQL

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   Set adoRST = Nothing

   ConfigureFlxLeaseList

   szSQL = "SELECT Tenants.SageAccountNumber, Name, UnitNumber " & _
          "From Tenants, LeaseDetails " & _
          "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber " & _
         "ORDER BY Tenants.SageAccountNumber;"

   PopulateTenantLookup adoConn, szSQL
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmCashbook.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim iRow As Integer
   Dim X As Boolean
   Dim adoConn As New ADODB.Connection

   If txtBC.text = "" Then GoTo FrmClose
        For iRow = 1 To flxStatementReconcile.Rows - 1
            If Val(flxStatementReconcile.TextMatrix(iRow, 8)) <> 0 And flxStatementReconcile.TextMatrix(iRow, 9) = "" Then
                bChangesMade = True
            End If
        Next
'         If bChangesMade Then
'             X = MsgBox("Do you want to save changes?", vbQuestion + vbYesNoCancel, "Data Saving")
'             If X = vbCancel Then Cancel = 1
'
'          Else
'             Exit Sub
'         End If
    'validation placed on 02 08 2016 by anol
        
'       If Not CheckTranDate_StDate And bChangesMade Then
'          ShowMsgInTaskBar "You cannot reconcile transactions with a date after the bank statement date", "Y", "N"
'          Exit Sub
'       End If
'
'       If txtStatementDate.text <> "" And bChangesMade Then
'           adoConn.Open getConnectionString
'          If Not IsBankStDtValid(adoConn) Then
'             MsgBox "Bank Reconciliation for this date has been done. Please select another date.", vbCritical + vbOKOnly, "Multiple Bank Reconciliation Date."
'             txtStatementDate.SetFocus
'             adoConn.Close
'             Set adoConn = Nothing
'             Exit Sub
'          End If
'          adoConn.Close
'          Set adoConn = Nothing
'       End If
      'End of validation
      
      '  Saving the entries but it will not book as reconciled
   If Val(txtProjClosingBal.text) = Val(lblClosingBalance.Caption) And Val(txtProjClosingBal.text) <> 0 And bChangesMade Then
      If MsgBox("Your bank reconciliation agrees. Do you wish to complete it?", vbQuestion + _
                 vbYesNo, "Bank Reconciliation") = vbNo Then
         GoTo FrmClose
      Else
         'validation placed on 02 08 2016 by anol
        
       If Not CheckTranDate_StDate And bChangesMade Then
          ShowMsgInTaskBar "You cannot reconcile transactions with a date after the bank statement date", "Y", "N"
           Cancel = 1
          Exit Sub
       End If
    
       If txtStatementDate.text <> "" And bChangesMade Then
           adoConn.Open getConnectionString
          If Not IsBankStDtValid(adoConn) Then
          'MsgBox " A Bank Reconciliation has been completed at this date. Please select another date.", vbCritical + vbOKOnly, "Invalid Bank Reconciliation Date."
             If MsgBox(" A Bank Reconciliation has been completed at this date. Do you want to select another date?", vbYesNo, "Invalid Bank Reconciliation Date.") = vbNo Then Exit Sub
             Cancel = 1
             FocusControl txtStatementDate
             adoConn.Close
             Set adoConn = Nothing
             Exit Sub
          End If
          adoConn.Close
          Set adoConn = Nothing
       End If
      'End of validation
'         X = SaveBankReconciliation(adoConn)
'         adoConn.Close
'         Set adoConn = Nothing
        adoConn.Open getConnectionString
        adoConn.BeginTrans
        X = SaveBankReconciliation(adoConn)
        If X = False Then
           '#
'           Dim sChoice As Single
'            sChoice = MsgBox("There are some reconciled transactions found dated after their reconciled statement date." + Chr(13) + _
'                          "Please contact PCM Support to correct these transactions before proceeding further." + Chr(13) + _
'                          "Click OK to print a list of these transactions.", vbCritical + vbOKCancel, _
'                          "Incorrect Data")
'            If sChoice = vbOK Then
'               ShowReport App.Path & szReportPath & "\TranDt_BankRecDate.rpt"
'            End If
           '#
           MsgBox "An error occured.Bank Reconciliation has not been saved.", vbInformation, "warning"
           adoConn.RollbackTrans
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
        adoConn.CommitTrans
        adoConn.Close
        Set adoConn = Nothing
   
         GoTo FrmClose
      End If
   End If

   'start saving only

   For iRow = 1 To flxStatementReconcile.Rows - 1
      If Val(flxStatementReconcile.TextMatrix(iRow, 8)) <> 0 And flxStatementReconcile.TextMatrix(iRow, 9) = "" And bChangesMade Then
         If MsgBox("You have an existing bank reconciliation in progress. Do you wish to save it?", vbYesNo, "Bank Reconciliation") = vbYes Then
            'adoConn.Open getConnectionString
            'validation placed on 02 08 2016 by anol
        
                    If Not CheckTranDate_StDate And bChangesMade Then
                       ShowMsgInTaskBar "You cannot reconcile transactions with a date after the bank statement date", "Y", "N"
                        Cancel = 1
                       Exit Sub
                    End If
            
                    If txtStatementDate.text <> "" And bChangesMade Then
                        adoConn.Open getConnectionString
                       If Not IsBankStDtValid(adoConn) Then
                          If MsgBox("Bank Reconciliation for this date has been done. Do you want to select another date?", vbYesNo, "Multiple Bank Reconciliation Date.") = vbNo Then Exit Sub
                          Cancel = 1
                          FocusControl txtStatementDate
                          adoConn.Close
                          Set adoConn = Nothing
                          Exit Sub
                       End If
                       adoConn.Close
                       Set adoConn = Nothing
                    End If
              'End of validation
                    adoConn.Open getConnectionString
                    SavedPreBankRecTrans adoConn
                    adoConn.Close
                    Set adoConn = Nothing
         End If
         Exit For
      End If
   Next iRow
 If Cancel = False Then
   'Call UnlockAllRecords
    
 End If
FrmClose:
'   Call WheelUnHook(Me.hwnd)
   Call UnlockAllRecords
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
   UnLoadForm Me
End Sub
Private Sub UnlockAllRecords()
    Dim adoConn As New ADODB.Connection
    If haveYouLockedAnyReccord = True Then
        adoConn.Open getConnectionString
        'adoConn.Execute "Delete FROM Recordlocking where screen='Cashbook' and clientID='" & txtClientList.Tag & "' And BankCode='" & txtBC.Tag & "' AND WorkStation='" & UCase(WS_Name) & "' AND USER='" & UCase(User) & "'"
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbReceipt Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
        adoConn.Close
        Set adoConn = Nothing
        haveYouLockedAnyReccord = False
    End If
    
End Sub
'Private Sub LoadClients(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT   CLIENTID, CLIENTNAME " & _
'           "FROM     CLIENT " & _
'           "ORDER BY CLIENTID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'   Dim Data() As String
'   'Below conditional statement has been added by anol 04 May 2015--implementation of consoliated check mark
'   If isConsolidateExists(adoConn) Then
'        'when there si any consilidated value found  -----true
'        TotalRow = adoRst.RecordCount
'        TotalCol = adoRst.Fields.count - 1
'        ReDim Data(TotalCol, TotalRow) As String
'        'anol 22 Feb 2015 adding consolidated test
'         Data(0, 0) = "Con"
'         Data(1, 0) = "Consolidated"
'        For i = 1 To TotalRow
'            For j = 0 To TotalCol
'                Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'            Next j
'            adoRst.MoveNext
'            If adoRst.EOF Then Exit For
'        Next i
'   Else
'        'when there is no consolidated value ----false
'        TotalRow = adoRst.RecordCount - 1
'        TotalCol = adoRst.Fields.count - 1
'        ReDim Data(TotalCol, TotalRow) As String
'        For i = 0 To TotalRow
'            For j = 0 To TotalCol
'                Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'            Next j
'            adoRst.MoveNext
'            If adoRst.EOF Then Exit For
'        Next i
'   End If
'   cboClientID.Column() = Data()
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
Private Function isConsolidateExists(adoConn As ADODB.Connection) As Boolean
'This function shall check if there is any value inputed as consolidated
'Written by anol 04 May 2015

   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT   consolidated " & _
           "FROM     tlbclientbanks " & _
           "where consolidated=true;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        isConsolidateExists = True
        Exit Function
   End If
End Function
Public Sub ConfigFlxStatementReconcile()
   Dim i As Integer, szHeader As String

   With flxStatementReconcile
      szHeader$ = "+|<Date|<TranID|<Type|<Account|<Reference|>ReceiptValue|>PaymentValue" & _
                  "|>Statement|<Reconciliation|ID|Flag|Statement date"
      .FormatString = szHeader
      .Clear
      .Rows = 2
     '' addmore 5 to right
      '.Cols = 13
       .Cols = 21
      .RowHeight(0) = 0

      .ColWidth(0) = 250
      ' For i = 1 To .Cols - 5
      For i = 1 To 8
         .ColWidth(i) = lblBankRec(i + 13).Left - lblBankRec(i + 12).Left
      Next i
      .ColWidth(i) = .Width + .Left - lblBankRec(i + 12).Left - 300
      .ColWidth(btRecColNo) = 0                                         ' 10) Column number defined in the global var
      .ColWidth(11) = 0
      .ColWidth(12) = 0 ' add more 5 col to right
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      .ColWidth(15) = 0
      .ColWidth(16) = 0
      .ColWidth(17) = 0 'Client ID
      .ColWidth(18) = 0
      .ColWidth(19) = 0 'for future use
      .ColWidth(20) = 0 'Property ID
   End With
End Sub

Private Sub Label1_Click(Index As Integer)
   If Index = 35 Then                               ' Date
      SortingGrid flxCashBook, 1, bSortingCol3, "Date"
      bSortingCol3 = IIf(bSortingCol3, False, True)
      Label1(35).FontBold = True
      Label1(36).FontBold = False
      Label1(37).FontBold = False
      Label1(42).FontBold = False
      Label1(43).FontBold = False
   End If

   If Index = 37 Then                               ' Account
      SortingGrid flxCashBook, 3, bSortingCol1
      bSortingCol1 = IIf(bSortingCol1, False, True)
      Label1(35).FontBold = False
      Label1(36).FontBold = False
      Label1(37).FontBold = True
      Label1(42).FontBold = False
      Label1(43).FontBold = False
   End If

   If Index = 36 Then                               ' Type of Trans.
      SortingGrid flxCashBook, 2, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      Label1(35).FontBold = False
      Label1(36).FontBold = True
      Label1(37).FontBold = False
      Label1(42).FontBold = False
      Label1(43).FontBold = False
   End If

   If Index = 42 Then                               ' Yes/No
      SortingGrid flxCashBook, 8, bSortingCol4
      bSortingCol4 = IIf(bSortingCol4, False, True)
      Label1(35).FontBold = False
      Label1(36).FontBold = False
      Label1(37).FontBold = False
      Label1(42).FontBold = True
      Label1(43).FontBold = False
   End If

   If Index = 43 Then                               ' Statement Date
      SortingGrid flxCashBook, 9, bSortingCol5 ', "Date"
      bSortingCol5 = IIf(bSortingCol5, False, True)
      Label1(35).FontBold = False
      Label1(36).FontBold = False
      Label1(37).FontBold = False
      Label1(42).FontBold = False
      Label1(43).FontBold = True
   End If
End Sub

Private Sub lblBankRec_Click(Index As Integer)
'   If txtStValue.Visible Then txtStValue_LostFocus
'
   If Index = 13 Then                               ' Tran. ID
      SortingGrid1 flxStatementReconcile, 1, bBRCol1, "Date"
      'Modified by anol 14 Jun 2016
      'SortingGrid flxStatementReconcile, 1, bBRCol1, "Date"
      bBRCol1 = IIf(bBRCol1, False, True)
      lblBankRec(13).FontBold = True
      lblBankRec(15).FontBold = False
      lblBankRec(16).FontBold = False
   End If

   If Index = 15 Then                               ' Tran. ID
      SortingGrid flxStatementReconcile, 3, bBRCol3
      bBRCol3 = IIf(bBRCol3, False, True)
      lblBankRec(15).FontBold = True
      lblBankRec(13).FontBold = False
      lblBankRec(16).FontBold = False
   End If

   If Index = 16 Then                               ' Account No
      SortingGrid flxStatementReconcile, 4, bBRCol3
      bBRCol3 = IIf(bBRCol3, False, True)
      lblBankRec(15).FontBold = False
      lblBankRec(13).FontBold = False
      lblBankRec(16).FontBold = True
   End If
End Sub
Private Function BolCompareDataST(szData1 As String, szData2 As String, dtDataType As String) As Boolean
   BolCompareDataST = False
  If szData1 = "" Or szData2 = "" Then Exit Function
   Select Case dtDataType
      Case "Integer"
         If Val(szData1) < Val(szData2) Then
            BolCompareDataST = True
         End If

      Case "Currency"
         If Val(szData1) < Val(szData2) Then
            BolCompareDataST = True
         End If

      Case "Date"
         If DateDiff("d", CDate(szData1), CDate(szData2)) > 0 Then
            BolCompareDataST = True
         End If
   
      Case Else
         If szData1 < szData2 Then
            BolCompareDataST = True
         End If
   End Select
End Function

Private Function BolCompareDataGT(szData1 As String, szData2 As String, dtDataType As String) As Boolean
   BolCompareDataGT = False
    If szData1 = "" Or szData2 = "" Then Exit Function
   Select Case dtDataType
      Case "Integer"
         If Val(szData1) > Val(szData2) Then
            BolCompareDataGT = True
         End If

      Case "Date"
         If DateDiff("d", CDate(szData1), CDate(szData2)) < 0 Then
            BolCompareDataGT = True
         End If
   
      Case Else
         If szData1 > szData2 Then
            BolCompareDataGT = True
         End If
   End Select
End Function
Private Sub SortingGrid1(flxGrid As MSHFlexGrid, iSortCol As Integer, bAscDsc As Boolean, Optional dtDataType As String)
   Dim i As Integer, j As Integer, c As Integer
   Dim szTemp() As String
   ReDim szTemp(flxGrid.Cols - 1) As String
   
   For i = 1 To flxGrid.Rows - 2
      If flxGrid.RowHeight(i) > 0 Then
         For j = i + 1 To flxGrid.Rows - 1
            If flxGrid.RowHeight(j) > 0 Then
            
               If Not bAscDsc Then                 'Sorting ascending order
'                  If flxGrid.TextMatrix(i, iSortCol) > flxGrid.TextMatrix(j, iSortCol) Then
                  If BolCompareDataGT(flxGrid.TextMatrix(i, iSortCol), flxGrid.TextMatrix(j, iSortCol), dtDataType) Then
                     For c = 0 To flxGrid.Cols - 1
                        szTemp(c) = flxGrid.TextMatrix(i, c)
                        flxGrid.TextMatrix(i, c) = flxGrid.TextMatrix(j, c)
                        flxGrid.TextMatrix(j, c) = szTemp(c)
                     Next c
                  End If
               End If

               If bAscDsc Then                 'Sorting decending order
'                  If flxGrid.TextMatrix(i, iSortCol) < flxGrid.TextMatrix(j, iSortCol) Then
                  If BolCompareDataST(flxGrid.TextMatrix(i, iSortCol), flxGrid.TextMatrix(j, iSortCol), dtDataType) Then
                     For c = 0 To flxGrid.Cols - 1
                        szTemp(c) = flxGrid.TextMatrix(i, c)
                        flxGrid.TextMatrix(i, c) = flxGrid.TextMatrix(j, c)
                        flxGrid.TextMatrix(j, c) = szTemp(c)
                     Next c
                  End If
               End If
            
            End If
         Next j
      End If
   Next i
End Sub
Private Sub optCbHRpt_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then cmdCbHRptCancel_Click
End Sub

Public Function CalculatedUnClearedBal(dtStart As Date, dtEnd As Date, dtState As Date) As Double
   Dim i As Long, cUCB As Currency
   
   With flxStatementReconcile
      For i = 1 To .Rows - 1
        'For batch receipt added by anol 19 Feb 2015 issu 523
         If IsDate(.TextMatrix(i, 1)) Then
            If .TextMatrix(i, 1) >= dtStart And .TextMatrix(i, 1) <= dtEnd And .TextMatrix(i, 0) <> "-" Then
'                If UCase(Trim((.TextMatrix(i, 4))) = UCase("Batch Receipt")) Then
'                    MsgBox "Batch Receipt" & .TextMatrix(i, 1) & CCur(.TextMatrix(i, 6))
'                End If
               If .TextMatrix(i, 9) <> "Full" And .TextMatrix(i, 9) <> "Part" Then
                  If .TextMatrix(i, 6) <> "" Then cUCB = cUCB + CCur(.TextMatrix(i, 6))      'RV
                  If .TextMatrix(i, 7) <> "" Then cUCB = cUCB - CCur(.TextMatrix(i, 7))      'PV
               End If
               If .TextMatrix(i, 9) = "Part" And .TextMatrix(i, 12) <> "" Then
                  If .TextMatrix(i, 12) <= dtState Then
                     If .TextMatrix(i, 6) <> "" Then cUCB = cUCB + CCur(.TextMatrix(i, 6)) + CCur(.TextMatrix(i, 8))
                     If .TextMatrix(i, 7) <> "" Then cUCB = cUCB - CCur(.TextMatrix(i, 7)) + CCur(.TextMatrix(i, 8))
                  End If
               End If
            End If
         End If
      Next i
   
  
    End With
   CalculatedUnClearedBal = cUCB
End Function

Public Function CalculatedClosingBal(dtEnd As Date) As Double
   Dim i As Long, cDr As Currency

   On Error Resume Next

   With flxStatementReconcile
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 1) <= dtEnd And .TextMatrix(i, 0) <> "-" Then
            cDr = cDr + CCur(.TextMatrix(i, 8))                         'Recon value
         End If
      Next i
   End With

   CalculatedClosingBal = cDr
End Function

Private Sub optReconciliation_Click(Index As Integer)
   If txtBC.text = "" Then Exit Sub
   
   Dim adoConn As New ADODB.Connection
   MousePointer = vbHourglass
   fmeLoading.Top = 4305
   fmeLoading.Left = 4508
   fmeLoading.Visible = True
   fmeLoading.Refresh
   
'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
'The following code of hiding the grid row is unnecessary as we reloading the data based on the
'option to load either all or unreconciled transactions only.
'The function for calculating summary is called here is redundancy as its called from LoadFlxStatementReconcile
   

   ConfigFlxStatementReconcile
   If adoConn.State = 0 Then
      adoConn.Open getConnectionString
   End If
   Call LoadFlxStatementReconcile(adoConn)
   If adoConn.State = 1 Then
       adoConn.Close
   End If
'END OF MODIFICATION
    MousePointer = vbDefault
    fmeLoading.Visible = False
End Sub

Private Sub tabCashbook_Click(PreviousTab As Integer)
   If yWorkingTabPayRpt < 255 And bChangesMade And tabCashbook.Tab <> yWorkingTabCashBook Then
      MsgBox "Please save your current job.", vbInformation + vbOKOnly, "Prestige"
      tabCashbook.Tab = yWorkingTabCashBook
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If tabCashbook.Tab = 2 Then
      'txtCBHDtFrm.SetFocus
   End If

   If tabCashbook.Tab = 3 Then      'Bank Reconciliation
      If txtBC.text = "" Then
         MsgBox "Please select Bank code.", vbInformation + vbOKOnly, "Bank Reconciliation"
         'issue 523
         txtStOpenBal.text = ""
         FocusControl cmdBC
      Else
         FocusControl txtStatementDate
         txtStatementDate.SelLength = Len(txtStatementDate.text)
'         'lebel change tlbBankReconcilation
'         Dim rstlbBankReconcilation As New ADODB.Recordset
'         rstlbBankReconcilation.Open "select * from tlbBankReconcilation where clientID='" & txtClientList.Tag & "' AND AccountNum='" & txtBC.text & "'", adoConn, adOpenKeyset, adLockReadOnly
'         If rstlbBankReconcilation.EOF Then
'                Label1(23).Caption = "Reconciled Cashbook Balance:"
'                Label1(24).Caption = "Reconciled Statement Balance:"
'         Else
'                Label1(23).Caption = "Statement Opening Balance:"
'                Label1(24).Caption = "Projected Closing Balance:"
'         End If
'         rstlbBankReconcilation.Close
      End If
   End If

   If tabCashbook.Tab = 4 Then      'Memo and File attachment
      If txtBC.text <> "" Then _
         Call LoadAttachmentFiles(cmbFiles, cmdBC.Tag, "BankMemo")
   End If

   adoConn.Close
   Set adoConn = Nothing
   yWorkingTabCashBook = tabCashbook.Tab
End Sub

Private Sub LoadRptAmtType(szValue As String, adoConn As ADODB.Connection, conCombo As Control)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   conCombo.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   conCombo.Column() = szaData()
End Sub

Private Sub LoadPayAmtType(szValue As String, adoConn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   cmbSPAmtType.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   cmbSPAmtType.Column() = szaData()
End Sub

Private Sub tabCashbook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tabCashbook.MousePointer = vbDefault
End Sub

Private Sub tabPayRpt_Click(PreviousTab As Integer)
'MsgBox "New:" & tabPayRpt.Tab & " Old:" & PreviousTab
   If yWorkingTabPayRpt < 255 And tabPayRpt.Tab <> yWorkingTabPayRpt And bChangesMade Then
      MsgBox "Please save your current job.", vbInformation + vbOKOnly, "Prestige"
      tabPayRpt.Tab = PreviousTab
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection

   If tabPayRpt.Tab = 1 Then
      ConfigureFlxSPayment
      ConfigureFlxSCrPoA

      If cmbSPSupplier.ListCount = 0 Then
         'Set the RDO Connections to the dataset
         adoConn.Open getConnectionString
         'Load All supplier in the dropdown combo
         LoadAllSupplierFlxGrd adoConn
         'Load all Bank account details in the dropdown combo
         LoadBankAccountInCombo adoConn
         'Load all Payment type in the dropdown combo
         LoadPayAmtType "RECEIPT AMOUNT TYPE", adoConn

         adoConn.Close
         Set adoConn = Nothing
      End If
   End If

   If tabPayRpt.Tab = 3 Then
      txtBkTrDate.text = Format(Now, "dd/mm/yyyy")
      Exit Sub
   End If

   If cmdNewBk(0).Enabled And cmdEditBk(0).Enabled Then
      HandleTextBoxesBk True, False
   End If

   yWorkingTabPayRpt = tabPayRpt.Tab
End Sub

Private Sub LoadBankAccountInCombo(ByVal adoConn As ADODB.Connection)
   On Error GoTo Error_Handler

   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, j As Integer
   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   iTotalRow = adoRST.RecordCount
   iTotalCol = adoRST.Fields.Count
   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String
   
   For i = 0 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields.Item(j).Value), "", adoRST.Fields.Item(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cmbSPBankAc.Column() = Data()

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub
   
Error_Handler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRST = Nothing
End Sub

Private Sub LoadAllSupplierFlxGrd(ByVal adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iTotalRow As Integer, j As Integer
   Dim i As Integer, iTotalCol As Integer, Data() As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT SupplierID, SupplierName  " & _
           "FROM Supplier " & _
           "ORDER BY SupplierName;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   iTotalRow = adoRST.RecordCount
   iTotalCol = adoRST.Fields.Count

   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String

   For i = 0 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields.Item(j).Value), "", adoRST.Fields.Item(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cmbSPSupplier.Column() = Data()

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRST = Nothing
End Sub

Private Sub txtBkTrAmt_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtBkTrAmt, KeyAscii, 2
End Sub

Private Sub txtBkTrAmt_LostFocus()
   If txtBkTrAmt.text = "" Then Exit Sub

   txtBkTrAmt.text = Format(txtBkTrAmt.text, "0.00")
End Sub

Private Sub txtCBHDtFrm_Change()
   TextBoxChangeDate txtCBHDtFrm
End Sub

Private Sub txtCBHDtFrm_GotFocus()
   SelTxtInCtrl txtCBHDtFrm
End Sub

Private Sub txtCBHDtFrm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtCBHDtTo
    End If
   TextBoxKeyPrsDate txtCBHDtFrm, KeyAscii
End Sub

Private Sub txtCBHDtFrm_LostFocus()
   TextBoxFormatDate txtCBHDtFrm
End Sub

Private Sub txtCBHDtTo_Change()
   TextBoxChangeDate txtCBHDtTo
End Sub

Private Sub txtCBHDtTo_GotFocus()
   SelTxtInCtrl txtCBHDtTo
End Sub

Private Sub txtCBHDtTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdCBHFilter
    End If

   TextBoxKeyPrsDate txtCBHDtTo, KeyAscii
End Sub

Private Sub txtCBHDtTo_LostFocus()
   TextBoxFormatDate txtCBHDtTo
End Sub

Private Sub txtCrPayment_GotFocus()
   SelTxtInCtrl txtCrPayment
   If Not lblAllocating(1).Visible Then
      iCurRow = flxSCrPoA.row
      HighLightRowFlxGrid flxSCrPoA, iCurRow
   Else
      iCurRow = flxSPayment.row
   End If
End Sub

Private Sub txtCrPayment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then FocusControl txtAllocatedDiff(1)
End Sub

Private Sub txtCrPayment_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtCrPayment, KeyAscii
End Sub

Private Sub txtCrPayment_LostFocus()
   txtCrPayment.text = Format(IIf(txtCrPayment.text = "", 0, txtCrPayment.text), "0.00")

   If Not lblAllocating(1).Visible Then
      If Val(flxSCrPoA.TextMatrix(iCurRow, 8)) < Val(txtCrPayment.text) Then
         MsgBox "Payment amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
         txtCrPayment.text = "0.00"
         flxSCrPoA.RowSel = iCurRow
         flxSCrPoA.row = iCurRow
         FocusControl txtCrPayment
         Exit Sub
      End If

      If Val(txtCrPayment.text) > 0 Then
         flxSCrPoA.TextMatrix(iCurRow, 9) = txtCrPayment.text
         txtAllocatedDiff(1).text = txtCrPayment.text
         Label10(5).Caption = flxSCrPoA.TextMatrix(iCurRow, 0)
         flxSCrPoA.Enabled = False

         lblAllocating(1).Caption = "Allocating...                      " & flxSCrPoA.TextMatrix(iCrPoARowSel, 1)
         lblAllocating(1).Visible = True
         Frame5(5).Enabled = False                     'Payment - Saving
         Frame5(1).Enabled = True                      'Allocation - Saving
      End If
   Else
'     Allocating in the invoice grid, $ txtCrPayment $ text box is in the Upper grid
      If Val(txtAllocatedDiff(1).text) < Val(txtCrPayment.text) Then
         MsgBox "Allocated amount cannot exceed allocation difference amount.", vbExclamation + vbOKOnly, "Warning"
         txtCrPayment.text = "0.00"
         FocusControl txtCrPayment
         flxSPayment.row = Label10(7).Caption
         Exit Sub
      End If
      If Val(flxSPayment.TextMatrix(iCurRow, 9)) < Val(txtCrPayment.text) Then
         MsgBox "Allocated amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
         txtCrPayment.text = "0.00"
         FocusControl txtCrPayment
         flxSPayment.row = Label10(7).Caption
         Exit Sub
      End If

      If Val(txtCrPayment.text) > 0 Then
         flxSPayment.TextMatrix(iCurRow, 10) = txtCrPayment.text
         txtAllocatedDiff(1).text = Val(txtAllocatedDiff(1).text) - Val(txtCrPayment.text)

         If Val(txtAllocatedDiff(1).text) = 0 Then
            cmdPayAllocateSave.Enabled = True
            FocusControl cmdPayAllocateSave
         Else
            cmdPayAllocateSave.Enabled = False
         End If

         flxSPayment.TextMatrix(iCurRow, 15) = "A"
         flxSPayment.TextMatrix(iCurRow, 16) = Label10(5).Caption
      End If
      flxSPayment.Enabled = True
   End If

   txtCrPayment.Visible = False
End Sub

Private Sub txtCrReceipt_GotFocus()
   SelTxtInCtrl txtCrReceipt
   If Not lblAllocating(0).Visible Then
      iCurRow = flxTCrPoA.row
      HighLightRowFlxGrid flxTCrPoA, iCurRow
   Else
      iCurRow = flxTReceipt.row
   End If
End Sub

Private Sub txtCrReceipt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        FocusControl txtAllocatedDiff(0)
    End If
End Sub

Private Sub txtCrReceipt_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtCrReceipt, KeyAscii
End Sub

Private Sub txtCrReceipt_LostFocus()
   txtCrReceipt.text = Format(IIf(txtCrReceipt.text = "", 0, txtCrReceipt.text), "0.00")

   If Not lblAllocating(0).Visible Then
      If Val(flxTCrPoA.TextMatrix(iCurRow, 8)) < Val(txtCrReceipt.text) Then
         MsgBox "Payment amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
         txtCrReceipt.text = "0.00"
         flxTCrPoA.RowSel = iCurRow
         flxTCrPoA.row = iCurRow
         FocusControl txtCrReceipt
         Exit Sub
      End If

      If Val(txtCrReceipt.text) > 0 Then
         flxTCrPoA.TextMatrix(iCurRow, 9) = txtCrReceipt.text
         txtAllocatedDiff(0).text = txtCrReceipt.text
         Label10(2).Caption = flxTCrPoA.TextMatrix(iCurRow, 0)
         flxTCrPoA.Enabled = False

         lblAllocating(0).Caption = "Allocating...                      " & flxTCrPoA.TextMatrix(iCrPoARowSel, 1)
         lblAllocating(0).Visible = True
         Frame5(0).Enabled = False                     'Receipt - Saving
         Frame5(4).Enabled = True                      'Allocation - Saving
      End If
'      flxTCrPoA.Enabled = True
   Else
'     Allocating in the invoice grid, $ txtCrReceipt $ text box is in the Upper grid
      If Val(txtAllocatedDiff(0).text) < Val(txtCrReceipt.text) Then
         MsgBox "Allocated amount cannot exceed allocation difference amount.", vbExclamation + vbOKOnly, "Warning"
         txtCrReceipt.text = "0.00"
         FocusControl txtCrReceipt
         flxTReceipt.row = Label10(3).Caption
         Exit Sub
      End If
      If Val(flxTReceipt.TextMatrix(iCurRow, 9)) < Val(txtCrReceipt.text) Then
         MsgBox "Allocated amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
         txtCrReceipt.text = "0.00"
         FocusControl txtCrReceipt
         flxTReceipt.row = Label10(3).Caption
         Exit Sub
      End If

      If Val(txtCrReceipt.text) > 0 Then
         flxTReceipt.TextMatrix(iCurRow, 10) = txtCrReceipt.text
         txtAllocatedDiff(0).text = Val(txtAllocatedDiff(0).text) - Val(txtCrReceipt.text)

         If Val(txtAllocatedDiff(0).text) = 0 Then
            cmdRptAllocateSave.Enabled = True
            FocusControl cmdRptAllocateSave
         Else
            cmdRptAllocateSave.Enabled = False
         End If

         flxTReceipt.TextMatrix(iCurRow, 15) = "A"
         flxTReceipt.TextMatrix(iCurRow, 16) = Label10(2).Caption
      End If
      flxTReceipt.Enabled = True
   End If

   txtCrReceipt.Visible = False
End Sub

Private Sub txtNetBk_Change(Index As Integer)
    If txtNetBk(0).text = "" Or txtVatBk(0).text = "" Then Exit Sub
    Dim tot As Integer
    tot = CInt(txtNetBk(0).text) + CInt(txtVatBk(0).text)
    txtTotalBk(0).text = Format(tot, "0.00")
End Sub

Private Sub txtNetBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 10 Then txtNetBk_LostFocus (0)

   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtNetBk_LostFocus(Index As Integer)
   txtVatBk(0).text = Format(IIf(txtNetBk(0).text = "", 0, Val(txtNetBk(0).text)) * (nTaxCode / 100), "0.00")
   txtNetBk(0).text = Format(txtNetBk(0).text, "0.00")
End Sub

Private Sub txtProjClosingBal_GotFocus()
   SelTxtInCtrl txtProjClosingBal
End Sub

Private Sub txtProjClosingBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl flxStatementReconcile
        flxStatementReconcile.col = 8
   End If
   DigitTextKeyPress txtProjClosingBal, KeyAscii
End Sub

Private Sub txtProjClosingBal_LostFocus()
   Label1(27).Caption = Format(UnclearedBalance, "0.00")
   txtProjClosingBal.text = Format(txtProjClosingBal.text, "0.00")
End Sub

Private Sub txtSPayment_Click()
   SelTxtInCtrl txtSPayment
End Sub

Private Sub txtSPayment_GotFocus()
   If Not bTotalRptTyped Then _
      txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) - CCur(txtSPayment.text), "0.00")

   SelTxtInCtrl txtSPayment
   iCurRow = flxTReceipt.row
   cGridRptTotal = cGridRptTotal - CCur(txtSPayment.text)
End Sub

Private Sub txtSPayment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        FocusControl txtSPayment
    End If
End Sub

Private Sub txtSPayment_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtSPayment, KeyAscii
End Sub

Private Sub txtSPayment_LostFocus()
   txtSPayment.text = Format(Val(txtSPayment.text), "0.00")

   If Val(flxSPayment.TextMatrix(iCurRow, 9)) < Val(txtSPayment.text) Then
      MsgBox "Payment amount exceeds amount outstanding.", vbExclamation + vbOKOnly, "Warning"
      txtSPayment.text = "0.00"
      FocusControl txtSPayment
      txtSPayment_GotFocus
      Exit Sub
   End If

   flxSPayment.TextMatrix(iCurRow, 10) = txtSPayment.text

   If bTotalPayTyped Then
      If Val(txtSPaymentTotal.text) < cGridSPTotal + CCur(txtSPayment.text) Then
         MsgBox "Payment amount entered would exceed total payment amount.", vbExclamation + vbOKOnly, "Supplier Payment"
         txtSPayment.text = "0.00"
         FocusControl txtSPayment
         SelTxtInCtrl txtSPayment
         Exit Sub
      End If
   Else
      txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + Val(txtSPayment.text), "0.00")
      If Val(txtSPayment.text) > 0 Then flxSCrPoA.Enabled = False
   End If
   cGridSPTotal = cGridSPTotal + CCur(txtSPayment.text)

   baChangesMade(iCurRow) = IIf(Val(flxSPayment.TextMatrix(iCurRow, 10)) > 0, True, False)
   txtSPayment.text = "0.00"

   txtSPayment.Visible = False
   flxSPayment.Enabled = True

   txtPaymentEntered.text = Format(TotalPaymentEntered, "0.00")
End Sub

Private Sub txtSPaymentTotal_LostFocus()
   If Trim(txtSPaymentTotal.text) = "" Then txtSPaymentTotal.text = "0.00"
   If (cGridSPTotal - CCur(txtSPaymentTotal.text) <> 0) And (cGridSPTotal + CCur(txtSPaymentTotal.text) <> 0) Then
      If CCur(txtSPaymentTotal.text) < cGridSPTotal Then
         MsgBox "Total payment amount can not be changed to less than the analysis total.", vbCritical + vbOKOnly, "Wrong amount entry"
         txtSPaymentTotal.text = CStr(Format(cGridSPTotal, "0.00"))
         Exit Sub
      End If
   End If
   txtSPaymentTotal.text = Format(txtSPaymentTotal.text, "0.00")
   If cTempPaymentAmt - CCur(txtSPaymentTotal.text) <> 0 Then flxSCrPoA.Enabled = False
End Sub

Private Sub txtStOpenBal_GotFocus()
    'added by anol 05052016
'   If Left(Label1(23).Caption, 10) = "Reconciled" Then
'        txtStOpenBal.text = txtAcBal
'   End If
   SelTxtInCtrl txtStOpenBal
End Sub

Private Sub txtStOpenBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtProjClosingBal.Enabled = True Then
        FocusControl txtProjClosingBal
    End If
End Sub

Private Sub txtStOpenBal_LostFocus()
   'If Not txtStOpenBal.Locked And txtStOpenBal.text <> "" Then 'modified by anol 25/04/2016
   If Not txtStOpenBal.Locked Then
      If Val(txtStOpenBal.text) <> Val(txtAcBal.text) Then
         'MsgBox "The statement opening balance must equal the reconciled Bank account balance", vbCritical + vbOKOnly, "Bank Reconciliation"'modified by anol 25/04/2016
         'Modified by anol 07 June 2016
         ' If Left(Label1(23).Caption, 10) = "Reconciled" Then
               ' MsgBox "The reconciled Cashbook Balance entered must equal the Cashbook A/C Balance shown in the system.", vbCritical + vbOKOnly, "Bank Reconciliation"
               If MsgBox("The opening balance amount entered does not equal the cashbook a/c balance shown in the system." & vbCrLf & " Do you wish to reconcile the opening balance amount entered?", vbCritical + vbYesNo, "Bank Reconciliation") = vbYes Then
'         Else
'                MsgBox "The statement opening balance must equal the reconciled cashbook balance.", vbCritical + vbOKOnly, "Bank Reconciliation"
'         End If
                   FocusControl txtProjClosingBal
                    
               Else
                    FocusControl txtStOpenBal
               End If
        ' txtStOpenBal.text = txtAcBal.text
        ' txtStOpenBal.text = ""'modified by anol 25/04/2016
      Else
         Dim strmsg As String
         If Left(Label1(23).Caption, 10) = "Reconciled" Then
            strmsg = "Please confirm that the reconciled cashbook balance entered equals" & Chr(13) & _
                   "the Cashbook A/C Balance shown in the system."
         Else
            strmsg = "Please confirm that the statement opening balance entered equals" & Chr(13) & _
                   "the reconciled Cashbook A/C Balance shown in the system."
         End If
         If MsgBox(strmsg, vbQuestion + vbYesNo, "Bank Reconciliation") = vbYes Then
            txtStOpenBal.Locked = True
         Else
            'txtStOpenBal.SetFocus
            FocusControl txtStatementDate
         End If
      End If
   End If
End Sub

Private Sub txtTReceipt_Click()
   SelTxtInCtrl txtTReceipt
End Sub

Private Sub txtTReceipt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        FocusControl flxTReceipt
   End If
End Sub

Private Sub txtTReceipt_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtTReceipt, KeyAscii
End Sub

Private Sub txtSPaymentTotal_Change()
   txtPaymentTotal.text = txtSPaymentTotal.text
   If Val(txtSPaymentTotal.text) = 0 Then
      bChangesMade = False     'there some no change has made in the form
      cmdSPSave.Enabled = False
      Exit Sub
   Else
      bChangesMade = True     'there some changes have made in the form
      cmdSPSave.Enabled = True
   End If
   cmbSPBankAc.Enabled = False
End Sub

Private Sub txtSPaymentTotal_GotFocus()
   cTempPaymentAmt = CCur(txtSPaymentTotal.text)
   SelTxtInCtrl txtSPaymentTotal
End Sub

Private Sub txtSPaymentTotal_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtSPaymentTotal, KeyAscii
End Sub

Private Sub txtSPaymentTotal_KeyUp(KeyCode As Integer, Shift As Integer)
   bTotalPayTyped = IIf(Val(txtSPaymentTotal.text) > 0, True, False)
End Sub

Private Sub txtTReceiptTotal_Change()
   txtReceiptTotal.text = txtTReceiptTotal.text
   If Val(txtTReceiptTotal.text) = 0 Then
      bChangesMade = False     'there some no change has made in the form
      cmdTRSave.Enabled = False
      Exit Sub
   Else
      bChangesMade = True     'there some changes have made in the form
      cmdTRSave.Enabled = True
   End If
   cmdBC.Enabled = False
End Sub

Private Sub txtTReceiptTotal_GotFocus()
   cTempReceiptAmt = CCur(txtTReceiptTotal.text)
   SelTxtInCtrl txtTReceiptTotal
End Sub

Private Sub txtTReceiptTotal_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtTReceiptTotal, KeyAscii
End Sub

Private Sub txtTReceiptTotal_KeyUp(KeyCode As Integer, Shift As Integer)
   bTotalRptTyped = IIf(Val(txtTReceiptTotal.text) > 0, True, False)
End Sub

Private Sub txtTReceiptTotal_LostFocus()
   If Trim(txtTReceiptTotal.text) = "" Then txtTReceiptTotal.text = "0.00"
   If (cGridRptTotal - CCur(txtTReceiptTotal.text) <> 0) And (cGridRptTotal + CCur(txtTReceiptTotal.text) <> 0) Then
      If CCur(txtTReceiptTotal.text) < cGridRptTotal Then
         MsgBox "Total receipt amount can not be changed to less than the analysis total.", vbCritical + vbOKOnly, "Wrong amount entry"
         txtTReceiptTotal.text = CStr(Format(cGridRptTotal, "0.00"))
         Exit Sub
      End If
   End If
   txtTReceiptTotal.text = Format(txtTReceiptTotal.text, "0.00")
   If cTempReceiptAmt - CCur(txtTReceiptTotal.text) <> 0 Then flxTCrPoA.Enabled = False
End Sub

Private Sub txtSPDate_Change()
   TextBoxChangeDate txtSPDate
End Sub

Private Sub txtSPDate_GotFocus()
   If txtSPDate.text = "dd/mm/yyyy" Then
      txtSPDate.text = ""
      Exit Sub
   End If
   If Len(txtSPDate.text) < 10 Then txtSPDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSPDate
End Sub

Private Sub txtSPDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSPDate, KeyAscii
End Sub

Private Sub txtSPDate_LostFocus()
   If txtSPDate.text <> "" Then TextBoxFormatDate txtSPDate
End Sub

Private Sub txtTRDate_Change()
   TextBoxChangeDate txtTRDate
End Sub

Private Sub txtTRDate_GotFocus()
   SelTxtInCtrl txtTRDate
End Sub

Private Sub txtTRDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtTRDate, KeyAscii
End Sub

Private Sub txtTRDate_LostFocus()
   TextBoxFormatDate txtTRDate
End Sub

Private Sub txtStatementDate_Change()
   'Added By Samrat. 16/01/2006
   TextBoxChangeDate txtStatementDate
End Sub

Private Sub txtStatementDate_GotFocus()
   SelTxtInCtrl txtStatementDate
End Sub

Private Sub txtStatementDate_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 16/01/2006
   If KeyAscii = 13 And txtStOpenBal.Enabled = True Then
        FocusControl txtStOpenBal
   End If
   TextBoxKeyPrsDate txtStatementDate, KeyAscii
End Sub

Private Sub txtStatementDate_LostFocus()
   TextBoxFormatDate txtStatementDate
   'adde by anol 02 08 2016
   Dim adoConn As New ADODB.Connection
   
        If txtStatementDate.text <> "" Then
               adoConn.Open getConnectionString
               Dim adoRST  As New ADODB.Recordset
               Dim szSQL   As String
               Dim r       As Integer
            'issue 523
            'Modified by anol 20 Jan 2015
               If txtClientList.text = "Consolidated" Then
                      szSQL = "SELECT ReconDate, BankCode " & _
                       "FROM tlbBankReconcilation " & _
                       "WHERE BankCode = '" & txtAccountName.text & "' AND clientID='" & txtClientList.Tag & "'AND " & _
                           "ReconDate >= #" & Format(txtStatementDate.text, "dd mmmm yyyy") & "# " & _
                       "Order BY ReconDate Desc;"
               Else
                    szSQL = "SELECT ReconDate, BankCode " & _
                       "FROM tlbBankReconcilation " & _
                       "WHERE BankCode = '" & txtAccountName.text & "' AND clientID='" & txtClientList.Tag & "'AND " & _
                           "ReconDate >= #" & Format(txtStatementDate.text, "dd mmmm yyyy") & "# " & _
                       "Order BY ReconDate Desc;"
               End If
            'Debug.Print szSQL
               adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            
               If Not adoRST.EOF Then
                    'MsgBox "Please enter a date after the last bank reconciliation date of :" & adoRst.Fields("ReconDate").Value & "  ", vbOKOnly, "Warning!!"
                    txtStatementDate.text = Format(adoRST.Fields("ReconDate").Value, "dd/MM/yyyy")
                    'txtStatementDate.SetFocus
               End If
            
               adoRST.Close
               Set adoRST = Nothing
'           If Not IsBankStDtValid(adoConn) Then
'              MsgBox "Bank Reconciliation for this date has been done. Please select another date.", vbCritical + vbOKOnly, "Multiple Bank Reconciliation Date."
'              txtStatementDate.SetFocus
'              adoConn.Close
'              Set adoConn = Nothing
'              Exit Sub
'           End If
           adoConn.Close
        Set adoConn = Nothing
        End If
        
        
        
'issue 523
'Modified by anol 21 Jan 2015
'   If Not txtStOpenBal.Locked Then
'      txtStOpenBal.text = ""
'      txtStOpenBal.SetFocus
'   End If
End Sub
'''''
'''''Private Function CheckDuplicateBankStDate() As Boolean
'''''   Dim i As Integer
'''''
'''''   CheckDuplicateBankStDate = True
'''''
'''''   For i = 0 To lstBankStDates.ListCount - 1
'''''      If lstBankStDates.List(i) = Trim(txtStatementDate.text) & "#" & txtAccountName.text Then
'''''         CheckDuplicateBankStDate = False
'''''         Exit For
'''''      End If
'''''   Next i
'''''End Function

'Private Sub txtStValue_KeyPress(KeyAscii As Integer)
'   Dim i As Integer
'
'   If KeyAscii = 45 Then KeyAscii = 0
'
'   With flxStatementReconcile
'      If KeyAscii = 13 Then
'         For i = .row + 1 To .Rows - 1
'            If .RowHeight(i) > 0 Then Exit For
'         Next i
'
'         If i <= .Rows - 1 Then
'            .row = i
'            .SetFocus
'         Else
'            cmdReconSave.SetFocus
'         End If
'      End If
'   End With
'End Sub

Private Sub LoadSupplier()
   Dim iRow As Integer
   iRow = 1

   flxLeaseList.Cols = 3
   flxLeaseList.ColWidth(0) = 1000
   flxLeaseList.ColWidth(1) = 2700
   flxLeaseList.ColWidth(2) = 700

   Label20(0).Width = 700
   Label20(1).Left = 50
   Label20(2).Width = 2600
   Label20(1).Left = Label20(0).Left + flxLeaseList.ColWidth(0)
   Label20(2).Width = 400
   Label20(2).Left = Label20(1).Left + flxLeaseList.ColWidth(1)
   
   txtTenantSearchID.Width = 800
   txtTenantSearchID.Left = 50
   
   txtTenantSearchName.Width = 2600
   txtTenantSearchName.Left = txtTenantSearchID.Left + flxLeaseList.ColWidth(0)
   
   txtTenantSearchUnitName.Width = 600
   txtTenantSearchUnitName.Left = txtTenantSearchName.Left + flxLeaseList.ColWidth(1)
   
   Dim rdoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL1, szSQL2, szSQL3 As String
      
   ' Error Handler
   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   rdoConn.Open getConnectionString

    szSQL1 = "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientID;"
           
    szSQL2 = "SELECT AgentID, AgentName " & _
           "FROM Agent " & _
           "ORDER BY AgentID;"

    szSQL3 = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier " & _
           "ORDER BY SupplierID;"
     
   rstRst.Open szSQL1, rdoConn, adOpenStatic, adLockReadOnly

 If rstRst.EOF Then GoTo NoRes
   
    flxLeaseList.Clear
    flxLeaseList.Rows = 2
    flxLeaseList.RowHeight(0) = 0
    
    '~~~Added By Senthuran~~~ Code to configuer Label Caption
      Label20(0).Caption = "Cod"
      Label20(1).Caption = "Name"
      Label20(2).Caption = "Type"

   While Not rstRst.EOF
      flxLeaseList.TextMatrix(iRow, 0) = rstRst!ClientID
      flxLeaseList.TextMatrix(iRow, 1) = rstRst!ClientName
      flxLeaseList.TextMatrix(iRow, 2) = "Client"
      rstRst.MoveNext
      flxLeaseList.AddItem ""
      iRow = iRow + 1
   Wend

    Set rstRst = Nothing
    rstRst.Open szSQL2, rdoConn, adOpenStatic, adLockReadOnly

    While Not rstRst.EOF
       flxLeaseList.TextMatrix(iRow, 0) = rstRst!AgentID
       flxLeaseList.TextMatrix(iRow, 1) = rstRst!AgentName
       flxLeaseList.TextMatrix(iRow, 2) = "Agent"
       rstRst.MoveNext
       flxLeaseList.AddItem ""
       iRow = iRow + 1
    Wend
    Set rstRst = Nothing
    rstRst.Open szSQL3, rdoConn, adOpenStatic, adLockReadOnly

    While Not rstRst.EOF
       flxLeaseList.TextMatrix(iRow, 0) = rstRst!SupplierID
       flxLeaseList.TextMatrix(iRow, 1) = rstRst!SupplierName
       flxLeaseList.TextMatrix(iRow, 2) = "Supplier"
       rstRst.MoveNext
       If Not rstRst.EOF Then flxLeaseList.AddItem ""
       iRow = iRow + 1
    Wend

NoRes:
   rstRst.Close
   rdoConn.Close
   Set rstRst = Nothing
   Set rdoConn = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
   
   rstRst.Close
   rdoConn.Close
   Set rstRst = Nothing
   Set rdoConn = Nothing
End Sub

Public Sub CopyBankTran()
   Dim szStr      As String
   Dim adoConn    As New ADODB.Connection
   Dim adoRstSrc  As New ADODB.Recordset
   Dim adoRstDsc  As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoConn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM  tlbBankPayment " & _
           "WHERE TRANS = '" & Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) & "' AND " & _
                 "TRAN_ID = '" & StrDigitVal(flxCashBook.TextMatrix(flxCashBook.row, 2)) & "';"
'Debug.Print szStr
   adoRstSrc.Open szStr, adoConn, adOpenStatic, adLockReadOnly
   
   szStr = "SELECT * " & _
           "FROM tlbBankPayment;"
'Debug.Print szStr
   adoRstDsc.Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   With adoRstDsc
'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = adoRstSrc.Fields.Item("ClientID").Value
      .Fields.Item("BANK_AC").Value = adoRstSrc.Fields.Item("BANK_AC").Value
      .Fields.Item("PropertyID").Value = adoRstSrc.Fields.Item("PropertyID").Value
      .Fields.Item("UNIT_ID").Value = adoRstSrc.Fields.Item("UNIT_ID").Value
      .Fields.Item("DESCRIPTION").Value = adoRstSrc.Fields.Item("DESCRIPTION").Value
      .Fields.Item("PROJ_REF").Value = adoRstSrc.Fields.Item("PROJ_REF").Value
      .Fields.Item("NOMINAL_CODE").Value = adoRstSrc.Fields.Item("NOMINAL_CODE").Value
      .Fields.Item("DEPT_ID").Value = adoRstSrc.Fields.Item("DEPT_ID").Value
      .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
      .Fields.Item("TAX_CODE").Value = adoRstSrc.Fields.Item("TAX_CODE").Value
      .Fields.Item("VAT").Value = adoRstSrc.Fields.Item("VAT").Value
      .Fields.Item("NET_AMOUNT").Value = adoRstSrc.Fields.Item("NET_AMOUNT").Value

      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BR" Then
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
      End If
      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BP" Then
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
      End If

      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Update
      .Close
   End With

   ShowMsgInTaskBar "The Transaction has been copied sucessfully.", "Y", "P"

   Set adoRstDsc = Nothing

   flxCashBook.Clear
   flxCashBook.Rows = 2

   LoadFlxCashBook adoConn
   CalDrCrCBHistory

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrHandler:
   MsgBox Err.description & ": System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRstDsc = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub CopyRevTransaction()
   Dim szStr      As String
   Dim adoConn    As New ADODB.Connection
   Dim adoRstSrc  As New ADODB.Recordset
   Dim adoRstDsc  As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoConn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM  tlbBankPayment " & _
           "WHERE TRANS = '" & Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) & "' AND " & _
                 "TRAN_ID = '" & StrDigitVal(flxCashBook.TextMatrix(flxCashBook.row, 2)) & "';"
'Debug.Print szStr
   adoRstSrc.Open szStr, adoConn, adOpenStatic, adLockReadOnly
   
   szStr = "SELECT * " & _
           "FROM tlbBankPayment;"
'Debug.Print szStr
   adoRstDsc.Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   With adoRstDsc
'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = adoRstSrc.Fields.Item("ClientID").Value
      .Fields.Item("BANK_AC").Value = adoRstSrc.Fields.Item("BANK_AC").Value
      .Fields.Item("PropertyID").Value = adoRstSrc.Fields.Item("PropertyID").Value
      .Fields.Item("UNIT_ID").Value = adoRstSrc.Fields.Item("UNIT_ID").Value
      .Fields.Item("DESCRIPTION").Value = adoRstSrc.Fields.Item("DESCRIPTION").Value
      .Fields.Item("PROJ_REF").Value = adoRstSrc.Fields.Item("PROJ_REF").Value
      .Fields.Item("NOMINAL_CODE").Value = adoRstSrc.Fields.Item("NOMINAL_CODE").Value
      .Fields.Item("DEPT_ID").Value = adoRstSrc.Fields.Item("DEPT_ID").Value
      .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
      .Fields.Item("TAX_CODE").Value = adoRstSrc.Fields.Item("TAX_CODE").Value
      .Fields.Item("VAT").Value = adoRstSrc.Fields.Item("VAT").Value
      .Fields.Item("NET_AMOUNT").Value = adoRstSrc.Fields.Item("NET_AMOUNT").Value

      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BR" Then
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
      End If
      If Left(flxCashBook.TextMatrix(flxCashBook.row, 2), 2) = "BP" Then
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
      End If

      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Update
      .Close
   End With

   ShowMsgInTaskBar "The Transaction has been copy reversed sucessfully.", "Y", "P"

   Set adoRstDsc = Nothing

   flxCashBook.Clear
   flxCashBook.Rows = 2

   LoadFlxCashBook adoConn
   CalDrCrCBHistory

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrHandler:
   MsgBox Err.description & ": System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRstDsc = Nothing
   adoConn.Close
   Set adoConn = Nothing
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


