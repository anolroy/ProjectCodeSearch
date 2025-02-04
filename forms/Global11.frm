VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGlobal1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Data"
   ClientHeight    =   9885
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Global11.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   12690
   Begin VB.Frame fraSelProp 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9960
      TabIndex        =   114
      Top             =   4080
      Width           =   3855
      Begin VB.CommandButton cmdOkSelProp 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   960
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00800000&
         Height          =   1335
         Index           =   2
         Left            =   75
         Top             =   75
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   2
         Height          =   1335
         Index           =   1
         Left            =   75
         Top             =   75
         Width           =   3735
      End
      Begin MSForms.ComboBox cboSelProp 
         Height          =   315
         Left            =   120
         TabIndex        =   117
         Top             =   480
         Width           =   3615
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "6376;556"
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
         Object.Width           =   "1411;4233"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Property:"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   115
         Top             =   120
         Width           =   1230
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
      Height          =   1515
      Left            =   3000
      TabIndex        =   57
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2672
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
   Begin VB.Frame Frame1 
      Caption         =   "Current Year Budget"
      Height          =   3180
      Left            =   120
      TabIndex        =   63
      Top             =   560
      Width           =   8955
      Begin VB.Frame fraInterestRates 
         BackColor       =   &H00FDEDED&
         Caption         =   "Lessee Finance Rates:"
         Height          =   2625
         Left            =   3120
         TabIndex        =   76
         Top             =   1560
         Visible         =   0   'False
         Width           =   4260
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInterestRates 
            Height          =   1755
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3096
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
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         Begin MSForms.CommandButton cmdDeleteInterest 
            Height          =   360
            Left            =   2285
            TabIndex        =   84
            Top             =   2160
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Delete"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdAddInterest 
            Height          =   360
            Left            =   120
            TabIndex        =   83
            Top             =   2160
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Add"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEditInterest 
            Height          =   360
            Left            =   1165
            TabIndex        =   82
            Top             =   2160
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Edit"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdCloseInterest 
            Height          =   360
            Left            =   3480
            TabIndex        =   81
            Top             =   2160
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Close"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.Label Label3 
            Height          =   270
            Index           =   10
            Left            =   2880
            TabIndex        =   80
            Top             =   200
            Width           =   1065
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Additional (%)"
            Size            =   "1879;476"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   79
            Top             =   200
            Width           =   1050
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Base Rate (%)"
            Size            =   "1852;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   78
            Top             =   200
            Width           =   975
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Date From"
            Size            =   "1720;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraDemandDaysB4Due 
         Caption         =   "Demand Notice Period"
         Height          =   825
         Left            =   120
         TabIndex        =   110
         Top             =   2280
         Width           =   4140
         Begin MSForms.TextBox txtDemandDaysB4Due 
            Height          =   285
            Left            =   1200
            TabIndex        =   112
            Top             =   315
            Width           =   495
            VariousPropertyBits=   142622747
            Size            =   "873;503"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Send Demands            Days Before Due Date"
            Height          =   210
            Index           =   5
            Left            =   105
            TabIndex        =   111
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame fraSCBdDtls 
         Caption         =   "Service Charge Budget Details:"
         Height          =   1095
         Left            =   4680
         TabIndex        =   106
         Top             =   240
         Width           =   4140
         Begin VB.CommandButton cmdYearlyService 
            Caption         =   "Service C&harge Budget"
            Height          =   345
            Left            =   1920
            TabIndex        =   109
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Service Charge Budget Total: "
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblSCTotal 
            Alignment       =   2  'Center
            Caption         =   "£0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   107
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame fraOthers 
         Caption         =   "Others"
         ForeColor       =   &H00004040&
         Height          =   1605
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   4140
         Begin VB.TextBox txtGlobalBankAccount 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   100
            Top             =   555
            Width           =   2055
         End
         Begin VB.CommandButton cmdExpandBankCode 
            Caption         =   "v"
            Height          =   285
            Left            =   3720
            TabIndex        =   99
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtFiYrEnd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   98
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtBIRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   97
            ToolTipText     =   "Latest base interest rate"
            Top             =   1200
            Width           =   1880
         End
         Begin VB.CommandButton cmdInterestRates 
            Caption         =   "- - -"
            Height          =   285
            Left            =   3600
            TabIndex        =   96
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Global Bank Account:"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   105
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Financial Year End:"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   104
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "VAT Rate:"
            Height          =   195
            Left            =   75
            TabIndex        =   103
            Top             =   885
            Width           =   915
         End
         Begin MSForms.ComboBox cboVatRate 
            Height          =   285
            Left            =   1680
            TabIndex        =   102
            Top             =   885
            Width           =   2295
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "4048;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Base Interest Rate:                         %"
            Height          =   195
            Left            =   75
            TabIndex        =   101
            Top             =   1200
            Width           =   1350
         End
      End
      Begin VB.Frame Frame8 
         Height          =   855
         Index           =   0
         Left            =   4680
         TabIndex        =   90
         Top             =   1380
         Width           =   4140
         Begin VB.CommandButton cmdYearlyInsurance 
            Caption         =   "&Insurance  Budget"
            Height          =   345
            Left            =   1920
            TabIndex        =   91
            Top             =   160
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance budget:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   94
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance budget total:"
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   93
            Top             =   560
            Width           =   1620
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "£0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   210
            Index           =   7
            Left            =   2700
            TabIndex        =   92
            Top             =   555
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Height          =   855
         Index           =   1
         Left            =   4680
         TabIndex        =   85
         Top             =   2280
         Width           =   4140
         Begin VB.CommandButton cmdYearlyRent 
            Caption         =   "&Rent  Budget"
            Height          =   345
            Left            =   1920
            TabIndex        =   86
            Top             =   160
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "£0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Index           =   8
            Left            =   2700
            TabIndex        =   89
            Top             =   555
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent budget total:"
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   9
            Left            =   105
            TabIndex        =   88
            Top             =   555
            Width           =   1260
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Rent budget:"
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   87
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.Frame fraSetInterestRates 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Set Finance Charge Rate"
         Height          =   2505
         Left            =   5520
         TabIndex        =   64
         Top             =   3600
         Visible         =   0   'False
         Width           =   4140
         Begin VB.TextBox txtDateFrom 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   68
            Top             =   360
            Width           =   1880
         End
         Begin VB.TextBox txtAdditionalRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   67
            Top             =   1240
            Width           =   1880
         End
         Begin VB.TextBox txtBaseRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   66
            Top             =   800
            Width           =   1880
         End
         Begin VB.TextBox txtRateDescription 
            Height          =   285
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   65
            Top             =   1680
            Width           =   1880
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   75
            Top             =   1240
            Width           =   1575
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Additional Rate (%)"
            Size            =   "2778;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   74
            Top             =   800
            Width           =   1050
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Base Rate (%)"
            Size            =   "1852;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   225
            Index           =   17
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   1605
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Date Applying From"
            Size            =   "2831;397"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmdSetIntRateSave 
            Height          =   360
            Left            =   2040
            TabIndex        =   72
            Top             =   2040
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Save"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdSetIntRateClose 
            Height          =   360
            Left            =   3245
            TabIndex        =   71
            Top             =   2040
            Width           =   675
            ForeColor       =   16384
            Caption         =   "Close"
            PicturePosition =   65543
            Size            =   "1191;635"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.Label Label3 
            Height          =   150
            Index           =   18
            Left            =   240
            TabIndex        =   70
            Top             =   2160
            Visible         =   0   'False
            Width           =   1320
            ForeColor       =   16777215
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "ADDING/EDITING"
            Size            =   "2328;265"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   150
            Index           =   19
            Left            =   240
            TabIndex        =   69
            Top             =   1680
            Width           =   1695
            ForeColor       =   16384
            BackColor       =   -2147483637
            VariousPropertyBits=   276824083
            Caption         =   "Additional Info"
            Size            =   "2990;265"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   455
      Left            =   7740
      TabIndex        =   61
      Top             =   6915
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2955
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   5212
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Monthly Payment Dates"
      TabPicture(0)   =   "Global11.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAutoSetup(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Quarterly Payment Dates"
      TabPicture(1)   =   "Global11.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Half Yearly payments"
      TabPicture(2)   =   "Global11.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Yearly payments"
      TabPicture(3)   =   "Global11.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Additional Payment Dates"
      TabPicture(4)   =   "Global11.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdMthPayDt"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdAutoSetup 
         BackColor       =   &H80000013&
         Caption         =   "Auto Date Fill"
         Enabled         =   0   'False
         Height          =   325
         Index           =   0
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   62
         Top             =   2560
         Width           =   1455
      End
      Begin VB.CommandButton cmdMthPayDt 
         BackColor       =   &H80000016&
         Caption         =   "Click here to enter additional Payment Sets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -72960
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Yearly Payment Date"
         Height          =   2295
         Left            =   -69480
         TabIndex        =   53
         Top             =   480
         Width           =   2535
         Begin VB.ComboBox cboM7 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD7 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Half Yearly Payment Dates"
         Enabled         =   0   'False
         Height          =   2295
         Left            =   -71160
         TabIndex        =   46
         Top             =   420
         Width           =   3135
         Begin VB.ComboBox cboD5 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM5 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD6 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM6 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "1st"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   2175
         Left            =   6120
         TabIndex        =   40
         Top             =   360
         Width           =   2655
         Begin VB.ComboBox cboDay9 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboDay10 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay11 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay12 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "September:"
            Height          =   195
            Left            =   360
            TabIndex        =   44
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "October:"
            Height          =   195
            Left            =   360
            TabIndex        =   43
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "November:"
            Height          =   195
            Left            =   360
            TabIndex        =   42
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "December:"
            Height          =   195
            Left            =   360
            TabIndex        =   41
            Top             =   1740
            Width           =   765
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   2175
         Left            =   3150
         TabIndex        =   36
         Top             =   360
         Width           =   2655
         Begin VB.ComboBox cboDay5 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboDay6 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay7 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay8 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "July:"
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   1260
            Width           =   300
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "May:"
            Height          =   195
            Left            =   360
            TabIndex        =   39
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "June:"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   780
            Width           =   360
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "August:"
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   1740
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Quarterly Payment Dates"
         Enabled         =   0   'False
         Height          =   2295
         Left            =   -73080
         TabIndex        =   23
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboD1 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM1 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD2 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM2 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboD3 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboM3 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboD4 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboM4 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "1st"
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "3rd"
            Height          =   195
            Left            =   360
            TabIndex        =   33
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "4th"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   1800
            Width           =   240
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   2175
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   2655
         Begin VB.ComboBox cboDay4 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox cboDay3 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay2 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay1 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "April:"
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   1740
            Width           =   375
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "March:"
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   1260
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "February:"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "January:"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdDemandTypes 
      Caption         =   "&Demand/Charge Types"
      Height          =   455
      Left            =   5835
      TabIndex        =   16
      Top             =   6915
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      Height          =   455
      Left            =   3930
      TabIndex        =   14
      Top             =   6915
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   455
      Left            =   2025
      TabIndex        =   56
      Top             =   6915
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Data"
      Height          =   455
      Left            =   120
      TabIndex        =   15
      Top             =   6915
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cboProperty 
      Bindings        =   "Global11.frx":04CE
      DataSource      =   "adoProperty"
      Height          =   330
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "PropertyName"
      BoundColumn     =   "PropertyID"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoProperty 
      Height          =   330
      Left            =   3360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Property"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSCDtls 
      Height          =   330
      Left            =   5880
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "SC"
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
   Begin MSAdodcLib.Adodc adoNC 
      Height          =   330
      Left            =   8640
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "NC"
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
   Begin MSForms.ComboBox cboClientList 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6800;503"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   8
      ListRows        =   20
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1587"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   113
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Default Payment Set:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   210
      Index           =   10
      Left            =   120
      TabIndex        =   59
      Top             =   3840
      Width           =   1650
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Property:"
      Height          =   195
      Index           =   4
      Left            =   4680
      TabIndex        =   58
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmGlobal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New ADODB.Connection
Dim MyForm As FRMSIZE

Dim Rst As New ADODB.Recordset
Dim SQLStr As String
Dim SCperSqFoot As Double

Dim bEditGlobalData As Boolean, bFlxInsEdit As Boolean, bNoGD As Boolean
Dim bIsChildBudget As Boolean, iCurFlxInsRow As Integer
Dim iNewEditMainBudget As Integer, iSCChildNewEdit As Integer
Dim iNewEditRCBudget As Integer, iCurFlxRCMainRow As Integer, iCurFlxRCMainCol As Integer
Dim iCurFlxMainRow As Integer, iCurFlxMainCol As Integer
Dim iCurFlxChildRow As Integer, iCurFlxChildCol As Integer
Dim iCurFlxInterestRatesRow As Integer
Dim cInsurance As Currency

Private Sub cboClientList_Click()
   Call LoadProperty
End Sub

Private Sub cboProperty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frmMMain.fraCmdButton.Enabled = True
      Unload Me
   End If
End Sub

Private Sub cboVatRate_LostFocus()
   Dim i, match As Integer
   match = 0
   If cboVatRate.text = "" Then Exit Sub
   For i = 0 To 12
       If cboVatRate.text = cboVatRate.List(i) Then
           match = 1
           Exit For
       End If
   Next i
   If match = 0 Then cboVatRate.text = ""
End Sub

Private Sub cboProperty_Change()
   bNoGD = False

   If Not GetData Then
      If (MsgBox("There is no Global Data setup for the property " & cboProperty.text & ". Would you like to set this up?", vbQuestion + vbYesNo, "No Global Data") = vbYes) Then
         Edit
         cmdCancel.Enabled = False

         bNoGD = IIf(DemandTypeExist, False, True)
      End If
   End If
End Sub

Private Function DemandTypeExist() As Boolean
   SQLStr = "SELECT * FROM DemandTypes WHERE PropertyID = '" & cboProperty.BoundText & "';"

   Conn.Open getConnectionString

   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   DemandTypeExist = IIf(Rst.EOF, False, True)

   Rst.Close
   Conn.Close
End Function

Private Sub cmdAddInterest_Click()
   fraSetInterestRates.Left = fraInterestRates.Left
   fraSetInterestRates.Top = fraInterestRates.Top
   fraSetInterestRates.Visible = True
   Label3(18).Caption = "ADDING"
   txtDateFrom.text = ""
   txtBaseRate.text = ""
   txtAdditionalRate.text = ""
   txtDateFrom.SetFocus

   cmdAddInterest.Enabled = False
End Sub

Private Sub cmdAutoSetup_Click(Index As Integer)
   Dim dtDate As Date, var

   On Error GoTo ErrorHandler

   var = InputBox("Please type the first payment date of the year. (dd/mm/yyyy)", "Frist Payment Date", "01/01/" & Year(Date))
   If var = "" Then Exit Sub

   dtDate = Format(var, "dd mmmm yyyy")

   SetAddDates dtDate

   Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
      cmdAutoSetup_Click (0)
   End If
End Sub

Private Sub SetAddDates(dtDate As Date)
   cboDay1.text = Format(dtDate, "dd")
   cboDay2.text = Format(dtDate, "dd")
   cboDay3.text = Format(dtDate, "dd")
   cboDay4.text = Format(dtDate, "dd")
   cboDay5.text = Format(dtDate, "dd")
   cboDay6.text = Format(dtDate, "dd")
   cboDay7.text = Format(dtDate, "dd")
   cboDay8.text = Format(dtDate, "dd")
   cboDay9.text = Format(dtDate, "dd")
   cboDay10.text = Format(dtDate, "dd")
   cboDay11.text = Format(dtDate, "dd")
   cboDay12.text = Format(dtDate, "dd")

   cboD1.text = Format(dtDate, "dd")
   cboD2.text = Format(dtDate, "dd")
   cboD3.text = Format(dtDate, "dd")
   cboD4.text = Format(dtDate, "dd")
   cboD5.text = Format(dtDate, "dd")
   cboD6.text = Format(dtDate, "dd")
   cboD7.text = Format(dtDate, "dd")

'Quarterly
   cboM1.text = Format(dtDate, "mmmm")
   cboM2.text = Format(DateAdd("m", 3, dtDate), "mmmm")
   cboM3.text = Format(DateAdd("m", 6, dtDate), "mmmm")
   cboM4.text = Format(DateAdd("m", 9, dtDate), "mmmm")

'Half yearly
   cboM5.text = Format(dtDate, "mmmm")
   cboM6.text = Format(DateAdd("m", 6, dtDate), "mmmm")

'Yearly
   cboM7.text = Format(dtDate, "mmmm")
End Sub

Private Sub cmdCancel_Click()
   Call GetData
   Call DisableBoxes
End Sub

Private Sub cmdClose_Click()
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub cmdCloseInterest_Click()
   Conn.Open getConnectionString

   txtBIRate.text = LastInterestRate

   fraInterestRates.Visible = False

   Conn.Close
'   Set Conn = Nothing
End Sub

Private Sub cmdDeleteInterest_Click()
   If iCurFlxInterestRatesRow < 1 Then Exit Sub

   If flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 5) = "YES" Then
      MsgBox "This Interest Rate cannot possible to delete.", vbCritical + vbOKOnly, "Delete Interest Rate"
      Exit Sub
   End If

   Dim szSQL As String

   If flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 5) = "NO" Then
      If MsgBox("Do you wish to delete the selected Interest Rate?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

      Conn.Open getConnectionString

      szSQL = "UPDATE InterestRates " & _
              "SET Active = False " & _
              "WHERE RateID = " & flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 0) & ";"

      Conn.Execute szSQL
      Conn.Close
'      Set Conn = Nothing
   End If
End Sub

Private Sub cmdDemandTypes_Click()
'   If cboProperty.text = "" Then
'      MsgBox "Please select the Property first.", vbCritical + vbOKOnly, "Demand/Charge Types"
'      Exit Sub
'   End If
'
'   If bNoGD And cboProperty.VisibleCount > 1 Then
'      If MsgBox("Would you like to copy demand types from another Property?", _
'                 vbQuestion + vbYesNo, "Demand Types") = vbYes Then
'         LoadSelProp
'
'         fraSelProp.Top = (frmGlobal.Height / 2) - (fraSelProp.Height / 2)
'         fraSelProp.Left = (frmGlobal.Width / 2) - (fraSelProp.Width / 2)
'         cboSelProp.SetFocus
'      Else
'         frmDemandTypes.szPropertyID = cboProperty.BoundText
'         Load frmDemandTypes
'         frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & cboProperty.text
'         cmdDemandTypes.Enabled = False
'         frmDemandTypes.Show
'         frmDemandTypes.SetFocus
'      End If
'   Else
'      frmDemandTypes.szPropertyID = cboProperty.BoundText
'      Load frmDemandTypes
'      frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & cboProperty.text
'      cmdDemandTypes.Enabled = False
'      frmDemandTypes.Show
'      frmDemandTypes.SetFocus
'   End If
End Sub

Private Sub LoadSelProp()
   Dim TotalRow As Integer, TotalCol As Integer, i As Integer, j As Integer
   Dim Data() As String

   SQLStr = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "WHERE ClientID IN " & _
               "(SELECT ClientID " & _
               " FROM Property " & _
               " WHERE PropertyID = '" & cboProperty.BoundText & "') AND " & _
            " PropertyID IN " & _
               "(SELECT PropertyID " & _
               " FROM DemandTypes " & _
               " GROUP BY PropertyID) " & _
            "ORDER BY PropertyID;"
'Debug.Print SQLStr

   Conn.Open getConnectionString
   
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Rst.EOF Then GoTo NoRes

   TotalRow = Rst.RecordCount
   TotalCol = Rst.Fields.count

   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(Rst.Fields(j).Value), "", Rst.Fields(j).Value)
      Next j
      Rst.MoveNext
      If Rst.EOF Then Exit For
   Next i
   cboSelProp.Column() = Data()
   cboSelProp.ListIndex = 0

NoRes:
   Rst.Close
   Conn.Close
End Sub

Private Sub cmdEdit_Click()
   If cboProperty.text = "" Then Exit Sub
   Call Edit
   bEditGlobalData = True
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdEditInterest_Click()
   If iCurFlxInterestRatesRow < 1 Then Exit Sub
   If flxInterestRates.TextMatrix(flxInterestRates.row, 0) = "" Then Exit Sub
   If flxInterestRates.TextMatrix(flxInterestRates.row, 5) = "NO" Then
      MsgBox "You cannot edit this rate as this is an expired interest rate.", vbInformation + vbOKOnly, "Edit Interest Rate"
      Exit Sub
   End If

   fraSetInterestRates.Left = fraInterestRates.Left
   fraSetInterestRates.Top = fraInterestRates.Top
   fraSetInterestRates.Visible = True
   txtDateFrom.text = flxInterestRates.TextMatrix(flxInterestRates.row, 2)
   txtBaseRate.text = flxInterestRates.TextMatrix(flxInterestRates.row, 3)
   txtAdditionalRate.text = flxInterestRates.TextMatrix(flxInterestRates.row, 4)
   txtRateDescription.text = flxInterestRates.TextMatrix(flxInterestRates.row, 6)

   txtDateFrom.SetFocus
   Label3(18).Caption = "EDITING"

   cmdEditInterest.Enabled = False
End Sub

Private Sub cmdExpandBankCode_Click()
   gridBankCode.Left = txtGlobalBankAccount.Left + Frame1.Left + fraOthers.Left
   gridBankCode.Top = txtGlobalBankAccount.Top + txtGlobalBankAccount.Height + Frame1.Top + fraOthers.Top + 5
   BankAccount
End Sub

Private Sub cmdExpandBankCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then gridBankCode.Visible = False
End Sub

Private Sub cmdInterestRates_Click()
   ConfigureFlxInterestRates

   Conn.Open getConnectionString

   Rst.Open "SELECT * FROM GlobalData WHERE PropertyID = '" & cboProperty.BoundText & "';", Conn, adOpenDynamic, adLockOptimistic
   If Rst.EOF Then
      MsgBox "Until you save the Property global data, you cannot interest rates.", vbInformation + vbOKOnly, "Global Data"
      cmdSave.SetFocus
      Rst.Close
      Conn.Close
'      Set Conn = Nothing
      Exit Sub
   End If
   Rst.Close
   LoadFlxInterestRates Conn

   Conn.Close
'   Set Conn = Nothing

   fraInterestRates.Left = 120
   fraInterestRates.Top = 240
   fraInterestRates.Visible = True
End Sub

Private Sub cmdMthPayDt_Click()
   Load frmPaymentDates
   frmPaymentDates.Show
   Me.Enabled = False
End Sub

Private Sub cmdOkSelProp_Click()
'   If cboSelProp.text <> "" Then
'      Dim adoRst As New ADODB.Recordset
'      Dim iMaxID As Integer
'
'      Conn.Open getConnectionString
'
'      adoRst.Open "SELECT MAX(ID) FROM DemandTypes;", Conn, adOpenStatic, adLockReadOnly
'      iMaxID = CInt(adoRst.Fields.Item(0).Value)
'      adoRst.Close
'
'      SQLStr = "SELECT * " & _
'               "FROM DemandTypes " & _
'               "WHERE PropertyID = '" & cboSelProp.Column(0) & "';"
'
'      adoRst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
'
'      Rst.Open "SELECT * " & _
'               "FROM DemandTypes;", Conn, adOpenDynamic, adLockOptimistic
'
'      While Not adoRst.EOF
'         With adoRst
'            Rst.AddNew
'            Rst.Fields.Item("Id").Value = iMaxID + 1
'            Rst.Fields.Item("Type").Value = .Fields.Item("Type").Value
'            Rst.Fields.Item("InvCrd").Value = .Fields.Item("InvCrd").Value
'            Rst.Fields.Item("Prefix").Value = .Fields.Item("Prefix").Value
'            Rst.Fields.Item("NominalCodeforAmount").Value = .Fields.Item("NominalCodeforAmount").Value
'            Rst.Fields.Item("NominalNameforAmount").Value = .Fields.Item("NominalNameforAmount").Value
'            Rst.Fields.Item("NominalCodeforVAT").Value = .Fields.Item("NominalCodeforVAT").Value
'            Rst.Fields.Item("NominalNameforVAT").Value = .Fields.Item("NominalNameforVAT").Value
'            Rst.Fields.Item("NominalCodeforTotal").Value = .Fields.Item("NominalCodeforTotal").Value
'            Rst.Fields.Item("NominalNameforTotal").Value = .Fields.Item("NominalNameforTotal").Value
'            Rst.Fields.Item("TransactionType").Value = .Fields.Item("TransactionType").Value
'            Rst.Fields.Item("CategoryCode").Value = .Fields.Item("CategoryCode").Value
'            Rst.Fields.Item("PaymentDates").Value = .Fields.Item("PaymentDates").Value
'            Rst.Fields.Item("spare1").Value = .Fields.Item("spare1").Value
'            Rst.Fields.Item("DTGroup").Value = .Fields.Item("DTGroup").Value
'            Rst.Fields.Item("DemandReportName").Value = .Fields.Item("DemandReportName").Value
'            Rst.Fields.Item("PropertyID").Value = cboProperty.BoundText
'            Rst.Update
'            .MoveNext
'         End With
'      Wend
'   End If
'
'   frmDemandTypes.szPropertyID = cboProperty.BoundText
'   Load frmDemandTypes
'   frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & cboProperty.text
'   cmdDemandTypes.Enabled = False
'   frmDemandTypes.Show
'   frmDemandTypes.SetFocus
'
'   fraSelProp.Left = 9720
'   Conn.Close
End Sub

Private Sub cmdSave_Click()
   Dim tempdate As String
   Dim VatCode As Integer

   'make sure all payment dates are entered.
   If MissingDate(cboDay1.text) = True Then Exit Sub
   If MissingDate(cboDay2.text) = True Then Exit Sub
   If MissingDate(cboDay3.text) = True Then Exit Sub
   If MissingDate(cboDay4.text) = True Then Exit Sub
   If MissingDate(cboDay5.text) = True Then Exit Sub
   If MissingDate(cboDay6.text) = True Then Exit Sub
   If MissingDate(cboDay7.text) = True Then Exit Sub
   If MissingDate(cboDay8.text) = True Then Exit Sub
   If MissingDate(cboDay9.text) = True Then Exit Sub
   If MissingDate(cboDay10.text) = True Then Exit Sub
   If MissingDate(cboDay11.text) = True Then Exit Sub
   If MissingDate(cboDay12.text) = True Then Exit Sub

   If MissingDate(cboD1.text) = True Then Exit Sub
   If MissingDate(cboD2.text) = True Then Exit Sub
   If MissingDate(cboD3.text) = True Then Exit Sub
   If MissingDate(cboD4.text) = True Then Exit Sub
   If MissingDate(cboD5.text) = True Then Exit Sub
   If MissingDate(cboD6.text) = True Then Exit Sub
   If MissingDate(cboD7.text) = True Then Exit Sub
   
   If MissingDate(cboM1.text) = True Then Exit Sub
   If MissingDate(cboM2.text) = True Then Exit Sub
   If MissingDate(cboM3.text) = True Then Exit Sub
   If MissingDate(cboM4.text) = True Then Exit Sub
   If MissingDate(cboM5.text) = True Then Exit Sub
   If MissingDate(cboM6.text) = True Then Exit Sub
   If MissingDate(cboM7.text) = True Then Exit Sub
   
   'validate the dates.
   tempdate = Format("January " & cboDay1.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay1.SetFocus
      Exit Sub
   End If

   tempdate = Format("February " & cboDay2.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay2.SetFocus
      Exit Sub
   End If

   tempdate = Format("March " & cboDay3.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay3.SetFocus
      Exit Sub
   End If

   tempdate = Format("April " & cboDay4.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay4.SetFocus
      Exit Sub
   End If

   tempdate = Format("May " & cboDay5.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay5.SetFocus
      Exit Sub
   End If

   tempdate = Format("June " & cboDay6.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay6.SetFocus
      Exit Sub
   End If

   tempdate = Format("July " & cboDay7.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay7.SetFocus
      Exit Sub
   End If

   tempdate = Format("August " & cboDay8.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay8.SetFocus
      Exit Sub
   End If

   tempdate = Format("September " & cboDay9.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay9.SetFocus
      Exit Sub
   End If

   tempdate = Format("October " & cboDay10.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay10.SetFocus
      Exit Sub
   End If

   tempdate = Format("November " & cboDay11.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay11.SetFocus
      Exit Sub
   End If

   tempdate = Format("December " & cboDay12.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay12.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM1.text & " " & cboD1.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD1.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM2.text & " " & cboD2.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD2.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM3.text & " " & cboD3.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD3.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM4.text & " " & cboD4.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD4.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM5.text & " " & cboD5.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD5.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM6.text & " " & cboD6.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD6.SetFocus
      Exit Sub
   End If

   tempdate = Format(cboM7.text & " " & cboD7.text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboD7.SetFocus
      Exit Sub
   End If

   Dim i As Integer

   If txtDemandDaysB4Due.text = "" Then
      MsgBox "Please enter the number of days demand to send before due.", vbCritical + vbOKCancel, "Demand Notice Period"
      txtDemandDaysB4Due.SetFocus
      Exit Sub
   End If

   If cboVatRate.text = "" Then
      MsgBox "Please enter the VAT Rate.", vbCritical + vbOKCancel, "VAT Rate"
      cboVatRate.SetFocus
      Exit Sub
   Else
       For i = 2 To 4
           If Mid(cboVatRate.text, i, 3) = " / " Then VatCode = Left(cboVatRate.text, i - 1)
       Next i
   End If

   If txtGlobalBankAccount.text = "" Then
      MsgBox "Please select the Global Bank Account Number.", vbCritical + vbOKCancel, "Bank Account"
      txtGlobalBankAccount.SetFocus
      Exit Sub
   End If

   Dim conn3 As New ADODB.Connection

'* Save records in the database
   conn3.Open getConnectionString

   Rst.Open "SELECT * " & _
            "FROM GlobalData " & _
            "WHERE PropertyID = '" & cboProperty.BoundText & "' ", _
                    conn3, adOpenDynamic, adLockOptimistic
   If Rst.EOF Then
       Rst.AddNew
   Else
       Rst.MoveFirst
   End If

   If cboProperty.text <> "" Then
       Rst!propertyID = cboProperty.BoundText
   Else
       MsgBox "Please select a property to continue", vbInformation, "Save global data"
       Rst.Close
       conn3.Close
       Exit Sub
   End If

   'Insurance has already saved, but if the insurance
   If cInsurance <> CCur(Label2(7).Caption) Then
      ChangeInsuraceAmount CCur(Label2(7).Caption), cboProperty.BoundText, conn3
   End If

'   If txtTotalArea.text <> "" Then Rst!TotalArea = CLng(txtTotalArea.text)

   If txtFiYrEnd.text <> "" Then Rst!SCYearEnd = txtFiYrEnd.text

   If txtGlobalBankAccount.text <> "" Then Rst!GlobalBankCode = txtGlobalBankAccount.text

   Rst!VatRate = VatCode

   If Rst!MonthlyDueDate1 <> cboDay1.text & " January" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate1, cboDay1.text & " January"
   Rst!MonthlyDueDate1 = cboDay1.text & " January"
   If Rst!MonthlyDueDate2 <> cboDay2.text & " February" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate2, cboDay2.text & " February"
   Rst!MonthlyDueDate2 = cboDay2.text & " February"
   If Rst!MonthlyDueDate3 <> cboDay3.text & " March" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate3, cboDay3.text & " March"
   Rst!MonthlyDueDate3 = cboDay3.text & " March"
   If Rst!MonthlyDueDate4 <> cboDay4.text & " April" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate4, cboDay4.text & " April"
   Rst!MonthlyDueDate4 = cboDay4.text & " April"
   If Rst!MonthlyDueDate5 <> cboDay5.text & " May" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate5, cboDay5.text & " May"
   Rst!MonthlyDueDate5 = cboDay5.text & " May"
   If Rst!MonthlyDueDate6 <> cboDay6.text & " June" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate6, cboDay6.text & " June"
   Rst!MonthlyDueDate6 = cboDay6.text & " June"
   If Rst!MonthlyDueDate7 <> cboDay7.text & " July" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate7, cboDay7.text & " July"
   Rst!MonthlyDueDate7 = cboDay7.text & " July"
   If Rst!MonthlyDueDate8 <> cboDay8.text & " August" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate8, cboDay8.text & " August"
   Rst!MonthlyDueDate8 = cboDay8.text & " August"
   If Rst!MonthlyDueDate9 <> cboDay9.text & " September" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate9, cboDay9.text & " September"
   Rst!MonthlyDueDate9 = cboDay9.text & " September"
   If Rst!MonthlyDueDate10 <> cboDay10.text & " October" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate10, cboDay10.text & " October"
   Rst!MonthlyDueDate10 = cboDay10.text & " October"
   If Rst!MonthlyDueDate11 <> cboDay11.text & " November" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate11, cboDay11.text & " November"
   Rst!MonthlyDueDate11 = cboDay11.text & " November"
   If Rst!MonthlyDueDate12 <> cboDay12.text & " December" Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate12, cboDay12.text & " December"
   Rst!MonthlyDueDate12 = cboDay12.text & " December"

   If Rst!QuarterlyDueDate1 <> cboD1.text & " " & cboM1.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate1, cboD1.text & " " & cboM1.text
   Rst!QuarterlyDueDate1 = cboD1.text & " " & cboM1.text
   If Rst!QuarterlyDueDate2 <> cboD2.text & " " & cboM2.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate2, cboD2.text & " " & cboM2.text
   Rst!QuarterlyDueDate2 = cboD2.text & " " & cboM2.text
   If Rst!QuarterlyDueDate3 <> cboD3.text & " " & cboM3.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate3, cboD3.text & " " & cboM3.text
   Rst!QuarterlyDueDate3 = cboD3.text & " " & cboM3.text
   If Rst!QuarterlyDueDate4 <> cboD4.text & " " & cboM4.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate4, cboD4.text & " " & cboM4.text
   Rst!QuarterlyDueDate4 = cboD4.text & " " & cboM4.text

   If Rst!HalfYearlyDueDate1 <> cboD5.text & " " & cboM5.text Then UpdateLeasePaymentDate conn3, 9, Rst!HalfYearlyDueDate1, cboD5.text & " " & cboM5.text
   Rst!HalfYearlyDueDate1 = cboD5.text & " " & cboM5.text
   If Rst!HalfYearlyDueDate2 <> cboD6.text & " " & cboM6.text Then UpdateLeasePaymentDate conn3, 9, Rst!HalfYearlyDueDate2, cboD6.text & " " & cboM6.text
   Rst!HalfYearlyDueDate2 = cboD6.text & " " & cboM6.text

   If Rst!YearlyDueDate <> cboD7.text & " " & cboM7.text Then UpdateLeasePaymentDate conn3, 11, Rst!YearlyDueDate, cboD7.text & " " & cboM7.text
   Rst!YearlyDueDate = cboD7.text & " " & cboM7.text

   Rst!NoOfDaysToSendDemandsB4Due = CInt(IIf(txtDemandDaysB4Due.text = "", 0, txtDemandDaysB4Due.text))
   Rst.Update

   Rst.Close

   ShowMsgInTaskBar "Your changes have been saved."

   Call DisableBoxes

   conn3.Close
End Sub

Private Sub UpdateLeasePaymentDate(dbConn As ADODB.Connection, iFrequency As Integer, szCurDate As String, szNewDate As String)
   If szCurDate = "" Then Exit Sub     'New Global Data Entry
   
   Dim adoRec As New ADODB.Recordset
   Dim dtNextDueDate As Date, dtNewDueDate As Date

   dtNextDueDate = CDate(Format(szCurDate, "dd mmmm yyyy"))
   dtNewDueDate = CDate(Format(szNewDate, "dd mmmm yyyy"))

    'Service Charge
   adoRec.Open "SELECT S.SCNextDueDate, L.UnitNumber, S.ServiceCharge " & _
                  "FROM LeaseDetails AS L, LServiceCharges AS S " & _
                  "WHERE L.Status = True And " & _
                     "(S.SCFrequency = " & iFrequency & " or " & _
                     "S.SCFrequency = " & iFrequency + 1 & ") And " & _
                     "L.SCPayable = 'Y' AND " & _
                     "S.LeaseID = L.LeaseID", _
                        dbConn, adOpenStatic, adLockReadOnly
   If adoRec.EOF Then
      adoRec.Close
      Set adoRec = Nothing
   Else
      adoRec.MoveFirst
      While Not adoRec.EOF
         If Format(adoRec!SCNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
            InThisProperty(dbConn, adoRec!UnitNumber) Then
            dbConn.Execute "UPDATE LServiceCharges " & _
                           "SET    SCNextDueDate = #" & CDate(Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!SCNextDueDate, "yyyy")) & "# " & _
                           "WHERE  ServiceCharge = '" & adoRec!ServiceCharge & "';"
         End If
         adoRec.MoveNext
      Wend
      adoRec.Close
   End If

   'Rent
   adoRec.Open "SELECT R.BRNextDueDate, L.UnitNumber, R.RentCharges " & _
                  "FROM LeaseDetails AS L, LRentCharges AS R " & _
                  "WHERE L.Status = True And " & _
                     "(R.BRFrequency = " & iFrequency & " or " & _
                     "R.BRFrequency = " & iFrequency + 1 & ") And " & _
                     "L.BRPayable = 'Y' And " & _
                     "R.LeaseID = L.LeaseID", _
                        dbConn, adOpenStatic, adLockReadOnly
   If adoRec.EOF Then
      adoRec.Close
      Set adoRec = Nothing
   Else
      adoRec.MoveFirst
      While Not adoRec.EOF
         If Format(adoRec!BRNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
            InThisProperty(dbConn, adoRec!UnitNumber) Then
            dbConn.Execute "UPDATE LRentCharges " & _
                           "SET    BRNextDueDate = #" & CDate(Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!BRNextDueDate, "yyyy")) & "# " & _
                           "WHERE  RentCharges = '" & adoRec!RentCharges & "';"
         End If
         adoRec.MoveNext
      Wend
      adoRec.Close
   End If

   'Insurance
   adoRec.Open "SELECT I.InsuranceNextDueDate, L.UnitNumber, I.InsCharges " & _
                  "FROM LeaseDetails AS L, LInsuranceCharges as I " & _
                  "WHERE L.Status = True And " & _
                     "(I.InsuranceFrequency = " & iFrequency & " or " & _
                     "I.InsuranceFrequency = " & iFrequency + 1 & ") And " & _
                     "L.InsurancePayable = 'Y' AND " & _
                     "I.LeaseID = L.LeaseID", _
                        dbConn, adOpenStatic, adLockReadOnly
   If adoRec.EOF Then
      adoRec.Close
      Set adoRec = Nothing
   Else
      adoRec.MoveFirst
      While Not adoRec.EOF
         If Format(adoRec!InsuranceNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
            InThisProperty(dbConn, adoRec!UnitNumber) Then
            dbConn.Execute "UPDATE LInsuranceCharges " & _
                           "SET    InsuranceNextDueDate = #" & CDate(Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!InsuranceNextDueDate, "yyyy")) & "# " & _
                           "WHERE  InsCharges = '" & adoRec!InsCharges & "';"
         End If
         adoRec.MoveNext
      Wend
      adoRec.Close
   End If
End Sub

Private Function InThisProperty(dbConn As ADODB.Connection, szUnitNumber As String) As Boolean
   Dim adoRec As New ADODB.Recordset
   
   adoRec.Open "SELECT * " & _
                  "FROM Units " & _
                  "WHERE Units.UnitNumber = '" & szUnitNumber & "' And " & _
                     "Units.PropertyID = '" & cboProperty.BoundText & "'", _
                        dbConn, adOpenDynamic, adLockOptimistic
   If adoRec.EOF Then
      InThisProperty = False
   Else
      InThisProperty = True
   End If

   adoRec.Close
   Set adoRec = Nothing
End Function

Private Sub cmdSetIntRateClose_Click()
   fraSetInterestRates.Visible = False

   cmdAddInterest.Enabled = True
   cmdEditInterest.Enabled = True
End Sub

Private Sub cmdSetIntRateSave_Click()
   Dim vChoice

   vChoice = MsgBox("Do you want to update the interest information?", vbQuestion + vbYesNoCancel, "Interest Rate")

   If vChoice = vbNo Then
      cmdSetIntRateClose_Click
      Exit Sub
   End If
   If vChoice = vbCancel Then
      Exit Sub
   End If

   If txtDateFrom.text = "" Then
      MsgBox "Please input the from date of the interest rate.", vbCritical + vbOKOnly, "Interest Rate"
      Exit Sub
   End If
   If txtBaseRate.text = "" Then
      MsgBox "Please input the Base rate of the interest rate.", vbCritical + vbOKOnly, "Interest Rate"
      Exit Sub
   End If
   If txtAdditionalRate.text = "" Then
      MsgBox "Please input the Additional rate of the interest rate.", vbCritical + vbOKOnly, "Interest Rate"
      Exit Sub
   End If

   Conn.Open getConnectionString
   Dim adoGD As New ADODB.Recordset

   If Label3(18).Caption = "ADDING" Then
      SQLStr = "SELECT * FROM InterestRates;"
   Else
      SQLStr = "SELECT * FROM InterestRates WHERE RateID = " & flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 0) & ";"
   End If
   Rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

   If Label3(18).Caption = "ADDING" Then
      Rst.AddNew
      adoGD.Open "SELECT BaseInterestRate FROM GlobalData WHERE PropertyID = '" & cboProperty.BoundText & "';", Conn, adOpenDynamic, adLockOptimistic

      adoGD!BaseInterestRate = Rst.RecordCount + 1
      adoGD.Update
      adoGD.Close
      Set adoGD = Nothing
   End If

   With Rst
      !propertyID = cboProperty.BoundText
      !DateFrom = CDate(txtDateFrom.text)
      !BaseRate = CSng(txtBaseRate.text)
      !AdditionalRate = CSng(txtAdditionalRate.text)
      !Active = True
      !RateDescription = IIf(IsNull(txtRateDescription.text), "", txtRateDescription.text)
      .Update
   End With

   MsgBox "Data has been updated.", vbInformation + vbOKOnly
   Rst.Close
'   Set Rst = Nothing

   LoadFlxInterestRates Conn

   Conn.Close
'   Set Conn = Nothing

   cmdSetIntRateClose_Click
End Sub

Private Sub cmdYearlyInsurance_Click()
   frmRentBudget1.sModule = "IB"              'Insurance Budget
   Load frmRentBudget1
   frmRentBudget1.Caption = "Insurance " + frmRentBudget1.Caption
   frmRentBudget1.Show
   Me.Enabled = False
End Sub

Private Sub cmdYearlyRent_Click()
   frmRentBudget1.sModule = "RB"              'Rent budget
   Load frmRentBudget1
   frmRentBudget1.Caption = "Rent " + frmRentBudget1.Caption
   frmRentBudget1.Show
   Me.Enabled = False
End Sub

Private Sub cmdYearlyService_Click()
   Load frmServiceCharge1
   frmServiceCharge1.Show
   Me.Enabled = False
End Sub

Private Sub flxInterestRates_Click()
   iCurFlxInterestRatesRow = flxInterestRates.row
End Sub

Private Sub HighLightRow(FlxGrid As MSHFlexGrid)
   Dim iCol As Integer, iRow As Integer
   Dim iCurCol As Integer, iCurRow As Integer

'   Saving current row and column selection
   iCurRow = FlxGrid.row
   iCurCol = FlxGrid.col

'  Clear any privious selection
   For iRow = 1 To FlxGrid.Rows - 1
      For iCol = 2 To FlxGrid.Cols - 1
         FlxGrid.col = iCol
         FlxGrid.row = iRow
         FlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iCol
   Next iRow

'  set the cellback color depends on current selection
   FlxGrid.row = iCurRow
   For iCol = 2 To FlxGrid.Cols - 1
      FlxGrid.col = iCol
      FlxGrid.CellBackColor = RGB(244, 244, 244)
   Next iCol

'  Set back original row and col selection, bacause setting the colback color it was changed
   FlxGrid.col = iCurCol
End Sub

Private Sub Form_Load()
   MyForm.Height = 7965 'Me.Height ' Remember the current size
   MyForm.Width = 9315 'Me.Width
   Me.Width = 9315
   Me.Height = 7965
   frmMMain.Arrange vbCascade
   Me.ZOrder 0

   Me.Caption = "Global Data"
   MousePointer = vbHourglass

   DisableBoxes

   Conn.Open getConnectionString
   PrepareList Conn
   Conn.Close
'   Set Conn = Nothing

   Call LoadProperty
   Call FillDaysMonths
   Call GetVATRates

   bEditGlobalData = False
   SSTab1.Tab = 0
   MousePointer = vbDefault
End Sub

Private Sub ConfigureFlxInterestRates()
   Dim szFlxHeader As String

   flxInterestRates.RowHeight(0) = 0
   flxInterestRates.Clear
   flxInterestRates.Cols = 7
   flxInterestRates.Rows = 2
   szFlxHeader$ = "<ID|<PropertyID|<DateFrom|>BaseRate|>AdditionalRate|<Active|<RateDescription"
   flxInterestRates.FormatString = szFlxHeader$

   flxInterestRates.ColWidth(0) = 0
   flxInterestRates.ColWidth(1) = 0
   flxInterestRates.ColWidth(2) = Label3(9).Left - Label3(8).Left
   flxInterestRates.ColWidth(3) = Label3(10).Left - Label3(9).Left
   flxInterestRates.ColWidth(4) = flxInterestRates.Width - flxInterestRates.Left - Label3(10).Left
   flxInterestRates.ColWidth(5) = 0
   flxInterestRates.ColWidth(6) = 0
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboClientList.Column() = Data()
   cboClientList.ListIndex = 0
   adoRst.Close

NoRes:
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Public Sub LoadProperty()
   Dim sSQLQuery_ As String
   adoProperty.ConnectionString = getConnectionString

   If cboClientList.Column(0) = "ALL" Then
      sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
                   "FROM PROPERTY "
   Else
      sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
                   "FROM PROPERTY " & _
                   "WHERE PROPERTY.ClientID = '" & cboClientList.Column(0) & "';"
   End If

   adoProperty.RecordSource = sSQLQuery_
   adoProperty.CommandType = adCmdText
   adoProperty.Refresh
End Sub

Public Function GetData() As Boolean
   Dim i As Integer, c As String, bEH As Boolean
   
   Conn.Open getConnectionString

   SQLStr = "SELECT * FROM GlobalData WHERE PropertyID = '" & cboProperty.BoundText & "' "
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Rst.RecordCount = 0 Then
       Rst.Close
       Conn.Close
       GetData = False
       Exit Function
   End If

'   If IsNull(Rst!TotalArea) Then txtTotalArea.text = "" Else txtTotalArea.text = Rst!TotalArea
   If IsNull(Rst!SCYearEnd) Then txtFiYrEnd.text = "" Else txtFiYrEnd.text = Rst!SCYearEnd
   If IsNull(Rst!GlobalBankCode) Then txtGlobalBankAccount.text = "" Else txtGlobalBankAccount.text = Rst!GlobalBankCode

   If Not IsNull(Rst!BaseInterestRate) Or Rst!BaseInterestRate <> "" Then
      txtBIRate.text = LastInterestRate
   Else
      txtBIRate.text = "0.0000"
   End If

   For i = 0 To cboVatRate.ListCount - 1
       c = cboVatRate.List(i)
       If CInt(Left(c, 2)) = Rst!VatRate Then
           cboVatRate.text = c
       End If
   Next i

   If IsNull(Rst!NoOfDaysToSendDemandsB4Due) = False Then txtDemandDaysB4Due.text = Rst!NoOfDaysToSendDemandsB4Due

   On Error GoTo DateErrorHandler

   bEH = True
   cboDay1.text = Left(Rst!MonthlyDueDate1, 2)
   cboDay2.text = Left(Rst!MonthlyDueDate2, 2)
   cboDay3.text = Left(Rst!MonthlyDueDate3, 2)
   cboDay4.text = Left(Rst!MonthlyDueDate4, 2)
   cboDay5.text = Left(Rst!MonthlyDueDate5, 2)
   cboDay6.text = Left(Rst!MonthlyDueDate6, 2)
   cboDay7.text = Left(Rst!MonthlyDueDate7, 2)
   cboDay8.text = Left(Rst!MonthlyDueDate8, 2)
   cboDay9.text = Left(Rst!MonthlyDueDate9, 2)
   cboDay10.text = Left(Rst!MonthlyDueDate10, 2)
   cboDay11.text = Left(Rst!MonthlyDueDate11, 2)
   cboDay12.text = Left(Rst!MonthlyDueDate12, 2)
   cboD2.text = Left(Rst!QuarterlyDueDate2, 2)
   cboM2.text = Right(Rst!QuarterlyDueDate2, Len(Rst!QuarterlyDueDate2) - 3)
   cboD1.text = Left(Rst!QuarterlyDueDate1, 2)
   cboM1.text = Right(Rst!QuarterlyDueDate1, Len(Rst!QuarterlyDueDate1) - 3)
   cboD3.text = Left(Rst!QuarterlyDueDate3, 2)
   cboM3.text = Right(Rst!QuarterlyDueDate3, Len(Rst!QuarterlyDueDate3) - 3)
   cboD4.text = Left(Rst!QuarterlyDueDate4, 2)
   cboM4.text = Right(Rst!QuarterlyDueDate4, Len(Rst!QuarterlyDueDate4) - 3)
   cboD5.text = Left(Rst!HalfYearlyDueDate1, 2)
   cboM5.text = Right(Rst!HalfYearlyDueDate1, Len(Rst!HalfYearlyDueDate1) - 3)
   cboD6.text = Left(Rst!HalfYearlyDueDate2, 2)
   cboM6.text = Right(Rst!HalfYearlyDueDate2, Len(Rst!HalfYearlyDueDate2) - 3)
   cboD7.text = Left(Rst!YearlyDueDate, 2)
   cboM7.text = Right(Rst!YearlyDueDate, Len(Rst!YearlyDueDate) - 3)
   bEH = False

DateErrorHandler:
   If bEH Then
      MsgBox "Please set a correct date for the default payment date set.", vbCritical + vbOKOnly, "Wrong Date"
   End If

   Label2(7).Caption = Format(TotalInsuranceProperty(cboProperty.BoundText), "£0.00")
   Label2(8).Caption = Format(TotalRCProperty(cboProperty.BoundText), "£0.00")
   lblSCTotal.Caption = Format(TotalSCTProperty(cboProperty.BoundText), "£0.00")
   
   cInsurance = CCur(Label2(7).Caption) 'save it to know at the end is there any modification of insurance

   Rst.Close

   Conn.Close

'   Set Rst = Nothing
'   Set Conn = Nothing
   GetData = True
   Exit Function
ErrH:
'   Set Rst = Nothing
'   Set Conn = Nothing
   MsgBox Err.Number & " - " & Err.description, vbOKOnly, "Error"
End Function

Private Function LastInterestRate() As String
   Dim sqlRST As New ADODB.Recordset

   SQLStr = "SELECT * FROM InterestRates " & _
            "WHERE PropertyID = '" & cboProperty.BoundText & "' " & _
            "ORDER BY DateFrom ASC;"
   sqlRST.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If sqlRST.EOF Then
      LastInterestRate = "0.0000"
   Else
      sqlRST.MoveLast
      LastInterestRate = Format(sqlRST!BaseRate, "0.0000")
      sqlRST.Close
   End If

   Set sqlRST = Nothing
End Function

Private Function TotalRCProperty(szPropertyID As String) As Double
   Dim Rst2 As New ADODB.Recordset

   SQLStr = "SELECT GlobalRC.PropertyID, SUM(GlobalRC.TotalBudget) AS TOTALRENT " & _
            "From GlobalRC " & _
            "WHERE GlobalRC.PropertyID = '" & szPropertyID & "' " & _
            "GROUP BY GlobalRC.PropertyID;"
'Debug.Print SQLStr
   Rst2.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
      TotalRCProperty = CDbl(Rst2!TOTALRENT)
   Else
      TotalRCProperty = 0
   End If

   Rst2.Close
   Set Rst2 = Nothing
End Function

Private Function TotalSCTProperty(szPropertyID As String) As Double
   Dim Rst2 As New ADODB.Recordset

   SQLStr = "SELECT GlobalSC.PropertyID, SUM(GlobalSC.TotalBudget) AS TOTALSC " & _
            "From GlobalSC " & _
            "WHERE GlobalSC.PropertyID = '" & szPropertyID & "' " & _
            "GROUP BY GlobalSC.PropertyID;"
'Debug.Print SQLStr
   Rst2.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
      TotalSCTProperty = CDbl(Rst2!TOTALSC)
   Else
      TotalSCTProperty = 0
   End If

   Rst2.Close
   Set Rst2 = Nothing
End Function


Private Function TotalInsuranceProperty(szPropertyID As String) As Double
   Dim Rst2 As New ADODB.Recordset

   SQLStr = "SELECT GlobalInsurance.PropertyID, SUM(GlobalInsurance.Amount) AS TOTALINSURANCE " & _
            "From GlobalInsurance " & _
            "WHERE GlobalInsurance.PropertyID = '" & szPropertyID & "' " & _
            "GROUP BY GlobalInsurance.PropertyID;"
   Rst2.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
      TotalInsuranceProperty = CDbl(Rst2!TotalInsurance)
   Else
      TotalInsuranceProperty = 0
   End If

   Rst2.Close
   Set Rst2 = Nothing
End Function

Private Sub LoadFlxInterestRates(Conn1 As ADODB.Connection)
   Dim i As Integer

   SQLStr = "SELECT * FROM InterestRates " & _
            "WHERE PropertyID = '" & cboProperty.BoundText & "' AND " & _
               "Active = TRUE;"

   Rst.Open SQLStr, Conn1, adOpenStatic, adLockReadOnly

   i = 1
   flxInterestRates.Rows = 2
   If Not Rst.EOF Then
      While Not Rst.EOF
         flxInterestRates.TextMatrix(i, 0) = Rst!RateID
         flxInterestRates.TextMatrix(i, 1) = Rst!propertyID
         flxInterestRates.TextMatrix(i, 2) = Rst!DateFrom
         flxInterestRates.TextMatrix(i, 3) = Format(Rst!BaseRate, "0.0000")
         flxInterestRates.TextMatrix(i, 4) = Format(Rst!AdditionalRate, "0.0000")
         flxInterestRates.TextMatrix(i, 5) = IIf(Rst!Active = True, "YES", "NO")
         flxInterestRates.TextMatrix(i, 6) = IIf(IsNull(Rst!RateDescription), "", Rst!RateDescription)

         i = i + 1
         Rst.MoveNext
         If Not Rst.EOF Then flxInterestRates.AddItem ""
      Wend
   Else
      flxInterestRates.Rows = 1
   End If
   flxInterestRates.row = 0
   flxInterestRates.col = 0

   Rst.Close
'   Set Rst = Nothing
End Sub

Public Function CheckDecimal(Value As String) As String
    Dim i As Integer
    Dim char As String
    Dim a As Integer

    a = 0
    If Asc(Mid(Value, 1, 1)) = 46 Then Value = "0" + Value
    For i = 2 To Len(Value)
        char = Mid(Value, i, 1)
        If Asc(char) = 46 And i = Len(Value) - 2 Then a = 1
        If Asc(char) = 46 And i = Len(Value) - 1 Then
            Value = Value + "0"
            a = 1
        End If
        If Asc(char) = 46 And i = Len(Value) Then
            Value = Value + "00"
            a = 1
        End If
    Next i
    If a = 0 Then Value = Value + ".00"
    CheckDecimal = Value
End Function

Public Sub DisableBoxes()
   cmdYearlyInsurance.Enabled = False
   cmdYearlyRent.Enabled = False
   cmdYearlyService.Enabled = False
   fraOthers.Enabled = False
   fraDemandDaysB4Due.Enabled = False

   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
   Frame6.Enabled = False
   Frame7.Enabled = False
   cmdAutoSetup(0).Enabled = False
   cmdMthPayDt.Enabled = False
   
   cmdEdit.Visible = True
   cmdSave.Visible = False
   cmdCancel.Visible = False
   cboProperty.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Resize()
'   Dim ScaleFactorX As Single, ScaleFactorY As Single
'
'   ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
'   ScaleFactorY = Me.Height / MyForm.Height
'   Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'   MyForm.Height = Me.Height ' Remember the current size
'   MyForm.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim X As Integer

   If cmdSave.Visible Then
      X = MsgBox("Do you want to save changes?", vbQuestion + vbYesNoCancel, "Data Saving")
      If X = vbCancel Then Cancel = 1
      If X = vbYes Then cmdSave_Click
   End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub gridBankCode_Click()
   Dim iRow As Integer
   iRow = gridBankCode.row
   txtGlobalBankAccount.text = gridBankCode.TextMatrix(iRow, 0)
   gridBankCode.Visible = False
End Sub

Private Sub gridBankCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim iRow As Integer

      iRow = gridBankCode.row
      txtGlobalBankAccount.text = gridBankCode.TextMatrix(iRow, 0)
      gridBankCode.Visible = False
   End If
End Sub

Private Sub mnuEdit_Click()
   Call Edit
End Sub

Public Sub Edit()
   cmdYearlyInsurance.Enabled = True
   cmdYearlyRent.Enabled = True
   cmdYearlyService.Enabled = True
   fraOthers.Enabled = True
   fraDemandDaysB4Due.Enabled = True

   Frame2.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
   Frame6.Enabled = True
   Frame7.Enabled = True
   cmdAutoSetup(0).Enabled = True
   cmdMthPayDt.Enabled = True

   cmdEdit.Visible = False
   cmdSave.Visible = True
   cmdCancel.Visible = True
   cboProperty.Enabled = False
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   If txtDateFrom.text <> "" Then TextBoxFormatDate txtDateFrom
End Sub

Private Sub txtDemandDaysB4Due_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Added By Samrat. 12/10/2006
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtDemandDaysB4Due, KA, 0
   KeyAscii = KA
End Sub

Private Sub txtDemandDaysB4Due_LostFocus()
   If txtDemandDaysB4Due.text <> "" Then If NumberCheck(txtDemandDaysB4Due.text) = False Then txtDemandDaysB4Due.text = ""
End Sub

Public Sub FillDaysMonths()
   Dim i As Integer
   Dim months(1 To 12)

   For i = 1 To 9
      cboD1.AddItem "0" & i
      cboD2.AddItem "0" & i
      cboD3.AddItem "0" & i
      cboD4.AddItem "0" & i
      cboD5.AddItem "0" & i
      cboD6.AddItem "0" & i
      cboD7.AddItem "0" & i

      cboDay1.AddItem "0" & i
      cboDay2.AddItem "0" & i
      cboDay3.AddItem "0" & i
      cboDay4.AddItem "0" & i
      cboDay5.AddItem "0" & i
      cboDay6.AddItem "0" & i
      cboDay7.AddItem "0" & i
      cboDay8.AddItem "0" & i
      cboDay9.AddItem "0" & i
      cboDay10.AddItem "0" & i
      cboDay11.AddItem "0" & i
      cboDay12.AddItem "0" & i
   Next i

   For i = 10 To 31
      cboD1.AddItem i
      cboD2.AddItem i
      cboD3.AddItem i
      cboD4.AddItem i
      cboD5.AddItem i
      cboD6.AddItem i
      cboD7.AddItem i

      cboDay1.AddItem i
      cboDay2.AddItem i
      cboDay3.AddItem i
      cboDay4.AddItem i
      cboDay5.AddItem i
      cboDay6.AddItem i
      cboDay7.AddItem i
      cboDay8.AddItem i
      cboDay9.AddItem i
      cboDay10.AddItem i
      cboDay11.AddItem i
      cboDay12.AddItem i
   Next i
   
   months(1) = "January"
   months(2) = "February"
   months(3) = "March"
   months(4) = "April"
   months(5) = "May"
   months(6) = "June"
   months(7) = "July"
   months(8) = "August"
   months(9) = "September"
   months(10) = "October"
   months(11) = "November"
   months(12) = "December"

   months(1) = "January"
   months(2) = "February"
   months(3) = "March"
   months(4) = "April"
   months(5) = "May"
   months(6) = "June"
   months(7) = "July"
   months(8) = "August"
   months(9) = "September"
   months(10) = "October"
   months(11) = "November"
   months(12) = "December"

'QUARTERLY
   For i = 1 To 3
      cboM1.AddItem months(i)
   Next i
   For i = 4 To 6
      cboM2.AddItem months(i)
   Next i
   For i = 7 To 9
      cboM3.AddItem months(i)
   Next i
   For i = 10 To 12
      cboM4.AddItem months(i)
   Next i
'HALF YEARLY
   For i = 1 To 6
      cboM5.AddItem months(i)
   Next i
   For i = 7 To 12
      cboM6.AddItem months(i)
   Next i
'YEARLY
   For i = 1 To 12
      cboM7.AddItem months(i)
   Next i
End Sub

Public Function MissingDate(text As String) As Boolean
   MissingDate = False
   If text = "" Then
       MsgBox "You must select all the payment dates", vbOKOnly + vbCritical, "Missing Payment Date"
       MissingDate = True
   End If
End Function

Public Function ValidDate(text As String) As Boolean
   ValidDate = True
   If IsDate(text) = False Then
       MsgBox "Invalid Date Selected.", vbOKOnly + vbCritical, "Invalid Date"
       ValidDate = False
   End If
End Function

Public Sub GetVATRates()
   Conn.Open getConnectionString

   'SQLStr = "SELECT VAT_CODE, VAT_RATE, VAT_RATE_NAME FROM SYS_VAT_FILE ORDER BY VAT_CODE"   'CHANGE TO SageLine50v12
   SQLStr = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM tlbVATCODE ORDER BY VAT_ID"
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   While Rst.EOF = False
      cboVatRate.AddItem Rst!VAT_ID & " / " & Rst!VAT_CODE & " / " & Rst!VAT_RATE
      Rst.MoveNext
   Wend

   Rst.Close
   Conn.Close
End Sub

Private Sub BankAccount()
   ' Error Handler
   On Error GoTo Error_Handler

   gridBankCode.Visible = True

   gridBankCode.TextMatrix(0, 0) = "Reference"
   gridBankCode.TextMatrix(0, 1) = "Name"
   gridBankCode.ColWidth(0) = 1200
   gridBankCode.ColWidth(1) = 2600

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Please setup bank account for the client."
   Else
      rRow = 1
      While Not adoRst.EOF
         gridBankCode.TextMatrix(rRow, 0) = adoRst.Fields.Item("BNC").Value
         gridBankCode.TextMatrix(rRow, 1) = adoRst.Fields.Item("BNN").Value
         gridBankCode.AddItem ""
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Load Bank Account in Demand"

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub txtGlobalBankAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      gridBankCode.SetFocus
      cmdExpandBankCode_Click
   End If
End Sub

Private Sub txtFiYrEnd_Change()
   TextBoxChangeDate txtFiYrEnd
End Sub

Private Sub txtFiYrEnd_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFiYrEnd, KeyAscii
End Sub

Private Sub txtFiYrEnd_LostFocus()
   TextBoxFormatDate txtFiYrEnd
End Sub
