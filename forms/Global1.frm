VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGlobalx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Data"
   ClientHeight    =   11115
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   19140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Global1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   19140
   Begin VB.PictureBox picBankCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   10125
      ScaleHeight     =   3180
      ScaleWidth      =   4260
      TabIndex        =   126
      Top             =   585
      Width           =   4290
      Begin VB.CommandButton cmdBankClose 
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
         Left            =   3915
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   70
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
         Height          =   2715
         Left            =   45
         TabIndex        =   134
         Top             =   405
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   4789
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
      Begin MSForms.Label Label9 
         Height          =   195
         Left            =   180
         TabIndex        =   128
         Top             =   90
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "Bank Code"
         Size            =   "2170;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   195
         Left            =   1680
         TabIndex        =   127
         Top             =   105
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Bank Name"
         Size            =   "2090;344"
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
         Index           =   0
         Left            =   90
         Top             =   90
         Width           =   3780
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   9405
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   115
      Top             =   8235
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
         TabIndex        =   116
         Top             =   70
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   117
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   118
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
         Top             =   90
         Width           =   5850
      End
   End
   Begin VB.Frame fraBudgetYear 
      BackColor       =   &H00FDEDED&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   12015
      TabIndex        =   103
      Top             =   5130
      Visible         =   0   'False
      Width           =   6540
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFinancialYears 
         Height          =   3105
         Left            =   120
         TabIndex        =   104
         Top             =   360
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   5477
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
      Begin MSForms.CommandButton cmdFYSet 
         Height          =   360
         Left            =   120
         TabIndex        =   110
         Top             =   3555
         Visible         =   0   'False
         Width           =   1155
         ForeColor       =   16384
         Caption         =   "Set"
         PicturePosition =   65543
         Size            =   "2037;635"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label Label7 
         Height          =   225
         Index           =   3
         Left            =   5160
         TabIndex        =   109
         Top             =   195
         Width           =   945
         ForeColor       =   16384
         BackColor       =   -2147483637
         VariousPropertyBits=   276824083
         Caption         =   "Frequency"
         Size            =   "1667;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   108
         Top             =   195
         Width           =   645
         ForeColor       =   16384
         BackColor       =   -2147483637
         VariousPropertyBits=   276824083
         Caption         =   "Financial Year"
         Size            =   "1138;714"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   107
         Top             =   195
         Width           =   855
         ForeColor       =   16384
         BackColor       =   -2147483637
         VariousPropertyBits=   276824083
         Caption         =   "Description"
         Size            =   "1508;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   106
         Top             =   195
         Width           =   555
         ForeColor       =   16384
         BackColor       =   -2147483637
         VariousPropertyBits=   276824083
         Caption         =   "Periods"
         Size            =   "979;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdFYClose 
         Height          =   360
         Left            =   5280
         TabIndex        =   105
         Top             =   3555
         Width           =   1155
         ForeColor       =   16384
         Caption         =   "Close"
         PicturePosition =   65543
         Size            =   "2037;635"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraSelProp 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9960
      TabIndex        =   94
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
         TabIndex        =   96
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
         TabIndex        =   97
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
         TabIndex        =   95
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Year Budget"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   135
      TabIndex        =   65
      Top             =   585
      Width           =   9915
      Begin VB.Frame Frame9 
         Caption         =   "Tax Payment"
         Height          =   690
         Left            =   4410
         TabIndex        =   145
         Top             =   2160
         Width           =   4920
         Begin VB.OptionButton optClienttosubmit 
            Caption         =   "Client To Pay Tax"
            Height          =   285
            Left            =   2430
            TabIndex        =   151
            Top             =   270
            Width           =   2175
         End
         Begin VB.OptionButton optAgentToSubmit 
            Caption         =   "Agent To Pay Tax"
            Height          =   285
            Left            =   135
            TabIndex        =   150
            Top             =   270
            Value           =   -1  'True
            Width           =   2040
         End
      End
      Begin VB.CheckBox chkRestricktedToBudget 
         Height          =   285
         Left            =   1800
         TabIndex        =   142
         Top             =   3105
         Width           =   240
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tax Return"
         Height          =   1950
         Left            =   4410
         TabIndex        =   135
         Top             =   225
         Width           =   4920
         Begin VB.CheckBox chkProduceVatReturn 
            Caption         =   "Produce Vat Return"
            Height          =   240
            Left            =   270
            TabIndex        =   144
            Top             =   270
            Width           =   2850
         End
         Begin VB.TextBox txtTaxInterval 
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
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   148
            Top             =   1260
            Width           =   480
         End
         Begin VB.TextBox txtLastCompletedTaxReturnDate 
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
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   147
            Top             =   945
            Width           =   2190
         End
         Begin VB.TextBox txtCurrentTaxPeriod 
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
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   149
            Top             =   1575
            Width           =   2190
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Basis (Scheme):  "
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   140
            Top             =   630
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Completed Tax return Date:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   6
            Left            =   180
            TabIndex        =   139
            Top             =   945
            Width           =   2265
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Return Interval :"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   9
            Left            =   165
            TabIndex        =   138
            Top             =   1275
            Width           =   1410
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Tax Period :"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   3
            Left            =   135
            TabIndex        =   137
            Top             =   1590
            Width           =   1395
         End
         Begin MSForms.ComboBox cboTaxBasis 
            Height          =   330
            Left            =   2520
            TabIndex        =   146
            Top             =   575
            Width           =   2220
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3916;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(In months)"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   7
            Left            =   3060
            TabIndex        =   136
            Top             =   1305
            Width           =   780
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "£0.00"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9690
         TabIndex        =   102
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdFinancialYear 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1755
         TabIndex        =   3
         Top             =   2655
         Width           =   2565
      End
      Begin VB.CommandButton cmdYearlyRent 
         Caption         =   "£0.00"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9690
         TabIndex        =   100
         Top             =   2145
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdYearlyInsurance 
         Caption         =   "£0.00"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9690
         TabIndex        =   99
         Top             =   1755
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdYearlyService 
         Caption         =   "£0.00"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9690
         TabIndex        =   98
         Top             =   1380
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame fraInterestRates 
         BackColor       =   &H00FDEDED&
         Caption         =   "Lessee Finance Rates:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   9630
         TabIndex        =   78
         Top             =   2250
         Visible         =   0   'False
         Width           =   4260
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInterestRates 
            Height          =   1755
            Left            =   120
            TabIndex        =   79
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
         Begin MSForms.CommandButton cmdDeleteInterest 
            Height          =   360
            Left            =   2285
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4410
         TabIndex        =   91
         Top             =   2880
         Width           =   4905
         Begin MSForms.TextBox txtDemandDaysB4Due 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   315
            Width           =   495
            VariousPropertyBits=   142622747
            MaxLength       =   4
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
            Left            =   345
            TabIndex        =   92
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame fraOthers 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   2205
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   4140
         Begin VB.CheckBox chkGlobalVat 
            Caption         =   "Check1"
            Height          =   285
            Left            =   1665
            TabIndex        =   141
            Top             =   1305
            Width           =   285
         End
         Begin VB.CommandButton cmdVatRate 
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
            Height          =   300
            Left            =   3720
            TabIndex        =   132
            Top             =   1305
            Width           =   300
         End
         Begin VB.TextBox txtVatRate 
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
            Left            =   2025
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   133
            Top             =   1305
            Width           =   1650
         End
         Begin VB.CommandButton cmdInterestRates 
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
            Height          =   300
            Left            =   3735
            TabIndex        =   131
            Top             =   1800
            Width           =   300
         End
         Begin VB.CommandButton cmdExpandBankCode 
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
            Height          =   300
            Left            =   3735
            TabIndex        =   2
            Top             =   765
            Width           =   300
         End
         Begin VB.TextBox txtGlobalBankAccount 
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
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   113
            Top             =   760
            Width           =   2145
         End
         Begin VB.TextBox txtBIRate 
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
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   112
            ToolTipText     =   "Latest base interest rate"
            Top             =   1800
            Width           =   2055
         End
         Begin MSForms.TextBox txtBudgetYears 
            Height          =   285
            Left            =   1665
            TabIndex        =   130
            Top             =   225
            Width           =   2205
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3889;503"
            SpecialEffect   =   0
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
            Left            =   75
            TabIndex        =   114
            Top             =   1800
            Width           =   2190
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Global Bank Account:"
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
            Left            =   75
            TabIndex        =   90
            Top             =   760
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financial Year End:"
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
            Left            =   75
            TabIndex        =   89
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Opted to Tax:"
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
            Left            =   75
            TabIndex        =   88
            Top             =   1275
            Width           =   945
         End
      End
      Begin VB.Frame fraSetInterestRates 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Set Finance Charge Rate"
         Height          =   2505
         Left            =   5520
         TabIndex        =   66
         Top             =   3915
         Visible         =   0   'False
         Width           =   4140
         Begin VB.TextBox txtDateFrom 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   70
            Top             =   360
            Width           =   1880
         End
         Begin VB.TextBox txtAdditionalRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   69
            Top             =   1240
            Width           =   1880
         End
         Begin VB.TextBox txtBaseRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   68
            Top             =   800
            Width           =   1880
         End
         Begin VB.TextBox txtRateDescription 
            Height          =   285
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   67
            Top             =   1680
            Width           =   1880
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Restricted to Budget:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   143
         Top             =   3150
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Global ID"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7470
         TabIndex        =   111
         Top             =   45
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Budget Year:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   101
         Top             =   2730
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   7980
      TabIndex        =   63
      Top             =   7395
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2955
      Left            =   135
      TabIndex        =   20
      Top             =   4365
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   5212
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
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
      TabCaption(0)   =   "Monthly Payment Dates"
      TabPicture(0)   =   "Global1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAutoSetup(0)"
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Quarterly Payment Dates"
      TabPicture(1)   =   "Global1.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Half Yearly payments"
      TabPicture(2)   =   "Global1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Yearly payments"
      TabPicture(3)   =   "Global1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Additional Payment Dates"
      TabPicture(4)   =   "Global1.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdMthPayDt"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdAutoSetup 
         BackColor       =   &H80000013&
         Caption         =   "Auto Date Fill"
         Enabled         =   0   'False
         Height          =   325
         Index           =   0
         Left            =   -74820
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   64
         Top             =   2560
         Width           =   1455
      End
      Begin VB.CommandButton cmdMthPayDt 
         BackColor       =   &H80000016&
         Caption         =   "Click here to enter additional Payment Sets"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -72975
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Yearly Payment Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -69480
         TabIndex        =   56
         Top             =   480
         Width           =   2535
         Begin VB.ComboBox cboM7 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD7 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Half Yearly Payment Dates"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -71160
         TabIndex        =   49
         Top             =   420
         Width           =   3135
         Begin VB.ComboBox cboD5 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM5 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD6 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM6 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
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
            Left            =   240
            TabIndex        =   55
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "1st"
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
            Left            =   240
            TabIndex        =   54
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
         Left            =   -68880
         TabIndex        =   43
         Top             =   360
         Width           =   3600
         Begin VB.ComboBox cboDay9 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboDay10 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay11 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay12 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "September:"
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
            Left            =   360
            TabIndex        =   47
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "October:"
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
            Left            =   360
            TabIndex        =   46
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "November:"
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
            Left            =   360
            TabIndex        =   45
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "December:"
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
            Left            =   360
            TabIndex        =   44
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
         Left            =   -71850
         TabIndex        =   39
         Top             =   360
         Width           =   2925
         Begin VB.ComboBox cboDay5 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboDay6 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay7 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay8 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "July:"
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
            Left            =   360
            TabIndex        =   48
            Top             =   1260
            Width           =   300
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "May:"
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
            Left            =   360
            TabIndex        =   42
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "June:"
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
            Left            =   360
            TabIndex        =   41
            Top             =   780
            Width           =   360
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "August:"
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
            Left            =   360
            TabIndex        =   40
            Top             =   1740
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Quarterly Payment Dates"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   1920
         TabIndex        =   26
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboD1 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM1 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD2 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM2 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboD3 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboM3 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboD4 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboM4 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "1st"
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
            Left            =   360
            TabIndex        =   38
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
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
            Left            =   360
            TabIndex        =   37
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "3rd"
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
            Left            =   360
            TabIndex        =   36
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "4th"
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
            Left            =   360
            TabIndex        =   35
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
         Left            =   -74910
         TabIndex        =   21
         Top             =   360
         Width           =   2970
         Begin VB.ComboBox cboDay4 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox cboDay3 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboDay2 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboDay1 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "April:"
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
            Left            =   360
            TabIndex        =   25
            Top             =   1740
            Width           =   375
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "March:"
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
            Left            =   360
            TabIndex        =   24
            Top             =   1260
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "February:"
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
            Left            =   360
            TabIndex        =   23
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "January:"
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
            Left            =   360
            TabIndex        =   22
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdDemandTypes 
      Caption         =   "&Demand/Charge Types"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   8730
      TabIndex        =   19
      Top             =   7650
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   3930
      TabIndex        =   17
      Top             =   7395
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   2025
      TabIndex        =   59
      Top             =   7395
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Data"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   120
      TabIndex        =   18
      Top             =   7395
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoProperty 
      Height          =   330
      Left            =   1395
      Top             =   10980
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
      Left            =   1845
      Top             =   9990
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
      Left            =   1890
      Top             =   10575
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
      Left            =   4095
      TabIndex        =   0
      Top             =   90
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
      Left            =   9360
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   5400
      TabIndex        =   125
      Top             =   90
      Width           =   3960
      VariousPropertyBits=   746604571
      Size            =   "6985;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   675
      TabIndex        =   124
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   93
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
      Left            =   210
      TabIndex        =   61
      Top             =   3750
      Width           =   1650
   End
   Begin VB.Label Label2 
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
      Index           =   4
      Left            =   4680
      TabIndex        =   60
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmGlobalx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form Name: frmGlobalx
Option Explicit

Dim MyForm As FRMSIZE

Dim rst As New ADODB.Recordset
Dim SCperSqFoot As Double

Dim bEditGlobalData As Boolean, bFlxInsEdit As Boolean, bNoGD As Boolean
Dim bIsChildBudget As Boolean, iCurFlxInsRow As Integer
Dim iNewEditMainBudget As Integer, iSCChildNewEdit As Integer
Dim iNewEditRCBudget As Integer, iCurFlxRCMainRow As Integer, iCurFlxRCMainCol As Integer
Dim iCurFlxMainRow As Integer, iCurFlxMainCol As Integer
Dim iCurFlxChildRow As Integer, iCurFlxChildCol As Integer
Dim iCurFlxInterestRatesRow As Integer
Dim cInsurance As Currency
Dim sTextBox As String
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   
   configGridFY
   adoConn.Open getConnectionString
           

   szSQL = "SELECT F.FYrID, F.FY_EndDate,FinancialYear,F.FY_Description " & _
           "FROM FinancialYear AS F, Property AS P " & _
           "WHERE P.PropertyID = '" & txtPropertyName.Tag & "' AND " & _
                 "F.ClientID = P.ClientID " & _
           "ORDER BY FY_EndDate DESC;"


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
      ShowMsgInTaskBar "Financial year has not been created.", "Y", "N"
   Else
        rRow = 1
        gridBankCode.Rows = rstRec.RecordCount + 1
        While Not rstRec.EOF
           gridBankCode.TextMatrix(rRow, 0) = ""
           gridBankCode.TextMatrix(rRow, 1) = Trim(rstRec.Fields.Item("FY_EndDate").Value)
           gridBankCode.TextMatrix(rRow, 2) = Trim(rstRec.Fields.Item("FY_Description").Value)
           gridBankCode.RowHeight(rRow) = 240
           rstRec.MoveNext
           'If Not rstRec.EOF Then gridBankCode.AddItem ""
           rRow = rRow + 1
        Wend
        gridBankCode.RowSel = 1
        picBankCode.Visible = True
        gridBankCode.row = 1
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
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
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   
   adoConn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = ""
'           flxClient.TextMatrix(rRow, 2) = ""
'           flxClient.RowHeight(rRow) = 240
'           flxClient.AddItem ""
'           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cboDay1_KeyPress(KeyAscii As Integer)
    FocusControl cmdAutoSetup(0)
End Sub

Private Sub cboTaxBasis_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtLastCompletedTaxReturnDate
    End If
End Sub

Private Sub chkGlobalVat_Click()
    Dim adoConn As New ADODB.Connection
    Dim rsGlobalData As New ADODB.Recordset
    Dim SQLStr As String
    If chkGlobalVat.Value = 1 Then
        cmdVatRate.Enabled = True
        Frame8.Enabled = True
        If txtVatRate.text = "" Then
            adoConn.Open getConnectionString
            SQLStr = "SELECT G.*, F.FinancialYear AS CBY, F.FYrID,F.FY_Description,V.VAT_ID,V.VAT_RATE,V.VAT_CODE " & _
            "FROM ((GlobalData AS G INNER JOIN Property AS P ON G.PropertyID = P.PropertyID) INNER JOIN tlbVatcode V ON G.Vatrate=V.VAT_ID )" & _
                  "LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
            "WHERE G.PropertyID = '" & txtPropertyName.Tag & "';"
            
                rsGlobalData.Open SQLStr, adoConn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData.EOF Then
                        txtVatRate.text = rsGlobalData!VAT_CODE & " / " & rsGlobalData!VAT_RATE
                End If
                rsGlobalData.Close
                Set rsGlobalData = Nothing
            adoConn.Close
            Set adoConn = Nothing
        End If
    Else
        txtVatRate.text = ""
        cmdVatRate.Enabled = False
        Frame8.Enabled = False
    End If
End Sub

Private Sub chkProduceVatReturn_Click()
    If chkProduceVatReturn.Value = 1 Then
        cboTaxBasis.Enabled = True
        txtLastCompletedTaxReturnDate.Enabled = True
        txtTaxInterval.Enabled = True
        txtCurrentTaxPeriod.Enabled = True
        optAgentToSubmit.Enabled = True
        optClienttosubmit.Enabled = True
        Frame9.Enabled = True
        optAgentToSubmit.Enabled = True
        optClienttosubmit.Enabled = True
        optAgentToSubmit.Value = True
        optClienttosubmit.Value = False
    
    
    Else
        cboTaxBasis.Enabled = False
        txtLastCompletedTaxReturnDate.Enabled = False
        txtTaxInterval.Enabled = False
        txtCurrentTaxPeriod.Enabled = False
        optAgentToSubmit.Enabled = False
        optClienttosubmit.Enabled = False
        
        
            Frame9.Enabled = False
            optAgentToSubmit.Enabled = False
            optClienttosubmit.Enabled = False
            optAgentToSubmit.Value = False
            optClienttosubmit.Value = False
    End If

End Sub

Private Sub cmdBankClose_Click()
    picBankCode.Visible = False
    SSTab1.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub cmdBudgetYears_Click()
    sTextBox = "3"
    picBankCode.Left = txtBudgetYears.Left + Frame1.Left + fraOthers.Left
    picBankCode.Top = txtBudgetYears.Top + txtBudgetYears.Height + Frame1.Top + fraOthers.Top + 5
    LoadGridFY
    gridBankCode.SetFocus
    cmdProperty.Enabled = False
    Frame1.Enabled = False
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    SSTab1.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub cmdproperty_Click()
    sTextBox = "2"
    picClient.Left = 3015
    picClient.Top = 70
    picClient.Visible = True
    LoadPropertyList
    SSTab1.Enabled = False
    Frame1.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdVatRate_Click()
   sTextBox = "5"
   picBankCode.Left = txtVatRate.Left + Frame1.Left + fraOthers.Left
   picBankCode.Top = txtVatRate.Top + txtVatRate.Height + Frame1.Top + fraOthers.Top + 5
   LoadVatCode
   FocusControl gridBankCode
End Sub

Private Sub Command2_Click()
    Frame1.Enabled = True
End Sub

Private Sub flxClient_Click()
        Frame1.Enabled = True
        SSTab1.Enabled = True
        Dim Conn As New ADODB.Connection
        Dim adoConn As New ADODB.Connection
        Dim adoRST As New ADODB.Recordset
        Dim szSQL As String
        If sTextBox = "1" Then
                clear_ALL
                chkGlobalVat.Value = 0
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = ""
                txtPropertyName.Tag = ""
                txtPropertyName.text = ""
                txtPropertyName.Tag = ""
                FocusControl cmdProperty
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtDemandDaysB4Due.Locked = True
                Frame1.Enabled = False
'                cmdBudgetYears.Enabled = False
                cmdExpandBankCode.Enabled = False
                cmdFinancialYear.Enabled = False
                txtGlobalBankAccount.Locked = True
                Conn.Open getConnectionString
                chkGlobalVat.Value = LoadVatOption(Conn)
                Conn.Close
                Set Conn = Nothing
                cmdVatRate.Enabled = False
                Call cboProperty_Change
                'Frame1.Enabled = False
'                fraBudgetYear.Top = 720
'                fraBudgetYear.Left = 2575
'                ConfigFlxFinancialYearsView
'                LoadFlxFinancialYearsByClient
'                flxFinancialYears.row = 0
                
'                FocusControl cmdBudgetYears
        End If
        If sTextBox = "3" Then
'                txtBudgetYears.text = flxClient.TextMatrix(flxClient.row, 1)
'                txtBudgetYears.Tag = Trim(flxClient.TextMatrix(flxClient.row, 1))
                cmdExpandBankCode.SetFocus
        End If
        picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub LastCompletedTaxReturnDate_Change()
     TextBoxChangeDate txtLastCompletedTaxReturnDate
End Sub

Private Sub txtCurrentTaxPeriod_Change()
    TextBoxChangeDate txtCurrentTaxPeriod
End Sub

Private Sub txtCurrentTaxPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl optAgentToSubmit
    End If
End Sub

Private Sub txtCurrentTaxPeriod_LostFocus()
    TextBoxFormatDate txtCurrentTaxPeriod
End Sub

Private Sub txtDemandDaysB4Due_Change()
    If IsNumeric(txtDemandDaysB4Due.text) = False Then
        txtDemandDaysB4Due.text = ""
    End If
End Sub

Private Sub txtDemandDaysB4Due_GotFocus()
    SelTxtInCtrl txtDemandDaysB4Due
End Sub

Private Sub txtLastCompletedTaxReturnDate_Change()
        TextBoxChangeDate txtLastCompletedTaxReturnDate
End Sub

Private Sub txtLastCompletedTaxReturnDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtTaxInterval
    End If
End Sub

Private Sub txtLastCompletedTaxReturnDate_LostFocus()
        TextBoxFormatDate txtLastCompletedTaxReturnDate
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
          Frame1.Enabled = True
          'Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
           End If
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
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
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
               flxClient.RowHeight(rRow) = 240
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub

Private Sub LoadRestrictedtoBudget(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT isRestrictedtoBudget FROM ShoppingCentre"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        chkRestricktedToBudget.Value = IIf(adoRST.Fields("isRestrictedtoBudget").Value = True, 1, 0)
   End If
   adoRST.Close
End Sub
Private Sub LoadCmbClient(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        txtClientList.text = adoRST.Fields("CLIENTNAME").Value
        txtClientList.Tag = adoRST.Fields("CLIENTID").Value
        adoRST.Close
   Else
        adoRST.Close
   End If
   
End Sub


Private Sub cboProperty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      'frmMMain.fraCmdButton.Enabled = True
      Unload Me
   End If
End Sub

'Private Sub cboVatRate_LostFocus()
'   Dim i, match As Integer
'   match = 0
'   If cboVatRate.text = "" Then Exit Sub
'   For i = 0 To 12
'       If cboVatRate.text = cboVatRate.List(i) Then
'           match = 1
'           Exit For
'       End If
'   Next i
'   If match = 0 Then cboVatRate.text = ""
'End Sub

Private Sub cboProperty_Change()
   bNoGD = False
   If txtPropertyName.Tag = "" Then
        MsgBox "Please select a Property", vbInformation, "Warning."
        FocusControl cmdProperty
        Exit Sub
   End If
   If Not GetData Then
      If (MsgBox("There is no Global Data setup for the property " & txtPropertyName.text & ". Would you like to set this up?", vbQuestion + vbYesNo, "No Global Data") = vbYes) Then
            cmdExpandBankCode.Enabled = True
            'Frame1.Enabled = True
            'Call Edit
            'cmdEdit_Click
            cmdCancel.Enabled = False
            txtDemandDaysB4Due.Locked = False
            'cboFiYrEnd.Locked = False
            'cmdBudgetYears.Enabled = True
            cmdExpandBankCode.Enabled = True
            cmdFinancialYear.Enabled = True
            txtGlobalBankAccount.Locked = False
            'cboVatRate.Locked = False
            cmdVatRate.Enabled = True
            bNoGD = IIf(DemandTypeExist, False, True)
            cmdCancel.Enabled = True
            cmdEdit_Click
            'Frame1.Enabled = True
      End If
   End If
End Sub

Private Function DemandTypeExist() As Boolean
   Dim SQLStr As String
   Dim Conn As New ADODB.Connection

   SQLStr = "SELECT * FROM DemandTypes WHERE PropertyID = '" & txtPropertyName.Tag & "';"

   Conn.Open getConnectionString

   rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   DemandTypeExist = IIf(rst.EOF, False, True)

   rst.Close
   Conn.Close
   Set Conn = Nothing
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

   'cmdAddInterest.Enabled = False
End Sub

Private Sub cmdAutoSetup_Click(Index As Integer)
   Dim DTdate As Date, var

   On Error GoTo ErrorHandler

   var = InputBox("Please type the first payment date of the year. (dd/mm/yyyy)", "Frist Payment Date", "01/01/" & Year(Date))
   If var = "" Then Exit Sub

   DTdate = Format(var, "dd mmmm yyyy")

   SetAddDates DTdate
   FocusControl cmdSave
   Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
      cmdAutoSetup_Click (0)
   End If
End Sub

Private Sub SetAddDates(DTdate As Date)
   cboDay1.text = Format(DTdate, "dd")
   cboDay2.text = Format(DTdate, "dd")
   cboDay3.text = Format(DTdate, "dd")
   cboDay4.text = Format(DTdate, "dd")
   cboDay5.text = Format(DTdate, "dd")
   cboDay6.text = Format(DTdate, "dd")
   cboDay7.text = Format(DTdate, "dd")
   cboDay8.text = Format(DTdate, "dd")
   cboDay9.text = Format(DTdate, "dd")
   cboDay10.text = Format(DTdate, "dd")
   cboDay11.text = Format(DTdate, "dd")
   cboDay12.text = Format(DTdate, "dd")

   cboD1.text = Format(DTdate, "dd")
   cboD2.text = Format(DTdate, "dd")
   cboD3.text = Format(DTdate, "dd")
   cboD4.text = Format(DTdate, "dd")
   cboD5.text = Format(DTdate, "dd")
   cboD6.text = Format(DTdate, "dd")
   cboD7.text = Format(DTdate, "dd")

'Quarterly
   cboM1.text = Format(DTdate, "mmmm")
   cboM2.text = Format(DateAdd("m", 3, DTdate), "mmmm")
   cboM3.text = Format(DateAdd("m", 6, DTdate), "mmmm")
   cboM4.text = Format(DateAdd("m", 9, DTdate), "mmmm")

'Half yearly
   cboM5.text = Format(DTdate, "mmmm")
   cboM6.text = Format(DateAdd("m", 6, DTdate), "mmmm")

'Yearly
   cboM7.text = Format(DTdate, "mmmm")
End Sub

Private Sub cmdCancel_Click()
   Call GetData
   Call DisableBoxes
   
End Sub

Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadflxClient
    SSTab1.Enabled = False
    Frame1.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClose_Click()
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub cmdCloseInterest_Click()
   Dim Conn As New ADODB.Connection

   Conn.Open getConnectionString

   txtBIRate.text = LastInterestRate(Conn)

   fraInterestRates.Visible = False

   Conn.Close
   Set Conn = Nothing
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

      Dim Conn As New ADODB.Connection

      Conn.Open getConnectionString

      szSQL = "UPDATE InterestRates " & _
              "SET Active = False " & _
              "WHERE RateID = " & flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 0) & ";"

      Conn.Execute szSQL
      Conn.Close
      Set Conn = Nothing
   End If
End Sub

Private Sub cmdDemandTypes_Click()
'   If txtPropertyName.Text = "" Then
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
'         frmDemandTypes.szPropertyID = txtPropertyName.tag
'         Load frmDemandTypes
'         frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & txtPropertyName.Text
'         cmdDemandTypes.Enabled = False
'         frmDemandTypes.Show
'         frmDemandTypes.SetFocus
'      End If
'   Else
'      frmDemandTypes.szPropertyID = txtPropertyName.tag
'      Load frmDemandTypes
'      frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & txtPropertyName.Text
'      cmdDemandTypes.Enabled = False
'      frmDemandTypes.Show
'      frmDemandTypes.SetFocus
'   End If
End Sub

Private Sub LoadSelProp()
   Dim Conn As New ADODB.Connection
   Dim TotalRow As Integer, TotalCol As Integer, i As Integer, j As Integer
   Dim Data() As String
   Dim SQLStr As String

   SQLStr = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "WHERE ClientID IN " & _
               "(SELECT ClientID " & _
               " FROM Property " & _
               " WHERE PropertyID = '" & txtPropertyName.Tag & "') AND " & _
            " PropertyID IN " & _
               "(SELECT PropertyID " & _
               " FROM DemandTypes " & _
               " GROUP BY PropertyID) " & _
            "ORDER BY PropertyID;"
'Debug.Print SQLStr

   Conn.Open getConnectionString
   
   rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If rst.EOF Then GoTo NoRes

   TotalRow = rst.RecordCount
   TotalCol = rst.Fields.Count

   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(rst.Fields(j).Value), "", rst.Fields(j).Value)
      Next j
      rst.MoveNext
      If rst.EOF Then Exit For
   Next i
   cboSelProp.Column() = Data()
   cboSelProp.ListIndex = 0
   
NoRes:
   rst.Close
   Conn.Close
   Set Conn = Nothing
End Sub

Private Sub cmdEdit_Click()
    If txtPropertyName.text = "" Then
        MsgBox "Please select a Property", vbInformation, "Warning"
        FocusControl cmdProperty
        Exit Sub
    End If
    fraBudgetYear.Visible = False
    txtDemandDaysB4Due.Locked = False
    'cboFiYrEnd.Locked = False
'    cmdBudgetYears.Enabled = True
    cmdExpandBankCode.Enabled = True
    cmdFinancialYear.Enabled = True
    txtGlobalBankAccount.Locked = False
    'cboVatRate.Locked = False
    cmdVatRate.Enabled = True
    fraOthers.Enabled = True
    Call Edit
    bEditGlobalData = True
    
'    FocusControl cmdBudgetYears
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
Private Sub clear_ALL()
    txtBudgetYears.text = ""
    txtBudgetYears.Tag = ""
    txtGlobalBankAccount.text = ""
    txtVatRate.text = ""
    txtVatRate.Tag = ""
    txtBIRate.text = ""
    cboDay1.ListIndex = -1
    cboDay2.ListIndex = -1
    cboDay3.ListIndex = -1
    cboDay4.ListIndex = -1
    cboDay5.ListIndex = -1
    cboDay6.ListIndex = -1
    cboDay7.ListIndex = -1
    cboDay8.ListIndex = -1
    cboDay9.ListIndex = -1
    cboDay10.ListIndex = -1
    cboDay11.ListIndex = -1
    cboDay12.ListIndex = -1
    
    
    cboD1.ListIndex = -1
    cboD2.ListIndex = -1
    cboD3.ListIndex = -1
    cboD4.ListIndex = -1
    cboD5.ListIndex = -1
    cboD6.ListIndex = -1
    cboD7.ListIndex = -1
    
    cboM1.ListIndex = -1
    cboM2.ListIndex = -1
    cboM3.ListIndex = -1
    cboM4.ListIndex = -1
    cboM5.ListIndex = -1
    cboM6.ListIndex = -1
    cboM7.ListIndex = -1
    
    cmdFinancialYear.Caption = ""
    txtDemandDaysB4Due.text = ""
End Sub
Private Sub cmdExpandBankCode_Click()
   sTextBox = "4"
   picBankCode.Left = txtGlobalBankAccount.Left + Frame1.Left + fraOthers.Left
   picBankCode.Top = txtGlobalBankAccount.Top + txtGlobalBankAccount.Height + Frame1.Top + fraOthers.Top + 5
   LoadBankAccount
End Sub

Private Sub cmdExpandBankCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then gridBankCode.Visible = False
End Sub

Private Sub cmdFinancialYear_Click()
   fraBudgetYear.Caption = "Budget Year:"
   cmdFYSet.Visible = True
   fraBudgetYear.Top = Frame1.Top + cmdFinancialYear.Top
   fraBudgetYear.Left = 540
   ConfigFlxFinancialYears
   LoadFlxFinancialYears
   flxFinancialYears.row = 0
End Sub

Private Sub cmdFYSet_Click()
   If flxFinancialYears.TextMatrix(flxFinancialYears.row, 0) = "" Then Exit Sub
   If flxFinancialYears.TextMatrix(flxFinancialYears.row, 1) = cmdFinancialYear.Caption Then Exit Sub
   
   If MsgBox("Do you wish to change the budget year?", vbQuestion + vbYesNo, "Current Budget Year") = vbNo Then Exit Sub

   cmdFinancialYear.Caption = flxFinancialYears.TextMatrix(flxFinancialYears.row, 1)
   Label4(1).Caption = flxFinancialYears.TextMatrix(flxFinancialYears.row, 0)

   Dim adoConn As New ADODB.Connection
   
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE Property AS P " & _
                   "SET P.CBY = '" & flxFinancialYears.TextMatrix(flxFinancialYears.row, 0) & "' " & _
                   "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"

   DisplayBudgetYears adoConn 'this is a display function on this form, after update is done.
    'issue 332 Lease details service charge not updating fixed by anol 20170310
   Update_SC_Lease adoConn
   adoConn.Close
   Set adoConn = Nothing

   fraBudgetYear.Visible = False
End Sub
Private Sub Update_SC_Lease(adoConn As ADODB.Connection)
   Dim rst     As New ADODB.Recordset
   Dim adoRST  As New ADODB.Recordset
   Dim szSQL   As String
   
   
On Error GoTo ErrHandler:

'Resolved By BOSL.
'Modified By Asif. Issue: 0000519. Date: 04-Jan-2015
'Updating the service charge budgets through SQL rather than iteration which is time consuming.

' Charging Method: 2
szSQL = "UPDATE LServiceCharges " & _
"SET " & _
"LServiceCharges.SCTotal = 0, " & _
"LServiceCharges.SCAmount = 0 " & _
"Where " & _
"LServiceCharges.ChargingMethod = 2;"

adoConn.Execute szSQL

szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
"Units AS U, Frequencies AS F,  Property AS P " & _
"SET " & _
"LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
"LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
"Where " & _
"cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
"L.LeaseID = LServiceCharges.LeaseID AND " & _
"LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
"P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
"LServiceCharges.ChargingMethod = 2;"

adoConn.Execute szSQL

' Charging Method: 4

szSQL = "UPDATE LServiceCharges " & _
"SET " & _
"LServiceCharges.SCTotal = 0, " & _
"LServiceCharges.SCAmount = 0, " & _
"LServiceCharges.CMFigure = 0 " & _
"Where " & _
"LServiceCharges.ChargingMethod = 4;"

adoConn.Execute szSQL

szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
"Units AS U, Frequencies AS F,  Property AS P " & _
"SET " & _
"LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
"LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
"LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
"Where " & _
"cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
"L.LeaseID = LServiceCharges.LeaseID AND " & _
"LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
"P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
"LServiceCharges.ChargingMethod = 4;"

adoConn.Execute szSQL

Exit Sub
ErrHandler:
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Could not update Service Charge Budget"
End Sub
Private Sub DisplayBudgetYears(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim rstSQL  As New ADODB.Recordset

   szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID " & _
           "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
           "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
'Debug.Print szSQL
'   szSQL
   rstSQL.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstSQL.EOF Then
      cmdFinancialYear.Caption = IIf(IsNull(rstSQL.Fields.Item("CBY").Value), "", rstSQL.Fields.Item("CBY").Value)
      Label4(1).Caption = IIf(IsNull(rstSQL.Fields.Item("FYrID").Value), "", rstSQL.Fields.Item("FYrID").Value) 'P.CBY = F.FYrID, so they are same
      cmdYearlyInsurance.Caption = Format(TotalInsuranceProperty(txtPropertyName.Tag, adoConn), "£0.00")
      cmdYearlyRent.Caption = Format(TotalRCProperty(txtPropertyName.Tag, adoConn), "£0.00")
      cmdYearlyService.Caption = Format(TotalSCTProperty(txtPropertyName.Tag, adoConn), "£0.00")
   End If

   rstSQL.Close
   Set rstSQL = Nothing
End Sub

Private Sub LoadFlxFinancialYears()
   Dim szSQL   As String
   Dim i       As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset

   adoConn.Open getConnectionString

   szSQL = "SELECT F.FYrID, F.FinancialYear, F.FY_Description, F.PeriodsCount, F.Freq " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.Status AND P.ClientID = F.ClientID AND " & _
                  "P.PropertyID = '" & txtPropertyName.Tag & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   While Not adoRST.EOF
      flxFinancialYears.TextMatrix(i, 0) = adoRST.Fields.Item(0).Value
      flxFinancialYears.TextMatrix(i, 1) = adoRST.Fields.Item(1).Value
      flxFinancialYears.TextMatrix(i, 2) = adoRST.Fields.Item(2).Value
      flxFinancialYears.TextMatrix(i, 3) = adoRST.Fields.Item(3).Value
      flxFinancialYears.TextMatrix(i, 4) = adoRST.Fields.Item(4).Value
      adoRST.MoveNext
      If Not adoRST.EOF Then flxFinancialYears.AddItem ""
      i = i + 1
   Wend

   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Sub LoadFlxFinancialYearsByClient()
   Dim szSQL   As String
   Dim i       As Integer
   Dim K       As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset

   adoConn.Open getConnectionString

   szSQL = "SELECT F.FYrID, F.FinancialYear, F.FY_Description, F.PeriodsCount, F.Freq,setascurrent " & _
           "FROM   FinancialYear AS F " & _
           "WHERE  F.Status AND " & _
                  "F.ClientID = '" & txtClientList.Tag & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   While Not adoRST.EOF
      flxFinancialYears.TextMatrix(i, 0) = adoRST.Fields.Item(0).Value
      flxFinancialYears.TextMatrix(i, 1) = adoRST.Fields.Item(1).Value
      flxFinancialYears.TextMatrix(i, 2) = adoRST.Fields.Item(2).Value
      flxFinancialYears.TextMatrix(i, 3) = adoRST.Fields.Item(3).Value
      flxFinancialYears.TextMatrix(i, 4) = adoRST.Fields.Item(4).Value
      flxFinancialYears.TextMatrix(i, 5) = adoRST.Fields.Item(5).Value
      If flxFinancialYears.TextMatrix(i, 5) = True Then
                 flxFinancialYears.row = i
                 For K = 1 To 5
                        flxFinancialYears.col = K
                        flxFinancialYears.CellFontBold = True
                 Next
      Else
                flxFinancialYears.row = i
                For K = 1 To 5
                        flxFinancialYears.col = K
                        flxFinancialYears.CellFontBold = False
                Next
      End If
      adoRST.MoveNext
      If Not adoRST.EOF Then flxFinancialYears.AddItem ""
      i = i + 1
   Wend

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub ConfigFlxFinancialYears()
   Dim szFlxHeader As String

   flxFinancialYears.RowHeight(0) = 0
   flxFinancialYears.Clear
   flxFinancialYears.Cols = 6
   flxFinancialYears.Rows = 2
   szFlxHeader$ = "<FYrID|<FinancialYear|<FY_Description|>PeriodsCount|<Freq"
   flxFinancialYears.FormatString = szFlxHeader$

   flxFinancialYears.ColWidth(0) = 0
   flxFinancialYears.ColWidth(1) = Label7(1).Left - Label7(0).Left
   flxFinancialYears.ColWidth(2) = Label7(2).Left - Label7(1).Left
   flxFinancialYears.ColWidth(3) = Label7(3).Left - Label7(2).Left
   flxFinancialYears.ColWidth(4) = flxFinancialYears.Width - flxFinancialYears.Left - Label7(3).Left
   flxFinancialYears.ColWidth(5) = 0

  
   fraBudgetYear.Visible = True
End Sub
Private Sub ConfigFlxFinancialYearsView()
   Dim szFlxHeader As String

   flxFinancialYears.RowHeight(0) = 0
   flxFinancialYears.Clear
   flxFinancialYears.Cols = 6
   flxFinancialYears.Rows = 2
   szFlxHeader$ = "<FYrID|<FinancialYear|<FY_Description|>PeriodsCount|<Freq"
   flxFinancialYears.FormatString = szFlxHeader$

   flxFinancialYears.ColWidth(0) = 0
   flxFinancialYears.ColWidth(1) = Label7(1).Left - Label7(0).Left
   flxFinancialYears.ColWidth(2) = Label7(2).Left - Label7(1).Left - 200
   flxFinancialYears.ColWidth(3) = 600
   flxFinancialYears.ColWidth(4) = 1100
   flxFinancialYears.ColWidth(5) = 0
  
   fraBudgetYear.Visible = True
End Sub


Private Sub cmdFYClose_Click()
   fraBudgetYear.Visible = False
End Sub

Private Sub cmdInterestRates_Click()
'   ConfigureFlxInterestRates
'
'   Dim Conn As New ADODB.Connection
'   Conn.Open getConnectionString
'
'   Rst.Open "SELECT * FROM GlobalData WHERE PropertyID = '" & txtPropertyName.Tag & "';", Conn, adOpenDynamic, adLockOptimistic
'   If Rst.EOF Then
'      MsgBox "Until you save the Property global data, you cannot interest rates.", vbInformation + vbOKOnly, "Global Data"
'      cmdSave.SetFocus
'      Rst.Close
'      Conn.Close
'      Set Conn = Nothing
'      Exit Sub
'   End If
'   Rst.Close
'   LoadFlxInterestRates Conn
'
'   Conn.Close
'   Set Conn = Nothing
'
'   fraInterestRates.Left = 120
'   fraInterestRates.Top = 240
'   fraInterestRates.Visible = True
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
'            Rst.Fields.Item("PropertyID").Value = txtPropertyName.tag
'            Rst.Update
'            .MoveNext
'         End With
'      Wend
'   End If
'
'   frmDemandTypes.szPropertyID = txtPropertyName.tag
'   Load frmDemandTypes
'   frmDemandTypes.Caption = frmDemandTypes.Caption & " - " & txtPropertyName.Text
'   cmdDemandTypes.Enabled = False
'   frmDemandTypes.Show
'   frmDemandTypes.SetFocus
'
'   fraSelProp.Left = 9720
'   Conn.Close
End Sub

Private Sub cmdSave_Click()
   Dim tempdate As String
   If txtVatRate.Tag = "" And chkGlobalVat.Value = 1 Then
        MsgBox "Please select a VAT code from the list", vbInformation, "Warning "
        FocusControl cmdVatRate
        Exit Sub
   End If
        

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
      MsgBox "Please enter the number of days demand to send before due.", vbCritical + vbOKOnly, "Demand Notice Period"
      txtDemandDaysB4Due.SetFocus
      Exit Sub
   End If

   If txtGlobalBankAccount.text = "" Then
        MsgBox "Please select the Global Bank Account Number.", vbCritical + vbOKOnly, "Bank Account"
        txtGlobalBankAccount.SetFocus
        Exit Sub
   End If

   Dim conn3 As New ADODB.Connection

'* Save records in the database
   conn3.Open getConnectionString
'   conn3.Execute "Update GlobalData set vatOptionEnabled=" & chkGlobalVat.Value & " WHERE PropertyID = '" & txtPropertyName.Tag & "'"
'   conn3.Execute "Update ShoppingCentre set isRestrictedtoBudget=" & chkRestricktedToBudget.Value & ""
'    conn3.Execute "Update GlobalData set chkProduceVatReturn ='" & chkProduceVatReturn.Value & "'" & _
'                        "Where PropertyID='" & txtPropertyName.Tag & "'"
   If chkProduceVatReturn.Value = 1 Then
    If IsDate(txtLastCompletedTaxReturnDate) = False Then
        MsgBox "Please enter Last Completed TaxReturn Date", vbInformation, "Warning"
    End If
    If IsDate(txtCurrentTaxPeriod) = False Then
        MsgBox "Please enter CurrentTaxPeriod", vbInformation, "Warning"
    End If
                If IsDate(txtLastCompletedTaxReturnDate) = True And IsDate(txtCurrentTaxPeriod) = True Then
                          conn3.Execute "Update GlobalData set TaxBasis='" & cboTaxBasis.Value & "', LastCompletedTaxReturnDate=#" & Format(txtLastCompletedTaxReturnDate.text, "dd MMM yyyy") & "#, " & _
                        "TaxInterval=" & Val(txtTaxInterval.text) & "," & _
                        "CurrentTaxPeriod = '" & Format(txtCurrentTaxPeriod, "dd MMM yyyy") & "', isAgentToSubmit =" & optAgentToSubmit.Value & " " & _
                        "Where PropertyID='" & txtPropertyName.Tag & "'"
                        
                        
                       
                  End If
 Else
     conn3.Execute "Update GlobalData set TaxBasis='', LastCompletedTaxReturnDate=null, " & _
                  "TaxInterval=null," & _
                  "CurrentTaxPeriod = null, isAgentToSubmit =false " & _
                  "Where PropertyID='" & txtPropertyName.Tag & "'"
      End If
                    
                    
   
   rst.Open "SELECT * " & _
            "FROM GlobalData " & _
            "WHERE PropertyID = '" & txtPropertyName.Tag & "' ", _
                    conn3, adOpenDynamic, adLockOptimistic
   If rst.EOF Then
       rst.AddNew
   Else
       rst.MoveFirst
   End If

   If txtPropertyName.text <> "" Then
       rst!propertyID = txtPropertyName.Tag
   Else
       MsgBox "Please select a property to continue", vbInformation, "Save global data"
       rst.Close
       conn3.Close
       Exit Sub
   End If

   'Insurance has already saved, but if the insurance
   If cInsurance <> CCur(cmdYearlyInsurance.Caption) Then
      ChangeInsuraceAmount CCur(cmdYearlyInsurance.Caption), txtPropertyName.Tag, conn3
   End If

   If txtBudgetYears.text <> "" Then rst!SCYearEnd = txtBudgetYears.text

   If txtGlobalBankAccount.text <> "" Then rst!GlobalBankCode = txtGlobalBankAccount.text
   If chkGlobalVat.Value = 1 Then
        rst!VatRate = txtVatRate.Tag
   Else
        rst!VatRate = -1
   End If
' If Not IsNull(Rst!isAgentToSubmit) Then
'        optAgentToSubmit.Value = Rst!isAgentToSubmit
'   Else
'        optAgentToSubmit.Value = True
'  End If
   If rst!MonthlyDueDate1 <> cboDay1.text & " January" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate1, cboDay1.text & " January"
   rst!MonthlyDueDate1 = cboDay1.text & " January"
   If rst!MonthlyDueDate2 <> cboDay2.text & " February" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate2, cboDay2.text & " February"
   rst!MonthlyDueDate2 = cboDay2.text & " February"
   If rst!MonthlyDueDate3 <> cboDay3.text & " March" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate3, cboDay3.text & " March"
   rst!MonthlyDueDate3 = cboDay3.text & " March"
   If rst!MonthlyDueDate4 <> cboDay4.text & " April" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate4, cboDay4.text & " April"
   rst!MonthlyDueDate4 = cboDay4.text & " April"
   If rst!MonthlyDueDate5 <> cboDay5.text & " May" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate5, cboDay5.text & " May"
   rst!MonthlyDueDate5 = cboDay5.text & " May"
   If rst!MonthlyDueDate6 <> cboDay6.text & " June" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate6, cboDay6.text & " June"
   rst!MonthlyDueDate6 = cboDay6.text & " June"
   If rst!MonthlyDueDate7 <> cboDay7.text & " July" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate7, cboDay7.text & " July"
   rst!MonthlyDueDate7 = cboDay7.text & " July"
   If rst!MonthlyDueDate8 <> cboDay8.text & " August" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate8, cboDay8.text & " August"
   rst!MonthlyDueDate8 = cboDay8.text & " August"
   If rst!MonthlyDueDate9 <> cboDay9.text & " September" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate9, cboDay9.text & " September"
   rst!MonthlyDueDate9 = cboDay9.text & " September"
   If rst!MonthlyDueDate10 <> cboDay10.text & " October" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate10, cboDay10.text & " October"
   rst!MonthlyDueDate10 = cboDay10.text & " October"
   If rst!MonthlyDueDate11 <> cboDay11.text & " November" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate11, cboDay11.text & " November"
   rst!MonthlyDueDate11 = cboDay11.text & " November"
   If rst!MonthlyDueDate12 <> cboDay12.text & " December" Then UpdateLeasePaymentDate conn3, 5, rst!MonthlyDueDate12, cboDay12.text & " December"
   rst!MonthlyDueDate12 = cboDay12.text & " December"

   If rst!QuarterlyDueDate1 <> cboD1.text & " " & cboM1.text Then UpdateLeasePaymentDate conn3, 7, rst!QuarterlyDueDate1, cboD1.text & " " & cboM1.text
   rst!QuarterlyDueDate1 = cboD1.text & " " & cboM1.text
   If rst!QuarterlyDueDate2 <> cboD2.text & " " & cboM2.text Then UpdateLeasePaymentDate conn3, 7, rst!QuarterlyDueDate2, cboD2.text & " " & cboM2.text
   rst!QuarterlyDueDate2 = cboD2.text & " " & cboM2.text
   If rst!QuarterlyDueDate3 <> cboD3.text & " " & cboM3.text Then UpdateLeasePaymentDate conn3, 7, rst!QuarterlyDueDate3, cboD3.text & " " & cboM3.text
   rst!QuarterlyDueDate3 = cboD3.text & " " & cboM3.text
   If rst!QuarterlyDueDate4 <> cboD4.text & " " & cboM4.text Then UpdateLeasePaymentDate conn3, 7, rst!QuarterlyDueDate4, cboD4.text & " " & cboM4.text
   rst!QuarterlyDueDate4 = cboD4.text & " " & cboM4.text

   If rst!HalfYearlyDueDate1 <> cboD5.text & " " & cboM5.text Then UpdateLeasePaymentDate conn3, 9, rst!HalfYearlyDueDate1, cboD5.text & " " & cboM5.text
   rst!HalfYearlyDueDate1 = cboD5.text & " " & cboM5.text
   If rst!HalfYearlyDueDate2 <> cboD6.text & " " & cboM6.text Then UpdateLeasePaymentDate conn3, 9, rst!HalfYearlyDueDate2, cboD6.text & " " & cboM6.text
   rst!HalfYearlyDueDate2 = cboD6.text & " " & cboM6.text

   If rst!YearlyDueDate <> cboD7.text & " " & cboM7.text Then UpdateLeasePaymentDate conn3, 11, rst!YearlyDueDate, cboD7.text & " " & cboM7.text
   rst!YearlyDueDate = cboD7.text & " " & cboM7.text

   rst!NoOfDaysToSendDemandsB4Due = CInt(IIf(txtDemandDaysB4Due.text = "", 0, txtDemandDaysB4Due.text))
   rst.Update

   rst.Close
    conn3.Execute "Update GlobalData set vatOptionEnabled=" & chkGlobalVat.Value & " WHERE PropertyID = '" & txtPropertyName.Tag & "'"
   conn3.Execute "Update ShoppingCentre set isRestrictedtoBudget=" & chkRestricktedToBudget.Value & ""
    conn3.Execute "Update GlobalData set chkProduceVatReturn ='" & chkProduceVatReturn.Value & "'" & _
                        "Where PropertyID='" & txtPropertyName.Tag & "'"
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
                           "SET    SCNextDueDate = #" & Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!SCNextDueDate, "yyyy") & "# " & _
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
                           "SET    BRNextDueDate = #" & Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!BRNextDueDate, "yyyy") & "# " & _
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
                           "SET    InsuranceNextDueDate = #" & Format(dtNewDueDate, "dd mmmm") & _
                                                               " " & _
                                                               Format(adoRec!InsuranceNextDueDate, "yyyy") & "# " & _
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
                     "Units.PropertyID = '" & txtPropertyName.Tag & "'", _
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
   Dim SQLStr As String

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

   Dim Conn As New ADODB.Connection
   Dim adoGD As New ADODB.Recordset

   Conn.Open getConnectionString

   If Label3(18).Caption = "ADDING" Then
      SQLStr = "SELECT * FROM InterestRates;"
   Else
      SQLStr = "SELECT * FROM InterestRates WHERE RateID = " & flxInterestRates.TextMatrix(iCurFlxInterestRatesRow, 0) & ";"
   End If
   rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

   If Label3(18).Caption = "ADDING" Then
      rst.AddNew
      adoGD.Open "SELECT BaseInterestRate FROM GlobalData WHERE PropertyID = '" & txtPropertyName.Tag & "';", Conn, adOpenDynamic, adLockOptimistic

      adoGD!BaseInterestRate = rst.RecordCount + 1
      adoGD.Update
      adoGD.Close
      Set adoGD = Nothing
   End If

   With rst
      !propertyID = txtPropertyName.Tag
      !dateFrom = CDate(txtDateFrom.text)
      !BaseRate = CSng(txtBaseRate.text)
      !AdditionalRate = CSng(txtAdditionalRate.text)
      !Active = True
      !RateDescription = IIf(IsNull(txtRateDescription.text), "", txtRateDescription.text)
      .Update
   End With

   MsgBox "Data has been updated.", vbInformation + vbOKOnly
   rst.Close
'   Set Rst = Nothing

   LoadFlxInterestRates Conn

   Conn.Close
   Set Conn = Nothing

   cmdSetIntRateClose_Click
End Sub

Private Sub cmdYearlyInsurance_Click()
   frmRentBudget.sModule = "IB"              'Insurance Budget
   Load frmRentBudget
   frmRentBudget.Caption = "Insurance " + frmRentBudget.Caption
   frmRentBudget.Show
   Me.Enabled = False
End Sub

Private Sub cmdYearlyRent_Click()
   frmRentBudget.sModule = "RB"              'Rent budget
   Load frmRentBudget
   frmRentBudget.Caption = "Rent " + frmRentBudget.Caption
   frmRentBudget.Show
   Me.Enabled = False
End Sub

Private Sub cmdYearlyService_Click()
   Load frmServiceCharge
   frmServiceCharge.Show
   Me.Enabled = False
End Sub



Private Sub flxInterestRates_Click()
   iCurFlxInterestRatesRow = flxInterestRates.row
End Sub

Private Sub HighLightRow(flxGrid As MSHFlexGrid)
   Dim iCol As Integer, iRow As Integer
   Dim iCurCol As Integer, iCurRow As Integer

'   Saving current row and column selection
   iCurRow = flxGrid.row
   iCurCol = flxGrid.col

'  Clear any privious selection
   For iRow = 1 To flxGrid.Rows - 1
      For iCol = 2 To flxGrid.Cols - 1
         flxGrid.col = iCol
         flxGrid.row = iRow
         flxGrid.CellBackColor = RGB(255, 255, 255)
      Next iCol
   Next iRow

'  set the cellback color depends on current selection
   flxGrid.row = iCurRow
   For iCol = 2 To flxGrid.Cols - 1
      flxGrid.col = iCol
      flxGrid.CellBackColor = RGB(244, 244, 244)
   Next iCol

'  Set back original row and col selection, bacause setting the colback color it was changed
   flxGrid.col = iCurCol
End Sub

Private Sub Form_Load()
    MyForm.Height = 7965 'Me.Height ' Remember the current size
    MyForm.Width = 9315 'Me.Width
    Me.Width = 9870
    Me.Height = 8475
    cboTaxBasis.AddItem ""
    cboTaxBasis.AddItem "Accruals Accounting"
    cboTaxBasis.AddItem "Cash Accounting"
    cboTaxBasis.ListIndex = 0
   
    chkProduceVatReturn.BackColor = MODULEBACKCOLOR
    Frame9.BackColor = MODULEBACKCOLOR
    Label2(5).BackColor = MODULEBACKCOLOR
    optAgentToSubmit.BackColor = MODULEBACKCOLOR
    optClienttosubmit.BackColor = MODULEBACKCOLOR
    chkProduceVatReturn.Value = 0
    Frame9.Enabled = False
    optAgentToSubmit.Enabled = False
    optClienttosubmit.Enabled = False
    optAgentToSubmit.Value = False
    optClienttosubmit.Value = False
    
    

'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   'Me.BackColor = MODULEBACKCOLOR
    Me.Caption = "Global Data"
   'Frame1.BackColor = MODULEBACKCOLOR
   'fraOthers.BackColor = Me.BackColor
   'fraDemandDaysB4Due.BackColor = Me.BackColor
'   fraSCBdDtls.BackColor = Me.BackColor
'   Frame8(0).BackColor = Me.BackColor
'   Frame8(1).BackColor = Me.BackColor
  ' SSTab1.BackColor = Me.BackColor
'   Label2(3).BackColor = Me.BackColor
'   Label2(11).BackColor = Me.BackColor
   'MousePointer = vbHourglass
    Frame1.BackColor = MODULEBACKCOLOR
    Frame8.BackColor = MODULEBACKCOLOR
    fraOthers.BackColor = MODULEBACKCOLOR
    fraDemandDaysB4Due.BackColor = MODULEBACKCOLOR
    Frame5.BackColor = MODULEBACKCOLOR
    Frame6.BackColor = MODULEBACKCOLOR
    Frame7.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Frame3.BackColor = MODULEBACKCOLOR
    Frame4.BackColor = MODULEBACKCOLOR
    Me.BackColor = MODULEBACKCOLOR
    Label20.BackColor = MODULEBACKCOLOR
    Label19.BackColor = MODULEBACKCOLOR
    Label18.BackColor = MODULEBACKCOLOR
    Label17.BackColor = MODULEBACKCOLOR
    Label24.BackColor = MODULEBACKCOLOR
    Label23.BackColor = MODULEBACKCOLOR
    Label22.BackColor = MODULEBACKCOLOR
    Label21.BackColor = MODULEBACKCOLOR
    Label28.BackColor = MODULEBACKCOLOR
    Label27.BackColor = MODULEBACKCOLOR
    Label26.BackColor = MODULEBACKCOLOR
    Label25.BackColor = MODULEBACKCOLOR
    Label11.BackColor = MODULEBACKCOLOR
    Label12.BackColor = MODULEBACKCOLOR
    Label13.BackColor = MODULEBACKCOLOR
    Label14.BackColor = MODULEBACKCOLOR
    Label15.BackColor = MODULEBACKCOLOR
    Label16.BackColor = MODULEBACKCOLOR
    chkGlobalVat.BackColor = MODULEBACKCOLOR
    chkRestricktedToBudget.BackColor = MODULEBACKCOLOR

   DisableBoxes

   Dim Conn As New ADODB.Connection

   Conn.Open getConnectionString
   LoadCmbClient Conn
  
    Conn.Execute "Update GlobalData set vatOptionEnabled=0 where isnull(vatOptionEnabled) "
   Conn.Close
   Set Conn = Nothing

'   Call LoadProperty
   Call FillDaysMonths
   'Call GetVATRates

   bEditGlobalData = False
   SSTab1.Tab = 0
'   MousePointer = vbDefault
    Call WheelHook(Me.hWnd)
End Sub
Private Function LoadVatOption(Conn As ADODB.Connection) As Integer
    Dim rsGlobalData As New ADODB.Recordset
    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtPropertyName.Tag & "'", Conn, adOpenStatic, adLockReadOnly
    If Not rsGlobalData.EOF Then
            LoadVatOption = IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value)
    End If
    rsGlobalData.Close
    Set rsGlobalData = Nothing
End Function
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

'Private Sub GetFinancialYearEndList(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT F.FYrID, F.FY_EndDate " & _
'           "FROM FinancialYear AS F, Property AS P " & _
'           "WHERE P.PropertyID = '" & txtPropertyName.Tag & "' AND " & _
'                 "F.ClientID = P.ClientID " & _
'           "ORDER BY FY_EndDate DESC;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'      For j = 0 To TotalCol
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboFiYrEnd.Clear
'   cboFiYrEnd.Column() = Data()
'   cboFiYrEnd.ListIndex = 0
'   adoRst.Close
'
'
'   cboFiYrEnd.ListIndex = 0
''MsgBox cboFiYrEnd.ListIndex
'NoRes:
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

'Private Sub PrepareList(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
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
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
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
'
'   cboClientList.Column() = Data()
'   cboClientList.ListIndex = 0
'   adoRst.Close
'
'NoRes:
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

'Public Sub LoadProperty()
'   Dim sSQLQuery_ As String
'   adoProperty.ConnectionString = getConnectionString
'
'   If txtClientList.tag = "ALL" Then
'      sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
'                   "FROM PROPERTY "
'   Else
'      sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
'                   "FROM PROPERTY " & _
'                   "WHERE PROPERTY.ClientID = '" & txtClientList.tag & "';"
'   End If
'
'   adoProperty.RecordSource = sSQLQuery_
'   adoProperty.CommandType = adCmdText
'   adoProperty.Refresh
'
'   Exit Sub
''   cboProperty.Index 1
'
'   'Resolved by BOSL
'   'issue 0000462: Global Data - Refreshing
'   'Modified by anol 25 Aug 2014
'   Dim Conn As New ADODB.Connection
'    Dim adoRst As New ADODB.Recordset
'   Conn.Open getConnectionString
'   adoRst.Open sSQLQuery_, Conn, adOpenStatic, adLockReadOnly
'   If adoRst.EOF = False Then
'     If Not IsNull(adoRst!PropertyName) Then
'        txtPropertyName.Text = adoRst("PROPERTYNAME").Value
'      End If
'   End If
'End Sub
Private Sub GetFinancialYearEnd(Conn As ADODB.Connection)
    Dim strSQL As String
    Dim rsFinancialYear As New ADODB.Recordset
    Dim rsFinancialYear2 As New ADODB.Recordset
    strSQL = "Select FY_Enddate from FinancialYear where ClientID='" & txtClientList.Tag & "' AND status  and setascurrent=true "
    rsFinancialYear.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    If Not rsFinancialYear.EOF Then
        txtBudgetYears.text = rsFinancialYear("FY_Enddate").Value
    Else
            strSQL = "Select FY_Enddate from FinancialYear where ClientID='" & txtClientList.Tag & "' AND status order by FY_Enddate DESC"
            rsFinancialYear2.Open strSQL, Conn, adOpenStatic, adLockReadOnly
            If Not rsFinancialYear2.EOF Then
                txtBudgetYears.text = rsFinancialYear2("FY_Enddate").Value
            Else
                txtBudgetYears.text = ""
            End If
            rsFinancialYear2.Close
            Set rsFinancialYear2 = Nothing
    End If
    rsFinancialYear.Close
    Set rsFinancialYear = Nothing

End Sub
Public Function GetData() As Boolean
   Dim i As Integer, c As String, bEH As Boolean
   Dim Conn As New ADODB.Connection
   Dim SQLStr As String

   Conn.Open getConnectionString
   GetFinancialYearEnd Conn
   LoadRestrictedtoBudget Conn 'added 2020-11-19 issue 889

   SQLStr = "SELECT G.*, F.FinancialYear AS CBY, F.FYrID,F.FY_Description,V.VAT_ID,V.VAT_RATE,V.VAT_CODE " & _
            "FROM ((GlobalData AS G INNER JOIN Property AS P ON G.PropertyID = P.PropertyID) LEFT JOIN tlbVatcode V ON G.Vatrate=V.VAT_ID )" & _
                  "LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
            "WHERE G.PropertyID = '" & txtPropertyName.Tag & "';"

   rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If rst.RecordCount = 0 Then
       rst.Close
       Conn.Close
       Set Conn = Nothing
       GetData = False
       clear_ALL
       Exit Function
   End If

   If Not IsNull(rst!SCYearEnd) Then
'       txtBudgetYears.Tag = Rst!SCYearEnd
'       txtBudgetYears.text = Rst!SCYearEnd
   End If
   If IsNull(rst!GlobalBankCode) Then txtGlobalBankAccount.text = "" Else txtGlobalBankAccount.text = rst!GlobalBankCode

   If Not IsNull(rst!BaseInterestRate) Or rst!BaseInterestRate <> "" Then
      txtBIRate.text = LastInterestRate(Conn)
   Else
      txtBIRate.text = "0.00"
   End If
   chkProduceVatReturn.Value = IIf((rst!chkProduceVatReturn) = True, 1, 0)
   If Not IsNull(rst!isAgentToSubmit) Then
        optAgentToSubmit.Value = rst!isAgentToSubmit
        optClienttosubmit.Value = Not optAgentToSubmit.Value
   Else
        optAgentToSubmit.Value = True
        optClienttosubmit.Value = Not optAgentToSubmit.Value
  End If
'Update GlobalData set isAgentToSubmit =true
'   For i = 0 To cboVatRate.ListCount - 1
'       c = cboVatRate.List(i)
'       If CInt(Left(c, 2)) = Rst!VatRate Then
'           cboVatRate.text = c
'       End If
'   Next i
    'Condition added by anol 2020-10-07
   If Not IsNull(rst!VAT_CODE) Then
         txtVatRate.text = rst!VAT_CODE & " / " & rst!VAT_RATE
         txtVatRate.Tag = rst!VAT_ID
         'chkGlobalVat.Value = 1
   Else
         txtVatRate.text = ""
         txtVatRate.Tag = ""
         chkGlobalVat.Value = 0
   End If
   
   'Adeed by anol 2021-09-15
'    chkProduceVatReturn.Value = Rst!chkProduceVatReturn
    cboTaxBasis.text = IIf(IsNull(rst!TaxBasis), "", rst!TaxBasis)
    txtLastCompletedTaxReturnDate.text = Format(rst!LastCompletedTaxReturnDate, "dd/MM/yyyy")
    txtTaxInterval.text = IIf(IsNull(rst!TaxInterval), "", rst!TaxInterval) 'Rst!TaxInterval
    txtCurrentTaxPeriod.text = IIf(IsNull(rst!CurrentTaxPeriod), "", rst!CurrentTaxPeriod) ' Rst!CurrentTaxPeriod
    optAgentToSubmit.Value = rst!isAgentToSubmit
'   chkProduceVatReturn
'TaxBasis
'LastCompletedTaxReturnDate
'TaxInterval
'CurrentTaxPeriod
'isAgentToSubmit
   
   If IsNull(rst!NoOfDaysToSendDemandsB4Due) = False Then txtDemandDaysB4Due.text = rst!NoOfDaysToSendDemandsB4Due
   On Error GoTo DateErrorHandler

   bEH = True
   cboDay1.text = Left(rst!MonthlyDueDate1, 2)
   cboDay2.text = Left(rst!MonthlyDueDate2, 2)
   cboDay3.text = Left(rst!MonthlyDueDate3, 2)
   cboDay4.text = Left(rst!MonthlyDueDate4, 2)
   cboDay5.text = Left(rst!MonthlyDueDate5, 2)
   cboDay6.text = Left(rst!MonthlyDueDate6, 2)
   cboDay7.text = Left(rst!MonthlyDueDate7, 2)
   cboDay8.text = Left(rst!MonthlyDueDate8, 2)
   cboDay9.text = Left(rst!MonthlyDueDate9, 2)
   cboDay10.text = Left(rst!MonthlyDueDate10, 2)
   cboDay11.text = Left(rst!MonthlyDueDate11, 2)
   cboDay12.text = Left(rst!MonthlyDueDate12, 2)
   cboD2.text = Left(rst!QuarterlyDueDate2, 2)
   cboM2.text = Right(rst!QuarterlyDueDate2, Len(rst!QuarterlyDueDate2) - 3)
   cboD1.text = Left(rst!QuarterlyDueDate1, 2)
   cboM1.text = Right(rst!QuarterlyDueDate1, Len(rst!QuarterlyDueDate1) - 3)
   cboD3.text = Left(rst!QuarterlyDueDate3, 2)
   cboM3.text = Right(rst!QuarterlyDueDate3, Len(rst!QuarterlyDueDate3) - 3)
   cboD4.text = Left(rst!QuarterlyDueDate4, 2)
   cboM4.text = Right(rst!QuarterlyDueDate4, Len(rst!QuarterlyDueDate4) - 3)
   cboD5.text = Left(rst!HalfYearlyDueDate1, 2)
   cboM5.text = Right(rst!HalfYearlyDueDate1, Len(rst!HalfYearlyDueDate1) - 3)
   cboD6.text = Left(rst!HalfYearlyDueDate2, 2)
   cboM6.text = Right(rst!HalfYearlyDueDate2, Len(rst!HalfYearlyDueDate2) - 3)
   cboD7.text = Left(rst!YearlyDueDate, 2)
   cboM7.text = Right(rst!YearlyDueDate, Len(rst!YearlyDueDate) - 3)
   bEH = False

DateErrorHandler:
   If bEH Then
      MsgBox "Please set a correct date for the default payment date set.", vbCritical + vbOKOnly, "Wrong Date"
   End If

   cmdFinancialYear.Caption = IIf(IsNull(rst!CBY), "", rst!CBY)
   Label4(1).Caption = IIf(IsNull(rst!FYrID), "", rst!FYrID)
   'Resolved by BOSL
    'issue 462 : Global Data - Refreshing
    'Modified by Anol 25 Aug 2014
   cmdYearlyInsurance.Caption = Format(TotalInsuranceProperty(txtPropertyName.Tag, Conn), "0.00")
   'End of modification
   'cmdYearlyInsurance.Caption = Format(TotalInsuranceProperty(txtPropertyName.tag, Conn), "£0.00")
   cmdYearlyRent.Caption = Format(TotalRCProperty(txtPropertyName.Tag, Conn), "£0.00")
   cmdYearlyService.Caption = Format(TotalSCTProperty(txtPropertyName.Tag, Conn), "£0.00")

   cInsurance = CCur(cmdYearlyInsurance.Caption) 'save it to know at the end is there any modification of insurance

   rst.Close

   Conn.Close

'   Set Rst = Nothing
   Set Conn = Nothing
   GetData = True
   Exit Function
ErrH:
'   Set Rst = Nothing
   Set Conn = Nothing
   MsgBox Err.Number & " - " & Err.description, vbOKOnly, "Error"
End Function

Private Function LastInterestRate(Conn As ADODB.Connection) As String
   Dim sqlRST As New ADODB.Recordset
   Dim SQLStr As String

   SQLStr = "SELECT * FROM InterestRates " & _
            "WHERE PropertyID = '" & txtPropertyName.Tag & "' " & _
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

Private Function TotalRCProperty(szPropertyID As String, Conn As ADODB.Connection) As Double
   Dim Rst2 As New ADODB.Recordset
   Dim SQLStr As String

   SQLStr = "SELECT GlobalRC.PropertyID, SUM(GlobalRC.TotalBudget) AS TOTALRENT " & _
            "From GlobalRC " & _
            "WHERE GlobalRC.PropertyID = '" & szPropertyID & "' AND " & _
                  "FinancialYear = '" & Label4(1).Caption & "' " & _
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

Private Function TotalSCTProperty(szPropertyID As String, Conn As ADODB.Connection) As Double
   Dim Rst2 As New ADODB.Recordset
   Dim SQLStr As String

   SQLStr = "SELECT GlobalSC.PropertyID, SUM(GlobalSC.TotalBudget) AS TOTALSC " & _
            "From GlobalSC " & _
            "WHERE GlobalSC.PropertyID = '" & szPropertyID & "' AND " & _
                  "FinancialYear = '" & Label4(1).Caption & "' " & _
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

Private Function TotalInsuranceProperty(szPropertyID As String, Conn As ADODB.Connection) As Double
   Dim Rst2 As New ADODB.Recordset
   Dim SQLStr As String

   SQLStr = "SELECT GlobalInsurance.PropertyID, SUM(GlobalInsurance.Amount) AS TOTALINSURANCE " & _
            "From GlobalInsurance " & _
            "WHERE GlobalInsurance.PropertyID = '" & szPropertyID & "' AND " & _
                  "FinancialYear = '" & Label4(1).Caption & "' " & _
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
   Dim SQLStr As String

   SQLStr = "SELECT * FROM InterestRates " & _
            "WHERE PropertyID = '" & txtPropertyName.Tag & "' AND " & _
               "Active = TRUE;"

   rst.Open SQLStr, Conn1, adOpenStatic, adLockReadOnly

   i = 1
   flxInterestRates.Rows = 2
   If Not rst.EOF Then
      While Not rst.EOF
         flxInterestRates.TextMatrix(i, 0) = rst!RateID
         flxInterestRates.TextMatrix(i, 1) = rst!propertyID
         flxInterestRates.TextMatrix(i, 2) = rst!dateFrom
         flxInterestRates.TextMatrix(i, 3) = Format(rst!BaseRate, "0.0000")
         flxInterestRates.TextMatrix(i, 4) = Format(rst!AdditionalRate, "0.0000")
         flxInterestRates.TextMatrix(i, 5) = IIf(rst!Active = True, "YES", "NO")
         flxInterestRates.TextMatrix(i, 6) = IIf(IsNull(rst!RateDescription), "", rst!RateDescription)

         i = i + 1
         rst.MoveNext
         If Not rst.EOF Then flxInterestRates.AddItem ""
      Wend
   Else
      flxInterestRates.Rows = 1
   End If
   flxInterestRates.row = 0
   flxInterestRates.col = 0

   rst.Close
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
'   cmdYearlyInsurance.Enabled = False
'   cmdYearlyRent.Enabled = False
'   cmdYearlyService.Enabled = False
'   fraOthers.Enabled = False
'   fraDemandDaysB4Due.Enabled = False
   Frame1.Enabled = False

   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
   Frame6.Enabled = False
   Frame7.Enabled = False
   fraOthers.Enabled = False
   
   cmdAutoSetup(0).Enabled = False
   cmdMthPayDt.Enabled = False
   
   cmdEdit.Visible = True
   cmdSave.Visible = False
   cmdCancel.Visible = False
   cmdProperty.Enabled = True
    
   cmdClientList.Enabled = True
   
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
    SSTab1.Enabled = True
    Frame1.Enabled = True
    If sTextBox = "4" Then
        txtGlobalBankAccount.text = gridBankCode.TextMatrix(gridBankCode.row, 1)
        FocusControl cmdVatRate
    ElseIf sTextBox = "3" Then
'        txtBudgetYears.text = gridBankCode.TextMatrix(gridBankCode.row, 1)
'        FocusControl cmdExpandBankCode
    ElseIf sTextBox = "5" Then
        txtVatRate.text = gridBankCode.TextMatrix(gridBankCode.row, 1) & " / " & gridBankCode.TextMatrix(gridBankCode.row, 2)
        txtVatRate.Tag = gridBankCode.TextMatrix(gridBankCode.row, 3)
        FocusControl txtDemandDaysB4Due
    End If
    picBankCode.Visible = False
    
   
End Sub

Private Sub gridBankCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      gridBankCode_Click
   End If
End Sub

Private Sub mnuEdit_Click()
   Call Edit
End Sub

Public Sub Edit()
   Frame1.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
   Frame6.Enabled = True
   Frame7.Enabled = True
   fraOthers.Enabled = True
   Frame8.Enabled = True
   cmdAutoSetup(0).Enabled = True
   cmdMthPayDt.Enabled = True

   cmdEdit.Visible = False
   cmdSave.Visible = True
   cmdCancel.Visible = True
   cmdProperty.Enabled = False
   cmdClientList.Enabled = False
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

'Public Sub GetVATRates()
'   Dim Conn As New ADODB.Connection
'   Dim SQLStr As String
'
'   Conn.Open getConnectionString
'
'   'SQLStr = "SELECT VAT_CODE, VAT_RATE, VAT_RATE_NAME FROM SYS_VAT_FILE ORDER BY VAT_CODE"   'CHANGE TO SageLine50v12
'   SQLStr = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM tlbVATCODE ORDER BY VAT_ID"
'   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
'
'   While Rst.EOF = False
'      cboVatRate.AddItem Rst!VAT_ID & " / " & Rst!VAT_CODE & " / " & Rst!VAT_RATE
'      Rst.MoveNext
'   Wend
'
'   Rst.Close
'   Conn.Close
'   Set Conn = Nothing
'End Sub
Private Sub configGridBankCode()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 4
   gridBankCode.TextMatrix(0, 0) = "Nominal Code"
   gridBankCode.TextMatrix(0, 1) = "Name"
   gridBankCode.ColWidth(0) = 60
   gridBankCode.ColWidth(1) = 1200
   gridBankCode.ColWidth(2) = 2600
   gridBankCode.ColWidth(3) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "Bank Code"
   Label8.Caption = "Bank Name"
   
End Sub
Private Sub configVatCode()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 4
'   gridBankCode.TextMatrix(0, 0) = "Nominal Code"
'   gridBankCode.TextMatrix(0, 1) = "Name"
   gridBankCode.ColWidth(0) = 60
   gridBankCode.ColWidth(1) = 1200
   gridBankCode.ColWidth(2) = 2000
   gridBankCode.ColWidth(3) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "Vat Code"
   Label8.Caption = "Vat Rate"
   
End Sub
Private Sub configGridFY()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 4
'   gridBankCode.TextMatrix(0, 0) = "Nominal Code"
'   gridBankCode.TextMatrix(0, 1) = "Name"
   gridBankCode.ColWidth(0) = 60
   gridBankCode.ColWidth(1) = 1200
   gridBankCode.ColWidth(2) = 2000
   gridBankCode.ColWidth(3) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "F. Year End"
   Label8.Caption = "F. Year Description"
   
End Sub
Private Sub LoadVatCode()
   ' Error Handler
   'On Error GoTo Error_Handler
   configVatCode

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM tlbVATCODE where IN_USE ORDER BY VAT_ID"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup vat Code."
      picBankCode.Visible = False
   Else
      gridBankCode.Rows = adoRST.RecordCount + 1
      rRow = 1
      While Not adoRST.EOF
         gridBankCode.TextMatrix(rRow, 1) = adoRST.Fields.Item("VAT_CODE").Value
         gridBankCode.TextMatrix(rRow, 2) = adoRST.Fields.Item("VAT_RATE").Value
         gridBankCode.TextMatrix(rRow, 3) = adoRST.Fields.Item("VAT_ID").Value
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
       picBankCode.Visible = True
       gridBankCode.row = 1
   End If

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Vat code"

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadBankAccount()
   ' Error Handler
   'On Error GoTo Error_Handler
   configGridBankCode

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = '" & txtClientList.Tag & "' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup bank account for this client : '" & txtClientList.text & "'", vbInformation, "Global bank account"
      picBankCode.Visible = False
   Else
      gridBankCode.Rows = adoRST.RecordCount + 1
      rRow = 1
      While Not adoRST.EOF
         gridBankCode.TextMatrix(rRow, 1) = adoRST.Fields.Item("BNC").Value
         gridBankCode.TextMatrix(rRow, 2) = adoRST.Fields.Item("BNN").Value
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
       picBankCode.Visible = True
       gridBankCode.row = 1
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

Private Sub txtGlobalBankAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdExpandBankCode_Click
   End If
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

        '        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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

Private Sub txtTaxInterval_KeyPress(KeyAscii As Integer)
    'sdss
    If KeyAscii = 13 Then
        FocusControl txtCurrentTaxPeriod
    End If
    DigitTextKeyPress txtTaxInterval, KeyAscii
End Sub

Private Sub txtTaxInterval_LostFocus()
        If Val(txtTaxInterval.text) >= 3 And Val(txtTaxInterval.text) < 12 Then
        txtCurrentTaxPeriod.text = DateAdd("m", txtTaxInterval.text, txtLastCompletedTaxReturnDate.text)
    Else
            txtTaxInterval.text = ""
            MsgBox "Please enter value between 3 and 12", vbInformation, "Warning"
    End If
End Sub
