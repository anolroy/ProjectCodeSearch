VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLeasee1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leaseholder"
   ClientHeight    =   13755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeasee1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13755
   ScaleWidth      =   13500
   Begin VB.PictureBox picPurchaseHistory 
      BackColor       =   &H00E5E5E5&
      Height          =   1605
      Left            =   12600
      ScaleHeight     =   1545
      ScaleWidth      =   3600
      TabIndex        =   321
      Top             =   4140
      Visible         =   0   'False
      Width           =   3660
      Begin VB.CommandButton cmdPrintHistOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   120
         TabIndex        =   316
         Top             =   1110
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrintHistCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1740
         TabIndex        =   317
         Top             =   1095
         Width           =   1200
      End
      Begin VB.TextBox txtStartDateR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   270
         MaxLength       =   80
         TabIndex        =   318
         Top             =   450
         Width           =   1200
      End
      Begin VB.TextBox txtEndDateR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1575
         MaxLength       =   80
         TabIndex        =   319
         Top             =   450
         Width           =   1200
      End
      Begin VB.CommandButton cmdClosePrintHIst 
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
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   322
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   615
         Index           =   21
         Left            =   45
         Top             =   360
         Width           =   3450
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   660
         Index           =   20
         Left            =   45
         Top             =   315
         Width           =   3450
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lessee Account History Report"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   100
         Left            =   90
         TabIndex        =   320
         Top             =   45
         Width           =   2400
      End
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
      Height          =   405
      Left            =   4020
      ScaleHeight     =   405
      ScaleWidth      =   4545
      TabIndex        =   60
      Top             =   2445
      Visible         =   0   'False
      Width           =   4545
      Begin VB.Label lblLoading 
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
         Left            =   225
         TabIndex        =   61
         Top             =   105
         Width           =   3960
      End
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   2220
      Left            =   12555
      TabIndex        =   300
      Top             =   6210
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         Height          =   2100
         Index           =   0
         Left            =   45
         ScaleHeight     =   2040
         ScaleWidth      =   3555
         TabIndex        =   301
         Top             =   50
         Width           =   3615
         Begin VB.CommandButton cmdCloseSearch 
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
            Left            =   3330
            Style           =   1  'Graphical
            TabIndex        =   306
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2025
            MaxLength       =   80
            TabIndex        =   305
            Top             =   1125
            Width           =   1380
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   80
            TabIndex        =   304
            Top             =   1125
            Width           =   1290
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   303
            Top             =   790
            Width           =   2685
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   302
            Top             =   450
            Width           =   2685
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2055
            TabIndex        =   309
            Top             =   1635
            Width           =   1200
         End
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   135
            TabIndex        =   307
            Top             =   1605
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   12
            Left            =   135
            TabIndex        =   308
            Top             =   1125
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   310
            Top             =   810
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   312
            Top             =   495
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Range"
            Height          =   195
            Index           =   9
            Left            =   765
            TabIndex        =   311
            Top             =   45
            Width           =   810
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0FFFF&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   30
            Left            =   0
            Top             =   260
            Width           =   3855
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   55
            Index           =   2
            Left            =   0
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   6
      Left            =   12705
      TabIndex        =   193
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descending Date Order"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   199
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintHistoryCancel 
         Caption         =   "&Cancel"
         Height          =   365
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintHistorySorted 
         Caption         =   "&OK"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ascending Date Order"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   195
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transaction Serial Number"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   194
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1215
         Index           =   18
         Left            =   120
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Sorting Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   -510
         TabIndex        =   198
         Top             =   0
         Width           =   2325
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   660
         Index           =   32
         Left            =   0
         Top             =   0
         Width           =   120
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   660
         Index           =   33
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   19
         Left            =   120
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtUnitNumber 
      Height          =   285
      Left            =   12480
      TabIndex        =   192
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Index           =   0
      Left            =   10380
      ScaleHeight     =   3945
      ScaleWidth      =   4905
      TabIndex        =   140
      Top             =   8835
      Visible         =   0   'False
      Width           =   4935
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
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   278
         Top             =   45
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3345
         Index           =   0
         Left            =   30
         TabIndex        =   276
         Top             =   600
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   5900
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
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Index           =   0
         Left            =   1905
         TabIndex        =   286
         Top             =   60
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
         Left            =   1215
         TabIndex        =   284
         Top             =   60
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
         Left            =   180
         TabIndex        =   282
         Top             =   45
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
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
         Index           =   2
         Left            =   45
         Top             =   45
         Width           =   4500
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   210
         Index           =   0
         Left            =   1890
         TabIndex        =   280
         Top             =   90
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1485
         TabIndex        =   272
         Top             =   315
         Width           =   3195
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5636;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   270
         Top             =   315
         Width           =   1380
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2434;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   2970
         TabIndex        =   274
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox fmeTenantLookup 
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
      Height          =   5295
      Left            =   135
      ScaleHeight     =   5265
      ScaleWidth      =   8880
      TabIndex        =   110
      Top             =   8865
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
         Left            =   8235
         TabIndex        =   113
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
         Left            =   8235
         TabIndex        =   112
         Top             =   45
         Width           =   300
      End
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
         Left            =   8625
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTenantLookup 
         Height          =   3990
         Left            =   30
         TabIndex        =   119
         Top             =   1215
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   7038
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   16777215
         BackColorSel    =   12648447
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Number"
         Height          =   195
         Index           =   2
         Left            =   3375
         TabIndex        =   292
         Top             =   660
         Width           =   900
      End
      Begin MSForms.TextBox txtSearchCompany 
         Height          =   285
         Left            =   3375
         TabIndex        =   116
         Top             =   870
         Width           =   2070
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3651;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyList 
         Height          =   285
         Left            =   855
         TabIndex        =   290
         Tag             =   "ALL"
         Top             =   360
         Width           =   7380
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "13017;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   855
         TabIndex        =   288
         Tag             =   "ALL"
         Top             =   45
         Width           =   7380
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "13017;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   285
         Left            =   7515
         TabIndex        =   118
         Top             =   870
         Width           =   1320
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2328;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchUnitName 
         Height          =   285
         Left            =   5475
         TabIndex        =   117
         Top             =   870
         Width           =   1995
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3519;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance"
         Height          =   195
         Index           =   4
         Left            =   7560
         TabIndex        =   141
         Top             =   660
         Width           =   840
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   122
         Top             =   660
         Width           =   165
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   1095
         TabIndex        =   121
         Top             =   660
         Width           =   405
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   3
         Left            =   5475
         TabIndex        =   120
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   14
         Left            =   45
         TabIndex        =   124
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   13
         Left            =   45
         TabIndex        =   123
         Top             =   45
         Width           =   465
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   30
         Top             =   660
         Width           =   8805
      End
      Begin MSForms.TextBox txtSearchName 
         Height          =   285
         Left            =   1095
         TabIndex        =   115
         Top             =   870
         Width           =   2250
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3969;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchTenant 
         Height          =   285
         Left            =   30
         TabIndex        =   114
         Top             =   870
         Width           =   1035
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1826;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtClientID 
      Height          =   285
      Left            =   12360
      TabIndex        =   106
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fmeTenant 
      Caption         =   "Lessee Information"
      Height          =   2595
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11955
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         TabIndex        =   253
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "C&opy Lessee"
         Height          =   385
         Left            =   4350
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         TabIndex        =   108
         Text            =   "0.00"
         Top             =   1486
         Width           =   3015
      End
      Begin VB.CommandButton cmdDeleteLessee 
         Caption         =   "&Delete Lessee"
         Height          =   385
         Left            =   8445
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Lessee"
         Height          =   385
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel Changes"
         Height          =   385
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Lessee"
         Height          =   385
         Left            =   5715
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Lessee"
         Height          =   385
         Left            =   2985
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   385
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtBalance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F9EE&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "0.00"
         Top             =   1173
         Width           =   3015
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   150
         Width           =   4140
         Begin VB.OptionButton optCurrentTenant 
            Caption         =   "Current"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   60
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optExTenant 
            Caption         =   "Ex-Lessee"
            Height          =   195
            Left            =   1680
            TabIndex        =   9
            Top             =   60
            Width           =   1155
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   195
            Left            =   3480
            TabIndex        =   10
            Top             =   60
            Width           =   825
         End
      End
      Begin VB.TextBox cboSageAccountNumber 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   96
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdTenantLookup 
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
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Share (%):"
         Height          =   195
         Index           =   8
         Left            =   6540
         TabIndex        =   254
         Top             =   1800
         Width           =   690
      End
      Begin MSForms.CommandButton cmdPropLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   191
         Top             =   500
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClientLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   190
         Top             =   150
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdUnitLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   189
         Top             =   860
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   188
         Top             =   840
         Width           =   675
      End
      Begin MSForms.CheckBox chkPrintDmd 
         Height          =   255
         Left            =   400
         TabIndex        =   187
         Top             =   1560
         Width           =   1440
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2540;450"
         Value           =   "1"
         Caption         =   "Print Demands:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkEmailSC 
         Height          =   285
         Left            =   355
         TabIndex        =   186
         Top             =   1815
         Width           =   1485
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2619;503"
         Value           =   "0"
         Caption         =   "Email S. Charge:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCombEmail 
         Height          =   255
         Left            =   4020
         TabIndex        =   159
         Top             =   1815
         Width           =   1560
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2752;450"
         Value           =   "0"
         Caption         =   "Combined Email:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkEmailSt 
         Height          =   255
         Left            =   2100
         TabIndex        =   158
         Top             =   1815
         Width           =   1560
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2752;450"
         Value           =   "0"
         Caption         =   "Email Statement:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkPrintSt 
         Height          =   255
         Left            =   4005
         TabIndex        =   156
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;450"
         Value           =   "0"
         Caption         =   "Print Statements:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkEmailDmd 
         Height          =   255
         Left            =   2220
         TabIndex        =   155
         Top             =   1560
         Width           =   1440
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2540;450"
         Value           =   "0"
         Caption         =   "Email Demands:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label lblLeaseChanged 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Left            =   5640
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Held (£):"
         Height          =   195
         Index           =   7
         Left            =   6540
         TabIndex        =   109
         Top             =   1486
         Width           =   1170
      End
      Begin MSForms.TextBox txtUnit 
         Height          =   315
         Left            =   8040
         TabIndex        =   18
         Top             =   822
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   315
         Left            =   8040
         TabIndex        =   17
         Top             =   471
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClient 
         Height          =   315
         Left            =   8040
         TabIndex        =   16
         Top             =   120
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance (£):"
         Height          =   195
         Index           =   6
         Left            =   6540
         TabIndex        =   15
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   4
         Left            =   6540
         TabIndex        =   14
         Top             =   460
         Width           =   645
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   3
         Left            =   6540
         TabIndex        =   13
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lessee ID:"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   12
         Top             =   1185
         Width           =   720
      End
      Begin MSForms.TextBox txtCompanyName 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   840
         Width           =   3945
         VariousPropertyBits=   746604571
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtName 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   460
         Width           =   3945
         VariousPropertyBits=   747653147
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   450
         Width           =   435
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Index           =   5
         Left            =   6540
         TabIndex        =   1
         Top             =   830
         Width           =   330
      End
      Begin MSForms.TextBox txtTenantID 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   1185
         Width           =   3945
         VariousPropertyBits=   746604575
         BackColor       =   15858158
         MaxLength       =   30
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin TabDlg.SSTab tabTenant 
      Height          =   5940
      Left            =   75
      TabIndex        =   11
      Top             =   2760
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   5
      TabsPerRow      =   8
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
      TabCaption(0)   =   "&Lessee Details"
      TabPicture(0)   =   "frmLeasee1.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeTenantAddress"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Defaults"
      TabPicture(1)   =   "frmLeasee1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(4)"
      Tab(1).Control(1)=   "txtSLControlName"
      Tab(1).Control(2)=   "Shape1(5)"
      Tab(1).Control(3)=   "Label18(0)"
      Tab(1).Control(4)=   "lblVatCode(0)"
      Tab(1).Control(5)=   "txtNominalCodeName"
      Tab(1).Control(6)=   "Label1(0)"
      Tab(1).Control(7)=   "txtNominalCode"
      Tab(1).Control(8)=   "Label18(1)"
      Tab(1).Control(9)=   "cmdSLC"
      Tab(1).Control(10)=   "txtSLControl"
      Tab(1).Control(11)=   "txtCodeVat"
      Tab(1).Control(12)=   "cmdTaxList"
      Tab(1).Control(13)=   "cmdNC"
      Tab(1).Control(14)=   "txtDefault(0)"
      Tab(1).Control(15)=   "cmdCancelDefaults"
      Tab(1).Control(16)=   "cmdEditDefaults"
      Tab(1).Control(17)=   "cmdSaveDefaults"
      Tab(1).Control(18)=   "txtDefault(1)"
      Tab(1).Control(19)=   "txtDefault(2)"
      Tab(1).Control(20)=   "txtDefault(3)"
      Tab(1).Control(21)=   "txtDefault(4)"
      Tab(1).Control(22)=   "txtDefault(5)"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "De&posit Held"
      TabPicture(2)   =   "frmLeasee1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape5"
      Tab(2).Control(1)=   "Label6(0)"
      Tab(2).Control(2)=   "Label6(1)"
      Tab(2).Control(3)=   "Label6(2)"
      Tab(2).Control(4)=   "Label6(3)"
      Tab(2).Control(5)=   "Label6(5)"
      Tab(2).Control(6)=   "Label6(6)"
      Tab(2).Control(7)=   "Label6(7)"
      Tab(2).Control(8)=   "Label6(8)"
      Tab(2).Control(9)=   "Label6(9)"
      Tab(2).Control(10)=   "Label6(4)"
      Tab(2).Control(11)=   "Label6(10)"
      Tab(2).Control(12)=   "flxDeposit"
      Tab(2).Control(13)=   "cmdPrintAccount"
      Tab(2).Control(14)=   "cmdDptNew"
      Tab(2).Control(15)=   "cmdDptSave"
      Tab(2).Control(16)=   "cmdDptEdit"
      Tab(2).Control(17)=   "cmdDptRefund"
      Tab(2).Control(18)=   "cmdDptPrint"
      Tab(2).Control(19)=   "cmdDptCancel"
      Tab(2).Control(20)=   "cmdDptExpenses"
      Tab(2).Control(21)=   "Frame1(3)"
      Tab(2).Control(22)=   "cmdNewDRefund"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "&Maintenance Entry"
      TabPicture(3)   =   "frmLeasee1.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(0)"
      Tab(3).Control(1)=   "fmeEventHistory"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lease &Agreement"
      TabPicture(4)   =   "frmLeasee1.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeTenancyDetails"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Account &History"
      TabPicture(5)   =   "frmLeasee1.frx":0956
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label72(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label1(8)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(7)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label1(1)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label1(2)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label1(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label1(6)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label1(3)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label1(4)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label1(35)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label1(33)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label1(39)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label1(32)"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "Label1(31)"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Label1(40)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Label1(41)"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Label1(34)"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Label1(37)"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Label1(36)"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "Label1(38)"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Label1(42)"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "flxACHistorySplit"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "flxACHistory"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "cmdCopyReceipt"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "cmdSentStByEmail"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).Control(25)=   "cmdPrintHistory"
      Tab(5).Control(25).Enabled=   0   'False
      Tab(5).Control(26)=   "cmdPrintStatement"
      Tab(5).Control(26).Enabled=   0   'False
      Tab(5).Control(27)=   "cmdPrintReceipt"
      Tab(5).Control(27).Enabled=   0   'False
      Tab(5).Control(28)=   "Command1"
      Tab(5).Control(28).Enabled=   0   'False
      Tab(5).Control(29)=   "Command2"
      Tab(5).Control(29).Enabled=   0   'False
      Tab(5).Control(30)=   "chkShowOutstanding"
      Tab(5).Control(30).Enabled=   0   'False
      Tab(5).Control(31)=   "cmdSearch"
      Tab(5).Control(31).Enabled=   0   'False
      Tab(5).ControlCount=   32
      TabCaption(6)   =   "Letters && Email"
      TabPicture(6)   =   "frmLeasee1.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdPrintWord"
      Tab(6).Control(1)=   "cmdPrintEmail"
      Tab(6).Control(2)=   "cmdResendEmail"
      Tab(6).Control(3)=   "cmdDelLetter"
      Tab(6).Control(4)=   "cmdEmail"
      Tab(6).Control(5)=   "cmdViewLetter"
      Tab(6).Control(6)=   "flxLetters"
      Tab(6).Control(7)=   "flxEmails"
      Tab(6).Control(8)=   "Label50(42)"
      Tab(6).Control(9)=   "Label50(41)"
      Tab(6).ControlCount=   10
      TabCaption(7)   =   "&Memo/Attachments"
      TabPicture(7)   =   "frmLeasee1.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame17(1)"
      Tab(7).Control(1)=   "Frame17(0)"
      Tab(7).Control(2)=   "Frame8"
      Tab(7).Control(3)=   "Frame17(2)"
      Tab(7).ControlCount=   4
      Begin VB.CommandButton cmdNewDRefund 
         Caption         =   "Full Refund"
         Height          =   315
         Left            =   -69960
         Style           =   1  'Graphical
         TabIndex        =   352
         ToolTipText     =   "Refund or Expenses"
         Top             =   5445
         Width           =   1530
      End
      Begin VB.Frame Frame1 
         Height          =   2085
         Index           =   3
         Left            =   -74955
         TabIndex        =   323
         Top             =   360
         Width           =   12210
         Begin VB.CommandButton cmdDptAmtType 
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
            Left            =   7875
            TabIndex        =   342
            Top             =   1320
            Width           =   300
         End
         Begin VB.CommandButton cmdFund 
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
            Left            =   4305
            TabIndex        =   338
            Top             =   1350
            Width           =   300
         End
         Begin VB.CommandButton cmdBank 
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
            Left            =   4320
            TabIndex        =   336
            Top             =   585
            Width           =   300
         End
         Begin VB.CommandButton cmdNCList 
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
            Left            =   4305
            TabIndex        =   337
            Top             =   975
            Width           =   300
         End
         Begin VB.Frame Frame1 
            Caption         =   "Grouping:"
            Height          =   780
            Index           =   1
            Left            =   8625
            TabIndex        =   324
            Top             =   1230
            Width           =   2775
            Begin VB.OptionButton optExitingGroup 
               Caption         =   "Add to Existing Group"
               Height          =   255
               Left            =   40
               TabIndex        =   362
               Top             =   420
               Width           =   1840
            End
            Begin VB.OptionButton optNewGroup 
               Caption         =   "Create a New Group"
               Height          =   255
               Left            =   40
               TabIndex        =   361
               Top             =   200
               Value           =   -1  'True
               Width           =   1815
            End
            Begin MSForms.ComboBox cboGroup 
               Height          =   285
               Left            =   1920
               TabIndex        =   343
               Top             =   420
               Width           =   780
               VariousPropertyBits=   679495709
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "1376;503"
               TextColumn      =   2
               ColumnCount     =   2
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "987;5000"
            End
         End
         Begin VB.TextBox txtOSDpt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00EDEDED&
            Height          =   285
            Left            =   8865
            Locked          =   -1  'True
            TabIndex        =   356
            Top             =   915
            Width           =   2220
         End
         Begin VB.CommandButton cmdSetDptType 
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
            Left            =   8295
            TabIndex        =   353
            Top             =   225
            Width           =   285
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   335
            Top             =   210
            Width           =   2790
         End
         Begin VB.CommandButton cmdSetAmtType 
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
            Left            =   8190
            TabIndex        =   360
            Top             =   1320
            Width           =   330
         End
         Begin VB.TextBox txtDptAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6075
            Locked          =   -1  'True
            TabIndex        =   341
            Top             =   930
            Width           =   2400
         End
         Begin VB.TextBox txtDptDetails 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6075
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   340
            Top             =   570
            Width           =   5010
         End
         Begin VB.CommandButton cmdDepositType 
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
            Left            =   7965
            TabIndex        =   339
            Top             =   225
            Width           =   300
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Download To CSV"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   150
            Index           =   0
            Left            =   5220
            TabIndex        =   363
            Top             =   1890
            Width           =   1770
         End
         Begin MSForms.TextBox txtDptAmtType 
            Height          =   285
            Left            =   6075
            TabIndex        =   359
            Top             =   1320
            Width           =   1800
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "3175;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFund 
            Height          =   285
            Left            =   1515
            TabIndex        =   348
            Top             =   1350
            Width           =   2790
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "4921;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBank 
            Height          =   285
            Left            =   1515
            TabIndex        =   344
            Top             =   585
            Width           =   2790
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "4921;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund:"
            Height          =   195
            Index           =   34
            Left            =   110
            TabIndex        =   334
            Top             =   1335
            Width           =   390
         End
         Begin MSForms.TextBox txtDNC 
            Height          =   285
            Index           =   1
            Left            =   2235
            TabIndex        =   346
            Top             =   975
            Width           =   2070
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "3651;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Type:"
            Height          =   195
            Index           =   1
            Left            =   4770
            TabIndex        =   333
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Account:"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   332
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O/S:"
            Height          =   195
            Index           =   9
            Left            =   8550
            TabIndex        =   331
            Top             =   930
            Width           =   300
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
            Height          =   195
            Index           =   7
            Left            =   4755
            TabIndex        =   330
            Top             =   930
            Width           =   585
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Index           =   6
            Left            =   4755
            TabIndex        =   329
            Top             =   570
            Width           =   870
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nominal Code:"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   328
            Top             =   975
            Width           =   1020
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type:"
            Height          =   195
            Index           =   11
            Left            =   4755
            TabIndex        =   327
            Top             =   1335
            Width           =   1005
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   326
            Top             =   210
            Width           =   375
         End
         Begin MSForms.TextBox txtDNC 
            Height          =   285
            Index           =   0
            Left            =   1515
            TabIndex        =   325
            Top             =   975
            Width           =   690
            VariousPropertyBits=   746604573
            BorderStyle     =   1
            Size            =   "1217;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDepositType 
            Height          =   285
            Left            =   6060
            TabIndex        =   350
            Top             =   225
            Width           =   1890
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "3334;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   5355
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   5175
         Width           =   1080
      End
      Begin VB.CheckBox chkShowOutstanding 
         Caption         =   "Show Outstanding only"
         Height          =   240
         Left            =   9360
         TabIndex        =   315
         Top             =   360
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.CommandButton Command2 
         Caption         =   "new"
         Height          =   375
         Left            =   10305
         TabIndex        =   314
         Top             =   5175
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "old"
         Height          =   375
         Left            =   9225
         TabIndex        =   313
         Top             =   5175
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame17 
         Caption         =   "Comment 2:"
         Height          =   690
         Index           =   2
         Left            =   -74910
         TabIndex        =   299
         Top             =   4905
         Width           =   11175
         Begin VB.CommandButton cmdRCCEdit2 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7440
            TabIndex        =   289
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCSave2 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   291
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCCancel2 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   293
            Top             =   240
            Width           =   1125
         End
         Begin MSForms.TextBox txtRCCComments2 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   180
            TabIndex        =   287
            Top             =   225
            Width           =   4890
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "8625;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Memo"
         Height          =   3165
         Left            =   -74910
         TabIndex        =   256
         Top             =   360
         Width           =   11175
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   2475
            Left            =   5040
            ScaleHeight     =   2475
            ScaleWidth      =   11025
            TabIndex        =   257
            Top             =   90
            Width           =   11025
            Begin VB.TextBox txtMemoAll 
               Height          =   2145
               Left            =   45
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   259
               Top             =   315
               Width           =   10980
            End
            Begin VB.CommandButton cmdCloseMemo 
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
               Left            =   10620
               Style           =   1  'Graphical
               TabIndex        =   258
               Top             =   0
               Width           =   390
            End
            Begin VB.Shape Shape4 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   240
               Index           =   3
               Left            =   45
               Top             =   30
               Width           =   10575
            End
            Begin MSForms.Label lblSea 
               Height          =   195
               Left            =   135
               TabIndex        =   260
               Top             =   0
               Visible         =   0   'False
               Width           =   1905
               VariousPropertyBits=   8388627
               Caption         =   "Details"
               Size            =   "3360;344"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   269
            Top             =   2745
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   7560
            TabIndex        =   267
            Top             =   2730
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   6375
            TabIndex        =   266
            Top             =   2730
            Width           =   1125
         End
         Begin VB.TextBox txtUnitMemo 
            Height          =   1245
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   262
            Top             =   210
            Width           =   10980
         End
         Begin VB.CommandButton cmdUnitMemoNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   5355
            TabIndex        =   265
            Top             =   2730
            Width           =   975
         End
         Begin VB.TextBox txtLeaseAnalysisID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10215
            TabIndex        =   261
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdVAMemo 
            Caption         =   "&View All Memo"
            Height          =   315
            Left            =   3825
            TabIndex        =   264
            Top             =   2730
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8730
            TabIndex        =   268
            Top             =   2730
            Width           =   1125
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridLeaseAnalysis 
            Height          =   945
            Left            =   90
            TabIndex        =   263
            Top             =   1755
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   1667
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            HighLight       =   2
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
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
            _Band(0).Cols   =   9
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            Height          =   195
            Index           =   12
            Left            =   9315
            TabIndex        =   297
            Top             =   1485
            Width           =   330
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Index           =   11
            Left            =   2025
            TabIndex        =   296
            Top             =   1485
            Width           =   840
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   10
            Left            =   900
            TabIndex        =   295
            Top             =   1485
            Width           =   345
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   294
            Top             =   1485
            Width           =   210
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   4
            Left            =   90
            Top             =   1485
            Width           =   10980
         End
      End
      Begin VB.CommandButton cmdPrintWord 
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
         Left            =   -65484
         Picture         =   "frmLeasee1.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   2880
         Width           =   410
      End
      Begin VB.CommandButton cmdPrintReceipt 
         Caption         =   "Print Receipt"
         Height          =   385
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   252
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdPrintEmail 
         Caption         =   "&Print Email"
         Height          =   315
         Left            =   -65100
         TabIndex        =   248
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton cmdResendEmail 
         Caption         =   "Resend Email"
         Height          =   315
         Left            =   -67050
         TabIndex        =   247
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   -72120
         TabIndex        =   246
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -72120
         TabIndex        =   245
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -72840
         TabIndex        =   244
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -72840
         TabIndex        =   243
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -72120
         TabIndex        =   242
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveDefaults 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66960
         TabIndex        =   235
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditDefaults 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -69000
         TabIndex        =   234
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelDefaults 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64920
         TabIndex        =   233
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -72840
         TabIndex        =   232
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNC 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   1320
         Width           =   320
      End
      Begin VB.CommandButton cmdTaxList 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   230
         Top             =   1800
         Width           =   320
      End
      Begin VB.TextBox txtCodeVat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68565
         Locked          =   -1  'True
         TabIndex        =   229
         Top             =   1800
         Width           =   1080
      End
      Begin VB.TextBox txtSLControl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   228
         Top             =   2280
         Width           =   1080
      End
      Begin VB.CommandButton cmdSLC 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   2280
         Width           =   320
      End
      Begin VB.Frame fmeTenancyDetails 
         Caption         =   "Lease Details"
         Height          =   3315
         Left            =   -74880
         TabIndex        =   205
         Top             =   1200
         Width           =   11115
         Begin MSForms.TextBox txtSERVICECHARGEFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   226
            Top             =   2250
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSCPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   225
            Top             =   1860
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtBASERENTFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   224
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBRPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   223
            Top             =   690
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent:"
            Height          =   195
            Index           =   21
            Left            =   6045
            TabIndex        =   222
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Frequency:"
            Height          =   195
            Index           =   22
            Left            =   6045
            TabIndex        =   221
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/C Frequency:"
            Height          =   195
            Index           =   25
            Left            =   6045
            TabIndex        =   220
            Top             =   2250
            Width           =   1065
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Charge:"
            Height          =   195
            Index           =   24
            Left            =   6045
            TabIndex        =   219
            Top             =   1860
            Width           =   1110
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            Height          =   195
            Index           =   18
            Left            =   315
            TabIndex        =   218
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Termination Date:"
            Height          =   195
            Index           =   20
            Left            =   315
            TabIndex        =   217
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Index           =   17
            Left            =   315
            TabIndex        =   216
            Top             =   720
            Width           =   750
         End
         Begin MSForms.TextBox txtStartDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   215
            Top             =   720
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEndDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   214
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            MaxLength       =   10
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TextBox13 
            Height          =   315
            Left            =   2340
            TabIndex        =   213
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtRentReviewDate 
            Height          =   315
            Left            =   7980
            TabIndex        =   212
            Top             =   1470
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TextBox15 
            Height          =   315
            Left            =   2340
            TabIndex        =   211
            Top             =   2250
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Type:"
            Height          =   195
            Index           =   19
            Left            =   315
            TabIndex        =   210
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Rent Review Date:"
            Height          =   195
            Index           =   23
            Left            =   6045
            TabIndex        =   209
            Top             =   1470
            Width           =   1680
         End
         Begin MSForms.TextBox txtLeaseId 
            Height          =   315
            Left            =   7995
            TabIndex        =   208
            Top             =   2850
            Visible         =   0   'False
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Ref:"
            Height          =   195
            Index           =   16
            Left            =   5955
            TabIndex        =   207
            Top             =   2910
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSForms.CheckBox chkHoldingOver 
            Height          =   345
            Left            =   285
            TabIndex        =   206
            Top             =   1470
            Width           =   2280
            VariousPropertyBits=   1015031839
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4022;609"
            Value           =   "0"
            Caption         =   "Holding Over"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attachment Files:"
         ForeColor       =   &H00000000&
         Height          =   690
         Index           =   0
         Left            =   -74925
         TabIndex        =   204
         Top             =   3495
         Width           =   11175
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   315
            Left            =   9870
            Style           =   1  'Graphical
            TabIndex        =   277
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   315
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   273
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   315
            Left            =   8685
            Style           =   1  'Graphical
            TabIndex        =   275
            Top             =   240
            Width           =   1110
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   180
            TabIndex        =   271
            Top             =   240
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
      Begin VB.Frame Frame17 
         Caption         =   "Comment 1:"
         Height          =   690
         Index           =   1
         Left            =   -74910
         TabIndex        =   203
         Top             =   4215
         Width           =   11175
         Begin VB.CommandButton cmdRCCCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   285
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   283
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7440
            TabIndex        =   281
            Top             =   240
            Width           =   1125
         End
         Begin MSForms.TextBox txtRCCComments 
            Height          =   315
            Left            =   180
            TabIndex        =   279
            Top             =   225
            Width           =   4890
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "8625;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdDelLetter 
         Caption         =   "&Delete Letter"
         Height          =   315
         Left            =   -69000
         TabIndex        =   200
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Email Letter"
         Height          =   315
         Left            =   -67242
         TabIndex        =   185
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame fmeEventHistory 
         BackColor       =   &H0000FFFF&
         Caption         =   "Maintenance History"
         Height          =   3135
         Left            =   -74970
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   11265
         Begin VB.TextBox dtpReportedDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   45
            Top             =   520
            Width           =   960
         End
         Begin VB.TextBox dtpRemindDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   9380
            MaxLength       =   10
            TabIndex        =   50
            Top             =   520
            Width           =   960
         End
         Begin VB.TextBox dtpDateCompleted 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   47
            Top             =   520
            Width           =   960
         End
         Begin VB.CommandButton cmdMType 
            Caption         =   "..."
            Height          =   315
            Left            =   1680
            TabIndex        =   44
            Top             =   520
            Width           =   255
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10360
            TabIndex        =   51
            Top             =   520
            Width           =   225
         End
         Begin VB.TextBox txtEventHistoryID 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7440
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CommandButton cmdNewEvent 
            Caption         =   "&New"
            Height          =   315
            Left            =   6120
            TabIndex        =   38
            Top             =   2760
            Width           =   915
         End
         Begin VB.CommandButton cmdEditEvent 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7470
            TabIndex        =   39
            Top             =   2775
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelEvent 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10170
            TabIndex        =   41
            Top             =   2775
            Width           =   915
         End
         Begin VB.CommandButton cmdSaveEvent 
            Caption         =   "&Save"
            Height          =   315
            Left            =   8820
            TabIndex        =   40
            Top             =   2775
            Width           =   915
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridEventHistory 
            Height          =   1845
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   3254
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
         Begin MSForms.TextBox txtEventTenantID 
            Height          =   315
            Left            =   3960
            TabIndex        =   86
            Top             =   120
            Visible         =   0   'False
            Width           =   1275
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "2249;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   315
            Left            =   2880
            TabIndex        =   46
            Top             =   510
            Width           =   2655
            VariousPropertyBits=   746604571
            Size            =   "4683;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTaskOwner 
            Height          =   315
            Left            =   6480
            TabIndex        =   48
            Top             =   510
            Width           =   1460
            VariousPropertyBits=   746604571
            Size            =   "2575;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtContact 
            Height          =   315
            Left            =   7920
            TabIndex        =   49
            Top             =   510
            Width           =   1460
            VariousPropertyBits=   746604571
            Size            =   "2575;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboEventType 
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   505
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
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label42 
            Caption         =   "Remind   Date:"
            Height          =   435
            Index           =   6
            Left            =   9380
            TabIndex        =   59
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label42 
            Caption         =   "Contact:"
            Height          =   255
            Index           =   5
            Left            =   7920
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label Label42 
            Caption         =   "Task Owner:"
            Height          =   255
            Index           =   4
            Left            =   6480
            TabIndex        =   57
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Date  Actioned:"
            Height          =   435
            Index           =   3
            Left            =   5520
            TabIndex        =   56
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label42 
            Caption         =   "Alarm"
            Height          =   195
            Index           =   7
            Left            =   10360
            TabIndex        =   55
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label42 
            Caption         =   "Reported Datet:"
            Height          =   435
            Index           =   1
            Left            =   1920
            TabIndex        =   54
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label42 
            Caption         =   "Event Type:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Description:"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   52
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Property Maintenance"
         Height          =   5085
         Index           =   0
         Left            =   -75000
         TabIndex        =   164
         Top             =   360
         Width           =   11505
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "View &Diary Entry"
            Height          =   355
            Left            =   5160
            TabIndex        =   172
            Top             =   4560
            Width           =   1395
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   9960
            TabIndex        =   171
            Top             =   4560
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   7680
            TabIndex        =   170
            Top             =   4560
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "View &Job"
            Height          =   355
            Left            =   3360
            TabIndex        =   169
            Top             =   4560
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   165
            Top             =   4455
            Width           =   2775
            Begin VB.OptionButton optDiary 
               Caption         =   "Diary Entries"
               Height          =   255
               Left            =   1440
               TabIndex        =   168
               Top             =   160
               Width           =   1215
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "Jobs"
               Height          =   255
               Left            =   720
               TabIndex        =   167
               Top             =   160
               Width           =   735
            End
            Begin VB.OptionButton optAll 
               Caption         =   "All"
               Height          =   255
               Left            =   120
               TabIndex        =   166
               Top             =   160
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   3765
            Left            =   120
            TabIndex        =   173
            Top             =   690
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   6641
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
            _Band(0).Cols   =   10
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   9120
            TabIndex        =   184
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   2820
            TabIndex        =   183
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Reported"
            Height          =   480
            Index           =   2
            Left            =   1830
            TabIndex        =   182
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   181
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   180
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9120
            TabIndex        =   179
            Top             =   255
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   8040
            TabIndex        =   178
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Item / Dairy Entry"
            Height          =   495
            Index           =   4
            Left            =   4200
            TabIndex        =   177
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   5400
            TabIndex        =   176
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   6840
            TabIndex        =   175
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   10215
            TabIndex        =   174
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdPrintStatement 
         Caption         =   "Print Statement"
         Height          =   385
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdPrintHistory 
         Caption         =   "Print Account History"
         Height          =   385
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   5160
         Width           =   1755
      End
      Begin VB.CommandButton cmdSentStByEmail 
         Caption         =   "Email Statement"
         Height          =   385
         Left            =   6875
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdCopyReceipt 
         Caption         =   "Copy"
         Height          =   385
         Left            =   10155
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   5160
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdDptExpenses 
         Caption         =   "New Expense"
         Height          =   315
         Left            =   -72525
         Style           =   1  'Graphical
         TabIndex        =   349
         ToolTipText     =   "Refund or Expenses"
         Top             =   5445
         Width           =   1170
      End
      Begin VB.CommandButton cmdViewLetter 
         Caption         =   "&Print Letter"
         Height          =   315
         Left            =   -64860
         TabIndex        =   138
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmdDptCancel 
         Caption         =   "Cancel Deposit"
         Height          =   315
         Left            =   -64290
         Style           =   1  'Graphical
         TabIndex        =   358
         Top             =   5445
         Width           =   1345
      End
      Begin VB.CommandButton cmdDptPrint 
         Caption         =   "Print Receipt"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   345
         Top             =   5445
         Width           =   1230
      End
      Begin VB.CommandButton cmdDptRefund 
         Caption         =   "Deposit Refund"
         Height          =   315
         Left            =   -71310
         Style           =   1  'Graphical
         TabIndex        =   351
         ToolTipText     =   "Refund or Expenses"
         Top             =   5445
         Width           =   1305
      End
      Begin VB.CommandButton cmdDptEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   -68520
         Style           =   1  'Graphical
         TabIndex        =   354
         Top             =   5445
         Width           =   1455
      End
      Begin VB.CommandButton cmdDptSave 
         Caption         =   "Save Deposit"
         Height          =   315
         Left            =   -67065
         Style           =   1  'Graphical
         TabIndex        =   355
         Top             =   5445
         Width           =   1275
      End
      Begin VB.CommandButton cmdDptNew 
         Caption         =   "New Deposit"
         Height          =   315
         Left            =   -73635
         Style           =   1  'Graphical
         TabIndex        =   347
         Top             =   5445
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrintAccount 
         Caption         =   "Print Account"
         Height          =   315
         Left            =   -65730
         Style           =   1  'Graphical
         TabIndex        =   357
         Top             =   5445
         Width           =   1350
      End
      Begin VB.Frame fmeTenantAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         Left            =   -74760
         TabIndex        =   19
         Top             =   360
         Width           =   10995
         Begin VB.CommandButton cmdCancelTenantAddress 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   9750
            TabIndex        =   82
            Top             =   4710
            Width           =   1035
         End
         Begin VB.CommandButton cmdSaveTenantAddress 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   7935
            TabIndex        =   81
            Top             =   4710
            Width           =   1035
         End
         Begin VB.CommandButton cmdEditTenantAddress 
            Caption         =   "&Update Contact Details"
            Height          =   315
            Left            =   135
            TabIndex        =   80
            Top             =   210
            Width           =   2430
         End
         Begin MSForms.CommandButton cmdSendMail 
            Height          =   330
            Index           =   1
            Left            =   10200
            TabIndex        =   202
            Top             =   3120
            Width           =   420
            Size            =   "732;591"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdSendMail 
            Height          =   330
            Index           =   0
            Left            =   3960
            TabIndex        =   201
            Top             =   3000
            Width           =   420
            Size            =   "732;591"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   " Alternative Address:"
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
            Index           =   34
            Left            =   6300
            TabIndex        =   85
            Top             =   615
            Width           =   1560
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   " Lessee Address:"
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
            Index           =   33
            Left            =   240
            TabIndex        =   84
            Top             =   585
            Width           =   1230
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice/Statement To:"
            Height          =   195
            Index           =   15
            Left            =   6150
            TabIndex        =   83
            Top             =   180
            Width           =   1560
         End
         Begin MSForms.ComboBox cboInvoiceTo 
            Height          =   315
            Left            =   8130
            TabIndex        =   79
            Top             =   210
            Width           =   2640
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "4657;556"
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
            Object.Width           =   "0"
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   36
            Left            =   6300
            TabIndex        =   78
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   37
            Left            =   6300
            TabIndex        =   77
            Top             =   2790
            Width           =   750
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Index           =   35
            Left            =   6300
            TabIndex        =   76
            Top             =   960
            Width           =   585
         End
         Begin MSForms.TextBox txtContact2 
            Height          =   315
            Left            =   7680
            TabIndex        =   66
            Top             =   900
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine1 
            Height          =   315
            Left            =   7680
            TabIndex        =   67
            Top             =   1290
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   68
            Top             =   1620
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine3 
            Height          =   315
            Left            =   7680
            TabIndex        =   69
            Top             =   1950
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine4 
            Height          =   315
            Left            =   7680
            TabIndex        =   70
            Top             =   2280
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillPostCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   71
            Top             =   2730
            Width           =   1815
            VariousPropertyBits=   746604571
            MaxLength       =   12
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillFax 
            Height          =   315
            Left            =   7680
            TabIndex        =   75
            Top             =   4290
            Width           =   2940
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5186;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillTelephone 
            Height          =   315
            Left            =   7680
            TabIndex        =   74
            Top             =   3900
            Width           =   2940
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5186;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   73
            Top             =   3510
            Width           =   2940
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5186;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail2 
            Height          =   315
            Left            =   7680
            TabIndex        =   72
            Top             =   3120
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Index           =   0
            Left            =   6300
            TabIndex        =   65
            Top             =   3210
            Width           =   405
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   40
            Left            =   6300
            TabIndex        =   64
            Top             =   3585
            Width           =   795
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   39
            Left            =   6300
            TabIndex        =   63
            Top             =   4320
            Width           =   285
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   38
            Left            =   6300
            TabIndex        =   62
            Top             =   3945
            Width           =   525
         End
         Begin MSForms.TextBox txtHOFax 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            Top             =   4200
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOTelephone 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   3810
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   3420
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail1 
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   3030
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   36
            Top             =   3060
            Width           =   405
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   30
            Left            =   240
            TabIndex        =   35
            Top             =   3435
            Width           =   795
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   34
            Top             =   4170
            Width           =   285
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   33
            Top             =   3795
            Width           =   525
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   31
            Top             =   2640
            Width           =   750
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   585
         End
         Begin MSForms.TextBox txtContact1 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   810
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   1200
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine2 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   1530
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine3 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine4 
            Height          =   315
            Left            =   1440
            TabIndex        =   24
            Top             =   2190
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOPostCode 
            Height          =   315
            Left            =   1440
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
            VariousPropertyBits=   746604571
            MaxLength       =   12
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   4065
            Index           =   0
            Left            =   120
            Top             =   690
            Width           =   4665
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   3945
            Left            =   6120
            Top             =   720
            Width           =   4665
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   2550
         Left            =   90
         TabIndex        =   89
         Top             =   855
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4498
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDeposit 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   134
         Top             =   2790
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4471
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLetters 
         Height          =   2265
         Left            =   -74925
         TabIndex        =   139
         Top             =   600
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   3995
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistorySplit 
         Height          =   1335
         Left            =   75
         TabIndex        =   142
         Top             =   3720
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmails 
         Height          =   2025
         Left            =   -74925
         TabIndex        =   249
         Top             =   3240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   3572
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales LedgerControl Account:"
         Height          =   195
         Index           =   1
         Left            =   -72300
         TabIndex        =   298
         Top             =   2340
         Width           =   2085
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emails:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   42
         Left            =   -74910
         TabIndex        =   251
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letters:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   41
         Left            =   -74925
         TabIndex        =   250
         Top             =   360
         Width           =   585
      End
      Begin MSForms.TextBox txtNominalCode 
         Height          =   285
         Left            =   -69960
         TabIndex        =   241
         Top             =   1320
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
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
         Caption         =   "Nominal Code:"
         Height          =   195
         Index           =   0
         Left            =   -72315
         TabIndex        =   240
         Top             =   1320
         Width           =   1020
      End
      Begin MSForms.TextBox txtNominalCodeName 
         Height          =   285
         Left            =   -68565
         TabIndex        =   239
         Top             =   1320
         Width           =   2055
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3625;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVatCode 
         Height          =   285
         Index           =   0
         Left            =   -69960
         TabIndex        =   238
         Top             =   1800
         Width           =   1080
         BackColor       =   16777215
         Size            =   "1905;503"
         BorderStyle     =   1
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         Height          =   195
         Index           =   0
         Left            =   -72270
         TabIndex        =   237
         Top             =   1845
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         Height          =   1935
         Index           =   5
         Left            =   -72840
         Top             =   960
         Width           =   6975
      End
      Begin MSForms.TextBox txtSLControlName 
         Height          =   285
         Left            =   -68565
         TabIndex        =   236
         Top             =   2280
         Width           =   2055
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3625;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   42
         Left            =   10440
         TabIndex        =   154
         Top             =   3465
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   38
         Left            =   5970
         TabIndex        =   153
         Top             =   3480
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Index           =   36
         Left            =   4065
         TabIndex        =   152
         Top             =   3480
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   37
         Left            =   4995
         TabIndex        =   151
         Top             =   3480
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   34
         Left            =   2100
         TabIndex        =   150
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   41
         Left            =   9525
         TabIndex        =   149
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   40
         Left            =   8625
         TabIndex        =   148
         Top             =   3480
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   31
         Left            =   75
         TabIndex        =   147
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   32
         Left            =   420
         TabIndex        =   146
         Top             =   3480
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   39
         Left            =   7605
         TabIndex        =   145
         Top             =   3480
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   33
         Left            =   1140
         TabIndex        =   144
         Top             =   3480
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   35
         Left            =   3480
         TabIndex        =   143
         Top             =   3480
         Width           =   285
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   10
         Left            =   -63570
         TabIndex        =   135
         Top             =   2490
         Width           =   1005
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Group"
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   4
         Left            =   3720
         TabIndex        =   133
         Top             =   630
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   132
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   6
         Left            =   7560
         TabIndex        =   131
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   5
         Left            =   6360
         TabIndex        =   130
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   1095
         TabIndex        =   129
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   128
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   7
         Left            =   8760
         TabIndex        =   127
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   8
         Left            =   9960
         TabIndex        =   126
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label72 
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   125
         Top             =   360
         Width           =   11400
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   4
         Left            =   -71070
         TabIndex        =   107
         Top             =   2490
         Width           =   1200
         ForeColor       =   0
         BackColor       =   -2147483646
         VariousPropertyBits=   8388627
         Caption         =   "Date"
         Size            =   "2117;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741828
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   9
         Left            =   -64860
         TabIndex        =   105
         Top             =   2490
         Width           =   855
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Refund/Exp"
         Size            =   "1508;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   8
         Left            =   -65730
         TabIndex        =   104
         Top             =   2490
         Width           =   495
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "O/S"
         Size            =   "873;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   7
         Left            =   -66555
         TabIndex        =   103
         Top             =   2490
         Width           =   615
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Amount"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   6
         Left            =   -68760
         TabIndex        =   102
         Top             =   2490
         Width           =   1470
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Description"
         Size            =   "2593;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   5
         Left            =   -70110
         TabIndex        =   101
         Top             =   2490
         Width           =   1710
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Bank Account"
         Size            =   "3016;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   3
         Left            =   -72450
         TabIndex        =   100
         Top             =   2490
         Width           =   1260
         ForeColor       =   0
         BackColor       =   -2147483646
         VariousPropertyBits=   8388627
         Caption         =   "TransactionType"
         Size            =   "2222;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741828
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   2
         Left            =   -72570
         TabIndex        =   99
         Top             =   2490
         Visible         =   0   'False
         Width           =   255
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "TransactionType"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   1
         Left            =   -73635
         TabIndex        =   98
         Top             =   2490
         Width           =   915
         ForeColor       =   0
         BackColor       =   -2147483646
         VariousPropertyBits=   8388627
         Caption         =   "Type"
         Size            =   "1614;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741828
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   195
         Index           =   0
         Left            =   -74385
         TabIndex        =   97
         Top             =   2490
         Width           =   225
         ForeColor       =   0
         BackColor       =   -2147483646
         VariousPropertyBits=   276824083
         Caption         =   "No"
         Size            =   "397;344"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741828
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00000000&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   -74895
         Top             =   2490
         Width           =   12150
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00004040&
         BorderWidth     =   3
         Height          =   1935
         Index           =   4
         Left            =   -72840
         Top             =   960
         Width           =   6975
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      Height          =   15
      Index           =   1
      Left            =   0
      Top             =   2670
      Width           =   12015
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Height          =   15
      Index           =   0
      Left            =   0
      Top             =   2670
      Width           =   12015
   End
End
Attribute VB_Name = "frmLeasee1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmLease1
Option Explicit

'Private Const COL_OUT_STANDING_REFUND = 10
'Private Const COL_IS_REFUND = 11
'Private Const COL_BANK_CODE = 13
'Private Const COL_NOMINAL_CODE = 14
'Private Const COL_AMOUNT_TYPE_CODE = 15
'Private Const COL_DEPOSIT_TYPE_CODE = 16
'Private Const COL_NC_NAME = 17
'Private Const COL_FUND = 18
'Private Const COL_RECON = 19

Private GROUP_NO     As String

Public LOAD_TENANT_TENANTID As String
Private CLIENT_ID As String
Private PROPERTY_ID As String

Dim NEWMODE_ As Boolean
Dim COPYMODE_ As Boolean
Dim SEARCHTenantMODE_ As Boolean
Dim M_HISTORY_NEW_ENTRY_ As Boolean
Dim BANK_PAYMENT_NEW_ENTRY_ As Boolean
Dim IMAGE_FILE_NAME_ As String
Dim bDEPOSIT_HELD As Boolean
Dim yDEPOSIT As Long                   'using for deposit(0-2), refund(3)
Dim bLeaseSetup As Boolean, cCurDepAmt As Currency, iGridRow As Integer
Dim bSortingCol1 As Boolean, bSortingCol2 As Boolean, bSortingCol3 As Boolean
Dim szSel      As String
Dim sText As String

Dim Lease_ANALYSIS_NEW_ENTRY  As Boolean

Dim szaTenantBalance()  As String
Dim cOriRAmt            As Currency
Dim dataProperty()      As String

Private Type SendDemandByEmail
   szLesseeID    As String
   szLesseeEmail As String
   colAtt        As Collection
End Type
Private uLessee   As SendDemandByEmail
Dim strSessionClientID As String
Dim strSessionPropertyID As String
Dim reportingDate As String
Dim sessionID As String
Dim bolgridTenantLookupRefresh As Boolean
Dim bEdit As Boolean
Dim bFullRefund As Boolean
Private Sub cmdDepositType_Click()
   LoadTypeGrid

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = txtDepositType.Left
   fraList(0).Top = txtDate.Top + 2500
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "Type"
End Sub
Private Function LoadDptAmtType()
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   'txtSearch1.Width = 1400
  ' txtSearch1.Left = 40

   'txtSearch2.Width = 2600
  ' txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Deposit Amount Type"
   lblSearch1(0).Caption = "Description"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
'   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code;"
  ' szSQL = "SELECT FundID,FundCode, FundName FROM FUND;"
 szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   


   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
            flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("SC").Value
            flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("V").Value
            flxSupplier(0).TextMatrix(iRows, 2) = "" 'adoRst.Fields.Item("FundID").Value
            flxSupplier(0).RowHeight(iRows) = 280
            If Not adoRST.EOF Then flxSupplier(0).AddItem ""
            flxSupplier(0).row = 1
            iRows = iRows + 1
            adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Function

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function
Private Function LoadTypeGrid()
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   'txtSearch1.Width = 1400
  ' txtSearch1.Left = 40

   'txtSearch2.Width = 2600
  ' txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   If yDEPOSIT = 1 Then
        lblSearch0(0).Caption = "Deposit Type"
   Else
        lblSearch0(0).Caption = "Expense Type"
   End If
   lblSearch1(0).Caption = "Description"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
'   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString
    szSQL = "SELECT SC.Code as C, SC.Value as V " & _
                "FROM SecondaryCode AS SC " & _
                "WHERE SC.PrimaryCode = 'RTYP' " & _
                "ORDER BY SC.Value;"

    If yDEPOSIT = 1 Then
        szSQL = "SELECT SC.Code as C, SC.Value as V " & _
                "FROM SecondaryCode AS SC " & _
                "WHERE SC.PrimaryCode = 'DPTYP' " & _
                "ORDER BY SC.Value;"
     ElseIf yDEPOSIT = 2 Then
        szSQL = "SELECT SC.Code as C, SC.Value as V " & _
                "FROM SecondaryCode AS SC " & _
                "WHERE SC.PrimaryCode = 'DPTYP' " & _
                "ORDER BY SC.Value;"
     ElseIf yDEPOSIT = 3 Or yDEPOSIT = 5 Then
        szSQL = "SELECT SC.Code as C, SC.Value as V " & _
                "FROM SecondaryCode AS SC " & _
                "WHERE SC.PrimaryCode = 'RTYP' " & _
                "ORDER BY SC.Value;"
     ElseIf yDEPOSIT = 4 Or yDEPOSIT = 31 Then
            szSQL = "SELECT SC.Code as C, SC.Value as V " & _
                "FROM SecondaryCode AS SC " & _
                "WHERE SC.PrimaryCode = 'EXPTYP' " & _
                "ORDER BY SC.Value;"
'     ElseIf yDEPOSIT = 5 Or yDEPOSIT = 32 Then
'        szSQL = "SELECT SC.Code as C, SC.Value as V " & _
'                "FROM SecondaryCode AS SC " & _
'                "WHERE SC.PrimaryCode = 'FRTYP' " & _
'                "ORDER BY SC.Value;"
     End If
           
   If flxDeposit.TextMatrix(flxDeposit.row, 6) = "Deposit" Then
        cmdDptRefund.Caption = "Deposit Refund"
   Else
        cmdDptRefund.Caption = "Expense Refund"
   End If
           
           

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   


   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
            flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("C").Value
            flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("V").Value
            flxSupplier(0).TextMatrix(iRows, 2) = "" 'adoRst.Fields.Item("FundID").Value
            flxSupplier(0).RowHeight(iRows) = 280
            If Not adoRST.EOF Then flxSupplier(0).AddItem ""
            flxSupplier(0).row = 1
            iRows = iRows + 1
            adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Function

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function
Private Sub cmdDptAmtType_Click()
   LoadDptAmtType

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = txtDptAmtType.Left
   fraList(0).Top = txtDptAmtType.Top + 2500
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "DptAmtType"
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdNewDRefund_Click()
   txtDepositType.text = ""
   txtDptAmount.text = txtOSDpt.text
   txtDptDetails.text = ""
   frmSecondaryCode.PRIMARY_CODE_SHOW = "FRTYP"
    
   ButtonHanlding RefundMode
   Call EnableLeaseHeldBoxed
   txtDate.text = Format(Date, "dd/mm/yyyy")
   'yDEPOSIT = 5
   yDEPOSIT = 3

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer
   FocusControl txtDate
'   Picture3.Visible = True
   'tabTenant.Enabled = False
    txtOSDpt.text = txtDeposit.text
    bFullRefund = True
'   Dim adoconn As New ADODB.Connection
'   Dim szSQL As String
'   adoconn.Open getConnectionString
'   szSQL = "SELECT (DptAmount-OSRefund) as amt " & _
'           "FROM TenantDeposit AS D LEFT JOIN tlbBankPayment AS B ON D.DepositID = B.TenantDeposit " & _
'           "Where D.TenantID = '" & txtTenantID.text & "' AND D.Deleted = False;"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   txtAmountOutStanding(0).text = "0.00"
'   If Not adoRst.EOF Then
'        txtAmountOutStanding(0).text = Format(adoRst("amt").Value, "0.00")
'        txtAmountOutStanding(1).text = Format(adoRst("amt").Value, "0.00")
'   End If
'   adoRst.Close
'   adoconn.Close
   
End Sub

'Private Sub cmdOK_Click(Index As Integer)
''    Picture3.Visible = False
'    tabTenant.Enabled = True
'    Dim adoRst As New ADODB.Recordset
'    Dim cTotalPay As Double
'    Dim szSQL As String
'    Dim adoconn As New ADODB.Connection
'    If Index = 2 Or Index = 3 Then
'        txtAmountOutStanding(0).text = "0.00"
'    End If
'    If Index = 0 Then
'         cTotalPay = Val(txtAmountOutStanding(1).text)
'         adoconn.Open getConnectionString
'        szSQL = "SELECT sum(DptAmount-OSRefund) as amt " & _
'                "FROM TenantDeposit AS D " & _
'                "Where D.TenantID = '" & txtTenantID.text & "' AND D.Deleted = False AND OSRefund>0 order by DepositDate desc;"
'        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'        If Not adoRst.EOF Then
'             txtAmountOutStanding(0).text = adoRst("amt").Value
'        End If
'        cTotalPay = txtAmountOutStanding(0).text
'        While Not cTotalPay > 0
''             cTotalPay = adoRst!DptAmount - adoRst!OSRefund
''             adoRst!OSRefund = adoRst!OSRefund + dblAmount
'
'              If cTotalPay >= CCur(adoRst.Fields.Item("OSRefund").Value) Then
'                    cTotalPay = cTotalPay - CCur(adoRst.Fields.Item("OSRefund").Value)
'                    adoRst.Fields.Item("OSRefund").Value = 0
'              Else
'                    adoRst.Fields.Item("OSRefund").Value = adoRst.Fields.Item("OSRefund").Value - cTotalPay
'                    cTotalPay = 0
'             End If
'
'             adoRst.Update
'             adoRst.MoveNext
'        Wend
'        adoRst.Close
'        adoconn.Close
'    End If
'End Sub

Private Sub cmdPrintAccount_Click()
On Error GoTo Err
   ' Passing the from and to date values to Crystal Reports
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TenantDepositHistory.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Report.ParameterFields(1).AddCurrentValue txtTenantID.text

   Load frmReport
   frmReport.LoadReportViewer Report
   Exit Sub
Err:
End Sub

Private Sub Label11_Click(Index As Integer)

On Error GoTo Err
    If Index = 0 Then
        Dim a
        Dim strFileName As String
        Dim i As Integer
        Dim j As Integer
        Dim newLine As String
        Dim strFile As String
        Dim FS

        strFileName = BrowseForFolder(Me.hWnd, "Select a Directory")
        If strFileName = "" Then Exit Sub
        strFileName = strFileName & "\TenantDepositHistory" & Format(Now, "yyyyMMddhhmmss") & ".csv"
            
        Dim iFileNo As Integer
        iFileNo = FreeFile
        'open the file for writing
        Open strFileName For Output As #iFileNo
        'please note, if this file already exists it will be overwritten!
         newLine = ""
         'Write to Specified file from Flex
            For i = 0 To flxDeposit.Rows - 1
                For j = 0 To flxDeposit.Cols - 1
                    newLine = newLine + flxDeposit.TextMatrix(i, j) + ","
                Next j
                newLine = newLine + vbCrLf
                 Print #iFileNo, newLine
                newLine = ""
            Next i

        Close #iFileNo
        MsgBox "File has been written"
      End If
    Exit Sub
Err:
    MsgBox Err.description
End Sub

Private Sub Label6_Click(Index As Integer)
     If Index >= 0 And Index <= 6 Then
            Label6(Index).FontBold = Not Label6(Index).FontBold
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            ConfigFlxDeposit
            LoadFlxDeposit adoConn, Index, IIf(Label6(Index).FontBold, "DESC", "ASC")
            adoConn.Close
    End If
End Sub

Private Sub optBoth_Click()
    bolgridTenantLookupRefresh = False
End Sub

Private Sub optCurrentTenant_Click()
    bolgridTenantLookupRefresh = False
End Sub

Private Sub optExTenant_Click()
    bolgridTenantLookupRefresh = False
End Sub



Private Sub txtAmountOutStanding_KeyPress(Index As Integer, KeyAscii As Integer)
     DigitTextKeyPress txtDptAmount, KeyAscii
End Sub

Private Sub txtDptAmount_GotFocus()
    If bFullRefund = True And bEdit = False Then
        txtDptAmount.text = txtOSDpt.text
        txtDptAmount.SelStart = 0
        txtDptAmount.SelLength = Len(txtDptAmount)
    End If
End Sub

'Private Sub txtAmountOutStanding_LostFocus(Index As Integer)
'    If Index = 1 Then
'            If Val(txtAmountOutStanding(1).text) > Val(txtAmountOutStanding(0).text) Then
'                txtAmountOutStanding(1).text = txtAmountOutStanding(0).text
'                MsgBox "Entered Outstanding amount cannot be greater than Total Outstanding amount", vbInformation, "Warning"
'            End If
'    End If
'End Sub

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtSearchFromD.text = ""
            txtSearchToD.text = ""
            txtSearchRef.text = ""
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If Len(txtSearchNo.text) > 0 Then
                Call LoadFlxACHistory(adoConn, "1")
            Else
                Call LoadFlxACHistory(adoConn, "")
            End If
            If Len(txtSearchNo.text) > 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
            Else
                cmdSearch.Caption = "Sea&rch"
            End If
             
            
             adoConn.Close
             Set adoConn = Nothing
    End If
End Sub

Private Sub txtSearchFromD_Change()
    TextBoxChangeDate txtSearchFromD
    txtSearchNo.text = ""
    txtSearchRef.text = ""
End Sub

Private Sub txtSearchFromD_GotFocus()
    SelTxtInCtrl txtSearchFromD
End Sub

Private Sub txtSearchFromD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchFromD, KeyAscii
End Sub

Private Sub txtSearchFromD_LostFocus()
    If txtSearchFromD.text <> "" Then
        TextBoxFormatDate txtSearchFromD
        txtSearchToD.text = txtSearchFromD.text
        SelTxtInCtrl txtSearchToD
     End If
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtSearchFromD.text = ""
            txtSearchToD.text = ""
            txtSearchNo.text = ""
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If Len(txtSearchRef.text) > 0 Then
                Call LoadFlxACHistory(adoConn, "2")
            Else
                Call LoadFlxACHistory(adoConn, "")
            End If
            If Len(txtSearchRef.text) > 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
            Else
                cmdSearch.Caption = "Sea&rch"
            End If
                
            
            adoConn.Close
            Set adoConn = Nothing
    End If
End Sub

Private Sub txtSearchToD_Change()
     TextBoxChangeDate txtSearchToD
     txtSearchNo.text = ""
     txtSearchRef.text = ""
End Sub

Private Sub txtSearchToD_GotFocus()
    SelTxtInCtrl txtSearchToD
End Sub

Private Sub txtSearchToD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchToD, KeyAscii
End Sub

Private Sub txtSearchToD_LostFocus()
    If txtSearchToD.text <> "" Then TextBoxFormatDate txtSearchToD
End Sub



Private Sub txtSupplierSearc_Change()
'   If Not bFormLoaded Then Exit Sub
'   SortTheGrid flxPurchHistory, txtClientIdlist, cmbPropertyHistory, txtSupplierSearc
'   flxPurchHistorySplit.Clear
'   flxPurchHistorySplit.Rows = 2
End Sub

Private Sub cmdSearchOK_Click()
    fraSearch.Visible = False
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "")
             cmdSearch.Caption = "Sea&rch"
        ElseIf Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "3")
             cmdSearch.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
             cmdSearch.Caption = "Clear Sea&rch"
             Call LoadFlxACHistory(adoConn, "4")
        End If

    
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdSearchCancel_Click()
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString

    fmeLoading.Visible = False
    fraSearch.Visible = False
'    adoconn.Close

End Sub
Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub


'Private Sub cboBankId_Change()
'   If Not IsNull(cboBankId) And cboBankId <> "" Then
'      txtBankSortCode.text = cboBankId.Column(2)
'      txtBranchName.text = cboBankId.Column(3)
'      txtBankAddress1.text = cboBankId.Column(4)
'      txtBankAddress2.text = cboBankId.Column(5)
'      txtBankAddress3.text = cboBankId.Column(6)
'      txtBankPostCode.text = cboBankId.Column(7)
'   End If
'End Sub

Private Sub cboClientList_Click()
   If fmeTenantLookup.Visible = False Then Exit Sub

   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
    Dim strWHR As String
'   If txtClientList.tag <> "ALL" Then
''     Filter properties
'
''     Get the count of the properties of the client
'      For i = 1 To UBound(dataProperty, 2)
'         If dataProperty(4, i) = txtClientList.tag Then
'            TotalRow = TotalRow + 1
'         End If
'      Next i
'
'      TotalCol = UBound(dataProperty, 1)
'      ReDim Data(TotalCol, TotalRow) As String
'
''     Load the properties in the combo
'      Data(0, 0) = "ALL"
'      Data(1, 0) = "All Properties"
'      j = 1
'      For i = 1 To UBound(dataProperty, 2)
'         If dataProperty(4, i) = txtClientList.tag Then
'            For K = 0 To TotalCol - 1
'               Data(K, j) = dataProperty(K, i)
'            Next K
'            j = j + 1
'         End If
'      Next i
'
'      cboPropertyList.Column() = Data()
'
''     Filter the Lessee
'   Else
''     Reload all Properties in the combo
'      Data = dataProperty
'      cboPropertyList.Column() = Data()
'   End If
'
'   Exit Sub
   
   Dim adoConn    As New ADODB.Connection
   Dim adoRST     As New ADODB.Recordset
   Dim szSQL      As String

   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

'   If txtClientList.Tag = "ALL" Then
'      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "ORDER BY PropertyID;"
'   Else
'      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'              "ORDER BY PropertyID;"
'   End If
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
'    cboPropertyList.ListIndex = 0
   'Modified by anol 25 Oct 2015
   'issue 571
'   lessee Module
'
'The Lessee records are not filtering correctly by client and by property. If the user has more than one client and more than
'
'one property for the same client and changes the filter selection for different client and properties the program is not
'
'filtering. Instead it is showing records for all lessees
'*************************************************************************************************
'  All tenants
'   If optBoth.Value Then
'      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
'              "FROM Tenants AS T LEFT JOIN " & _
'                   "[" & _
'                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
'                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
'                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                      "L.Status = TRUE "
'
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "
'
'      If txtClientList.tag <> "ALL" Then
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
'         Else
'            szSQL = szSQL + "WHERE "
'         End If
'         szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'      Else
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' "
'         End If
'      End If
'
'      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'
'      PopulateTenantLookup szSQL, adoConn
'   End If
'
''  Current tenants Only
'   If optCurrentTenant.Value Then
'      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
'              "FROM Tenants AS T LEFT JOIN " & _
'                   "[" & _
'                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
'                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
'                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                      "L.Status = TRUE "
'
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') "
'
'      If txtClientList.tag <> "ALL" Then
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'         End If
'         szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
'      Else
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'         End If
'      End If
'
'      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'
'      PopulateTenantLookup szSQL, adoConn
'   End If
'
''  Deleted tenants only
'   If optExTenant.Value Then
'       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                    "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
'               "FROM Tenants AS T LEFT JOIN " & _
'                    "[" & _
'                    "SELECT U.UnitName, L.SageAccountNumber " & _
'                    "From Units AS U, LeaseDetails AS L, Property AS P, P.PropertyID, P.ClientID " & _
'                    "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                       "L.Status = TRUE "
'
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'                 "WHERE (T.Comments) IS NOT NULL OR T.Comments<>'' "
'
'      If txtClientList.tag <> "ALL" Then
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'         End If
'         szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
'      Else
'         If txtPropertyList.tag <> "ALL" Then
'            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'         End If
'      End If
'
'      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'
'      PopulateTenantLookup szSQL, adoConn
'   End If
'**************************************************************************************************

'#########################################################################################
'  All tenants
   If optBoth.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                   "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE OCCUPIDE_ = FALSE "
    
              If txtClientList.Tag <> "ALL" Then
                 If txtPropertyList.Tag <> "ALL" Then
                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' AND "
                 Else
                    szSQL = szSQL + "AND "
                 End If
                 szSQL = szSQL + "IQ.ClientID = '" & txtClientList.Tag & "' "
              Else
                 If txtPropertyList.Tag <> "ALL" Then
                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' "
                 End If
              End If
   End If

'  Current tenants Only
   If optCurrentTenant.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber  " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                   "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE "
            If txtClientList.Tag <> "ALL" Then
                'If cboPropertyList.ListIndex > -1 Then 'if and else condition added by anol 17 Sep 2015 issue 571 note 1174
                        If txtPropertyList.Tag <> "ALL" Then
                           szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' AND "
                        Else
                           szSQL = szSQL + "AND "
                        End If
'                Else
'                       cboPropertyList.ListIndex = 0
'                       szSQL = szSQL + "AND "
'                End If
                 szSQL = szSQL + "IQ.ClientID = '" & txtClientList.Tag & "' "
              Else
                 If txtPropertyList.Tag <> "ALL" Then
                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' "
                 End If
              End If
                 
   End If

'  Deleted tenants only
   If optExTenant.Value Then
'   "SELECT SageAccountNumber, CompanyName " & _
'             "FROM Tenants " & _
'             "WHERE Tenants.SageAccountNumber NOT IN " & _
'                 "(SELECT LeaseDetails.SageAccountNumber " & _
'                 "FROM LeaseDetails " & _
'                 "WHERE Status=True) AND " & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
'             "ORDER BY SageAccountNumber"
szSQL = ""
       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                    "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
               "FROM Tenants AS T INNER JOIN " & _
                    "[" & _
                    "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                    "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                    "WHERE U.UnitNumber = L.UnitNumber AND " & _
                        "L.Status = FALSE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "
      
     If txtClientList.Tag <> "ALL" Then
        'If cboPropertyList.ListIndex > -1 Then 'added by anol 09 July 2015
                If txtPropertyList.Tag <> "ALL" Then
                    strWHR = strWHR + "WHERE IQ.PropertyID = '" & txtPropertyList.Tag & "' AND "
                End If
                If strWHR <> "" Then
                    strWHR = strWHR + "AND "
                Else
                    strWHR = strWHR + "WHERE "
                End If
                strWHR = strWHR + "IQ.ClientID = '" & txtClientList.Tag & "' "
'         Else
'            strWHR = strWHR + "WHERE IQ.ClientID = '" & txtClientList.Tag & "' "
'         End If

      Else
            ''If cboPropertyList.ListIndex > -1 Then
                If txtPropertyList.Tag <> "ALL" Then
                      strWHR = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.Tag & "' "
                End If
'                If strWhr <> "" Then
'                    strWhr = strWhr + "AND "
'                Else
'                    strWhr = strWhr + "WHERE "
'                End If
'                szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'             Else
'                 szSQL = szSQL + "Where IQ.ClientID = '" & txtClientList.tag & "' "
             ''End If
      End If
   End If
   'Modified by anol 09 July 2015
   'Error was happening  on exlesse option
   ' I am changing the filer clause as it was not containg 'Where' (It was 'AND')
'           If txtClientList.tag <> "ALL" Then
'                 If txtPropertyList.tag <> "ALL" Then
'                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
'                 Else
'                    szSQL = szSQL + "AND "
'                 End If
'                 szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'              Else
'                 If txtPropertyList.tag <> "ALL" Then
'                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'                 End If
'              End If
'      If txtClientList.tag <> "ALL" Then
'        If cboPropertyList.ListIndex > -1 Then 'added by anol 09 July 2015
'                If txtPropertyList.tag <> "ALL" Then
'                       If InStr(1, szSQL, "WHERE") = 0 Then
'                           szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
'                       Else
'                           szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
'                       End If
'                Else
'                   If InStr(1, szSQL, "WHERE") = 0 Then
'                       szSQL = szSQL + "WHERE "
'                   Else
'                       szSQL = szSQL + "AND "
'                   End If
'                End If
'              szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'         Else
'            szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
'         End If
'
'      Else
'            If cboPropertyList.ListIndex > -1 Then
'                If txtPropertyList.tag <> "ALL" Then
'                    If InStr(1, szSQL, "WHERE") = 0 Then
'                        szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' "
'                    Else
'                       szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
'                    End If
'                End If
'                szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'             Else
'                 szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
'             End If
'      End If
      'End of modification
              szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
        
              PopulateTenantLookup szSQL, adoConn
              UpdateBalance
   '#########################################################################################
'NoRes:
'   adoRst.Close
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

Private Sub cboClientList_GotFocus()
   'SelTxtInCtrl cboClientList
End Sub

Private Sub cboDepositType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtDptDetails.SetFocus
    End If
End Sub

'Private Sub cboPropertyList_Click()
'   If fmeTenantLookup.Visible = False Then Exit Sub
'   Dim i          As Integer
'    Dim strWHR As String
''  Reset the grid - open all hidden rows
''Resolved by BOSL
''issue 445,note 972,When user selects a property on the lessee search it is taking a very long time to load the list of lessees for the property.
'
''''   For i = 1 To gridTenantLookup.Rows - 1
''''      gridTenantLookup.RowHeight(i) = 240
''''   Next i
''''
''''   If txtPropertyList.tag <> "ALL" Then
''''      For i = 1 To gridTenantLookup.Rows - 1
''''         If gridTenantLookup.TextMatrix(i, 5) <> txtPropertyList.tag Then
''''            gridTenantLookup.RowHeight(i) = 0
''''         End If
''''      Next i
''''   Else              'Check client
''''      If txtClientList.tag <> "ALL" Then
''''         For i = 1 To gridTenantLookup.Rows - 1
''''            If gridTenantLookup.TextMatrix(i, 6) <> txtClientList.tag Then
''''               gridTenantLookup.RowHeight(i) = 0
''''            End If
''''         Next i
''''      End If
''''   End If
'
'
'   'Exit Sub
'    Dim adoConn As New ADODB.Connection
'    Dim szSQL As String
'   'Dim strWHR As String
'    Dim szWhere As String
'    If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
'       szWhere = ""
'
'    If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
'       szWhere = "AND LA.CLIENTID = '" & txtClientList.Tag & "' "
'
'    If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
'       szWhere = "AND LA.PROPERTYID = '" & txtPropertyList.Tag & "' "
'
'    If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
'       szWhere = "AND LA.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
'                          "AND LA.CLIENTID = '" & txtClientList.Tag & "' "
'
'
'   adoConn.Open getConnectionString
'
'''  All tenants
''   If optBoth.Value Then
''      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
''                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
''              "FROM Tenants AS T LEFT JOIN " & _
''                   "[" & _
''                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
''                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
''                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
''                      "L.Status = TRUE "
''
''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "
''
''      If txtClientList.tag <> "ALL" Then
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
''         Else
''            szSQL = szSQL + "WHERE "
''         End If
''         szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
''      Else
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "WHERE IQ.PropertyID = '" & txtPropertyList.tag & "' "
''         End If
''      End If
''
''      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
''
''      PopulateTenantLookup szSQL, adoConn
''
''   End If
''
'''  Current tenants Only
''   If optCurrentTenant.Value Then
''      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
''                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
''              "FROM Tenants AS T LEFT JOIN " & _
''                   "[" & _
''                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
''                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
''                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
''                      "L.Status = TRUE "
''
''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
''                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') "
''
''      If txtClientList.tag <> "ALL" Then
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''         End If
''         szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
''      Else
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''         End If
''      End If
''
''      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'''Debug.Print szSQL
''      PopulateTenantLookup szSQL, adoConn
''
''   End If
''
'''  Deleted tenants only
''   If optExTenant.Value Then
''       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
''                    "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
''               "FROM Tenants AS T LEFT JOIN " & _
''                    "[" & _
''                    "SELECT U.UnitName, L.SageAccountNumber " & _
''                    "From Units AS U, LeaseDetails AS L, Property AS P, P.PropertyID, P.ClientID " & _
''                    "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
''                       "L.Status = TRUE "
''
''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
''                 "WHERE (T.Comments) IS NOT NULL OR T.Comments<>'' "
''
''      If txtClientList.tag <> "ALL" Then
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''         End If
''            szSQL = szSQL + "AND IQ.ClientID = '" & txtClientList.tag & "' "
''      Else
''         If txtPropertyList.tag <> "ALL" Then
''            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''         End If
''      End If
''
''      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
''
''      PopulateTenantLookup szSQL, adoConn
''
''   End If
'   '#########################################################################################
''  All tenants
'
'
'   If optBoth.Value Then
'      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                   "'' AS Balance, " & _
'                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
'              "FROM Tenants AS T LEFT JOIN " & _
'                   "[" & _
'                   "SELECT U.UnitName, L.SageAccountNumber, " & _
'                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'                   "FROM Units AS U, LeaseDetails AS L, " & _
'                        "Property AS P " & _
'                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
'                      " U.PropertyID = P.PropertyID "
''L.Status = TRUE AND rem on 20171208
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'                 "WHERE OCCUPIDE_ = FALSE "
'   End If
'
''  Current tenants Only
''Modified by anol 23 Aug 2015
''L.Status = TRUE AND ommited
''AND IsNull(IQ.TerminateDate)
'   If optCurrentTenant.Value Then
'''      strWHR = ""
'''      If txtClientList.Tag <> "ALL" Then
'''           ' If cboPropertyList.ListIndex > -1 Then 'added by anol 09 July 2015
'''                    If txtPropertyList.Tag <> "ALL" Then
'''                        strWHR = strWHR + " IQ.PropertyID = '" & txtPropertyList.Tag & "' "
'''                    End If
'''                    If strWHR <> "" Then
'''                          strWHR = strWHR + "AND "
'''                    End If
'''                    strWHR = strWHR + " IQ.ClientID = '" & txtClientList.Tag & "' "
''''            Else
''''                   strWHR = strWHR + " IQ.ClientID = '" & txtClientList.Tag & "' "
''''            End If
'''
'''      Else
'''            'If cboPropertyList.ListIndex > -1 Then
'''                    If txtPropertyList.Tag <> "ALL" Then
'''
'''                          strWHR = strWHR + " IQ.PropertyID = '" & txtPropertyList.Tag & "' "
'''                    End If
'''             'End If
'''      End If
'''
'''      If strWHR <> "" Then
'''            strWHR = strWHR + " AND "
'''      End If
'''      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'''                   "'' AS Balance, " & _
'''                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'''                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber  " & _
'''              "FROM Tenants AS T LEFT JOIN " & _
'''                   "[" & _
'''                   "SELECT U.UnitName, L.SageAccountNumber, L.TerminateDate, " & _
'''                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'''                   "FROM Units AS U, LeaseDetails AS L, " & _
'''                        "Property AS P " & _
'''                   "WHERE U.UnitNumber = L.UnitNumber AND L.Status AND " & _
'''                      "U.PropertyID = P.PropertyID "
''''added AND L.Status on 20171208 because it was showing double lesseID where some of them had status expired
'''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'''                 "WHERE " & strWHR & " ((T.Comments) IS NULL OR T.Comments = '') AND IsNull(IQ.TerminateDate) AND OCCUPIDE_ = FALSE "
'''             strWHR = ""
'
'    'Implemented by anol 20180325, issse 556
'    szSQL = "Select LA.SageAccountNumber, LA.CompanyName,LA.UnitName,  Balance,Notes,LA.PropertyID,LA.ClientID, LA.UnitNumber,LA.Status  FROM (SELECT T.SageAccountNumber, IQ.UnitName, '' AS Balance, IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber,IQ.Status,CompanyName  FROM Tenants AS T LEFT JOIN (SELECT U.UnitName, L.SageAccountNumber,L.TerminateDate, " & _
'    "P.PropertyID, P.ClientID, U.UnitNumber,L.Status FROM Units AS U, LeaseDetails AS L, Property AS P WHERE U.UnitNumber = L.UnitNumber AND " & _
'    "U.PropertyID = P.PropertyID AND L.Status = TRUE) AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber WHERE ((T.Comments) IS NULL OR T.Comments = '') AND  " & _
'    "OCCUPIDE_ = FALSE ORDER BY T.SageAccountNumber ) as LA LEFT JOIN (  " & _
'    "Select  S.SageAccountNumber from LeaseDetails as S INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM " & _
'    "LeaseDetails GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where   IQ.MaxOfStartDate=S.StartDate AND S.status=false) " & _
'    "as IM ON IM.Sageaccountnumber=LA.sageaccountnumber where IM.Sageaccountnumber IS null " & szWhere & " order by LA.sageaccountnumber; "
'
'           '
'           Debug.Print szSQL
'
'   End If
'
''  Deleted tenants only
'   If optExTenant.Value Then
''   "SELECT SageAccountNumber, CompanyName " & _
''             "FROM Tenants " & _
''             "WHERE Tenants.SageAccountNumber NOT IN " & _
''                 "(SELECT LeaseDetails.SageAccountNumber " & _
''                 "FROM LeaseDetails " & _
''                 "WHERE Status=True) AND " & _
''                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
''             "ORDER BY SageAccountNumber"
'       szSQL = ""
'       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                   "'' AS Balance, " & _
'                    "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'                    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
'               "FROM Tenants AS T INNER JOIN " & _
'                    "[" & _
'                    "SELECT U.UnitName, L.SageAccountNumber, " & _
'                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'                    "FROM Units AS U, LeaseDetails AS L, " & _
'                        "Property AS P " & _
'                    "WHERE U.UnitNumber = L.UnitNumber AND " & _
'                        "L.Status = FALSE AND U.PropertyID = P.PropertyID "
'
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "
''                    If txtClientList.tag <> "ALL" Then
''                         If txtPropertyList.tag <> "ALL" Then
''                            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
''                         Else
''                            szSQL = szSQL + "AND "
''                         End If
''                         szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
''                      Else
''                         If txtPropertyList.tag <> "ALL" Then
''                            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''                         End If
''                      End If
'      strWHR = ""
'      If txtClientList.Tag <> "ALL" Then
'          'If cboPropertyList.ListIndex > -1 Then 'added by anol 09 July 2015
'                If txtPropertyList.Tag <> "ALL" Then
'                    strWHR = strWHR + "WHERE IQ.PropertyID = '" & txtPropertyList.Tag & "' "
'                End If
'                If strWHR <> "" Then
'                    strWHR = strWHR + "AND "
'                Else
'                    strWHR = strWHR + "WHERE "
'                End If
'                strWHR = strWHR + "IQ.ClientID = '" & txtClientList.Tag & "' "
''         Else
''            strWHR = strWHR + "WHERE IQ.ClientID = '" & txtClientList.Tag & "' "
''         End If
'
'      Else
'            'If cboPropertyList.ListIndex > -1 Then
'                If txtPropertyList.Tag <> "ALL" Then
'                      strWHR = strWHR + "WHERE IQ.PropertyID = '" & txtPropertyList.Tag & "' "
'                End If
''                If strWhr <> "" Then
''                    strWhr = strWhr + "AND "
''                Else
''                    strWhr = strWhr + "WHERE "
''                End If
''                'strWhr = strWhr + "IQ.ClientID = '" & txtClientList.tag & "' "
''             End If
'      End If
'      'strWhr = ""
'      szSQL = szSQL + strWHR
'   End If
''   If InStr(1, szSQL, "WHERE") = 0 Then
''
''   End If
''optExTenant.Value = False  added and I have repleced 'AND' by 'WHERE'
''On 09 July 2015
'
''                     If txtClientList.tag <> "ALL" Then
''                         If txtPropertyList.tag <> "ALL" Then
''                            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' AND "
''                         Else
''                            szSQL = szSQL + "AND "
''                         End If
''                         szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
''                      Else
''                         If txtPropertyList.tag <> "ALL" Then
''                            szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyList.tag & "' "
''                         End If
''                      End If
'
'             ' szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'
'              PopulateTenantLookup szSQL, adoConn
'              UpdateBalance
'   '#########################################################################################
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub cboPropertyList_GotFocus()
'   SelTxtInCtrl cboPropertyList
'End Sub

Private Sub chkCombEmail_Click()
   If chkCombEmail.Value Then
      chkEmailDmd.Value = True
      chkEmailSt.Value = True
   End If
End Sub

Private Sub chkEmailDmd_Click()
   If chkEmailDmd.Value = False Then
      chkEmailSC.Value = False
'      chkCombEmail.Value = False
      chkEmailSt.Value = False
   End If
End Sub

Private Sub chkEmailSC_Click()
   If chkEmailSC.Value Then
      chkEmailDmd.Value = True
   End If
End Sub

Private Sub chkEmailSt_Click()
   If chkEmailSt.Value = False Then
      chkCombEmail.Value = False
   End If
End Sub

Private Sub cmbBank_Click()
   cmdNCList.SetFocus
End Sub



Private Sub chkShowOutstanding_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFlxACHistory(adoConn, "")
    adoConn.Close
End Sub

Private Sub cmbDptAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdFund.SetFocus
    End If
End Sub

Private Sub cmdAddDiary_Click()
   If txtTenantID.text = "" Then Exit Sub

   With frmMaintananceDairy
      .CallingForm = "L"          'Calling from lessee form
      .isEdit = False
      .RecordType = "D"
      .lblJobName.Caption = "Diary Name"
      .Label1.Caption = "Diary Entry No."
      Load frmMaintananceDairy
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With

'   Me.Enabled = False
End Sub

Private Sub cmdBank_Click()
   LoadBankGrid

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "BANK"         'BANK MODE
End Sub

Private Sub cmdCancel_Click()
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
'   If NEWMODE_ Then SageCustomerAccCombo cboSageAccountNumber

   NEWMODE_ = False
   SEARCHTenantMODE_ = True

   txtName.Enabled = True
   cmdTenantLookup.Enabled = True
   cmdTenantLookup.Visible = True

   If txtTenantID.text = "" Then Exit Sub

   tabTenant.Enabled = True

   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth

   ComponentInFrameClearMode Me, fmeTenant, ClearOnlyTextBoxes
   txtTenantID.BackColor = &HF1F9EE
   txtTenantID.Locked = True
   cmdTenantLookup.Visible = True
End Sub

'Private Sub cmdCancelBank_Click()
'   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
'End Sub

Private Sub cmdCancelTenantAddress_Click()
   ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
End Sub

Private Sub cmdClientList_Click()
    fmeTenantLookup.Enabled = False
    LoadflxClient

   'tabTenant.Enabled = False 'it is already false by other lessee grid
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = fmeTenantLookup.Left + 500 'tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = fmeTenantLookup.Top + 200 'tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "Client"
End Sub
Private Sub LoadflxClient()
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 70
   flxSupplier(0).ColWidth(1) = 1500
   flxSupplier(0).ColWidth(2) = 2700
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
Private Sub cmdClientLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmClientNew4
   frmClientNew4.LOAD_CLINT_CLIENTID = txtClientID.text
   frmClientNew4.Show
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachmentInCombo cmbFiles, "Tenants", txtTenantID.text
   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdCloseMemo_Click()
   Picture2.Visible = False
   txtUnitMemo.SetFocus
   cmdVAMemo.Visible = True
End Sub



Private Sub cmdFund_Click()
    LoadFundGrid

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = txtDptAmount.Top + 2500
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "FUND"         'BANK MODE
End Sub

Private Sub cmdPropertyList_Click()
    fmeTenantLookup.Enabled = False
    LoadflxProperty

   'tabTenant.Enabled = False 'it is already false by other lessee grid
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = fmeTenantLookup.Left + 500 'tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = fmeTenantLookup.Top + 200 'tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "Property"
End Sub
Private Sub LoadflxProperty()
    flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 70
   flxSupplier(0).ColWidth(1) = 1500
   flxSupplier(0).ColWidth(2) = 2700
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

Private Sub cmdRCCCancel2_Click()
    If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
    txtRCCComments2.Locked = True
    cmdRCCEdit2.Enabled = True
    cmdRCCSave2.Enabled = False
    cmdRCCCancel2.Enabled = False
    Dim adoConn As New ADODB.Connection
   Dim rsComment As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsComment.Open "Select RCCComments2 from tenants where SageAccountNumber='" & txtTenantID.text & "' and (isnull(Comments) Or Comments='')", adoConn, adOpenStatic, adLockReadOnly
   If Not rsComment.EOF Then
       txtRCCComments2.text = IIf(IsNull(rsComment("RCCComments2").Value), "", rsComment("RCCComments2").Value)
   End If
   rsComment.Close
   Set rsComment = Nothing
   adoConn.Close
   Set adoConn = Nothing
   FocusControl txtRCCComments2
End Sub

Private Sub cmdRCCEdit2_Click()
    txtRCCComments2.Locked = False
    cmdRCCEdit2.Enabled = False
    cmdRCCSave2.Enabled = True
    cmdRCCCancel2.Enabled = True
    FocusControl txtRCCComments2
End Sub

Private Sub cmdRCCSave2_Click()
    If SaveComments("Tenants", "RCCComments2", txtRCCComments2.text, "SageAccountNumber", txtTenantID.text) Then
      ShowMsgInTaskBar "The comments2 have been saved successfully."
   End If
   txtRCCComments2.Locked = True
   cmdRCCEdit2.Enabled = True
   cmdRCCSave2.Enabled = False
   cmdRCCCancel2.Enabled = False
End Sub

Private Sub cmdSearch_Click()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        fraSearch.Left = 7335
        fraSearch.Top = 6210
        
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        If cmdSearch.Caption = "Clear Sea&rch" Then
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             fmeLoading.Visible = False
             cmdSearch.Caption = "Sea&rch"
             fraSearch.Visible = False
             Call LoadFlxACHistory(adoConn, "")
        Else
            If fraSearch.Visible = False Then
                fraSearch.Visible = True
                txtSearchNo.SetFocus
            Else
                fraSearch.Visible = False
            End If
        End If
        adoConn.Close
        Set adoConn = Nothing
End Sub

Private Sub Command1_Click()
    Debug.Print time
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFlxACHistory_old(adoConn)
    adoConn.Close
    Debug.Print time
   
End Sub

Private Sub Command2_Click()
    Debug.Print time
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFlxACHistory(adoConn, "")
    adoConn.Close
    Debug.Print time
End Sub

Private Sub flxDeposit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         flxDeposit_RowColChange
    End If
End Sub

Private Sub flxSupplier_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplier_Click (Index)
    End If
End Sub

Private Sub gridLeaseAnalysis_RowColChange()
      txtLeaseAnalysisID.text = gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 1)
      txtUnitMemo.text = gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 5)
      cmdUnitMemoEdit.Enabled = True
      cmdDelete.Enabled = True
End Sub
Private Sub gridLeaseAnalysis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridLeaseAnalysis.ToolTipText = gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.MouseRow, gridLeaseAnalysis.MouseCol)
End Sub
Private Sub cmdCopy_Click()
   If txtTenantID.text = "" Then
      ShowMsgInTaskBar "Please select a Lessee to continue."
      Exit Sub
   End If

   COPYMODE_ = True
   NEWMODE_ = True
   SEARCHTenantMODE_ = False

   txtClient.text = ""
   txtProperty.text = ""
   txtUnit.text = ""
   txtDeposit.text = ""

   txtClient.Enabled = True
   txtProperty.Enabled = True
   txtUnit.Enabled = True
   txtDeposit.Enabled = True

   txtTenantID.Enabled = True
   txtTenantID.Locked = False
   txtName.Enabled = True
   txtCompanyName.Enabled = True

   cmdNew.Enabled = False
   cmdEdit.Enabled = False
   cmdCopy.Enabled = False
   cmdDeleteLessee.Enabled = False

   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   cmdClose.Enabled = True

'   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
   ComponentInFrameClearMode Me, fmeEventHistory, ClearBoth
'   ConfigurFlxEventHistory
   ComponentInFrameClearMode Me, Frame8, ClearBoth
   ComponentInFrameClearMode Me, Frame17(1), ClearBoth
End Sub

Private Sub cmdCopyReceipt_Click()
   If flxACHistory.row < 1 Then
      MsgBox "Please select a transaction from the grid.", vbInformation + vbOKOnly, "Selection"
      flxACHistory.SetFocus
      Exit Sub
   End If
'
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + cmdCopyReceipt.Top + _
'                      Me.Top + tabTenant.Top + 1160
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + _
'                       tabTenant.Left + cmdCopyReceipt.Left + 80

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "SRR" Then
      frmPopUpMenu.CallingFrom "LESSEE_SRR"
   Else
      frmPopUpMenu.CallingFrom "LESSEE_" & Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2)
      
      If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI" Or _
            Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SC" Then
         Dim adoConn As New ADODB.Connection
         Dim adoRST As New ADODB.Recordset
         Dim szSQL As String

         adoConn.Open getConnectionString
         
         szSQL = "SELECT DemandID " & _
                 "FROM DemandRecords " & _
                 "WHERE DmdSlNo = " & _
                     StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & " AND " & _
                    "TransactionType = " & _
                     IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", 1, 2) & ";"
         adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         frmPopUpMenu.TRANSACITON_ID_LONG = CLng(adoRST.Fields.Item(0).Value)
         frmPopUpMenu.TRANS_TYPE = IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", "INV", "CRN")

         adoRST.Close
         Set adoRST = Nothing
         adoConn.Close
         Set adoConn = Nothing
      End If
      If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR" Or _
            Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SA" Then
         frmPopUpMenu.CallingFrom "LESSEE_" & Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2)
      End If
   End If

   frmPopUpMenu.Show
End Sub

Private Sub cmdDelete_Click()
      Picture2.Visible = False
      If gridLeaseAnalysis.row = 0 Then
          ShowMsgInTaskBar "Please select a memo from the list", "Y"
          If gridLeaseAnalysis.Enabled = True Then
               gridLeaseAnalysis.SetFocus
          End If
          Exit Sub
      End If
      If MsgBox("Are you sure to delete memo?", vbQuestion + vbYesNo, "Delete Memo") = vbNo Then Exit Sub
      Dim adoConn As New ADODB.Connection
      adoConn.Open getConnectionString
      adoConn.Execute "DELETE from MemoDetails where MemoID=" & Val(gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 1)) & " and sageaccountNumber='" & txtTenantID.text & "'"
      
      adoConn.Close
      MsgBox "Memo has been deleted successfully", vbInformation + vbOKOnly, "Delete Memo"
      
      PopulateGridLeaseAnalysis
      
      cmdUnitMemoNew.Enabled = True
      cmdUnitMemoEdit.Enabled = True
      cmdUnitMemoSave.Enabled = False
      cmdUnitMemoCancel.Enabled = False
      gridLeaseAnalysis.Enabled = True
      gridLeaseAnalysis.row = 0
      txtUnitMemo.text = ""
      txtLeaseAnalysisID.text = ""
      txtUnitMemo.Locked = True
      fmeTenant.Enabled = True
      cmdDelete.Enabled = False
      Picture2.Visible = True
      txtMemoAll.text = ""
      Call ViewMemo
      txtMemoAll.SetFocus
End Sub
Private Sub ViewMemo()
   'Issue 488
   'Added by anol 04 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtTenantID.text & "' And  MemoType='Lease' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
  Dim strTemp As String
   While Not rstLeaseAnalysis_.EOF
         If Len(rstLeaseAnalysis_!UpdateTime) > 0 Then
               strTemp = " -  "
         Else
               strTemp = ""
         End If
         If Len(txtMemoAll.text) > 0 Then txtMemoAll.text = txtMemoAll.text & vbCrLf & vbCrLf
         txtMemoAll.text = txtMemoAll.text & Left(rstLeaseAnalysis_!UpdateTime, 11) & strTemp & rstLeaseAnalysis_!UserName & vbCrLf & vbCrLf & IIf(IsNull(rstLeaseAnalysis_!MemoDescription) = True, "", rstLeaseAnalysis_!MemoDescription)
         rstLeaseAnalysis_.MoveNext
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   cmdCloseMemo.Refresh
End Sub
Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtTenantID.text, "Tenants"
   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdDeleteLessee_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a Lessee to continue.", vbInformation, "Delete Lessee"
      Exit Sub
   End If

   

'Check for is this Lessee occupying any unit or not
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szNewID As String

   adoConn.Open getConnectionString

'------------------------------------LEASE RECORDS-------------------------------------------------------------
   szSQL = "SELECT Comments " & _
           "FROM Tenants, LeaseDetails " & _
           "WHERE Tenants.SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      MsgBox "This lessee cannot be deleted because a lease agreement exists for the lessee." & Chr(10) & _
             "You must terminated the lease agreement before you delete this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"

      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------DEMAND RECORDS-------------------------------------------------------------
   szSQL = "SELECT DemandID " & _
           "FROM DemandRecords " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "UPDATE_SAGE = True;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      MsgBox "This lessee cannot be deleted because demands exits that have not been updated to SAGE." & Chr(10) & _
             "You must update these demands into SAGE before you can delete this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------RECEIPT TABLE-------------------------------------------------------------
   szSQL = "SELECT * " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "(OSAmount > 0 OR " & _
               "IsSageUpdate = FALSE OR " & _
               "ReceiptView = TRUE);"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      If adoRST!OSAmount > 0 Or adoRST!ReceiptView Then
         MsgBox "This lessee cannot be deleted because are some outstanding amounts in the receipt." & Chr(10) & _
                "You must clear down all outstanding amount of this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"
      ElseIf Not adoRST!IsSageUpdate Then
      'Message changes by anol 21 July 2015
         MsgBox "This lessee record cannot be deleted because there are transactions that have been recorded against it.", vbExclamation + vbOKOnly, "Deleting Lessee"
      End If
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------DELETE LESSEE-------------------------------------------------------------
   szSQL = "SELECT Comments " & _
           "FROM Tenants " & _
           "WHERE Tenants.SageAccountNumber = '" & txtTenantID.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If InStr("DELETED", adoRST.Fields.Item("Comments").Value) > 0 Then
      MsgBox "The lessee has been deleted already.", vbExclamation + vbOKOnly, "Deleting Lessee"
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If
   If MsgBox("Do you sure to delete this Lessee?", vbQuestion + vbYesNo, "Delete Lessee") = vbNo Then Exit Sub
'  Create New id for the Lessee to release the SAGE id
   szNewID = DelTenantID(adoConn)

   UpdateTenantComment adoConn

'  SWAP Lessee's old SAGE id with New ID
'  First swap the id into Tenants table
   SwapID "Tenants", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "DemandRecords", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "LeaseDetails", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "tlbReceipt", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "tblPoA", "SageAccountNumber", txtTenantID.text, szNewID, adoConn

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   MsgBox "The Lessee has been deleted successfully.", vbInformation + vbOKOnly, "Deleting Lessee"

   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
'   SageCustomerAccCombo txtTenantID

   tabTenant.Enabled = False
   cmdTenantLookup.Enabled = True
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
End Sub

Private Sub UpdateTenantComment(adoConn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE Tenants " & _
           "SET Comments ='DELETED' " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "';"
   adoConn.Execute szSQL
End Sub

Private Sub SwapID(szTableName As String, szTableField As String, szOldID As String, szNewID As String, adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT " & szTableField & " " & _
           "FROM " & szTableName & " " & _
           "WHERE " & szTableField & " = '" & szOldID & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   While Not adoRST.EOF
      adoRST.Fields.Item(szTableField).Value = szNewID
      adoRST.Update
      adoRST.MoveNext
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub cmdDelLetter_Click()
   If MsgBox("Do you wish to delete the letter?", vbQuestion + vbYesNo, "Deleting letters") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim iRow As Integer

   adoConn.Open getConnectionString
   
   For iRow = 1 To flxLetters.Rows - 1
      If flxLetters.TextMatrix(iRow, 0) = "X" Then
         adoConn.Execute "DELETE * FROM tlbLetterReports WHERE Id = " & Val(flxLetters.TextMatrix(iRow, 1)) & ";"
      End If
   Next iRow

   LoadFlxLetter adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdDptCancel_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer
   Call disableLeaseHeldBoxed
   bEdit = False
   ButtonHanlding DefaultMode
End Sub

Private Sub cmdDptDelete_Click()
   Dim connCon As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   If MsgBox("Are you sure to delete current deposit?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   connCon.Open getConnectionString

   ButtonHanlding DefaultMode
   populateGroupCombo connCon

   szSQL = "SELECT * " & _
           "FROM TenantDeposit " & _
           "WHERE RefundRef = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"

   adoRST.Open szSQL, connCon, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
         MsgBox "This deposit could not be deleted. This deposit has refund deposit.", vbCritical + vbOKOnly, "Delete Not Possible"
         adoRST.Close
         Set adoRST = Nothing
         connCon.Close
         Set connCon = Nothing
         Exit Sub
   End If

   adoRST.Close

   connCon.Execute "UPDATE TenantDeposit SET Status = false, Deleted = true WHERE DepositID = '" & _
            flxDeposit.TextMatrix(flxDeposit.row, 1) & "' OR RefundRef = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
   LoadFlxDeposit connCon, 0, "ASC"

   Set adoRST = Nothing
   connCon.Close
   Set connCon = Nothing

   MsgBox "Deposit has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"
End Sub

Private Sub cmdDptEdit_Click()
   Call EnableLeaseHeldBoxed
   bEdit = True
   If Val(txtDptAmount.text) <> Val(txtOSDpt.text) And flxDeposit.TextMatrix(flxDeposit.row, 7) = "Deposit" Then
      MsgBox "You cannot edit this transaction. Refund has been booked against this deposit.", vbCritical + vbOKOnly, "Edit Deposit"
      Exit Sub
   End If
'  Check the bank reconciliation
   If flxDeposit.TextMatrix(flxDeposit.row, 19) <> "" Then
      MsgBox "You cannot edit this transaction. Bank receipt of the deposit has been reconciled.", vbCritical + vbOKOnly, "Edit Deposit"
      Exit Sub
   End If

   cCurDepAmt = CCur(txtDptAmount.text)
   ButtonHanlding EditMode
    bEdit = True
   If Left(flxDeposit.TextMatrix(flxDeposit.row, 4), 1) = "D" Then
        yDEPOSIT = 1
        'txtDataMode.text = "1"
   ElseIf Left(flxDeposit.TextMatrix(flxDeposit.row, 4), 1) = "E" Then
        yDEPOSIT = 4
   ElseIf Left(flxDeposit.TextMatrix(flxDeposit.row, 4), 1) = "R" Then
        yDEPOSIT = 3
        ' txtDataMode.text = "3"
   End If
   If flxDeposit.TextMatrix(flxDeposit.row, 4) = "RF" Then
        bFullRefund = True
   Else
        bFullRefund = False
   End If
   If flxDeposit.TextMatrix(flxDeposit.row, 7) = "Refund" Then
      cOriRAmt = CCur(txtDptAmount.text)
      'yDEPOSIT = 31
   End If
'   If flxDeposit.TextMatrix(flxDeposit.row, 6) = "Full Refund" Then
'      cOriRAmt = CCur(txtDptAmount.text)
      'yDEPOSIT = 32
'   End If
   If flxDeposit.TextMatrix(flxDeposit.row, 7) = "Expenses" Then
      cOriRAmt = CCur(txtDptAmount.text)
      'yDEPOSIT = 41
   End If

   iGridRow = flxDeposit.row
End Sub

Private Sub cmdDptRefund_Click()
   If flxDeposit.TextMatrix(flxDeposit.row, 7) <> "Deposit" Then
      MsgBox "Select a deposit to refund.", vbInformation + vbOKOnly, "Refund - Deposit"
      Exit Sub
   End If

   If Val(flxDeposit.TextMatrix(iGridRow, 10)) <= 0 Then
      MsgBox "Deposit has been refunded fully.", vbExclamation + vbOKOnly, "Refund - Deposit"
      Exit Sub
   End If

    
    
    
   txtDepositType.text = ""
   txtDptAmount.text = txtOSDpt.text
   txtDptDetails.text = ""
   ButtonHanlding RefundMode
   Call EnableLeaseHeldBoxed
   txtDate.text = Format(Date, "dd/mm/yyyy")
   yDEPOSIT = 3

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer
   FocusControl txtDate

End Sub

Private Sub cmdDptExpenses_Click()
    Call EnableLeaseHeldBoxed
    If flxDeposit.TextMatrix(flxDeposit.row, 7) <> "Deposit" Then
       MsgBox "Please select a deposit to book an expense.", vbInformation + vbOKOnly, "Expenses - Deposit"
       Exit Sub
    End If
    yDEPOSIT = 4
    txtDepositType.text = ""
    txtDepositType.Tag = ""
    txtDptDetails.text = ""
    txtDptAmount.text = ""
    txtDptDetails.text = ""
 
    txtDptAmount.text = ""
    txtDptAmtType.text = ""
    txtDptAmtType.Tag = ""
    txtDptAmount.text = ""
    txtDptAmount.text = txtOSDpt.text

   ButtonHanlding ExpensesMode
   txtDate.text = Format(Date, "dd/mm/yyyy")
   

   'Load default Bank Code and Nominal Code
   FocusControl txtDate
   If Not bLeaseSetup Then Exit Sub

   txtDptAmount.text = ""

End Sub

Private Sub cmdDptNew_Click()
   Frame1(3).Enabled = True
   Call EnableLeaseHeldBoxed
   ButtonHanlding NewEntryMode
   optNewGroup.Value = True
   FocusControl txtDate

   yDEPOSIT = 1
'   txtDataMode.text = "1"

   'Load default Bank Code and Nominal Code
   If Not bLeaseSetup Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT C.spare1, C.spare2, N.Name " & _
           "FROM Client AS C LEFT OUTER JOIN NominalLedger AS N ON C.spare2 = N.Code " & _
           "WHERE C.ClientID = '" & txtClientID.text & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   txtBank.Tag = IIf(IsNull(adoRST.Fields.Item("spare1").Value), "", adoRST.Fields.Item("spare1").Value)
   txtDNC(1).ToolTipText = IIf(IsNull(adoRST.Fields.Item("spare2").Value), "", adoRST.Fields.Item("spare2").Value)
   txtDNC(1).text = IIf(IsNull(adoRST.Fields.Item("Name").Value), "", adoRST.Fields.Item("Name").Value)
   adoRST.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdDptPrint_Click()
   If flxDeposit.TextMatrix(iGridRow, 1) = "" Or flxDeposit.TextMatrix(iGridRow, 11) <> "" Then
      If flxDeposit.TextMatrix(iGridRow, 11) = "" Then
'         MsgBox "This is a refund.", vbInformation + vbOKOnly, "Print - Deposit"
'      Else
         MsgBox "Please select a deposit transaction from the grid.", vbCritical + vbOKOnly, "Print - Deposit"
      End If
      'Exit Sub
   End If
   ' Passing the from and to date values to Crystal Reports
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\DepositReceipt.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue flxDeposit.TextMatrix(iGridRow, 1)
   Report.ParameterFields(2).AddCurrentValue CInt(flxDeposit.TextMatrix(iGridRow, 12))
   Report.ParameterFields(3).AddCurrentValue txtTenantID.text

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub UpdateDepositHeldBalance()
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    adoConn.Open getConnectionString
     szSQL = "SELECT TenantDeposit.TenantID, (SUM(DptAmount) - " & _
                 "iif(isnull(DepositRefund.TotalDepositRefund),0,DepositRefund.TotalDepositRefund)) as TotalDepositHeld " & _
           "FROM TenantDeposit LEFT JOIN [SELECT SUM(DptAmount) AS TotalDepositRefund, TenantID " & _
              "FROM TenantDeposit " & _
              "WHERE DptRefund = True and Status = true and Deleted = false " & _
              "GROUP BY tenantid]. AS DepositRefund ON  TenantDeposit.TenantID = DepositRefund.TenantID " & _
           "WHERE TenantDeposit.TenantID = '" & txtTenantID.text & "' AND " & _
                 "TenantDeposit.Deleted = FALSE AND DptRefund = FALSE " & _
           "GROUP BY TenantDeposit.TenantID, DepositRefund.TotalDepositRefund;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      txtDeposit.text = "0.00"
   Else
      txtDeposit.text = Format(IIf(IsNull(adoRST.Fields.Item("TotalDepositHeld").Value), 0, adoRST.Fields.Item("TotalDepositHeld").Value), "0.00")
   End If

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
End Sub

Private Function MAXSL(sz As String) As Long
    Dim adoConn       As New ADODB.Connection
    Dim rsRst       As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsRst.Open "Select max(DepositSL) as SL from TenantDeposit where DepositTypePrefix='" & sz & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsRst.EOF Then
            MAXSL = IIf(IsNull(rsRst("SL")), 0, rsRst("SL")) + 1
    End If
    rsRst.Close
    adoConn.Close
End Function
Private Sub cmdDptSave_Click()
   'If yDEPOSIT = 0 Then Exit Sub
   If txtBank.text = "" Then
      MsgBox "Please select a Bank Code from the drop down list.", vbCritical + vbOKOnly, "Deposit"
      cmdBank.SetFocus
      Exit Sub
   End If
   If txtDNC(0).text = "" Then
      MsgBox "Please select a Nominal Code from the drop down list.", vbCritical + vbOKOnly, "Deposit"
      cmdNCList.SetFocus
      Exit Sub
   End If
   If txtDate.text = "" Then
      MsgBox "Please enter the Date.", vbCritical + vbOKOnly, "Deposit"
      FocusControl txtDate
      Exit Sub
   End If
   If txtDptAmtType.text = "" Then
      MsgBox "Please select a Amount Type from the list", vbCritical + vbOKOnly, "Deposit"
      FocusControl cmdDptAmtType
      Exit Sub
   End If
   If txtDptAmount.text = "" Or Val(txtDptAmount.text) <= 0 And bFullRefund = False Then
      MsgBox "Please enter correct Amount.", vbCritical + vbOKOnly, "Deposit"
      txtDptAmount.SetFocus
      Exit Sub
   End If
   If txtDepositType.Tag = "" Then
      MsgBox "Please select type.", vbCritical + vbOKOnly, "Deposit Type"
      cmdDepositType.SetFocus
      Exit Sub
   End If
   If optExitingGroup.Value = True And cboGroup.text = "" Then
      MsgBox "Type the group number.", vbCritical + vbOKOnly, "Group Number"
      optExitingGroup.Enabled = True
      FocusControl cboGroup
      Exit Sub
   End If
   If txtFund.text = "" Then
      MsgBox "Select the fund.", vbCritical + vbOKOnly, "fund Number"
      cmdFund.SetFocus
      Exit Sub
   End If
   If txtOSDpt.text = "" Then
        txtOSDpt.text = "0.00"
   End If
'   If cboGroup.text = "" Then
'        MsgBox "Select the Group.", vbCritical + vbOKOnly, "Group Number"
'        'cboGroup.SetFocus
'        FocusControl cboGroup
'        Exit Sub
'   End If
  
   Call disableLeaseHeldBoxed
   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
   Dim szSQL         As String
   Dim btYesNo       As Byte
   Dim szID          As String
    
   adoConn.Open getConnectionString
   btYesNo = 0
'///////////////////////////////////// ADD NEW DEPOSIT /////////////////////////////////////////////////////
        If bEdit = False Then 'this means add new/save mode
                  Dim GP_ID As Integer
            
                  szSQL = "SELECT MAX(GROUPNO) AS GPNO FROM TENANTDEPOSIT WHERE TenantID = '" & cboSageAccountNumber.text & "';"
                  adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
            
                  If (optNewGroup.Value = True) Then
                     GP_ID = CInt(IIf(IsNull(adoRST.Fields.Item("GPNO").Value), 0, adoRST.Fields.Item("GPNO").Value))
                     GP_ID = GP_ID + 1
                  Else
                     GP_ID = CInt(cboGroup.text)
                  End If
                  adoRST.Close
                  
                   If yDEPOSIT = 1 Then
                      
                
                      szSQL = "SELECT * FROM TENANTDEPOSIT;"
                      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                      With adoRST
                         .AddNew
                         szID = UniqueID()
                         .Fields.Item("DepositID").Value = szID
                         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
                         .Fields.Item("BankCode").Value = txtBank.Tag
                         .Fields.Item("NominalCode").Value = txtDNC(0).text
                         .Fields.Item("DepositDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                         .Fields.Item("DptType").Value = txtDepositType.Tag
                         .Fields.Item("DptAmtType").Value = txtDptAmtType.text
                         .Fields.Item("DptDetails").Value = txtDptDetails.text
                         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)
                         .Fields.Item("DptVatRate").Value = 0
                         .Fields.Item("DptRefund").Value = False
                         .Fields.Item("BCNAME").Value = txtBank.text
                         .Fields.Item("NCNAME").Value = txtDNC(1).text
                         .Fields.Item("BankTransaction").Value = False               '? what is this for?
                         .Fields.Item("OSRefund").Value = CCur(txtDptAmount.text)
                         .Fields.Item("GroupNo").Value = GP_ID
                         .Fields.Item("TransactionID").Value = "D" & CStr(TransID("D") + 1)
                         .Fields.Item("DepositTypePrefix").Value = "DP"
                         .Fields.Item("DepositSL").Value = MAXSL("DP")
                         .Fields.Item("FundID").Value = txtFund.Tag
                         .Update
                         .Close
                      End With
                      Set adoRST = Nothing
                      LoadFlxDeposit adoConn, 0, "ASC"
                      btYesNo = MsgBox("Deposit has been saved successfully." & Chr(13) & "Do you want to print the receipt now?", vbInformation + vbYesNo, "Deposit")
                       FocusControl cmdDptNew
                '     Export the transaction as Bank Recipt
                      Export2BankReceipt szID, adoConn
                   End If
                '////////////////////////////////////// EDIT DEPOSIT ///////////////////////////////////////////////////////
                   If yDEPOSIT = 2 Then
                      Dim RemovedGrpNo As String
                      If (GROUP_NO <> cboGroup.text) Then
                         RemovedGrpNo = GROUP_NO
                         GROUP_NO = cboGroup.text
                      End If
                
                      szSQL = "SELECT * FROM TENANTDEPOSIT " & _
                              "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
                      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                      With adoRST
                         If Left(flxDeposit.TextMatrix(iGridRow, 3), 1) <> "D" Then
                            'rem by anol 2021-12-04
                          '  adoconn.Execute "UPDATE TenantDeposit " & _
                              '              "SET OSRefund = OSRefund + " & .Fields.Item("DptAmount").Value & " - " & _
                                '                 Val(txtDptAmount.text) & " " & _
                                '                 "WHERE DepositID = " & .Fields.Item("RefundRef").Value & ";"
                         End If
                         .Fields.Item("BankCode").Value = txtBank.Tag
                         .Fields.Item("NominalCode").Value = txtDNC(0).text
                         .Fields.Item("DepositDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                         .Fields.Item("DptType").Value = txtDepositType.Tag
                         .Fields.Item("DptAmtType").Value = txtDptAmtType.text
                         .Fields.Item("DptDetails").Value = txtDptDetails.text
                         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)
                         .Fields.Item("DptVatRate").Value = 0
                         .Fields.Item("BCNAME").Value = txtBank.text
                         .Fields.Item("NCNAME").Value = txtDNC(1).text
                         .Fields.Item("OSRefund").Value = CCur(IIf(txtOSDpt.text = "", 0, txtOSDpt.text))
                         .Fields.Item("FundID").Value = txtFund.Tag
                          .Fields.Item("GROUPNO").Value = cboGroup.text
                         .Update
                         .Close
                      End With

                      Set adoRST = Nothing
                      LoadFlxDeposit adoConn, 0, "ASC"
                      MsgBox "The Modifications have been saved successfully.", vbInformation, "Modifications Saved"
                      btYesNo = vbNo
                
                      UpdateBankReceipt flxDeposit.TextMatrix(iGridRow, 1), adoConn
                   End If
                '////////////////////////////////////// REFUND DEPOSIT /////////////////////////////////////////////////////
                   If yDEPOSIT = 3 Then
                      szSQL = "SELECT * FROM TENANTDEPOSIT;"
                      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                      With adoRST
                         .AddNew
                         szID = UniqueID()
                         .Fields.Item("DepositID").Value = szID
                         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
                         .Fields.Item("BankCode").Value = txtBank.Tag
                         .Fields.Item("NominalCode").Value = txtDNC(0).text
                         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
                         .Fields.Item("DptType").Value = txtDepositType.Tag
                         .Fields.Item("DptAmtType").Value = txtDptAmtType.text                     'TRANSACTION AMT TYPE
                         .Fields.Item("DptDetails").Value = txtDptDetails.text
                         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
                         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
                         .Fields.Item("RefundRef").Value = flxDeposit.TextMatrix(iGridRow, 1) 'Deposit ID
                         .Fields.Item("BCNAME").Value = txtBank.text
                         .Fields.Item("NCNAME").Value = txtDNC(1).text
                         .Fields.Item("GroupNo").Value = GP_ID 'cboGroup.text
                         .Fields.Item("DptRefund").Value = True
                          If bFullRefund = False Then
                                 .Fields.Item("TransactionID").Value = "R" & CStr(TransID("R") + 1)
                                 .Fields.Item("DepositTypePrefix").Value = "DR"
                                 .Fields.Item("DepositSL").Value = MAXSL("DR")
                          Else
                                .Fields.Item("TransactionID").Value = "RF"
                                .Fields.Item("DepositTypePrefix").Value = "FR"
                                .Fields.Item("DepositSL").Value = MAXSL("FR")
                          End If
                         
                         .Fields.Item("FundID").Value = txtFund.Tag
                         .Update
                         .Close
                      End With
                      Set adoRST = Nothing
                      If bFullRefund = False Then
                                adoConn.Execute "UPDATE TenantDeposit " & _
                                      "SET OSRefund = " & _
                                           Val(flxDeposit.TextMatrix(iGridRow, 10)) - Val(txtDptAmount.text) & " " & _
                                           "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
                       Else
                                     Dim cTotalPay As Double
                                     Dim cTotalPayTemp As Double
                                     Dim rsUpdateOsAmount As New ADODB.Recordset
                                             szSQL = "SELECT sum(OSRefund) as amt " & _
                                                     "FROM TenantDeposit AS D " & _
                                                     "Where D.TenantID = '" & txtTenantID.text & "' AND D.Deleted = False AND OSRefund>0 ;"
                                                 rsUpdateOsAmount.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                                 
                                                  szSQL = "SELECT OSRefund,ParentRFID " & _
                                                     "FROM TenantDeposit AS D " & _
                                                     "Where D.TenantID = '" & txtTenantID.text & "' AND D.Deleted = False AND OSRefund>0 order by DepositDate desc;"
                                                     adoRST.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
                                                  cTotalPayTemp = IIf(IsNull(rsUpdateOsAmount!amt), 0, rsUpdateOsAmount!amt)
                                                If cTotalPayTemp > 0 Then
                                                         cTotalPay = txtDptAmount.text
                                                         While cTotalPay > 0 And Not adoRST.EOF
                                                               If cTotalPay >= CCur(adoRST.Fields.Item("OSRefund").Value) Then
                                                                     cTotalPay = cTotalPay - CCur(adoRST.Fields.Item("OSRefund").Value)
                                                                     adoRST.Fields.Item("OSRefund").Value = 0
                                                                     adoRST.Fields.Item("ParentRFID").Value = szID ' Here I am putting full refund ID to all deposits
                                                               Else
                                                                     adoRST.Fields.Item("OSRefund").Value = adoRST.Fields.Item("OSRefund").Value - cTotalPay
                                                                      adoRST.Fields.Item("ParentRFID").Value = szID
                                                                     cTotalPay = 0
                                                              End If
                                                 
                                                              adoRST.Update
                                                              adoRST.MoveNext
                                                         Wend
                                               End If
                                               adoRST.Close
                                               rsUpdateOsAmount.Close
                      End If
                      LoadFlxDeposit adoConn, 0, "ASC"
                      MsgBox "Refund has been saved successfully.", vbInformation, "Refund Saved"
                      btYesNo = vbNo
                      Export2BankPayment szID, adoConn
                   End If
                   '////////////////////////////////////// EXPENSES  ///////////////////////////////////////////////////////////
                   If yDEPOSIT = 4 Then
                      szSQL = "SELECT * FROM TENANTDEPOSIT;"
                      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                      With adoRST
                         .AddNew
                         szID = UniqueID()
                         .Fields.Item("DepositID").Value = szID
                         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
                         .Fields.Item("BankCode").Value = txtBank.Tag
                         .Fields.Item("NominalCode").Value = txtDNC(0).text
                         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
                         .Fields.Item("DptType").Value = txtDepositType.Tag
                         .Fields.Item("DptAmtType").Value = txtDptAmtType.text                     'TRANSACTION AMT TYPE
                         .Fields.Item("DptDetails").Value = txtDptDetails.text
                         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'EXPENSES
                         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
                         .Fields.Item("RefundRef").Value = flxDeposit.TextMatrix(iGridRow, 1) 'Deposit ID
                         .Fields.Item("BCNAME").Value = txtBank.text
                         .Fields.Item("NCNAME").Value = txtDNC(1).text
                         .Fields.Item("GroupNo").Value = cboGroup.text
                         .Fields.Item("DptRefund").Value = True
                         .Fields.Item("TransactionID").Value = "E" & CStr(TransID("E") + 1)
                         .Fields.Item("DepositTypePrefix").Value = "EX"
                         .Fields.Item("DepositSL").Value = MAXSL("EX")
                         .Fields.Item("FundID").Value = txtFund.Tag
                         .Update
                         .Close
                      End With
                      Set adoRST = Nothing
                      adoConn.Execute "UPDATE TenantDeposit " & _
                                      "SET OSRefund = " & _
                                           Val(flxDeposit.TextMatrix(iGridRow, 10)) - Val(txtDptAmount.text) & " " & _
                                           "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
                      LoadFlxDeposit adoConn, 0, "ASC"
                      MsgBox "The Expenses has been saved successfully.", vbInformation, "Expenses Saved"
                      btYesNo = vbNo
                
                      Export2BankPayment szID, adoConn
                   End If

    
        Else '////////////////////////////////////// EDIT REFUND /////////////////////////////////////////////////////
        
                       If yDEPOSIT = 1 Or yDEPOSIT = 3 Then
                                  szSQL = "SELECT * FROM TENANTDEPOSIT " & _
                                          "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
                                  adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                  With adoRST
                                     .Fields.Item("TenantID").Value = cboSageAccountNumber.text
                                     .Fields.Item("BankCode").Value = txtBank.Tag
                                     .Fields.Item("NominalCode").Value = txtDNC(0).text
                                     .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
                                     .Fields.Item("DptType").Value = txtDepositType.Tag
                                     .Fields.Item("DptAmtType").Value = txtDptAmtType.text                     'TRANSACTION AMT TYPE
                                     .Fields.Item("DptDetails").Value = txtDptDetails.text
                                     .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
                                     .Fields.Item("OSRefund").Value = CCur(txtOSDpt.text)                  'REFUND TOTAL
                                     .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
                                     szID = IIf(IsNull(.Fields.Item("RefundRef").Value), "", .Fields.Item("RefundRef").Value)                                   'Deposit ID
                                     .Fields.Item("BCNAME").Value = txtBank.text
                                     .Fields.Item("NCNAME").Value = txtDNC(1).text
                                     .Fields.Item("GroupNo").Value = cboGroup.text
                                     If yDEPOSIT = 1 Then
                                        .Fields.Item("DptRefund").Value = False
                                     Else
                                        .Fields.Item("DptRefund").Value = True
                                     End If
                                     .Fields.Item("FundID").Value = txtFund.Tag
                                     If yDEPOSIT = 1 Then
                                           .Fields.Item("TransactionID").Value = "D1"
                                     ElseIf yDEPOSIT = 3 Then
                                           ' .Fields.Item("TransactionID").Value = "R1"
                                     End If
                                     .Update
                                     .Close
                                  End With
                                  Set adoRST = Nothing
                                  If bFullRefund = False Then
                                                 adoConn.Execute "UPDATE TenantDeposit " & _
                                                  "SET OSRefund = OSRefund + " & _
                                                       cOriRAmt - Val(txtDptAmount.text) & " " & _
                                                       "WHERE DepositID = '" & szID & "';"
                                  End If
                                  If bFullRefund = True And yDEPOSIT = 3 Then
                                    'cascade os amount update
                                            Dim cAmtToDeduct As Double
                                            Dim cAmtToDeductTemp As Double
'                                            Dim rsUpdateOsAmount As New ADODB.Recordset
'                                            szSQL = "SELECT sum(DptAmount-OSRefund) as amt " & _
'                                                     "FROM TenantDeposit AS D " & _
'                                                     "Where D.TenantID = '" & txtTenantID.text & "' AND D.Deleted = False AND dptRefund=false AND  DptAmount-OSRefund>0 group by DepositDate order by DepositDate desc;"
'                                            rsUpdateOsAmount.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                                            If Not rsUpdateOsAmount.EOF Then
'                                                cAmtToDeductTemp = IIf(IsNull(rsUpdateOsAmount("amt").Value), 0, rsUpdateOsAmount("amt").Value)
'                                            End If
'
'                                            rsUpdateOsAmount.Close
                                             cAmtToDeductTemp = cOriRAmt - Val(txtDptAmount.text)
                                             'adoRst.Close flxDeposit.TextMatrix(iGridRow, 1) is the parent full Refund ID
                                             'which is now selecting all related deposites now which shall be updated with new os refund amount, flxDeposit.TextMatrix(iGridRow, 1) is the parent Deposit ID
                                            szSQL = "SELECT OSRefund,DptAmount FROM TenantDeposit AS D " & _
                                                     "Where D.TenantID = '" & txtTenantID.text & "' AND ParentRFID ='" & _
                                                     flxDeposit.TextMatrix(iGridRow, 1) & "' AND D.Deleted = False AND dptRefund=false order by DepositDate desc;"
                                            adoRST.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
                                                     
                                            If cAmtToDeductTemp > 0 Then
                                                         cAmtToDeduct = cAmtToDeductTemp
'                                                         cAmtToDeduct = cOriRAmt - Val(txtDptAmount.text)
'adoRst.MoveFirst

                                                         While cAmtToDeduct > 0 And Not adoRST.EOF
                                                               If cAmtToDeduct >= adoRST.Fields.Item("DptAmount").Value - adoRST.Fields.Item("OSRefund").Value Then
                                                                     cAmtToDeduct = cAmtToDeduct - adoRST.Fields.Item("DptAmount").Value + adoRST.Fields.Item("OSRefund").Value
                                                                     adoRST.Fields.Item("OSRefund").Value = 0
                                                               Else
                                                                     adoRST.Fields.Item("OSRefund").Value = adoRST.Fields.Item("DptAmount").Value - cAmtToDeduct
                                                                     cAmtToDeduct = 0
                                                              End If
                                                 
                                                              adoRST.Update
                                                              adoRST.MoveNext
                                                         Wend
'                                                         While cAmtToDeduct > 0 And Not adoRst.EOF
'                                                               If cAmtToDeduct >= adoRst.Fields.Item("DptAmount").Value - adoRst.Fields.Item("OSRefund").Value Then
'                                                                     cAmtToDeduct = cAmtToDeduct - adoRst.Fields.Item("OSRefund").Value
'                                                                     adoRst.Fields.Item("OSRefund").Value = 0
'                                                               Else
'                                                                     adoRst.Fields.Item("OSRefund").Value = cAmtToDeduct
'                                                                     cAmtToDeduct = 0
'                                                              End If
'
'                                                              adoRst.Update
'                                                              adoRst.MoveNext
'                                                         Wend
                                             End If
                                             adoRST.Close
                                  End If
                                  'ShowMsgInTaskBar "Refund has been saved successfully.", "Y", "P"
                                  LoadFlxDeposit adoConn, 0, "ASC"
                                  If yDEPOSIT = 1 Then
                                        UpdateBankReceipt flxDeposit.TextMatrix(iGridRow, 1), adoConn
                                        MsgBox "Tenant Deposit has been saved successfully.", vbInformation, "Tenant Deposit Saved"
                                  Else
                                        UpdateBankPayment flxDeposit.TextMatrix(iGridRow, 1), adoConn
                                        MsgBox "Refund has been saved successfully.", vbInformation, "Refund Saved"
                                  End If
                                  'UpdateBankPayment flxDeposit.TextMatrix(iGridRow, 1), adoconn
                       End If
                       '////////////////////////////////////// EDIT EXPENSES /////////////////////////////////////////////////////
                   If yDEPOSIT = 4 Then
                          szSQL = "SELECT * FROM TENANTDEPOSIT " & _
                                  "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 1) & "';"
                          adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                    
                          With adoRST
                             .Fields.Item("TenantID").Value = cboSageAccountNumber.text
                             .Fields.Item("BankCode").Value = txtBank.Tag
                             .Fields.Item("NominalCode").Value = txtDNC(0).text
                             .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
                             .Fields.Item("DptType").Value = txtDepositType.Tag
                             .Fields.Item("DptAmtType").Value = txtDptAmtType.text                     'TRANSACTION AMT TYPE
                             .Fields.Item("DptDetails").Value = txtDptDetails.text
                             .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
                             .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
                             szID = .Fields.Item("RefundRef").Value                                     'Deposit ID
                             .Fields.Item("BCNAME").Value = txtBank.text
                             .Fields.Item("NCNAME").Value = txtDNC(1).text
                             .Fields.Item("GroupNo").Value = cboGroup.text
                             .Fields.Item("DptRefund").Value = True
                             .Fields.Item("FundID").Value = txtFund.Tag
                             .Fields.Item("TransactionID").Value = "E1"
                             .Update
                             .Close
                          End With
                          Set adoRST = Nothing
                
                          adoConn.Execute "UPDATE TenantDeposit " & _
                                      "SET OSRefund = OSRefund + " & _
                                           cOriRAmt - Val(txtDptAmount.text) & " " & _
                                           "WHERE DepositID = '" & szID & "';"
                        'ShowMsgInTaskBar "Expenses has been updated successfully.", "Y", "P"
                        LoadFlxDeposit adoConn, 0, "ASC"
                        MsgBox "Expenses has been saved successfully.", vbInformation, "Expenses Saved"
                        UpdateBankPayment flxDeposit.TextMatrix(iGridRow, 1), adoConn
                   End If
                '///////////////////////////////////////////////////////////////////////////////////////////////////////////
    End If
        bEdit = False
        bFullRefund = False
        '   LoadFlxDeposit adoConn
        LoadComboes adoConn
        populateGroupCombo adoConn
        
        If yDEPOSIT = 2 And btYesNo = 6 Then cmdDptPrint_Click           'Edit Deposit
        
        If yDEPOSIT = 1 And btYesNo = 6 Then          'Deposit
              flxDeposit.row = flxDeposit.Rows - 1
              iGridRow = flxDeposit.Rows - 1
            
              cmdDptPrint_Click
        End If
        adoConn.Close
        adoConn.Open getConnectionString
        adoConn.BeginTrans
        '  Export Transactions to Nominal Ledger (NLPosting table)
        If Export_BPnBR_2_NL(adoConn) = True Then
            adoConn.CommitTrans
            'ShowMsgInTaskBar "The Transaction has been saved.", "Y", "P"
            MsgBox "The Transaction has been saved.", vbInformation, "Transaction Saved"
        Else
            adoConn.RollbackTrans
            MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
        End If
        txtFund.text = ""
        txtFund.Tag = ""
        adoConn.Close
        Set adoConn = Nothing
        Call UpdateDepositHeldBalance
        ButtonHanlding DefaultMode
        FocusControl cmdDptNew
End Sub

Private Sub Export2BankPayment(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset
   
   szStr = "SELECT * FROM tlbBankPayment;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = txtClientID.text
      .Fields.Item("BANK_AC").Value = txtBank.Tag
      .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
      .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
      .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
      .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
      .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
      .Fields.Item("DEPT_ID").Value = txtFund.Tag
      .Fields.Item("TRAN_DATE").Value = txtDate.text
      .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
      .Fields.Item("TAX_CODE").Value = "T9"
      .Fields.Item("VAT").Value = 0
      .Fields.Item("TransactionType").Value = 11
      .Fields.Item("TRANS").Value = "BP"
      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Fields.Item("TenantDeposit").Value = szID
      .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
      If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
         .Fields.Item("CT").Value = "C"
      End If

      .Update

      If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = txtBank.Tag
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("CT").Value = "C"
         .Fields.Item("TenantDeposit").Value = szID
         .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
         .Update
      End If

      .Close
   End With

   Set adoRST = Nothing
End Sub

Private Sub Export2BankReceipt(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset

   szStr = "SELECT * FROM tlbBankPayment;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = txtClientID.text
      .Fields.Item("BANK_AC").Value = txtBank.Tag
      .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
      .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
      .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
      .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
      .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
      .Fields.Item("DEPT_ID").Value = txtFund.Tag
      .Fields.Item("TRAN_DATE").Value = txtDate.text
      .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
      .Fields.Item("TAX_CODE").Value = "T9"
      .Fields.Item("VAT").Value = 0
      .Fields.Item("TransactionType").Value = 12
      .Fields.Item("TRANS").Value = "BR"
      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Fields.Item("TenantDeposit").Value = szID
       .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")

      If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
         .Fields.Item("CT").Value = "C"
      End If

      .Update

      If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = txtBank.Tag
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text  '& - Adams01
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("CT").Value = "C"
         .Fields.Item("TenantDeposit").Value = szID
         .Update
      End If

      .Close
   End With

   Set adoRST = Nothing
End Sub

Private Sub UpdateBankPayment(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset
   
   szStr = "SELECT * FROM tlbBankPayment " & _
           "WHERE TenantDeposit = '" & szID & "' AND " & _
                 "TransactionType = 11;"
'   szStr = "SELECT * FROM tlbBankPayment " & _
'           "WHERE TenantDeposit = '" & szID & "'"


   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      
      If Not adoRST.EOF Then
        'here delete from NLposting
         adoConn.Execute "UPDATE NLPOSTING SET DeleteFlag=1 where PARENT_RECORD ='" & adoRST.Fields.Item("MY_ID").Value & "'"
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = txtBank.Tag
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("TenantDeposit").Value = szID
          .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
         If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
            .Fields.Item("CT").Value = "C"
         End If
          .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
          .Fields.Item("NLpost").Value = False
         .Update
         
      End If
   adoRST.Close
      szStr = "SELECT * FROM tlbBankPayment " & _
              "WHERE TenantDeposit = '" & szID & "' AND " & _
                    "TransactionType = 12;"

      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      If Not adoRST.EOF Then                                         'Contra Transation FOUND
            'Here Delete From NLPosting
            adoConn.Execute "UPDATE NLPOSTING SET DeleteFlag=1 where  PARENT_RECORD ='" & adoRST.Fields.Item("MY_ID").Value & "'"
                If .Fields.Item("BANK_AC").Value = .Fields.Item("NOMINAL_CODE").Value Then
                   If txtBank.Tag <> txtDNC(0).text Then
                      adoConn.Execute "DELETE * FROM tlbBankPayment " & _
                                      "WHERE TenantDeposit = '" & szID & "' AND " & _
                                            "TransactionType = 11;"
                      .Close
                   Else
                      .Fields.Item("ClientID").Value = txtClientID.text
                      .Fields.Item("BANK_AC").Value = txtBank.Tag
                      .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
                      .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
                      .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
                      .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
                      .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
                      .Fields.Item("DEPT_ID").Value = txtFund.Tag
                      .Fields.Item("TRAN_DATE").Value = txtDate.text
                      .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
                      .Fields.Item("TAX_CODE").Value = "T9"
                      .Fields.Item("VAT").Value = 0
                      .Fields.Item("TransactionType").Value = 12
                      .Fields.Item("TRANS").Value = "BR"
                      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
                      .Fields.Item("CT").Value = "C"
                      .Fields.Item("TenantDeposit").Value = szID
                      .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
                      .Fields.Item("NLpost").Value = False

                      .Update
                      .Close
                   End If
                End If
      Else              'IF CONTRA NOT FOUND
                     If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
                        szStr = "SELECT * FROM tlbBankPayment;"

                        .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

                        .AddNew
                        .Fields.Item("MY_ID").Value = UniqueID()
                        .Fields.Item("ClientID").Value = txtClientID.text
                        .Fields.Item("BANK_AC").Value = txtBank.Tag
                        .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
                        .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
                        .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
                        .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
                        .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
                        .Fields.Item("DEPT_ID").Value = txtFund.Tag
                        .Fields.Item("TRAN_DATE").Value = txtDate.text
                        .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
                        .Fields.Item("TAX_CODE").Value = "T9"
                        .Fields.Item("VAT").Value = 0
                        .Fields.Item("TransactionType").Value = 12
                        .Fields.Item("TRANS").Value = "BR"
                        .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
                        .Fields.Item("CT").Value = "C"
                        .Fields.Item("TenantDeposit").Value = szID
                        .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")

                        .Update
                        .Close
                     End If
                  End If
   End With

   Set adoRST = Nothing
End Sub

Private Sub UpdateBankReceipt(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset
   adoConn.Execute "Update NLPosting N,TenantDeposit T,tlbbankpayment B SET DeleteFlag=true where  N.Trans_ID=B.tran_ID AND T.DepositID=B.TenantDeposit " & _
                    "AND B.TenantDeposit='" & szID & "' and  TRANSACTION_TYPE in (11,12)"
                    
   szStr = "SELECT * FROM tlbBankPayment " & _
           "WHERE TenantDeposit = '" & szID & "' AND " & _
                 "TransactionType = 12;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

      If Not adoRST.EOF Then
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = txtBank.Tag
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("TenantDeposit").Value = szID
         .Fields.Item("NLPost").Value = False
         .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
         If txtBank.Tag = txtDNC(0).text Then     'Contra Transation
            .Fields.Item("CT").Value = "C"
         End If

         .Update
         .Close
      End If

      szStr = "SELECT * FROM tlbBankPayment " & _
              "WHERE TenantDeposit = '" & szID & "' AND " & _
                    "TransactionType = 11;"

      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      If Not adoRST.EOF Then                                         'Contra Transation FOUND
         If .Fields.Item("BANK_AC").Value = .Fields.Item("NOMINAL_CODE").Value Then
            If txtBank.Tag <> txtDNC(0).text Then
               adoConn.Execute "DELETE * FROM tlbBankPayment " & _
                               "WHERE TenantDeposit = '" & szID & "' AND " & _
                                     "TransactionType = 11;"
               .Close
            Else
               .Fields.Item("ClientID").Value = txtClientID.text
               .Fields.Item("BANK_AC").Value = txtBank.Tag
               .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
               .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
               .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
               .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
               .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
               .Fields.Item("DEPT_ID").Value = txtFund.Tag
               .Fields.Item("TRAN_DATE").Value = txtDate.text
               .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
               .Fields.Item("TAX_CODE").Value = "T9"
               .Fields.Item("VAT").Value = 0
               .Fields.Item("TransactionType").Value = 11
               .Fields.Item("TRANS").Value = "BP"
               .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
               .Fields.Item("CT").Value = "C"
               .Fields.Item("TenantDeposit").Value = szID
               .Fields.Item("NLPost").Value = False
               .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
               .Update
               .Close
            End If
         End If
      Else              'IF CONTRA NOT FOUND
         If txtBank.Tag = txtDNC(0).text Then
            szStr = "SELECT * FROM tlbBankPayment;"

            .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
            .AddNew
            .Fields.Item("MY_ID").Value = UniqueID()
            .Fields.Item("ClientID").Value = txtClientID.text
            .Fields.Item("BANK_AC").Value = txtBank.Tag
            .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
            .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
            .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
            .Fields.Item("PROJ_REF").Value = txtDepositType.text & " - " & txtTenantID.text
            .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
            .Fields.Item("DEPT_ID").Value = txtFund.Tag
            .Fields.Item("TRAN_DATE").Value = txtDate.text
            .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
            .Fields.Item("TAX_CODE").Value = "T9"
            .Fields.Item("VAT").Value = 0
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
            .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
            .Fields.Item("CT").Value = "C"
            .Fields.Item("TenantDeposit").Value = szID
            .Fields.Item("NLPost").Value = False
            .Fields.Item("postingDate").Value = Format(Now, "dd/MMM/yyyy")
            .Update
            .Close
         End If
      End If
   End With
'TenantDeposit -tlbbankpayment
'TenantDeposit depositID
   Set adoRST = Nothing
'   adoconn.Execute "Update NLPosting N,TenantDeposit T,tlbbankpayment B SET DeleteFlag=true where T.DepositID=B.TenantDeposit " & _
'                    "AND B.TRAN_ID=N.TRANS_ID AND  B.TenantDeposit='" & szID & "' and  TRANSACTION_TYPE in ( 11,12)"
End Sub

Private Function TransID(DRE As String) As Integer
   Dim iRow As Integer

   TransID = 0
   For iRow = 1 To flxDeposit.Rows - 1
      If Left(flxDeposit.TextMatrix(iRow, 2), 1) = DRE Then
         If Val(Mid(flxDeposit.TextMatrix(iRow, 2), 2, Len(flxDeposit.TextMatrix(iRow, 2)))) > TransID Then _
            TransID = Val(Mid(flxDeposit.TextMatrix(iRow, 2), 2, Len(flxDeposit.TextMatrix(iRow, 2))))
      End If
   Next iRow
End Function

Private Sub cmdEdit_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a Lessee to continue.", vbInformation, "Edit Lessee"
      Exit Sub
   End If
   NEWMODE_ = False
   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, fmeTenant, EditMode
   txtTenantID.Locked = True 'issue 343 resolved by anol 20170322
   txtName.SetFocus
   tabTenant.Enabled = False
   cmdTenantLookup.Enabled = False
End Sub

'Private Sub cmdEditBank_Click()
'   BANK_PAYMENT_NEW_ENTRY_ = False
'   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, EditMode
'   cboBankId.Locked = True
'   txtBankACNumber.Locked = True
'End Sub

Private Sub cmdCancelDefaults_Click()
   txtNominalCode.text = txtDefault(0).text
   txtNominalCodeName.text = txtDefault(1).text
   txtCodeVat.text = txtDefault(2).text
   txtSLControl.text = txtDefault(3).text
   txtSLControlName.text = txtDefault(4).text
   lblVatCode(0).Caption = txtDefault(5).text

   cmdNC.Enabled = False
   cmdTaxList.Enabled = False
   cmdSLC.Enabled = False
   fmeTenant.Enabled = True

   cmdSaveDefaults.Enabled = False
   cmdCancelDefaults.Enabled = False
   cmdEditDefaults.Enabled = True
End Sub

Private Sub cmdEditDefaults_Click()
   txtDefault(0).text = txtNominalCode.text
   txtDefault(1).text = txtNominalCodeName.text
   txtDefault(2).text = txtCodeVat.text
   txtDefault(3).text = txtSLControl.text
   txtDefault(4).text = txtSLControlName.text
   txtDefault(5).text = lblVatCode(0).Caption
   
   cmdNC.Enabled = True
   cmdTaxList.Enabled = True
   cmdSLC.Enabled = True

   cmdSaveDefaults.Enabled = True
   cmdCancelDefaults.Enabled = True
   cmdEditDefaults.Enabled = False

   fmeTenant.Enabled = False
End Sub

Private Sub cmdEditMHistory_Click()
   If gridMaintenanceHistory.TextMatrix(1, 0) = "" Then Exit Sub

   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      frmMaintenanceJob.isEdit = True
      frmMaintenanceJob.CallingForm = "L"                         'Lessee
      frmMaintenanceJob.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintenanceJob
      frmMaintenanceJob.ZOrder 0
      frmMaintenanceJob.Show
   Else
      frmMaintananceDairy.isEdit = True
      frmMaintananceDairy.CallingForm = "L"                        'Lessee
      frmMaintananceDairy.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintananceDairy
      frmMaintananceDairy.ZOrder 0
      frmMaintananceDairy.Show
   End If
   Me.Enabled = False
End Sub

Private Sub cmdEditTenantAddress_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a lessee to continue.", vbCritical + vbOKOnly, "Lessee - Address"
      cmdTenantLookup.SetFocus
      Exit Sub
   End If
   ComponentInFrameEnableMode Me, fmeTenantAddress, EditMode
   If txtContact1.Enabled = True Then
        txtContact1.SetFocus
   End If
End Sub

Private Sub cmdEmail_Click()
   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub
   Call ChangeReportODBC
   If szFromEmail = "" Or szSMTPserver = "" Then
      ShowMsgInTaskBar "Company email or SMTP server IP has not been setup.", "Y", "N"
      Exit Sub
   End If
   If cboInvoiceTo.Value = "B" Then
      If txtEmail2.text = "" Then
         ShowMsgInTaskBar "Lessee's email address is missing.", "Y", "N"
         Exit Sub
      End If
   Else
      If txtEmail1.text = "" Then
         ShowMsgInTaskBar "Lessee's email address is missing.", "Y", "N"
         Exit Sub
      End If
   End If
   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
   Dim szTemp As String
   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
   If FileExists(szTemp) Then
      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If

   Dim adoConn       As New ADODB.Connection
   Dim szID          As String
   Dim bEmailResult  As Boolean
   Dim szSQL         As String

   uLessee.szLesseeID = txtTenantID.text
   If cboInvoiceTo.Value = "B" Then
      uLessee.szLesseeEmail = txtEmail2.text
   Else
      uLessee.szLesseeEmail = txtEmail1.text
   End If

   szID = SelectedIDs

   If Len(szID) > 1 Then
      adoConn.Open getConnectionString

      adoConn.Execute "UPDATE tlbLetterReports SET isPrint = '';"

      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE Id IN (" & szID & ");"
      adoConn.Execute szSQL

      adoConn.Close
      Set adoConn = Nothing
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   szSQL = txtTenantID.text & "_" & UniqueID() & ".pdf"

   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing

'                    Attach the PDF in the email
   SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL

   bEmailResult = SendDemandByE_Mail("General Letter", "Please find the letter in the attachment.", "General Letter")

   If bEmailResult Then
      ShowMsgInTaskBar "Email sent.", "Y", "P"
   Else
      ShowMsgInTaskBar "No email sent.", "Y", "N"
   End If
End Sub

Private Sub SaveAttachment(szFile As String)
   Set uLessee.colAtt = New Collection
   uLessee.colAtt.Add szFile
End Sub

Private Function SelectedIDs() As String
   Dim i As Integer

   For i = 1 To flxLetters.Rows - 1
      If (flxLetters.TextMatrix(i, 0) = "X") Then
        SelectedIDs = SelectedIDs & flxLetters.TextMatrix(i, 1) & ", "
      End If
   Next i
   
   If Len(SelectedIDs) > 2 Then SelectedIDs = Left(SelectedIDs, Len(SelectedIDs) - 2)
End Function

Private Sub cmdGridTenantLookup_Click()
   fmeTenantLookup.Visible = False
   fmeTenant.Enabled = True
   tabTenant.Enabled = True
   cmdTenantLookup.Visible = True
   cmdTenantLookup.Enabled = True
   
   FocusControl cmdTenantLookup
End Sub

Private Sub cmdGridUnitLookup_Click(Index As Integer)
    If Index = 0 Then
        fraList(0).Visible = False
        fmeTenantLookup.Enabled = True
        tabTenant.Enabled = True
    End If
End Sub

Private Sub cmdNC_Click()
   Dim szSQL As String

   szSQL = "SELECT S.Code, S.Name " & _
              "FROM NominalLedger AS S " & _
              "WHERE S.Posting AND " & _
                    "S.ClientID = '" & txtClientID.text & "';"
   
   szSel = "NC"
   LoadNC szSQL

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Left = tabTenant.Left + txtNominalCode.Left
   fraList(0).Top = tabTenant.Top + txtNominalCode.Top + txtNominalCode.Height + 10
   'fraList(0).Width = 3520
  ' Picture1.Width = fraList(0).Width - 80
  ' flxSupplier(0).Width = fraList(0).Width
   'fraList(0).Height = 2805
   'Picture1.Height = fraList(0).Height - 80
   'flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
End Sub

Private Sub LoadNC(szSQL As String)
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "NAME"
   flxSupplier(0).RowHeight(0) = 0

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1900
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40
   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)


   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "NAME"
   lblSearch2(0).Visible = False

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   flxSupplier(0).Sort = 1
   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdNCList_Click()
   LoadNominalCode

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Width = 4815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = 4695
   fraList(0).Left = tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "DH"         'Deposit Held
End Sub

Private Sub cmdNew_Click()
   NEWMODE_ = True
   SEARCHTenantMODE_ = False
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   
   adoConn.Open getConnectionString
   
   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If
   
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   tabTenant.Enabled = False
   ComponentInFrameEnableMode Me, fmeTenant, NewEntryMode
   chkPrintDmd.Value = True
   
   txtTenantID.Locked = False
   txtTenantID.BackColor = &H80000005
   cmdTenantLookup.Visible = False

   If UBound(szaChoice) > 0 Then
      If szaChoice(2) <> "" Then
         If InStr(szaChoice(2), "L") > 0 Then
            txtName.SetFocus
         End If
      Else
         txtTenantID.Locked = False
         txtTenantID.SetFocus
      End If
   End If

   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
'   ConfigurFlxEventHistory
   ComponentInFrameClearMode Me, fmeEventHistory, ClearBoth
   ComponentInFrameClearMode Me, Frame8, ClearBoth
   ComponentInFrameClearMode Me, Frame17(1), ClearBoth
End Sub

Private Sub ClearLesseeAddress(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
      txtTenantID.text = ""
      txtName.text = ""
      txtCompanyName.text = ""
'      cboSageAccountNumber.text = ""
      txtClient.text = ""
      txtProperty.text = ""
      txtUnit.text = ""
'      txtBankCode.text = ""
      txtDeposit.text = "0.00"
      txtBalance1.text = "0.00"
      txtContact1.text = ""
      txtHOAddressLine1.text = ""
      txtHOAddressLine2.text = ""
      txtHOAddressLine3.text = ""
      txtHOAddressLine4.text = ""
      txtHOPostCode.text = ""
      txtEmail1.text = ""
      txtDirectLine1.text = ""
      txtHOTelephone.text = ""
      txtHOFax.text = ""
      txtContact.text = ""
      txtBillAddressLine1.text = ""
      txtBillAddressLine2.text = ""
      txtBillAddressLine3.text = ""
      txtBillAddressLine4.text = ""
      txtBillPostCode.text = ""
      txtEmail2.text = ""
      txtDirectLine2.text = ""
      txtBillTelephone.text = ""
      txtBillFax.text = ""
   
   Case CearEntryComponents.ClearOnlyComboBoxes
   
   Case CearEntryComponents.ClearBoth
      ClearLesseeAddress ClearOnlyTextBoxes
      ClearLesseeAddress ClearOnlyComboBoxes
   End Select
End Sub

'Private Sub cmdNewBank_Click()
'   BANK_PAYMENT_NEW_ENTRY_ = True
'   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, NewEntryMode
'   cboBankId.Locked = False
'   txtBankACNumber.Locked = False
'End Sub
'
'Private Sub cmdNewEvent_Click()
'   M_HISTORY_NEW_ENTRY_ = True
'   ComponentInFrameEnableMode me, fmeEventHistory, NewEntryMode
'End Sub

Private Sub cmdNewMHistory_Click()
   If txtTenantID.text = "" Then Exit Sub

'   Load frmMaintenanceJob
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
        With frmMaintenanceJob
            .isEdit = True
           .CallingForm = "L"          'Calling from lessee form
           .RecordType = "J"
           .lblJobName.Caption = "Job Name"
           .Label1.Caption = "Job No."
           .txtRef.Enabled = True
           .UpdateRow = gridMaintenanceHistory.row
           .Frame1.Enabled = False
           .Show
           .ZOrder 0
        End With
    Else
        ShowMsgInTaskBar "Please select a job."
    End If
'   Me.Enabled = False
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdPrintEmail_Click()
   If flxEmails.TextMatrix(flxEmails.row, 6) = "" Then Exit Sub
   If flxEmails.TextMatrix(flxEmails.row, 0) = "" Then Exit Sub
   Call ChangeReportODBC
   Dim szPath        As String
   Dim bLesseeEmail  As Boolean
   Dim szaDateTime() As String
   Dim szBody        As String
   Dim szSubject     As String
   Dim szAddress     As String
   Dim szLine        As String
    On Error GoTo Err
   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
   bLesseeEmail = FileExists(szPath)

   If bLesseeEmail Then
      Open szPath For Input As #2

      Do Until EOF(2)
         Line Input #2, szLine
         If InStr(szLine, "Email sent on:") > 0 Then
            szLine = Mid(szLine, 15)
            szaDateTime = Split(szLine, "#")

            If UBound(szaDateTime) > 1 Then
               If flxEmails.TextMatrix(flxEmails.row, 6) = szaDateTime(2) Then
                  Line Input #2, szLine
                  szAddress = Mid(szLine, 15)            'Address
                  Line Input #2, szLine
                  szSubject = Mid(szLine, 15)            'Subject

                  Close #2
                  szBody = GetTheBody()
                  Exit Do
               End If
            End If
         End If
      Loop
      
      Close #2
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Email_Print.rpt")

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue szAddress
   Report.ParameterFields(2).AddCurrentValue szaDateTime(0)
   Report.ParameterFields(3).AddCurrentValue szaDateTime(1)
   Report.ParameterFields(4).AddCurrentValue szSubject
   Report.ParameterFields(5).AddCurrentValue szBody

   Load frmReport
   frmReport.LoadReportViewer Report
   Exit Sub
Err:
   MsgBox Err.description, vbOKOnly, "Warning"
End Sub

Private Sub cmdPrintHistory_Click()

   
   On Error GoTo Err
'   If Val(txtBalance1.text) = 0 Then
'      ShowMsgInTaskBar "Statement will not be printed as account balance is 0.", "Y", "N"
'      Exit Sub
'   End If

   Dim adoConn As New ADODB.Connection
   Dim szSQL   As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
   Dim strReportName   As String
   Call ChangeReportODBC
   adoConn.Open getConnectionString

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = '' " & _
           "WHERE  spare2 = 'Y';"
   adoConn.Execute szSQL

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = 'Y' " & _
           "WHERE  SageAccountNumber = ('" & txtTenantID.text & "');"
   adoConn.Execute szSQL
   
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
   Dim rsStatementTemplate As New ADODB.Recordset
   rsStatementTemplate.Open "Select LesseeAccTemplate from client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
   If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("LesseeAccTemplate").Value), "", rsStatementTemplate("LesseeAccTemplate").Value)
   End If
   If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
   Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
   End If
   rsStatementTemplate.Close
   
   
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue optCurrentTenant.Value
   Report.ParameterFields(2).AddCurrentValue False
   Report.ParameterFields(3).AddCurrentValue "1"
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub
Err:
   MsgBox Err.description
   
'   Exit Sub
'
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'   Dim rep As New frmReport
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
'
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue txtTenantID.text
'   Report.ParameterFields(2).AddCurrentValue txtName.text
'   Report.ParameterFields(3).AddCurrentValue txtClient.text
'   Report.ParameterFields(4).AddCurrentValue txtProperty.text
'   Report.ParameterFields(5).AddCurrentValue txtUnit.text
'   Report.ParameterFields(6).AddCurrentValue CDbl(txtBalance1.text)
'
''MsgBox CCur(txtAcBal.text)
'   Load rep
'   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintHistorySorted_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As New frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtTenantID.text
   Report.ParameterFields(2).AddCurrentValue txtName.text
   Report.ParameterFields(3).AddCurrentValue txtClient.text
   Report.ParameterFields(4).AddCurrentValue txtProperty.text
   Report.ParameterFields(5).AddCurrentValue txtUnit.text
   Report.ParameterFields(6).AddCurrentValue CDbl(txtBalance1.text)
   Report.ParameterFields(7).AddCurrentValue IIf(optSortingAccountHistory(0), 0, IIf(optSortingAccountHistory(1), 1, 2))

'MsgBox CCur(txtAcBal.text)
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintJobSheet_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
      
   Report.ParameterFields(1).AddCurrentValue gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3)
   
   If (gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB") Then
      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "JOB SHEET"
   Else
      Report.ParameterFields(2).AddCurrentValue "Diary Entry"
      Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
   End If
      
   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdPrintReceipt_Click()
   If flxACHistory.row > 0 And _
         InStr(flxACHistorySplit.TextMatrix(flxACHistorySplit.row, 2), "Receipt") > 0 Then
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report
      Dim rep As frmReport
      Call ChangeReportODBC
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Receipt.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      If Report.HasSavedData Then Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue CLng(flxACHistorySplit.TextMatrix(flxACHistorySplit.row, 1))

      Set rep = New frmReport
      Load rep
      rep.LoadReportViewer Report
   End If
End Sub

Private Sub cmdPrintStatement_Click()
    'Print Lessee Statement
   On Error GoTo Err
'   If Val(txtBalance1.text) = 0 Then
'      ShowMsgInTaskBar "Statement will not be printed as account balance is 0.", "Y", "N"
'      Exit Sub
'   End If

   Dim adoConn As New ADODB.Connection
   Dim szSQL   As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
   Dim strReportName   As String
   Call ChangeReportODBC
   adoConn.Open getConnectionString

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = '' " & _
           "WHERE  spare2 = 'Y';"
   adoConn.Execute szSQL

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = 'Y' " & _
           "WHERE  SageAccountNumber = ('" & txtTenantID.text & "');"
   adoConn.Execute szSQL
   
   Dim rsStatementTemplate As New ADODB.Recordset
   rsStatementTemplate.Open "Select LesseeTemplate from client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
   If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("LesseeTemplate").Value), "", rsStatementTemplate("LesseeTemplate").Value)
   End If
   If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
   End If
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False
   Report.ParameterFields(2).AddCurrentValue False
   Report.ParameterFields(3).AddCurrentValue "1"
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub
Err:
   MsgBox Err.description
End Sub


Private Sub cmdPrintWord_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String
   Dim rstReport As ADODB.Recordset
   Dim rstRst As New ADODB.Recordset

   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub

   adoConn.Open getConnectionString
   
   szSQL = "UPDATE tlbLetterReports SET isPrint = '';"
   adoConn.Execute szSQL
   
   For i = 1 To flxLetters.Rows - 1
    If (flxLetters.TextMatrix(i, 0) = "X") Then
      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE id = " & flxLetters.TextMatrix(i, 1) & ";"
      adoConn.Execute szSQL
    End If
   Next i

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   szSQL = txtTenantID.text & "_" & UniqueID() & ".DOC"
   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTWordForWindows
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing

   OpenFile szSQL, DB_PATH & "\AllStuff\Temp\"

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdPropLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmProperty2
   frmProperty2.LOAD_PROPERTY_PROPERTYID = Left(txtProperty, 4)
   frmProperty2.CLIENT_NAME = Trim(txtClient.text)
   frmProperty2.Show
End Sub

Private Sub cmdRCCCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   txtRCCComments.Locked = True
   cmdRCCEdit.Enabled = True
   cmdRCCSave.Enabled = False
   cmdRCCCancel.Enabled = False
   Dim adoConn As New ADODB.Connection
   Dim rsComment As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsComment.Open "Select RCCComments from tenants where SageAccountNumber='" & txtTenantID.text & "' and (isnull(Comments) Or Comments='')", adoConn, adOpenStatic, adLockReadOnly
   If Not rsComment.EOF Then
       txtRCCComments.text = IIf(IsNull(rsComment("RCCComments").Value), "", rsComment("RCCComments").Value)
   End If
   rsComment.Close
   Set rsComment = Nothing
   adoConn.Close
   Set adoConn = Nothing
   FocusControl txtRCCComments
End Sub

Private Sub cmdRCCEdit_Click()
   txtRCCComments.Locked = False
   cmdRCCEdit.Enabled = False
   cmdRCCSave.Enabled = True
   cmdRCCCancel.Enabled = True
   FocusControl txtRCCComments
End Sub

Private Sub cmdRCCSave_Click()
   If SaveComments("Tenants", "RCCComments", txtRCCComments.text, "SageAccountNumber", txtTenantID.text) Then
      ShowMsgInTaskBar "The comments have been saved successfully."
   End If
   txtRCCComments.Locked = True
   cmdRCCEdit.Enabled = True
   cmdRCCSave.Enabled = False
   cmdRCCCancel.Enabled = False
End Sub

Private Function GetTheBody() As String
   Dim szLine     As String
   Dim szPath     As String
   Dim szaDateTime() As String
                  
   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
   Open szPath For Input As #2

   Do Until EOF(2)
      Line Input #2, szLine
      If InStr(szLine, "Email sent on:") > 0 Then
         szLine = Mid(szLine, 15)
         szaDateTime = Split(szLine, "#")
         
         If UBound(szaDateTime) > 1 Then
            If flxEmails.TextMatrix(flxEmails.row, 6) = szaDateTime(2) Then
               Line Input #2, szLine
               Line Input #2, szLine
               Line Input #2, szLine
               While Not EOF(2) And InStr(szLine, "*****") = 0
            'Debug.Print szLine
                  GetTheBody = GetTheBody + szLine + (Chr(10) & Chr(13))
                  Line Input #2, szLine
               Wend
               If InStr(szLine, "*****") > 0 Then
                  GetTheBody = GetTheBody + Mid(szLine, 1, Len(szLine) - 5)
               End If
               Exit Do
            End If
         End If
      End If
   Loop
   
   Close #2
End Function

Private Sub cmdResendEmail_Click()
   If flxEmails.TextMatrix(flxEmails.row, 6) = "" Then Exit Sub
   If flxEmails.TextMatrix(flxEmails.row, 0) = "" Then Exit Sub
   If MsgBox("Would you like to send the email again?", vbQuestion + vbYesNo, "Email") = vbNo Then Exit Sub
   Call ChangeReportODBC
   Dim bEmailResult     As Boolean
   Dim szBody           As String

   szBody = GetTheBody()
   bEmailResult = SendEmail(szFromEmail, flxEmails.TextMatrix(flxEmails.row, 3), _
                           flxEmails.TextMatrix(flxEmails.row, 4), _
                           szBody, , , _
                           , Me.txtTenantID.text, "Just Email")
   If bEmailResult Then
      ShowMsgInTaskBar "Email sent.", "Y", "P"

      SavingEmailInformation szBody
      LoadFlxEmails
   Else
      ShowMsgInTaskBar "No email sent.", "Y", "N"
   End If

End Sub

Private Sub SavingEmailInformation(szBody As String)
   Dim szLine  As String
   Dim szPath  As String

   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
'  even the file does not exists, system will create the file and start adding text at the bottom of the file
   Open szPath For Append As #1

   szLine = "Email sent on:" & Format(Now, "dd/mm/yyyy") & "#" & Format(Now, "hh:nn") & "#" & UniqueID() & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Address:" & flxEmails.TextMatrix(flxEmails.row, 3) & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Subject:" + flxEmails.TextMatrix(flxEmails.row, 4) & Chr$(13) & Chr$(10)
   szLine = szLine + szBody
   szLine = szLine + "*****"

   Print #1, szLine
   Close #1
End Sub

Private Sub cmdSave_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a lessee to continue.", vbExclamation, "No Lessee Selected"
      Exit Sub

   ElseIf txtName.text = "" Then
      MsgBox "Please enter a Lessee Name to continue.", vbExclamation, "No Lessee Name"
      txtName.SetFocus
      txtName.text = ""
      Exit Sub
   End If

   If chkCombEmail.Value And (Not chkEmailDmd.Value Or Not chkEmailSt.Value) Then
      MsgBox "Email Demand and Email Statement both have to be selected.", vbInformation + vbOKOnly, "Combined Email"
      If Not chkEmailDmd.Value Then chkEmailDmd.SetFocus
      If Not chkEmailSt.Value Then chkEmailSt.SetFocus
      Exit Sub
   End If

   If txtTenantID.text <> cboSageAccountNumber.text Then cboSageAccountNumber.text = txtTenantID.text

   If txtCompanyName.text = "" Then txtCompanyName.text = txtName.text

   If txtDeposit.text = "" Then txtDeposit.text = "0.00"
   If txtBalance1.text = "" Then txtBalance1.text = "0.00"

   If SaveTenantInformation Then
       'MsgBox cboSageAccountNumber.text
      If COPYMODE_ Then
         cmdSaveTenantAddress_Click
         COPYMODE_ = False
      End If
       bolgridTenantLookupRefresh = False
      ShowMsgInTaskBar "The lessee record has been saved successfully."
      NEWMODE_ = False
      ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
      SEARCHTenantMODE_ = True
      tabTenant.Enabled = True
      cmdTenantLookup.Enabled = True
      cmdTenantLookup.Visible = True
      'added by anol 27 Jun 2015
      tabTenant.Tab = 0
      FocusControl cmdEditTenantAddress
   Else
      txtTenantID.SetFocus
   End If

   If Not NEWMODE_ Then lblLeaseChanged.Caption = txtTenantID.text
End Sub

Public Function PopulateTenantLookup(ByVal sSQLQuery_ As String, adoConn As ADODB.Connection)
   Dim iRow As Integer

   iRow = 1

   ConfigGridTenantLookup
' szHeader$ = "<Sage A/C|<Name|<Address|>Balance||Client|Property|Unit"
'   gridTenantLookup.FormatString = szHeader$
'
'   gridTenantLookup.ColWidth(0) = lblTenantSort(1).Left - lblTenantSort(0).Left
'   gridTenantLookup.ColWidth(1) = lblTenantSort(2).Left - lblTenantSort(1).Left
'   gridTenantLookup.ColWidth(2) = lblTenantSort(3).Left - lblTenantSort(2).Left
'   gridTenantLookup.ColWidth(3) = 1200
'   gridTenantLookup.ColWidth(4) = 0
'   gridTenantLookup.ColWidth(5) = 0
'   gridTenantLookup.ColWidth(6) = 0
'   gridTenantLookup.ColWidth(7) = 0
   lblLoading.Caption = "Please wait while loading..."
   fmeLoading.Visible = True
   fmeLoading.Top = 2715
   fmeLoading.Left = 3560
   fmeLoading.ZOrder 0
   lblLoading.ZOrder 0
   fmeLoading.Refresh
   populateGrid adoConn, sSQLQuery_, gridTenantLookup
   fmeLoading.Visible = False
   
'   Dim rRow As Integer, iRec As Integer
'   Dim adoRST As New ADODB.Recordset
'   Dim szSQL As String
'
'   Set adoconn = New ADODB.Connection
'   adoconn.Open getConnectionString
'
'
'   adoRST.Open sSQLQuery_, adoconn, adOpenStatic, adLockReadOnly
'   Dim iRows As Integer
'   gridTenantLookup.Rows = 2
'   iRows = 1
'   While Not adoRST.EOF
'      gridTenantLookup.TextMatrix(iRows, 0) = adoRST.Fields(0).Value
'      gridTenantLookup.TextMatrix(iRows, 1) = adoRST.Fields(1).Value
'      gridTenantLookup.TextMatrix(iRows, 2) = adoRST.Fields(2).Value
'      gridTenantLookup.TextMatrix(iRows, 3) = adoRST.Fields(3).Value
'      gridTenantLookup.TextMatrix(iRows, 4) = adoRST.Fields(4).Value
'      gridTenantLookup.TextMatrix(iRows, 5) = adoRST.Fields(5).Value
'      gridTenantLookup.TextMatrix(iRows, 6) = adoRST.Fields(6).Value
'      'gridTenantLookup.TextMatrix(iRows, 7) = adoRST.Fields(7).Value
'      If Not adoRST.EOF Then gridTenantLookup.AddItem ""
'      iRows = iRows + 1
'      adoRST.MoveNext
'   Wend
'
'   Set adoRST = Nothing
   
End Function

'Private Sub cmdSaveBank_Click()
'   Dim sSQLQuery As String, sWhere As String
'   Dim adoConn As New ADODB.Connection
'   Dim oResultSet As New ADODB.Recordset
'
'   If cboBankId.text = "" Then
'      MsgBox "Please select a bank from the drop down list.", vbCritical + vbOKOnly, "Bank Payment Details"
'      cboBankId.SetFocus
'      Exit Sub
'   End If
'   If cboPaymentMethod.text = "" Then
'      MsgBox "Please select a payment method from the drop down list.", vbCritical + vbOKOnly, "Bank Payment Details"
'      cboPaymentMethod.SetFocus
'      Exit Sub
'   End If
'   If txtBankACName.text = "" Then
'      MsgBox "Please enter bank account name.", vbCritical + vbOKOnly, "Bank Payment Details"
'      txtBankACName.SetFocus
'      Exit Sub
'   End If
'   If txtBankACNumber.text = "" Then
'      MsgBox "Please enter bank account number.", vbCritical + vbOKOnly, "Bank Payment Details"
'      txtBankACNumber.SetFocus
'      Exit Sub
'   End If
'
'   adoConn.Open getConnectionString
'
'   txtBankTenantID.text = txtTenantID.text
'
'   sSQLQuery = "SELECT BankTenantID, BankID, BankACNumber, BankACName, " & _
'                  "BankSortCode, IsDefaultAC, PaymentMethod, BacsRef " & _
'               "FROM TenantBankDetails "
'   sWhere = " WHERE BankTenantID = '" & txtTenantID.text & "' AND " & _
'              "BankID = '" & cboBankId.Value & "' AND " & _
'              "BankACNumber = '" & txtBankACNumber.text & "'"
'
'   oResultSet.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
'
'   With oResultSet
'      If BANK_PAYMENT_NEW_ENTRY_ And .EOF Then .AddNew
''else it is updating data
'      !BankTenantID = txtBankTenantID.text
'      !BankID = cboBankId.Value
'      !BankACNumber = txtBankACNumber.text
'      !BankACName = txtBankACName.text
'      !BankSortCode = txtBankSortCode.text
'      !IsDefaultAC = chkIsDefaultAC.Value
'      !PaymentMethod = cboPaymentMethod.Value
'      !BacsRef = txtBACSRef.text
'      .Update
'      .Close
'   End With
'
'   Set oResultSet = Nothing
'
'   populateGrid adoConn, "SELECT BankTenantID, BankID, BankACNumber, BankACName, " & _
'                           "BankSortCode, IsDefaultAC, PaymentMethod, BacsRef " & _
'                         "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'", gridBank
'   adoConn.Close
'   Set adoConn = Nothing
'
''   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
'End Sub
'
'Private Sub cmdSaveEvent_Click()
'   Dim adoConn As New ADODB.Connection
'   Dim oResultSet As New ADODB.Recordset
'   Dim szHeader As String, sSQLQuery As String
'
'   If cboEventType.Value = "" Then
'      MsgBox "Please select the event type.", vbCritical + vbOKOnly, "Event History"
'      cboEventType.SetFocus
'      Exit Sub
'   End If
'   If dtpReportedDate.text = "" Then
'      MsgBox "Please enter the reported date.", vbCritical + vbOKOnly, "Event History"
'      dtpReportedDate.SetFocus
'      Exit Sub
'   End If
'   If txtDescription.text = "" Then
'      MsgBox "Please enter the description of the event.", vbCritical + vbOKOnly, "Event History"
'      txtDescription.SetFocus
'      Exit Sub
'   End If
'   If txtTaskOwner.text = "" Then
'      MsgBox "Please enter the name of the task owner of the event.", vbCritical + vbOKOnly, "Event History"
'      txtTaskOwner.SetFocus
'      Exit Sub
'   End If
'   If txtContact.text = "" Then
'      MsgBox "Please enter the name of the contact of the event.", vbCritical + vbOKOnly, "Event History"
'      txtContact.SetFocus
'      Exit Sub
'   End If
'   adoConn.Open getConnectionString
'
'   ' Event Type
'   txtEventHistoryID.text = txtTenantID.text & "-" & cboEventType.Value & "-" & dtpReportedDate.text
'   txtEventTenantID.text = txtTenantID.text
'
'   sSQLQuery = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                      "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm, REMINDER_ID " & _
'               "FROM TenantEventHistory " & _
'               "WHERE EventHistoryID = '" & txtEventHistoryID.text & "'"
'
'   oResultSet.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
'
'   With oResultSet
'      If M_HISTORY_NEW_ENTRY_ And .EOF Then .AddNew
''else it is updating data
'      !EventHistoryID = txtEventHistoryID.text
'      !EventTenantID = txtTenantID.text
'      !EventType = cboEventType.Value
'      !ReportedDate = dtpReportedDate.text
'      !description = txtDescription.text
'      !DateCompleted = IIf(dtpDateCompleted.text = "", Null, dtpDateCompleted.text)
'      !TaskOwner = txtTaskOwner.text
'      !Contact = txtContact.text
'      !RemindDate = IIf(dtpRemindDate.text = "", Null, dtpRemindDate.text)
'
'      !Alarm = IIf(chkAlarm.Value = 1, True, False)
'      If chkAlarm.Value = 1 Then
'         If M_HISTORY_NEW_ENTRY_ Then
'            !Reminder_ID = NewReminder(Format(CDate(!RemindDate), "YYYYMMDD"), "083000", txtDescription.text, "TenantEventHistory", txtEventHistoryID.text)
'         Else
'            UpdateReminder !Reminder_ID, Format(CDate(!RemindDate), "YYYYMMDD"), "083000", txtDescription.text
'         End If
'      Else
'         ClearReminder !Reminder_ID
'      End If
'
'      .Update
'      .Close
'   End With
'
'   Set oResultSet = Nothing
'
''   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
''   sSQLQuery = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
''                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
''                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
''   populateGridSimply adoConn, sSQLQuery, gridEventHistory, szHeader
'   LoadGridMaintenanceHistory adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'   ComponentInFrameEnableMode me, fmeEventHistory, DefaultMode
'End Sub

Private Sub cmdSaveDefaults_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset
   Dim szSQL   As String

   adoConn.Open getConnectionString

   szSQL = "SELECT SLControl, DefaultNC, VAT_CODE " & _
           "FROM  Tenants " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "';"
   
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With adoRST
      .Fields.Item("DefaultNC").Value = IIf(IsNull(txtNominalCode.text), "", txtNominalCode.text)
      .Fields.Item("VAT_CODE").Value = IIf(IsNull(lblVatCode(0).Caption), "", lblVatCode(0).Caption)
      .Fields.Item("SLControl").Value = IIf(IsNull(txtSLControl.text), "", txtSLControl.text)
      .Update
      .Close
   End With

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   cmdNC.Enabled = False
   cmdTaxList.Enabled = False
   cmdSLC.Enabled = False
   fmeTenant.Enabled = True

   cmdSaveDefaults.Enabled = False
   cmdCancelDefaults.Enabled = False
   cmdEditDefaults.Enabled = True

   ShowMsgInTaskBar "It has been saved successfully", "Y", "P"
End Sub

Private Sub cmdSaveTenantAddress_Click()
   Dim adoConn As New ADODB.Connection
'   'validating email
'   Dim szErrMsg As String
'
'   If Trim(txtEmail1.text) <> "" Then
'      If Not ValidateEmail(txtEmail1.text, szErrMsg) Then
'         MsgBox szErrMsg, vbCritical + vbOKOnly, "Lessee Email"
'         SelTxtInCtrl txtEmail1
'         txtEmail1.SetFocus
'      End If
'   End If
'   'End of validation
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
   Dim rsTenants As New ADODB.Recordset
   sSQLQuery = "SELECT SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                  "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                  "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                  "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                  "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                  "TenantMemo, Balance, Deposite, BankCode, spare1, EmailDmd " & _
                "FROM TENANTS " & _
                "WHERE TenantID = '" & txtTenantID.text & "' and isCurrent"

   If PostToDBUsingADODB(Me, fmeTenantAddress, adoConn, sSQLQuery, False) Then
      If Not COPYMODE_ Then MsgBox "The contact details of the Lessee has been updated successfully", vbInformation
'       UpdateSAGECustomerAddress
   Else
       MsgBox "Error occured while updating the contact information", vbInformation
   End If
   
   sSQLQuery = "SELECT SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                  "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                  "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                  "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                  "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                  "TenantMemo, Balance, Deposite, BankCode, spare1, EmailDmd,LastModifiedby,LastModifiedDate " & _
                "FROM TENANTS " & _
                "WHERE TenantID = '" & txtTenantID.text & "' and isCurrent"
rsTenants.Open sSQLQuery, adoConn, adOpenKeyset, adLockOptimistic
If Not rsTenants.EOF Then
     With rsTenants
     !LastModifiedBy = User
     !LastModifiedDate = Now
     .Update
     End With
     
     rsTenants.Close
End If
     
     
                
''    rsTenants.Open sSQLQuery, adoConn, adOpenKeyset, adLockOptimistic
''    If Not rsTenants.EOF Then
''        With rsTenants
'''            !SageAccountNumber = txtTenantID.text
'''            !TenantID = txtTenantID.text
''            !Name = txtName.text
''            !CompanyName = txtCompanyName.text
''            !Contact1 = txtContact1.text
''            !Email1 = txtEmail1.text
''            !DirectLine1 = txtDirectLine1.text
''            !Contact2 = txtContact2.text
''            !Email2 = txtEmail2.text
''            !DirectLine2 = txtDirectLine2.text
''            !HOAddressLine1 = txtHOAddressLine1.text
''            !HOAddressLine2 = txtHOAddressLine2.text
''            !HOAddressLine3 = txtHOAddressLine3.text
''            !HOAddressLine4 = txtHOAddressLine4.text
''            !HOPostCode = txtHOPostCode.text
''            !HOTelephone = txtHOTelephone.text
''            !HOFax = txtHOFax.text
''            !BillAddressLine1 = txtBillAddressLine1.text
''            !BillAddressLine2 = txtBillAddressLine2.text
''            !BillAddressLine3 = txtBillAddressLine3.text
''            !BillAddressLine4 = txtBillAddressLine4.text
''            !BillPostCode = txtBillPostCode.text
''            !BillTelephone = txtBillTelephone.text
''            !BillFax = txtBillFax.text
''            !InvoiceTo = cboInvoiceTo.text
''            '!TenantMemo = txtInvoiceTo.text
''            '!balance = txtBalance.text
''            !Deposite = txtDeposit.text
''            !BankCode = txtBank.text
''            '!spare1=txtInvoiceTo.text
''            !EmailDmd = IIf(chkEmailDmd.Value = 0, False, True)
''            .Update
''        End With
''     End If
        adoConn.Close
        Set adoConn = Nothing
        
        ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
End Sub

Private Function SendDemandByE_Mail(szSub, szBody, szType) As Boolean
   Dim i As Integer

   SendDemandByE_Mail = SendEmail(szFromEmail, Trim(uLessee.szLesseeEmail), _
                                  szSub, _
                                  szBody, , , _
                                  uLessee.colAtt, uLessee.szLesseeID, szType)
End Function

Private Sub cmdSendMail_Click(Index As Integer)
   If Index = 0 Then
      If txtEmail1.text = "" Then Exit Sub
      Load frmSendMail
      frmSendMail.lblRecipientAddress.Caption = txtEmail1.text
   End If
   If Index = 1 Then
      If txtEmail2.text = "" Then Exit Sub
      Load frmSendMail
      frmSendMail.lblRecipientAddress.Caption = txtEmail2.text
   End If

   frmSendMail.Show
   Me.Enabled = False
End Sub

Private Sub cmdSentStByEmail_Click()
    On Error GoTo Err
   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
'   Dim szTemp As String
'   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
'   If FileExists(szTemp) Then
'      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
'      Exit Sub
'   End If
   Call ChangeReportODBC
   uLessee.szLesseeID = txtTenantID.text
   If cboInvoiceTo.Value = "B" Then
      uLessee.szLesseeEmail = txtEmail2.text
   Else
      uLessee.szLesseeEmail = txtEmail1.text
   End If

   Dim szPath        As String
   Dim i             As Integer
   Dim bEmailResult  As Boolean

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
'Newly added by anol 2023-08-25
   ConfigureSMTP adoConn
   
   adoConn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = '' " & _
                   "WHERE  spare2 = 'Y';"

   adoConn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = 'Y' " & _
                   "WHERE  SageAccountNumber = ('" & txtTenantID.text & "');"

   adoConn.Close
   Set adoConn = Nothing

   '  Create the pdf file name of the statement
   szPath = txtTenantID.text & "_" & UniqueID() & ".pdf"

   CreatePDF_Statement txtTenantID.text, DB_PATH & "\AllStuff\Temp\" & szPath

   '  Attaching the statement to the email
   SaveAttachment DB_PATH & "\AllStuff\Temp\" & szPath

   bEmailResult = SendDemandByE_Mail("Account Statement", "Please find attachment your account statement.", "Account Statement")

   If bEmailResult Then
      MsgBox "Email sent.", vbInformation + vbOKOnly, "Lessee Statement"
   Else
      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Lessee Statement"
   End If
   Exit Sub
Err:
   MsgBox Err.description
Exit Sub
'
'
'
'
'   Dim szEmailRecepent  As String
'   Dim colAtt           As New Collection
'
'   If cboInvoiceTo.Value = "B" Then
'      szEmailRecepent = txtEmail2.text
'   Else
'      szEmailRecepent = txtEmail1.text
'   End If
'
'   Dim szPath        As String
'   Dim bEmailResult  As Boolean
'   Dim adoConn       As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
''   CreateNonExistsFolder App.Path & "\Temp"
''   szPath = App.Path & "\Temp\" & txtTenantID.text & "_" & UniqueID() & ".pdf"
'   szPath = DB_PATH & "\AllStuff\Temp\" & txtTenantID.text & "_" & UniqueID() & ".pdf"
'
'   CreatePDF_Statement txtTenantID.text, szPath, adoConn
'
'   '  Attaching the demand to the email
'   If colAtt.count > 0 Then colAtt.Remove (1)
'   colAtt.Add szPath
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   bEmailResult = SendEmail(szFromEmail, szEmailRecepent, _
'                  "Account Statement", "Please see the attached of your account statement for details.", , , _
'                  colAtt, txtTenantID.text, "Account Statement")
'
'   MousePointer = vbDefault
'   If bEmailResult Then
'      MsgBox "Email sent.", vbInformation + vbOKOnly, "Demands"
'   Else
'      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Demands"
'   End If
End Sub

Private Sub CreatePDF_Statement(szLessee As String, szFileName As String)
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim strReportName As String
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Dim rsStatementTemplate As New ADODB.Recordset
   rsStatementTemplate.Open "Select LesseeTemplate from client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
   If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("LesseeTemplate").Value), "", rsStatementTemplate("LesseeTemplate").Value)
   End If
   If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
   End If
   rsStatementTemplate.Close
   adoConn.Close
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False          'Only outstanding statement
   'Report.ParameterFields(2).AddCurrentValue False          'specific lessee
   'Modified by anol 11 April 2016
   Report.ParameterFields(2).AddCurrentValue True          'specific lessee
   Report.ParameterFields(3).AddCurrentValue szLessee       'Lessee id

'   Transfer report into PDF file
   Report.ExportOptions.DiskFileName = szFileName
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing
End Sub
'
'Private Sub CreatePDF_Statement(szLessee As String, szFileName As String, adoConn As ADODB.Connection)
'   Dim szSQL      As String
'   Dim dLesBal    As Double
'   Dim adoRst     As New ADODB.Recordset
'   Dim reportApp  As New CRAXDRT.Application
'   Dim Report     As CRAXDRT.Report
'
'   szSQL = "SELECT T.Name, U.UnitName, P.PropertyName, C.ClientName " & _
'           "FROM Tenants AS T,  LeaseDetails AS L, Units AS U, Property AS P, Client AS C " & _
'           "WHERE Status = TRUE AND T.SageAccountNumber = '" & szLessee & "' AND " & _
'               "T.SageAccountNumber = L.SageAccountNumber AND " & _
'               "L.UnitNumber = U.UnitNumber AND " & _
'               "U.PropertyID = P.PropertyID AND " & _
'               "P.ClientID = C.ClientID;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      Set adoRst = Nothing
'      Exit Sub
'   End If
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   If Report.HasSavedData Then Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue szLessee
'   Report.ParameterFields(2).AddCurrentValue adoRst.Fields.Item("Name").Value
'   Report.ParameterFields(3).AddCurrentValue adoRst.Fields.Item("ClientName").Value
'   Report.ParameterFields(4).AddCurrentValue adoRst.Fields.Item("PropertyName").Value
'   Report.ParameterFields(5).AddCurrentValue adoRst.Fields.Item("UnitName").Value
'   dLesBal = CDbl(LesseeAccountBalance(adoConn, szLessee))
'   Report.ParameterFields(6).AddCurrentValue dLesBal
'
'   adoRst.Close
'   Set adoRst = Nothing
'
'   'Transfer report into PDF file
'   Report.ExportOptions.DiskFileName = szFileName
'   Report.ExportOptions.DestinationType = crEDTDiskFile
'   Report.ExportOptions.FormatType = crEFTPortableDocFormat
'   Report.ExportOptions.PDFExportAllPages = True
'   Report.Export False
'   Set Report = Nothing
'End Sub

Private Sub cmdSetAmtType_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

'   LoadRptAmtType cmbDptAmtType, "RECEIPT AMOUNT TYPE", adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdSetDptType_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If yDEPOSIT = 1 Then
      frmSecondaryCode.PRIMARY_CODE_SHOW = "DPTYP"
   Else
      If yDEPOSIT = 3 Then _
         frmSecondaryCode.PRIMARY_CODE_SHOW = "RTYP"
      If yDEPOSIT = 4 Or yDEPOSIT = 31 Then _
         frmSecondaryCode.PRIMARY_CODE_SHOW = "EXPTYP"
       If yDEPOSIT = 5 Or yDEPOSIT = 32 Then _
         frmSecondaryCode.PRIMARY_CODE_SHOW = "RTYP"
   End If
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   If yDEPOSIT = 1 Then
'      LoadRptAmtType cboDepositType, "DEPOSIT TYPE", adoConn
   Else
'      If yDEPOSIT = 3 Then _
'         LoadRptAmtType cboDepositType, "REFUND TYPE", adoConn
'      If yDEPOSIT = 4 Then _
'         LoadRptAmtType cboDepositType, "EXPENSES TYPE", adoConn
   End If
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadRptAmtType(conComboBox As Control, szValue As String, adoConn As ADODB.Connection)
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

   conComboBox.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   conComboBox.Column() = szaData()
End Sub

Private Sub cmdSLC_Click()
   If Not bLeaseSetup Then
      ShowMsgInTaskBar "No lease has been setup", "Y", "P"
      Exit Sub
   End If
   
   Dim szSQL As String
   
   szSQL = "SELECT Code AS NCode, CAName " & _
           "FROM NominalLedger " & _
           "WHERE NOT CAFixed AND " & _
                 "ClientID = '" & txtClientID.text & "' AND " & _
                 "NOT ISNULL(CAName) AND CAType = 'S';"

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Left = tabTenant.Left + txtSLControl.Left
   fraList(0).Top = tabTenant.Top + txtSLControl.Top + txtSLControl.Height + 10
   fraList(0).Width = 3520
   Shape4(2).Width = fraList(0).Width - 200
  ' Picture1.Width = fraList(0).Width - 80
   'flxSupplier(0).Width = Picture1.Width - 40
  ' fraList(0).Height = 1965
   'Picture1.Height = fraList(0).Height - 80
  ' flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "SLC"

   LoadControlAccount szSQL
End Sub

Private Sub LoadControlAccount(szSQL As String)
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
'   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "NAME"
'   flxSupplier(0).TextMatrix(0, 2) = "Control Account"
   flxSupplier(0).RowHeight(0) = 0

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1900
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40
   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "NAME"
   lblSearch2(0).Visible = False

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
'         flxSupplier(0).TextMatrix(rRow, 2) = rstRec.Fields.Item(2).Value
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   flxSupplier(0).Sort = 1
   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdTaxList_Click()
   szSel = "TAX"
   LoadVAT

   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList(0).Left = lblVatCode(0).Left + tabTenant.Left
   fraList(0).Top = cmdTaxList.Top + cmdTaxList.Height + tabTenant.Top
'   fraList(0).Width = 3520
'   Picture1.Width = fraList(0).Width - 80
'   flxSupplier(0).Width = Picture1.Width - 40
'   fraList(0).Height = 2805
'   Picture1.Height = fraList(0).Height - 80
   'flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
End Sub

Private Sub LoadVAT()
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "RATE"

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1000
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40

   txtSearch2.Width = 1000
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "RATE"
   lblSearch2(0).Visible = False

   flxSupplier(0).RowHeight(0) = 0

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   rstRec.Open "SELECT VAT_CODE, VAT_RATE FROM tlbVatCode;", adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      flxSupplier(0).TextMatrix(0, 0) = "VAT Code"
      flxSupplier(0).TextMatrix(0, 1) = "VAT Rate"

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!VAT_CODE
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec!VAT_RATE
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Function SQLTenantList() As String
         Dim szWhere As String
        If optBoth.Value Then
                 szWhere = ""
                 If txtClientList.Tag <> "ALL" Then
                        If txtPropertyList.Tag <> "ALL" Then
                           szWhere = szWhere + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' AND "
                        Else
                           szWhere = szWhere + "AND "
                        End If
                        szWhere = szWhere + "Client.ClientID = '" & txtClientList.Tag & "' "
                 Else
                        If txtPropertyList.Tag <> "ALL" Then
                              szWhere = szWhere + "AND IQ.PropertyID = '" & txtPropertyList.Tag & "' "
                        End If
                 End If
                
                SQLTenantList = "SELECT T.SageAccountNumber, T.Name, T.CompanyName, IQ.UnitName, '' AS Balance, IIf(IsNull(T.Comments),'CURRENT','DELETED') AS Notes, IQ.PropertyName, Client.ClientName, IQ.UnitNumber,  LeaseID " & _
                        "FROM (Tenants AS T LEFT JOIN [SELECT U.UnitName, L.SageAccountNumber, P.PropertyName, P.ClientID, U.UnitNumber,P.PropertyID,L.LeaseID FROM Units AS U, LeaseDetails AS L, Property AS P WHERE " & _
                        "U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID ]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber) LEFT JOIN Client " & _
                        "ON IQ.ClientID = Client.ClientID WHERE (((T.Comments) Is Null Or (T.Comments)='') AND ((T.OCCUPIDE_)=False))"
                 SQLTenantList = SQLTenantList + szWhere
                 
                 'this fine tuning is done by anol 20230705
              SQLTenantList = SQLTenantList + szWhere
              SQLTenantList = "SELECT SQL1.SageAccountNumber, SQL1.Name, SQL1.CompanyName, SQL1.UnitName, SQL1.Notes, SQL1.PropertyName, SQL1.ClientName, SQL1.UnitNumber, SQL1.LeaseID, round(SQL2.Amt,2) as amt" & _
                        " FROM" & _
                        " (" & _
                        SQLTenantList & _
                        " ) AS SQL1" & _
                        " LEFT JOIN" & _
                        " (" & _
                        "     SELECT SageAccountNumber,  SUM(Switch(type=1, Amount, type=23, Amount, type=2, -Amount, type=3, -Amount, type=4, -Amount)) AS Amt" & _
                        "     FROM tlbReceipt" & _
                        "     GROUP BY SageAccountNumber" & _
                        " ) AS SQL2" & _
                        " ON SQL1.SageAccountNumber = SQL2.SageAccountNumber"
              SQLTenantList = SQLTenantList + " ORDER BY T.SageAccountNumber;"
              
                
           End If
        'AND IsNull(TerminateDate) has been added by anol 14 May 2015
        'The bug was that it was showing termindated lease
        'Changed join type LEFT to inner
        '  Current tenants Only
        '  Current tenants Only
        
        'Changed join type LEFT to inner  by anol 26 Jun 2015
        'Modified by anol 23 Aug 2015
        'L.Status = TRUE AND ommited
        'AND IsNull(IQ.TerminateDate)
           If optCurrentTenant.Value Then
           'comment out date 26/10/2015
        
        'I am showing property Name and clientName in this SQL
        'Here I need to find the max status if that is true that means current, no need to deduct that,if max is false then we need to deduct that
             szWhere = ""
            'iscurrent and CurrUnit is populating while you click the cmdTenant command button
             SQLTenantList = "SELECT T.SageAccountNumber, T.Name, T.CompanyName, Units.UnitName, IIf(IsNull(T.Comments),'CURRENT','DELETED') AS Notes, Property.PropertyName, " & _
            "Client.ClientName, Units.UnitNumber, '0' as LeaseID FROM Tenants AS T LEFT JOIN ((Units LEFT JOIN Property ON Units.PropertyID = Property.PropertyID) LEFT " & _
            "JOIN Client ON Property.ClientID = Client.ClientID) ON T.CurrUnit = Units.UnitNumber where T.iscurrent=true "

             If txtClientList.Tag <> "ALL" Then
                    If txtPropertyList.Tag <> "ALL" Then
                       szWhere = szWhere + "AND Property.PropertyID = '" & txtPropertyList.Tag & "' AND "
                    Else
                       szWhere = szWhere + "AND "
                    End If
                    szWhere = szWhere + "Client.ClientID = '" & txtClientList.Tag & "' "
              Else
                    If txtPropertyList.Tag <> "ALL" Then
                        szWhere = szWhere + "AND Property.PropertyID = '" & txtPropertyList.Tag & "' "
                    End If
              End If
              'this fine tuning is done by anol 20230705
              SQLTenantList = SQLTenantList + szWhere
              SQLTenantList = "SELECT SQL1.SageAccountNumber, SQL1.Name, SQL1.CompanyName, SQL1.UnitName, SQL1.Notes, SQL1.PropertyName, SQL1.ClientName, SQL1.UnitNumber, SQL1.LeaseID, round(SQL2.Amt,2) as amt" & _
                        " FROM" & _
                        " (" & _
                        SQLTenantList & _
                        " ) AS SQL1" & _
                        " LEFT JOIN" & _
                        " (" & _
                        "     SELECT SageAccountNumber,  SUM(Switch(type=1, Amount, type=23, Amount, type=2, -Amount, type=3, -Amount, type=4, -Amount)) AS Amt" & _
                        "     FROM tlbReceipt" & _
                        "     GROUP BY SageAccountNumber" & _
                        " ) AS SQL2" & _
                        " ON SQL1.SageAccountNumber = SQL2.SageAccountNumber"
              SQLTenantList = SQLTenantList + " ORDER BY T.SageAccountNumber;"
                 Debug.Print SQLTenantList
           End If
        
        '  Deleted tenants only
           If optExTenant.Value Then
        '   "SELECT SageAccountNumber, CompanyName " & _
        '             "FROM Tenants " & _
        '             "WHERE Tenants.SageAccountNumber NOT IN " & _
        '                 "(SELECT LeaseDetails.SageAccountNumber " & _
        '                 "FROM LeaseDetails " & _
        '                 "WHERE Status=True) AND " & _
        '                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
        '             "ORDER BY SageAccountNumber"
        'rem by anol 20170121
'               szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'                           "'' AS Balance, " & _
'                            "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'                            "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
'                       "FROM Tenants AS T INNER JOIN " & _
'                            "[" & _
'                            "SELECT U.UnitName, L.SageAccountNumber, " & _
'                                  "P.PropertyID, P.ClientID, U.UnitNumber " & _
'                            "FROM Units AS U, LeaseDetails AS L, " & _
'                                "Property AS P " & _
'                            "WHERE U.UnitNumber = L.UnitNumber AND " & _
'                                "L.Status = FALSE AND U.PropertyID = P.PropertyID "
'
'              szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'                         "ORDER BY T.SageAccountNumber;"
'implemented by anol 20170121
'if Lestest status true that means this is not an expired lessee. it is current. so you need to find the last status first
'Implemented on 20180514
'        szSQL = "Select SageAccountNumber, CompanyName, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) as UnitName ," & _
'                        "(Select PropertyID from units where IM.Unitnumber=units.Unitnumber) as PropertyID ," & _
'                        "(Select Property.ClientID from units,Property where Property.PropertyID=Units.PropertyID AND  IM.Unitnumber=units.Unitnumber) as ClientID ," & _
'                        " '' AS Balance,IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes,Property.PropertyID, Property.ClientID " & _
'             "FROM (  Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
'             "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
'             "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false) as IM ORDER BY SageAccountNumber"
'
'             szSQL = "SELECT IM.SageAccountNumber, IM.CompanyName, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) AS UnitName, IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes, " & _
'             "Property.PropertyName , Property.clientID FROM (Property INNER JOIN ([Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
'           "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
'             "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false]. AS IM INNER JOIN Units " & _
'             "ON IM.UnitNumber = Units.UnitNumber) ON Property.PropertyID = Units.PropertyID) INNER JOIN Tenants ON IM.SageAccountNumber = Tenants.SageAccountNumber " & _
'            "ORDER BY IM.SageAccountNumber;"
               szWhere = ""
                SQLTenantList = "SELECT IM.SageAccountNumber,tenants.Name, IM.CompanyName, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) AS UnitName, IIf(IsNull(Tenants.Comments),'CURRENT','DELETED') AS Notes, " & _
                "Property.PropertyName, Client.ClientName,units.Unitnumber,LeaseID FROM ((Property INNER JOIN ([Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber,S.LeaseID  from LeaseDetails as S INNER JOIN " & _
                "(SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails GROUP BY LeaseDetails.SageAccountNumber) as IQ ON " & _
                "IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false]. AS IM INNER JOIN Units ON IM.UnitNumber = Units.UnitNumber) ON " & _
                "Property.PropertyID = Units.PropertyID) INNER JOIN Tenants ON IM.SageAccountNumber = Tenants.SageAccountNumber) INNER JOIN Client ON Property.ClientID = Client.ClientID "
                

                If txtClientList.Tag <> "ALL" Then
                        If txtPropertyList.Tag <> "ALL" Then
                            szWhere = szWhere + "WHERE Property.PropertyID = '" & txtPropertyList.Tag & "' "
                        End If
                        If szWhere <> "" Then
                            szWhere = szWhere + "AND "
                        Else
                            szWhere = szWhere + "WHERE "
                        End If
                        szWhere = szWhere + "Client.ClientID = '" & txtClientList.Tag & "' "
            
                  Else
                        If txtPropertyList.Tag <> "ALL" Then
                              szWhere = szWhere + "WHERE Property.PropertyID = '" & txtPropertyList.Tag & "' "
                        End If
            
                  End If
                   'SQLTenantList = SQLTenantList + szWhere + " ORDER BY IM.SageAccountNumber;"
                   
                    'this fine tuning is done by anol 20230705
              SQLTenantList = SQLTenantList + szWhere
              SQLTenantList = "SELECT SQL1.SageAccountNumber, SQL1.Name, SQL1.CompanyName, SQL1.UnitName, SQL1.Notes, SQL1.PropertyName, SQL1.ClientName, SQL1.UnitNumber, SQL1.LeaseID, round(SQL2.Amt,2) as amt" & _
                        " FROM" & _
                        " (" & _
                        SQLTenantList & _
                        " ) AS SQL1" & _
                        " LEFT JOIN" & _
                        " (" & _
                        "     SELECT SageAccountNumber,  SUM(Switch(type=1, Amount, type=23, Amount, type=2, -Amount, type=3, -Amount, type=4, -Amount)) AS Amt" & _
                        "     FROM tlbReceipt" & _
                        "     GROUP BY SageAccountNumber" & _
                        " ) AS SQL2" & _
                        " ON SQL1.SageAccountNumber = SQL2.SageAccountNumber"
              SQLTenantList = SQLTenantList + " ORDER BY SQL1.SageAccountNumber;"
              
              
             
 'following one is the wrong sql ANOL 20180514
'                 szSQL = "SELECT  LeaseDetails.SageAccountNumber, " & _
'                        "Tenants.CompanyName, UnitName, '' AS Balance,IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes," & _
'                        "Property.PropertyID, Property.ClientID " & _
'                        "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
'                        "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
'                      "LeaseDetails.Status = false And " & _
'                      "Units.PropertyId = Property.PropertyID And " & _
'                      "Property.ClientID = Client.ClientID AND " & _
'                      "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
'                      "" & szWhere & " " & _
'                  "ORDER BY " & szOrderBy & ";"
           End If
End Function
Private Sub cmdTenantLookup_Click()
   If cmdSave.Enabled Then Exit Sub
'   txtPropertyList.Tag = "ALL"
'   txtPropertyList.text = "ALL"
    fmeTenantLookup.Left = txtName.Left
    fmeTenantLookup.Top = optCurrentTenant.Top + 80
    fmeTenantLookup.Visible = True
    fmeTenantLookup.ZOrder 0
    cmdTenantLookup.Enabled = False
    cmdTenantLookup.Visible = False
    cmdTenantLookup.Refresh
    
    
    FocusControl txtSearchTenant
    
   If strSessionPropertyID = "" Then
        txtPropertyList.Tag = "ALL"
        txtPropertyList.text = "ALL"
   Else
        txtPropertyList.Tag = strSessionPropertyID
        txtPropertyList.text = strSessionPropertyID
   End If
'   If frmMMain.Leasee4_LesseList_isUptoDate = False Then
'        Conn.Open getConnectionString
'        TenantAccountBalance Conn
'        Conn.Close
'        frmMMain.Leasee4_LesseList_isUptoDate = True
'   End If
'   txtClientList.Tag = "ALL"
'   txtClientList.text = "ALL"
 'issue  402 the client/property selection should remain active until the user changes it or closes the lease details form'by anol 20170608
   If strSessionClientID = "" Then
        txtClientList.Tag = "ALL"
        txtClientList.text = "ALL"
   Else
        txtClientList.Tag = strSessionClientID
        txtClientList.text = strSessionClientID
   End If
   
   fmeTenant.Enabled = False
   tabTenant.Enabled = False
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   'rem by anol 2023-07-05
'   If bolgridTenantLookupRefresh = True Then
'        fmeTenantLookup.Visible = True
'        cmdTenantLookup.Enabled = True
'        Exit Sub
'   Else
'        bolgridTenantLookupRefresh = False
'   End If
   
'   If txtSearchTenant.text = "" Then
   txtSearchTenant.text = ""
   adoConn.Open getConnectionString

 'this is setting up the status and unit in the tentants table for pop up grid issue 656 CurrUnit and iscurrent filed has been newly created for this quick search they needs
 ' to be updated by the following update query before search to get the result with status and unit
   adoConn.Execute "UPDATE Tenants T SET T.iscurrent=true, T.CurrUnit=NULL"
   adoConn.Execute "UPDATE Tenants T INNER JOIN LeaseDetails L ON T.SageAccountNumber = L.SageAccountNumber SET iscurrent=status where L.status=false"
   adoConn.Execute "UPDATE Tenants T SET T.iscurrent=false where Not isnull(T.Comments) OR T.Comments<>''"
   adoConn.Execute "UPDATE Tenants T INNER JOIN LeaseDetails L ON T.SageAccountNumber = L.SageAccountNumber SET iscurrent=status,CurrUnit=L.UnitNumber where L.status=true;"
   
   'PrepareList adoConn             'prepare the list of clients and properties in dwopdown comboes
   'anol on 28 Feb 2016
   'I am not building balances any more ' rem by anol 2023-07-05
'   If frmMMain.Leasee1_LesseList_isUptoDate = False Then
'        lblLoading.Caption = "Please wait while Building lessee balances..."
'        fmeLoading.Visible = True
'        TenantAccountBalance adoconn
'        fmeLoading.Visible = False
'        frmMMain.Leasee1_LesseList_isUptoDate = True
'   End If
   adoConn.Close
   Set adoConn = Nothing
   Call FilterTenantsList("")   'Loading the search grid lessee list
   bolgridTenantLookupRefresh = True
'           Dim szWhere As String
'           Dim szOrderBy As String
'
'           szOrderBy = "LeaseDetails.SageAccountNumber ASC"
'
'           If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
'              szWhere = ""
'
'           If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
'              szWhere = "AND LA.CLIENTID = '" & txtClientList.Tag & "' "
'
'           If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
'              szWhere = "AND LA.PROPERTYID = '" & txtPropertyList.Tag & "' "
'
'           If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
'              szWhere = "AND LA.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
'                                 "AND LA.CLIENTID = '" & txtClientList.Tag & "' "
'
'           Dim Filter As String
'           If Len(txtSearchTenant.text) > 0 Then
'              txtSearchName.text = ""
'              txtSearchUnitName.text = ""
'              Filter = " SageAccountNumber LIKE '%" + UCase(txtSearchTenant.text) + "*'"
'
'           End If
'
'           If Len(txtSearchName.text) > 0 Then
'              txtSearchTenant.text = ""
'              txtSearchUnitName.text = ""
'              Filter = " CompanyName LIKE '%" + UCase(txtSearchName.text) + "*'"
'           End If
'
'           If Len(txtSearchUnitName.text) > 0 Then
'              txtSearchTenant.text = ""
'              txtSearchName.text = ""
'              Filter = " UnitName LIKE '%" + UCase(txtSearchUnitName.text) + "*'"
'           End If
'          'All tenants
'
'
'            ConfigGridTenantLookup
'            fmeLoading.Visible = True
'            fmeLoading.Top = 2715
'            fmeLoading.Left = 3560
'            fmeLoading.ZOrder 0
'            lblLoading.ZOrder 0
'            fmeLoading.Refresh
'
'            Dim adoRst As New ADODB.Recordset
'            adoRst.Open SQLTenantList, adoConn, adOpenStatic, adLockReadOnly
'
'
'           Dim iRow As Integer
'           iRow = 1
'           gridTenantLookup.Rows = adoRst.RecordCount + 1
'           While Not adoRst.EOF
'              gridTenantLookup.TextMatrix(iRow, 0) = adoRst!SageAccountNumber
'              gridTenantLookup.TextMatrix(iRow, 1) = adoRst!Name
'              gridTenantLookup.TextMatrix(iRow, 2) = adoRst!CompanyName
'              gridTenantLookup.TextMatrix(iRow, 3) = IIf(IsNull(adoRst!UnitNumber), "", adoRst!UnitNumber)
'              gridTenantLookup.TextMatrix(iRow, 4) = IIf(IsNull(adoRst!Notes), "", adoRst!Notes)
'              gridTenantLookup.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!PropertyName), "", adoRst!PropertyName)
'              gridTenantLookup.TextMatrix(iRow, 6) = IIf(IsNull(adoRst!ClientName), "", adoRst!ClientName)
'              adoRst.MoveNext
'              iRow = iRow + 1
'           Wend
'
'            adoRst.Close
'            Set adoRst = Nothing
'
'            fmeLoading.Visible = False
'
'            UpdateBalance
'
'           If gridTenantLookup.Rows > 1 Then
'                gridTenantLookup.row = 1
'           End If
'           adoConn.Close
'           Set adoConn = Nothing

   
   cmdTenantLookup.Enabled = True
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer
   
   For i = 1 To gridTenantLookup.Rows - 1
      For j = 0 To UBound(szaTenantBalance, 2) - 1
         If gridTenantLookup.TextMatrix(i, 0) = szaTenantBalance(0, j) Then
            'gridTenantLookup.TextMatrix(i, 3) = Format(szaTenantBalance(1, j), "0.000")
            gridTenantLookup.TextMatrix(i, 7) = RoundingNumber2(szaTenantBalance(1, j))
           ' gridTenantLookup.TextMatrix(i, 3) = RoundingNumber(szaTenantBalance(1, j), 2)
            'gridTenantLookup.TextMatrix(i, 3) = szaTenantBalance(1, j)
            
            Exit For
         End If
      Next j
      If j = UBound(szaTenantBalance, 2) Then gridTenantLookup.TextMatrix(i, 7) = "0.00"
   Next i
End Sub
'Public Function Roundingnumber2(a As String) As String
'    Roundingnumber2 = Left(Format(a, "0.000"), InStr(Format(a, "0.000"), ".") + 2)
'End Function
Private Sub PrepareList(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler

'*************************************** CLIENT ********************************************
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

'*************************************** PROPERTY ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode, ClientID " & _
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
'   ReDim dataProperty(TotalCol, TotalRow) As String
'
'   dataProperty(0, 0) = "ALL"
'   dataProperty(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'      For j = 0 To TotalCol - 1
'         dataProperty(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboPropertyList.Column() = dataProperty()
'   cboPropertyList.ListIndex = 0
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
End Sub

Private Sub cmdUnitLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmUnits2
   frmUnits2.LOAD_UNIT_UNITID = txtUnitNumber.text
   frmUnits2.Show
End Sub

Private Sub cmdUnitMemoCancel_Click()
   'Issue 488
   'Modified by anol 04 Oct 2014
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   'MemoButtonEnable False
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdDelete.Enabled = False
   txtUnitMemo.Locked = True
   gridLeaseAnalysis.Enabled = True
   txtUnitMemo.text = ""
   Picture2.Visible = True
   cmdVAMemo.Enabled = True
   cmdVAMemo.Visible = False
   txtMemoAll.SetFocus
End Sub

Private Sub cmdUnitMemoEdit_Click()
   'Modified by Anol 02 Nov 2014
   'Issue 488  Memo and attachment not saving a record of the date of the memo entry
      Picture2.Visible = False
       If txtLeaseAnalysisID.text = "" Then
            cmdVAMemo.Enabled = True
      Else
            cmdVAMemo.Enabled = False
      End If
      cmdVAMemo.Visible = True
      If txtLeaseAnalysisID.text = "" Then
          ShowMsgInTaskBar "Please select the memo you would like to edit", "Y"
          If gridLeaseAnalysis.Enabled = True Then
               gridLeaseAnalysis.SetFocus
          End If
          Exit Sub
      End If
      
      cmdUnitMemoNew.Enabled = False
      cmdUnitMemoEdit.Enabled = False
      cmdUnitMemoSave.Enabled = True
      cmdUnitMemoCancel.Enabled = True

      fmeTenant.Enabled = False
      txtUnitMemo.Locked = False
      Lease_ANALYSIS_NEW_ENTRY = False
   If txtUnitMemo.Enabled = True Then
      txtUnitMemo.SetFocus
   End If
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtUnitMemo.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub cmdUnitMemoNew_Click()
   Lease_ANALYSIS_NEW_ENTRY = True
   fmeTenant.Enabled = False
   cmdUnitMemoNew.Enabled = False
   cmdUnitMemoEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdUnitMemoSave.Enabled = True
   cmdUnitMemoCancel.Enabled = True
   gridLeaseAnalysis.Enabled = False
   txtUnitMemo.Locked = False
   Picture2.Visible = False
   txtUnitMemo.text = ""
   txtUnitMemo.SetFocus
End Sub
Private Sub ConfigGridLeaseAnalysis()
   'Issue 488
   'Added by anol 03 Nov 2014
   Dim szHeader As String
   gridLeaseAnalysis.Clear
   gridLeaseAnalysis.Rows = 1
   gridLeaseAnalysis.Cols = 7
    szHeader$ = "<SL|<Date|<Description|>User"
    gridLeaseAnalysis.FormatString = szHeader$
   gridLeaseAnalysis.TextMatrix(0, 0) = "SL"
   gridLeaseAnalysis.TextMatrix(0, 1) = "MemoID"
   gridLeaseAnalysis.TextMatrix(0, 2) = "MemoType"
   gridLeaseAnalysis.TextMatrix(0, 3) = "SageAccountNumber"
   gridLeaseAnalysis.TextMatrix(0, 4) = "UpdateTime"
   gridLeaseAnalysis.TextMatrix(0, 5) = "MemoDescription"
   gridLeaseAnalysis.TextMatrix(0, 6) = "UserName"
     
   gridLeaseAnalysis.ColWidth(0) = 600
   gridLeaseAnalysis.ColWidth(1) = 0
   gridLeaseAnalysis.ColWidth(2) = 0
   gridLeaseAnalysis.ColWidth(3) = 0
   gridLeaseAnalysis.ColWidth(4) = Label33(11).Left - Label33(10).Left
   gridLeaseAnalysis.ColWidth(5) = Label33(12).Left - Label33(11).Left
   gridLeaseAnalysis.ColWidth(6) = 1550
End Sub
Public Sub PopulateGridLeaseAnalysis()
   'Issue 488
   'Added by anol 03 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtTenantID.text & "' And  MemoType='Lease' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
   Dim iRow As Integer
   iRow = 1

   gridLeaseAnalysis.Clear
   gridLeaseAnalysis.Rows = 1
   gridLeaseAnalysis.Cols = 7
   gridLeaseAnalysis.RowHeight(0) = 0
   If rstLeaseAnalysis_.EOF = True Then
       gridLeaseAnalysis.Rows = 2
   End If
   
   While Not rstLeaseAnalysis_.EOF
      gridLeaseAnalysis.AddItem ""
      gridLeaseAnalysis.TextMatrix(iRow, 0) = iRow
      gridLeaseAnalysis.TextMatrix(iRow, 1) = rstLeaseAnalysis_!MemoID
      gridLeaseAnalysis.TextMatrix(iRow, 2) = rstLeaseAnalysis_!MemoType 'height 0
      gridLeaseAnalysis.TextMatrix(iRow, 3) = rstLeaseAnalysis_!SageAccountNumber 'height 0
      gridLeaseAnalysis.TextMatrix(iRow, 4) = rstLeaseAnalysis_!UpdateTime
      gridLeaseAnalysis.TextMatrix(iRow, 5) = rstLeaseAnalysis_!MemoDescription
      gridLeaseAnalysis.TextMatrix(iRow, 6) = rstLeaseAnalysis_!UserName
      rstLeaseAnalysis_.MoveNext
      iRow = iRow + 1
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   If iRow > 0 Then
      gridLeaseAnalysis.row = 0
   End If
End Sub
Private Function SaveLeaseAnalysis() As Boolean
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim conMemo As New ADODB.Connection
   Dim rstLease_ As New ADODB.Recordset
   conMemo.Open getConnectionString
   Dim sSQLQuery_ As String
   Dim sSQLFilter As String
   If Not Lease_ANALYSIS_NEW_ENTRY Then
       sSQLFilter = "WHERE MemoID = " & Val(gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 1)) & " AND Memotype='Lease' AND SageAccountNumber = '" & txtTenantID.text & "'"
   Else
       sSQLFilter = ""
   End If
   sSQLQuery_ = "SELECT * " & _
                "FROM MemoDetails " & sSQLFilter
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenDynamic, adLockOptimistic
   If Lease_ANALYSIS_NEW_ENTRY Then rstLeaseAnalysis_.AddNew
   If Lease_ANALYSIS_NEW_ENTRY = False Then
      rstLeaseAnalysis_!MemoID = txtLeaseAnalysisID.text
   Else
      rstLeaseAnalysis_!MemoID = NewMemoID()
   End If
   
   rstLeaseAnalysis_!MemoType = "Lease"
   rstLeaseAnalysis_!SageAccountNumber = txtTenantID.text
   rstLeaseAnalysis_!MemoDescription = IIf(txtUnitMemo.text <> "", txtUnitMemo.text, "")
   rstLeaseAnalysis_!UpdateTime = Now
   rstLeaseAnalysis_!UserName = frmMMain.SystemUserName
   rstLeaseAnalysis_.Update
   rstLeaseAnalysis_.Close
   Set rstLease_ = Nothing
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   SaveLeaseAnalysis = True
End Function
Private Function NewMemoID() As Integer
   Dim conMemo As New ADODB.Connection
   conMemo.Open getConnectionString
   Dim szSQL As String
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT MAX(MemoID) AS x   " & _
                 "FROM MemoDetails ;"
   rstSet.Open szSQL, conMemo, adOpenStatic, adLockReadOnly

   NewMemoID = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
   rstSet.Close
   Set rstSet = Nothing
   conMemo.Close
End Function
Private Sub cmdUnitMemoSave_Click()
   cmdVAMemo.Visible = False
   If Len(txtUnitMemo.text) = 0 Then
      ShowMsgInTaskBar "Please enter description of memo", "Y"
      If txtUnitMemo.Enabled = True Then
         txtUnitMemo.SetFocus
      End If
      Exit Sub
   End If
   If gridLeaseAnalysis.row = 0 And Lease_ANALYSIS_NEW_ENTRY = False Then
       ShowMsgInTaskBar "Please select a memo from list", "Y"
       Exit Sub
   End If
   
   If SaveLeaseAnalysis Then
      ShowMsgInTaskBar "The memo has been saved successfully."
      PopulateGridLeaseAnalysis
   Else
      ShowMsgInTaskBar "Could not save lease analysis", , "N"
   End If
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdUnitMemoCancel.Enabled = False
   gridLeaseAnalysis.Enabled = True
   gridLeaseAnalysis.row = 0
   txtUnitMemo.text = ""
   txtLeaseAnalysisID.text = ""
   txtUnitMemo.Locked = True
   fmeTenant.Enabled = True
   cmdDelete.Enabled = False
   cmdVAMemo.Enabled = True
   Picture2.Visible = True
   txtMemoAll.text = ""
   Call ViewMemo
   txtMemoAll.SetFocus
End Sub

Private Sub cmdViewLetter_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String
   Dim rstReport As ADODB.Recordset
   Dim rstRst As New ADODB.Recordset

   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub

   adoConn.Open getConnectionString
   
   szSQL = "UPDATE tlbLetterReports SET isPrint = '';"
   adoConn.Execute szSQL
   
   For i = 1 To flxLetters.Rows - 1
    If (flxLetters.TextMatrix(i, 0) = "X") Then
      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE id = " & flxLetters.TextMatrix(i, 1) & ";"
      adoConn.Execute szSQL
    End If
   Next i

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdVAMemo_Click()
   Picture2.Visible = True
   txtMemoAll.text = ""
   Call ViewMemo
   cmdVAMemo.Visible = False
   txtMemoAll.SetFocus
End Sub

Private Sub dtpDateCompleted_Change()
   TextBoxChangeDate dtpDateCompleted
End Sub

Private Sub dtpDateCompleted_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpDateCompleted, KeyAscii
End Sub

Private Sub dtpDateCompleted_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpDateCompleted
End Sub

Private Sub dtpRemindDate_Change()
   TextBoxChangeDate dtpRemindDate
End Sub

Private Sub dtpRemindDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpRemindDate, KeyAscii
End Sub

Private Sub dtpRemindDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpRemindDate
End Sub

Private Sub dtpReportedDate_Change()
   TextBoxChangeDate dtpReportedDate
End Sub

Private Sub dtpReportedDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpReportedDate, KeyAscii
End Sub

Private Sub dtpReportedDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpReportedDate
End Sub

Private Sub flxACHistory_Click()
   Dim iCurRowHeight As Integer
   Dim iRow          As Integer
   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
   Dim szSQL         As String

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "" Then GoTo ChildGrid

'****************************************************** EXPANDING THE GRID *********************************
   iRow = flxACHistory.row
   iCurRowHeight = flxACHistory.RowHeight(iRow)

   If flxACHistory.col = 0 Then
      If flxACHistory.TextMatrix(iRow, 0) = "-" Then Exit Sub
      If flxACHistory.TextMatrix(iRow, 0) = "+" And flxACHistory.RowHeight(iRow + 1) = 0 Then
      If flxACHistory.TextMatrix(iRow, 0) = "" Then Exit Sub
         flxACHistory.TextMatrix(iRow, 0) = ">"
         For iRow = iRow + 1 To flxACHistory.Rows - 1
            If flxACHistory.TextMatrix(iRow, 0) = "+" Or flxACHistory.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxACHistory.TextMatrix(iRow, 0) = "-" Then flxACHistory.RowHeight(iRow) = iCurRowHeight
         Next iRow
      ElseIf flxACHistory.TextMatrix(iRow, 0) = ">" And flxACHistory.RowHeight(iRow + 1) = iCurRowHeight Then
         flxACHistory.TextMatrix(iRow, 0) = "+"
         For iRow = iRow + 1 To flxACHistory.Rows - 1
            If flxACHistory.TextMatrix(iRow, 0) = "+" Or flxACHistory.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxACHistory.TextMatrix(iRow, 0) = "-" Then flxACHistory.RowHeight(iRow) = 0
         Next iRow
      End If
   End If
'***********************************************************************************************************
   'HighLightRowFlxGrid flxACHistory, flxACHistory.row

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "-" Then Exit Sub

ChildGrid:
'  Displaying the splits ************************************************************************************
   
   ConfigFlxACHistorySplit
   adoConn.Open getConnectionString

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SC" Then
      szSQL = "SELECT S.*,F.* " & _
              "FROM tlbReceipt AS R, DemandRecords AS D, DemandSplitRecords AS S ,Fund F " & _
              "WHERE R.DemandRef = D.DemandID AND " & _
                  "D.SageAccountNumber = '" & Trim(txtTenantID.text) & "' AND " & _
                  "D.DemandID = S.DemandID AND F.FundID=S.SageDepartment AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", 1, 2) & " AND " & _
                  "R.SlNumber = " & StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & " " & _
              "ORDER BY S.SplitID;"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = adoRST.Fields.Item("SplitID").Value
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("DueDate").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundCode").Value 'SageDepartment is  FundID
            .TextMatrix(iRow, 5) = adoRST.Fields.Item("NominalCodeforAmount").Value
            .TextMatrix(iRow, 6) = adoRST.Fields.Item("DateFrom").Value
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("DateTo").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRST.Close
      End With
   End If

   If (Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SA") And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "SRR" Then
      szSQL = "SELECT S.*, R.SlNumber " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S " & _
              "WHERE R.TransactionID = S.RptHeader AND " & _
                  "R.SageAccountNumber = '" & Trim(txtTenantID.text) & "' AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR", 3, 4) & " AND " & _
                  "R.SlNumber = " & StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = adoRST.Fields.Item("SlNumber").Value
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2) '
            'Due date needs to show from the tlbReceipt Due date
            'Modified by anol 06 oct 2015
            .TextMatrix(iRow, 3) = flxACHistory.TextMatrix(flxACHistory.row, 3) 'IIf(IsNull(adoRst.Fields.Item("DueDate").Value), "", adoRst.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundID").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = ""
            .TextMatrix(iRow, 7) = ""
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "SRR" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S " & _
              "WHERE R.TransactionID = S.RptHeader AND " & _
                  "R.Type = 23 AND " & _
                  "R.SlNumber = " & Mid(flxACHistory.TextMatrix(flxACHistory.row, 1), 4, _
                                    Len(flxACHistory.TextMatrix(flxACHistory.row, 1)) - 3) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = ""
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 3) = IIf(IsNull(adoRST.Fields.Item("DueDate").Value), "", adoRST.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundID").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = ""
            .TextMatrix(iRow, 7) = ""
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxDeposit_Click()
   flxDeposit_RowColChange
   Call SquezeExpand
End Sub

Private Sub flxDeposit_RowColChange()
'   On Error Resume Next

   iGridRow = flxDeposit.row

   ButtonHanlding GridRowOnSelection

   txtBank.Tag = flxDeposit.TextMatrix(flxDeposit.row, 13) ' NO need after 10
   txtBank.text = flxDeposit.TextMatrix(flxDeposit.row, 20) ' NO need after 10
   txtDNC(0).text = flxDeposit.TextMatrix(flxDeposit.row, 14) ' NO need after 10 Deposit Nominal Code
   txtDNC(1).text = flxDeposit.TextMatrix(flxDeposit.row, 17) ' NO need after 10 Deposit Nominal Name
   txtDate.text = flxDeposit.TextMatrix(flxDeposit.row, 6) 'date
   txtDepositType.text = flxDeposit.TextMatrix(flxDeposit.row, 16) ' NO need after 10
   txtDepositType.Tag = flxDeposit.TextMatrix(flxDeposit.row, 22)  ' NO need after 10
   txtDptDetails.text = flxDeposit.TextMatrix(flxDeposit.row, 8) 'No 8 is consistent as
   If flxDeposit.TextMatrix(flxDeposit.row, 9) <> "" Then
      txtDptAmount.text = flxDeposit.TextMatrix(flxDeposit.row, 9) ' NO need after 10
   Else
      txtDptAmount.text = flxDeposit.TextMatrix(flxDeposit.row, 11) ' NO need after 10
   End If
   txtDptAmtType.text = flxDeposit.TextMatrix(flxDeposit.row, 15) ' NO need after 10
   txtOSDpt.text = flxDeposit.TextMatrix(flxDeposit.row, 10) ' NO need after 10
   cboGroup.text = flxDeposit.TextMatrix(flxDeposit.row, 12) ' NO need after 10
   GROUP_NO = flxDeposit.TextMatrix(flxDeposit.row, 13) ' NO need after 10
   txtFund.Tag = flxDeposit.TextMatrix(flxDeposit.row, 18) ' NO need after 10
   txtFund.text = flxDeposit.TextMatrix(flxDeposit.row, 21) ' NO need after 10
   'Label6(5).Caption = flxDeposit.TextMatrix(flxDeposit.row, 7) & " Type "
   If flxDeposit.TextMatrix(flxDeposit.row, 3) = "Deposit" Then
        cmdDptRefund.Caption = "Deposit Refund"
   Else
        cmdDptRefund.Caption = "Expense Refund"
   End If
End Sub

Private Sub flxEmails_RowColChange()
   SelectOnly1RowFlxGrid flxEmails, flxEmails.row
End Sub

Private Sub flxLetters_Click()
    Dim i As Integer
    i = Select1RowFlxGrid(flxLetters, flxLetters.row, 0)
End Sub

Private Sub flxSupplier_Click(Index As Integer)
   If szSel = "TAX" Then
      lblVatCode(0).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtCodeVat.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   If szSel = "NC" Then
      txtNominalCode.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtNominalCodeName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      FocusControl cmdDepositType
   End If
   If szSel = "SLC" Then
      txtSLControl.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtSLControlName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   If szSel = "DH" Then
      txtDNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtDNC(1).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      tabTenant.Enabled = True
      FocusControl cmdFund
   End If
   If szSel = "BANK" Then
      txtBank.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtBank.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      tabTenant.Enabled = True
      cmdNCList.SetFocus
   End If
   If szSel = "FUND" Then
      txtFund.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
      txtFund.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      tabTenant.Enabled = True
      FocusControl cmdDepositType
   End If
   If szSel = "Type" Then
        tabTenant.Enabled = True
        txtDepositType.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtDepositType.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        FocusControl txtDptDetails
   End If
    If szSel = "DptAmtType" Then
        tabTenant.Enabled = True
        txtDptAmtType.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        txtDptAmtType.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        FocusControl cmdDptSave
   End If
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
        fmeTenantLookup.Enabled = True
        bolgridTenantLookupRefresh = False
        Call FilterTenantsList("")
        cmdPropertyList.SetFocus
   End If
    If szSel = "Property" Then
        txtPropertyList.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtPropertyList.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
        fmeTenantLookup.Enabled = True
        'cboPropertyList_Click
        bolgridTenantLookupRefresh = False
        Call FilterTenantsList("")
        txtSearchTenant.SetFocus
   End If
   flxSupplier(0).Clear
   fraList(0).Visible = False
End Sub

Private Sub fmeTenant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Me.MousePointer = vbArrow
End Sub

Private Sub Form_Activate()
   Dim iRow As Integer
    Dim adoConn As New ADODB.Connection
   
    sessionID = GetTimeStamp
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
     'I am not building balances any more ' rem by anol 2023-07-05
''    adoconn.Open getConnectionString
''    'Load all tenants with balance in gridTenantLookup
''    'I am moving this thing to form active because the user Needs to see the balance is building and no wait to see the form
''    'lblLoading.caption=Please wait while loading...
''
''
''    If frmMMain.Leasee1_LesseList_isUptoDate = False Then
''        lblLoading.Caption = "Please wait while Building lessee balances..."
''        fmeLoading.Visible = True
''        Me.Refresh
''        TenantAccountBalance adoconn
''        frmMMain.Leasee1_LesseList_isUptoDate = True
''        fmeLoading.Visible = False
''    End If
''    adoconn.Close
''    Set adoconn = Nothing
   
   If LOAD_TENANT_TENANTID <> "" Then
      'issue 488
      'Added by Anol 03 Nov 2014
      LoadTenantByTenantID
      PopulateGridLeaseAnalysis
      Call ViewMemo
   End If
End Sub
Private Sub disableLeaseHeldBoxed()
    cmdBank.Enabled = False
    txtBank.Enabled = False
    txtDNC(0).Enabled = False
    txtDNC(1).Enabled = False
    cmdNewDRefund.Enabled = True
    txtDepositType.Enabled = False
    txtDptAmtType.Enabled = False
    cmdNCList.Enabled = False
    txtDate.Enabled = False
    cmdDepositType.Enabled = False
    cmdSetDptType.Enabled = False
    txtDptDetails.Enabled = False
    txtDptAmount.Enabled = False
    txtOSDpt.Enabled = False
    cmdDptAmtType.Enabled = False
    cmdSetAmtType.Enabled = False
    txtFund.Enabled = False
    cmdFund.Enabled = False
End Sub
Private Sub EnableLeaseHeldBoxed()
    cmdBank.Enabled = True
    txtBank.Enabled = True
    cmdNewDRefund.Enabled = False
    txtDNC(0).Enabled = True
    txtDNC(1).Enabled = True
    txtDepositType.Enabled = True
    txtDptAmtType.Enabled = True
    cmdNCList.Enabled = True
    txtDate.Enabled = True
    cmdDepositType.Enabled = True
    cmdSetDptType.Enabled = True
    txtDptDetails.Enabled = True
    txtDptAmount.Enabled = True
    txtOSDpt.Enabled = True
    cmdDptAmtType.Enabled = True
    cmdSetAmtType.Enabled = True
    txtFund.Enabled = True
    cmdFund.Enabled = True
End Sub
'Private Sub UpdateDatabase_LeaseDetails(adoConn As ADODB.Connection)
'    On Error GoTo Err
'    Dim rst1 As New ADODB.Recordset
'    rst1.Open "Select IncreamentalID from LeaseDetails", adoConn
'    rst1.Close
'    Exit Sub
'Err:
'    adoConn.Execute "ALTER TABLE LeaseDetails ADD COLUMN IncrementalID Long;"
'    adoConn.Execute "Update LeaseDetails set IncrementalID=Left(right(LeaseID,12),8)"
'End Sub
Private Sub Form_Load()
'   MousePointer = vbHourglass
    'szaTenantBalance
'    Frame1(3).Enabled = False
   Call disableLeaseHeldBoxed
   If WS_Name = "PCM-DEV2" Then
        Command1.Visible = True
        Command2.Visible = True
   End If
   'issue 488
   'Modified by anol 04 Nov 2014
   ConfigGridLeaseAnalysis
   Picture2.Top = 180
   Picture2.Left = 40
   'End of modification
   Me.Height = tabTenant.Top + tabTenant.Height + 555
   Me.Width = 12615 ' 11775
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fmeTenant.BackColor = Me.BackColor
   tabTenant.BackColor = Me.BackColor
   optCurrentTenant.BackColor = fmeTenant.BackColor
   optExTenant.BackColor = fmeTenant.BackColor
   optBoth.BackColor = fmeTenant.BackColor
   Frame13.BackColor = Me.BackColor
   optCurrentTenant.BackColor = Me.BackColor
   optExTenant.BackColor = Me.BackColor
   optBoth.BackColor = Me.BackColor
   fmeLoading.Top = 4305
   fmeLoading.Left = 4508
   Me.Caption = "Lessee"
   tabTenant.Tab = 0
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameEnableMode Me, fmeTenancyDetails, DefaultMode

   txtSearchTenant.Enabled = True
   COPYMODE_ = False
   NEWMODE_ = False
   SEARCHTenantMODE_ = True
   bDEPOSIT_HELD = False

'   SageCustomerAccCombo cboSageAccountNumber
   TenantTabEnabled False

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   'Call UpdateDatabase_LeaseDetails(adoConn)
 'end of addition
'    Populate the codes
   PopulateCodes adoConn

'  Load all tenants with balance in gridTenantLookup
'I am moving this ti to form active because the user Needs to see the ballance is building and no wait to see the form
''   If frmMMain.Leasee1_LesseList_isUptoDate = False Then
''        TenantAccountBalance adoConn
''        frmMMain.Leasee1_LesseList_isUptoDate = True
''   End If

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
'   MousePointer = vbDefault
End Sub

Private Sub PopulateLessee(adoConn As ADODB.Connection)
'   Dim szSQL As String
'   Dim rstSQL As New ADODB.Recordset
End Sub

Private Sub ConfigFlxDeposit()
    Dim szHeader As String, iCol As Integer
    
    flxDeposit.Clear
    flxDeposit.Cols = 23
    flxDeposit.Rows = 2
    flxDeposit.RowHeight(0) = 0
    flxDeposit.ColWidth(0) = 280
    flxDeposit.ColWidth(1) = 0
    flxDeposit.ColWidth(3) = Label6(2).Left - Label6(1).Left
    flxDeposit.ColAlignment(2) = vbLeftJustify
    flxDeposit.ColWidth(4) = 0 'Label6(3).Left - Label6(2).Left
    flxDeposit.ColWidth(5) = Label6(4).Left - Label6(3).Left
    flxDeposit.ColWidth(6) = Label6(5).Left - Label6(4).Left
    flxDeposit.ColWidth(7) = Label6(6).Left - Label6(5).Left
    flxDeposit.ColWidth(8) = Label6(7).Left - Label6(6).Left
    flxDeposit.ColWidth(9) = Label6(8).Left - Label6(7).Left
    flxDeposit.ColWidth(10) = Label6(9).Left - Label6(8).Left
    '   For iCol = 3 To flxDeposit.Cols - 12
    '      flxDeposit.ColWidth(iCol) = Label6(iCol + 1).Left - Label6(iCol).Left
    '      flxDeposit.ColAlignment(iCol) = vbLeftJustify
    '   Next iCol
    '    flxDeposit.ColWidth(9) = 0
    flxDeposit.ColWidth(11) = 800
    flxDeposit.ColWidth(12) = 1000
    'flxDeposit.ColWidth(iCol - 1) = flxDeposit.Width + flxDeposit.Left - Label6(iCol - 1).Left - 280   'col = 9
    flxDeposit.ColWidth(13) = 0
    flxDeposit.ColWidth(14) = 0
    flxDeposit.ColWidth(15) = 0
    flxDeposit.ColWidth(16) = 0
    flxDeposit.ColWidth(17) = 0
    flxDeposit.ColWidth(18) = 0
    flxDeposit.ColWidth(19) = 0
    
    flxDeposit.ColWidth(19) = 0
    flxDeposit.ColWidth(20) = 0
    flxDeposit.ColWidth(21) = 0 'Deposit type code
    flxDeposit.ColWidth(22) = 0
End Sub

Private Sub ConfigFlxACHistorySplit()
   Dim szHeader As String, iCol As Integer

   flxACHistorySplit.Clear
   flxACHistorySplit.Cols = 13
   flxACHistorySplit.Rows = 2
   flxACHistorySplit.RowHeight(0) = 0

   flxACHistorySplit.ColWidth(0) = 0
   For iCol = 1 To flxACHistorySplit.Cols - 2
      flxACHistorySplit.ColWidth(iCol) = Label1(30 + iCol + 1).Left - Label1(30 + iCol).Left
   Next iCol

   flxACHistorySplit.ColWidth(iCol) = flxACHistorySplit.Width + flxACHistorySplit.Left - Label1(42).Left - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   strSessionPropertyID = ""
   strSessionClientID = ""
   LOAD_TENANT_TENANTID = ""
   'frmMMain.fraCmdButton.Enabled = True
   bolgridTenantLookupRefresh = False
   UnLoadForm Me
   Unload Me
End Sub

Private Sub ConfigGridTenantLookup()
   Dim szHeader   As String
   Dim i          As Integer

   gridTenantLookup.Clear
   gridTenantLookup.Rows = 2
   'gridTenantLookup.Cols = 8
   gridTenantLookup.Cols = 9
   
   gridTenantLookup.Visible = True
   gridTenantLookup.RowHeight(0) = 0
   gridTenantLookup.row = 0

   For i = 0 To gridTenantLookup.Cols - 1
      gridTenantLookup.col = i
      gridTenantLookup.CellFontBold = True
   Next i

   szHeader$ = "<Sage A/C|<Name|<Address|>Balance||Client|Property|Unit"
   gridTenantLookup.FormatString = szHeader$

   gridTenantLookup.ColWidth(0) = lblTenantSort(1).Left - lblTenantSort(0).Left
   gridTenantLookup.ColWidth(1) = lblTenantSort(2).Left - lblTenantSort(1).Left
   gridTenantLookup.ColWidth(2) = lblTenantSort(3).Left - lblTenantSort(2).Left
   gridTenantLookup.ColWidth(3) = lblTenantSort(4).Left - lblTenantSort(3).Left
   gridTenantLookup.ColWidth(4) = 0
   gridTenantLookup.ColWidth(5) = 0
   gridTenantLookup.ColWidth(6) = 0
   gridTenantLookup.ColWidth(7) = 1100
   gridTenantLookup.ColWidth(8) = 0
End Sub

'Private Sub gridBank_Click()
'   populateControl Me, gridBank
'End Sub

Private Sub LoadTenantByTenantID()
   Dim sSQLQuery_ As String, szHeader As String

   SEARCHTenantMODE_ = False
   fmeTenantLookup.Visible = False
   lblLoading.Caption = "Please wait while loading..."
   fmeLoading.Visible = True
   fmeLoading.Refresh

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   'Populate the Tenant Header
   PopulateTenantInformation adoConn, LOAD_TENANT_TENANTID
   lblLeaseChanged.Caption = LOAD_TENANT_TENANTID

   ' Populate Bank Details
''   ConfigureFlxBank
''   szHeader$ = "<BankTenantID|<BankID|<BankACName|<BankSortCode|<BankACNumber|<PaymentMethod|<BacsRef|<IsDefaultAC"
''   sSQLQuery_ = "SELECT BankTenantID, BankID, BankACName, BankSortCode, BankACNumber, " & _
''                           "PaymentMethod, BacsRef, IsDefaultAC " & _
''                         "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'"
''   populateGridSimply adoConn, sSQLQuery_, gridBank, szHeader

'   ' Populate Event History
'   ConfigurFlxEventHistory
'   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
'   sSQLQuery_ = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
'                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
'   populateGridSimply adoConn, sSQLQuery_, gridEventHistory, szHeader
   LoadGridMaintenanceHistory adoConn
   Call LoadFlxACHistory(adoConn, "")

   '' LOAD Tenant DETAIL INFORMATION
   RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo

   fmeLoading.Visible = False
   adoConn.Close
   Set adoConn = Nothing

   ' SET OTHERS
   SEARCHTenantMODE_ = True
   TenantTabEnabled True
End Sub

Private Sub gridMaintenanceHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridMaintenanceHistory.ToolTipText = gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.MouseRow, gridMaintenanceHistory.MouseCol)
End Sub

Private Sub gridMaintenanceHistory_RowColChange()
   populateControl Me, gridMaintenanceHistory
End Sub

Private Sub gridMaintenanceHistory_Click()
   If (gridMaintenanceHistory.row > 0 And gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) <> "") Then
      cmdEditMHistory.Enabled = True
   Else
      cmdEditMHistory.Enabled = False
   End If
End Sub

Private Sub gridTenantLookup_Click()
   Dim szSQL As String, szHeader As String
'   Debug.Print time
   If gridTenantLookup.row = 0 Then
        fmeTenantLookup.Visible = False
        SEARCHTenantMODE_ = False
        
        fmeTenant.Enabled = True
        tabTenant.Enabled = True
        txtNominalCode.text = ""
        txtNominalCodeName.text = ""
        lblVatCode(0).Caption = ""
        txtCodeVat.text = ""
        txtSLControl.text = ""
        txtSLControlName.text = ""
        cmdTenantLookup.Visible = True
        FocusControl cmdTenantLookup
        'Need to clear all controls bcoz it is holding all the values of previous selection
        Exit Sub
   End If
   SEARCHTenantMODE_ = False
   fmeTenantLookup.Visible = False
   fmeTenant.Enabled = True
   tabTenant.Enabled = True
   
   'cmdTenantLookup.Enabled = False
   lblLoading.Caption = "Please wait while loading..."
   fmeLoading.Visible = True
   fmeLoading.Refresh

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   adoConn.Open getConnectionString

'   Populate the Tenant Header
   If gridTenantLookup.TextMatrix(gridTenantLookup.row, 4) = "CURRENT" Then  'And optCurrentTenant.Value rem by anol 20170124' here current means not deleted
'      txtProperty.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 5)
'      txtClient.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 6)
'      txtUnit.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 2)
      If gridTenantLookup.TextMatrix(gridTenantLookup.row, 8) <> "0" Then
         txtLeaseId.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 8)
      Else
         txtLeaseId.text = ""
      End If
      PopulateTenantInformation adoConn, gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      'PopulateLeaseInformation adoConn, gridTenantLookup.TextMatrix(gridTenantLookup.row, 8)
      lblLeaseChanged.Caption = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      txtProperty.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 5)
      txtClient.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 6)
      txtUnit.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 3)

      
   Else 'I think this else part is unreachable beacese we are not laoding deleted tenant
      txtTenantID.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      lblLeaseChanged.Caption = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      txtName.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 1)
      txtCompanyName.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 2)
      txtProperty.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 5)
      txtClient.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 6)
      txtUnit.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 3)
   End If

   szSQL = "SELECT N.Code, N.Name, S.CAName, S.Code AS NCode, V.VAT_CODE, V.VAT_RATE " & _
           "FROM   ((Tenants AS T LEFT JOIN NominalLedger AS N ON T.DefaultNC = N.Code) LEFT JOIN " & _
                  "( SELECT Code, CAName " & _
                   " FROM   NominalLedger AS N " & _
                   " WHERE  N.ClientID = '" & txtClientID.text & "') AS S ON T.SLControl = S.Code) " & _
                  "LEFT JOIN tlbVatCode AS V ON T.VAT_CODE = V.VAT_CODE " & _
           "WHERE  N.ClientID = '" & txtClientID.text & "' AND " & _
                  "T.SageAccountNumber = '" & txtTenantID.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      txtNominalCode.text = IIf(IsNull(adoRST.Fields.Item("Code").Value), "", adoRST.Fields.Item("Code").Value)
      txtNominalCodeName.text = IIf(IsNull(adoRST.Fields.Item("Name").Value), "", adoRST.Fields.Item("Name").Value)
      lblVatCode(0).Caption = IIf(IsNull(adoRST.Fields.Item("VAT_CODE").Value), "", adoRST.Fields.Item("VAT_CODE").Value)
      txtCodeVat.text = IIf(IsNull(adoRST.Fields.Item("VAT_RATE").Value), "", adoRST.Fields.Item("VAT_RATE").Value)
      txtSLControl.text = IIf(IsNull(adoRST.Fields.Item("NCode").Value), "", adoRST.Fields.Item("NCode").Value)
      txtSLControlName.text = IIf(IsNull(adoRST.Fields.Item("CAName").Value), "", adoRST.Fields.Item("CAName").Value)
   Else
      txtNominalCode.text = ""
      txtNominalCodeName.text = ""
      lblVatCode(0).Caption = ""
      txtCodeVat.text = ""
      txtSLControl.text = ""
      txtSLControlName.text = ""
   End If

   adoRST.Close

'    Populate Bank Details
''   ConfigureFlxBank
''   szHeader$ = "<BankTenantID|<BankID|<BankACName|<BankSortCode|<BankACNumber|<PaymentMethod|<BacsRef|<IsDefaultAC"
''   szSQL = "SELECT BankTenantID, BankID, BankACName, BankSortCode, BankACNumber, " & _
''              "PaymentMethod, BacsRef, IsDefaultAC " & _
''           "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'"
''   populateGridSimply adoConn, szSQL, gridBank, szHeader

'  Populate Lessee Account history
   
   Call LoadFlxACHistory(adoConn, "")
  
   LoadFlxLetter adoConn
' Debug.Print time
'   txtBalance.text = Format(AccountBalance(adoconn), "0.00")
    txtBalance1.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 7)

''    Populate Event History
'   ConfigurFlxEventHistory
'   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
'   szSQL = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
'                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
'   populateGridSimply adoConn, szSQL, gridEventHistory, szHeader
   LoadGridMaintenanceHistory adoConn

'    LOAD Tenant DETAIL INFORMATION
   RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo

   szSQL = "SELECT TenantDeposit.TenantID, (SUM(DptAmount) - " & _
                 "iif(isnull(DepositRefund.TotalDepositRefund),0,DepositRefund.TotalDepositRefund)) as TotalDepositHeld " & _
           "FROM TenantDeposit LEFT JOIN [SELECT SUM(DptAmount) AS TotalDepositRefund, TenantID " & _
              "FROM TenantDeposit " & _
              "WHERE DptRefund = True and Status = true and Deleted = false " & _
              "GROUP BY tenantid]. AS DepositRefund ON  TenantDeposit.TenantID = DepositRefund.TenantID " & _
           "WHERE TenantDeposit.TenantID = '" & txtTenantID.text & "' AND " & _
                 "TenantDeposit.Deleted = FALSE AND DptRefund = FALSE " & _
           "GROUP BY TenantDeposit.TenantID, DepositRefund.TotalDepositRefund;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      txtDeposit.text = "0.00"
   Else
      txtDeposit.text = Format(IIf(IsNull(adoRST.Fields.Item("TotalDepositHeld").Value), 0, adoRST.Fields.Item("TotalDepositHeld").Value), "0.00")
   End If

   adoRST.Close
   Set adoRST = Nothing

   fmeLoading.Visible = False

'    SET OTHERS
   SEARCHTenantMODE_ = True
   TenantTabEnabled True
   tabTenant.Tab = 0

   DeletedMode gridTenantLookup.TextMatrix(gridTenantLookup.row, gridTenantLookup.Cols - 1)

   populateGroupCombo adoConn

   adoConn.Close
   Set adoConn = Nothing

   LoadFlxEmails
    'issue 488
   'Added by Anol 03 Nov 2014
   PopulateGridLeaseAnalysis
   'fixed by anol 09 sep 2015
   txtMemoAll.text = ""
   Call ViewMemo
'    Debug.Print time
   cmdTenantLookup.Visible = True
   FocusControl cmdTenantLookup
   
End Sub
Public Sub ConfiggridMaintenanceHistory1()
'added by anol 20161120
   Dim iColumn    As Integer
   Dim szHeader   As String

   szHeader$ = "T|<Value|<ReportedDate|<Ref|<Job_DiaryName|<TaskOwner|" & _
               "<ReportedBy|<AssignedTo|<RemindDate|''|<DateCompleted|" & _
               ">BudgetCost|<ExpectedStartDate|<ExpectedCompletionDate|" & _
               "<Detail|>ActualCost|<ReportedBy|<AssignedIL|<ReportedIS|" & _
               "<RemindTime|<Urgent|<MaintenanceType|<ReportedFrom|" & _
               "<FundID|<OverrideBudget|<FYrID|<BudgetPassed|" & _
               "<PropertyID|<ClientID|<EmailAdd|<PropertyID|<SupplierID|<SupplierName|ClientName"
'  0|1|2|3|4|5
'  6|7|8|9
'  10|11|12
'  13|14|15|16|17
'  18|19|20|21
'  22|23|24|25
'  26|27|28|29

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = 35
   gridMaintenanceHistory.RowHeight(0) = 0
   gridMaintenanceHistory.FormatString = szHeader$

   For iColumn = 1 To 11
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn - 1).Left - 70

   For iColumn = 12 To gridMaintenanceHistory.Cols
      gridMaintenanceHistory.ColWidth(iColumn) = 0
   Next iColumn
End Sub
Public Sub LoadGridMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection)
   Dim rstMHistory_ As New ADODB.Recordset
   Dim szSQL As String
'
'   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
'                "H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
'                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
'                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
'                "H.Detail, H.ActualCost, H.ReportedBy, " & _
'                "H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent, " & _
'                "H.MaintenanceType, H.ReportedFrom " & _
'           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
'           "WHERE H.ReportedBy = '" & txtTenantID.text & "' AND " & _
'               "S.Code = H.MaintenanceType AND " & _
'               "S.PrimaryCode = 'MTYP' " & _
'           "ORDER BY H.ReportedDate DESC;"
'Debug.Print szSQL
'modified by anol 20161120, view job is not working
   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID, '', P.PropertyName , '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S , Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE H.ReportedBy = '" & txtTenantID.text & "' AND  U.UnitNumber = L.UnitNumber AND U.PropertyID= H.PropertyID AND  H.PropertyID = P.PropertyID AND H.ReportedBy=L.SageAccountNumber AND " & _
               "S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' " & _
           "ORDER BY H.ReportedDate DESC;"
           
           
   rstMHistory_.Open szSQL, conMHistory_, adOpenStatic, adLockReadOnly

   ConfigGridMaintenanceHistory rstMHistory_

   If rstMHistory_.EOF Then
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
      Exit Sub
   Else
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
   End If
'   ConfiggridMaintenanceHistory1
   populateGridDefinedHeader2 conMHistory_, szSQL, gridMaintenanceHistory

   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.col = 0
End Sub
Private Function populateGridDefinedHeader2(ByVal adoConn As ADODB.Connection, ByVal sSQLQuery As String, ByVal gridMain As MSHFlexGrid, Optional RowHeight As Integer = 240) As Integer
   Dim adoRST As New ADODB.Recordset

   adoRST.Open sSQLQuery, adoConn, adOpenStatic, adLockOptimistic

   populateGridDefinedHeader2 = adoRST.RecordCount

   gridMain.Rows = 2
   gridMain.RowHeight(1) = RowHeight
   If adoRST.EOF Then
       adoRST.Close
       Set adoRST = Nothing
       Exit Function
   End If

   Dim i As Integer, j As Integer

'   gridMain.AddItem ""

   For i = 0 To adoRST.RecordCount - 1
      For j = 0 To adoRST.Fields.Count - 1
         gridMain.TextMatrix(i + 1, j) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
      Next j
      'issue 572 2018/05/09 by anol list taking long time to load
      'gridMain.RowHeight(i + 1) = RowHeight
      adoRST.MoveNext
      If Not adoRST.EOF Then gridMain.AddItem ""
   Next i
   gridMain.ColWidth(6) = 0
   adoRST.Close
   Set adoRST = Nothing

   Exit Function

Error_Handler:
   MsgBox "An Error occurred while populating the grid"
End Function
Private Function AccountBalance(ByVal adoConn As ADODB.Connection) As Currency
   Dim iRow As Integer, cDr As Currency, cCr As Currency

   For iRow = 1 To flxACHistory.Rows - 1
'If Val(flxACHistory.TextMatrix(iRow, 7)) <> 0 Then
'MsgBox ""
'End If
      cDr = cDr + Round(CCur(IIf(flxACHistory.TextMatrix(iRow, 7) = "", 0, flxACHistory.TextMatrix(iRow, 7))), 2)
      cCr = cCr + Round(CCur(IIf(flxACHistory.TextMatrix(iRow, 8) = "", 0, flxACHistory.TextMatrix(iRow, 8))), 2)
'Debug.Print flxACHistory.TextMatrix(iRow, 7) + "   -" + flxACHistory.TextMatrix(iRow, 8)
   Next iRow

   AccountBalance = cDr - cCr
   Exit Function

'  Calculating Lessee's balance from the history grid. thats why this function does not execute the following code.

'  NOT IN USE THE FOLLWING CODE -------->>>>>>>>>>

'--------------------------------------------------------------------------------------------------------------------
   Dim szSQL As String, cAllOSInv As Currency, cAllPoA As Currency, cAllCredit As Currency
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT SageAccountNumber, SUM(OSAmount) AS OSI " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "ReceiptView = TRUE AND TYPE <> 2 " & _
           "GROUP BY SageAccountNumber;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllOSInv = CCur(IIf(IsNull(adoRST.Fields.Item("OSI").Value), 0, adoRST.Fields.Item("OSI").Value))
   adoRST.Close

   szSQL = "SELECT SageAccountNumber, SUM(OSAllocation) AS OSA " & _
           "FROM tblPoA " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "PoAView = TRUE " & _
           "GROUP BY SageAccountNumber;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllPoA = CCur(IIf(IsNull(adoRST.Fields.Item("OSA").Value), 0, adoRST.Fields.Item("OSA").Value))
   adoRST.Close

   szSQL = "SELECT SageAccountNumber, SUM(OSAmount) AS OSI " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "ReceiptView = TRUE AND TYPE = 2 " & _
           "GROUP BY SageAccountNumber;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllCredit = CCur(IIf(IsNull(adoRST.Fields.Item("OSI").Value), 0, adoRST.Fields.Item("OSI").Value))
   adoRST.Close
   Set adoRST = Nothing

   cAllCredit = cAllCredit + cAllPoA

   AccountBalance = cAllOSInv - cAllCredit
End Function

Private Sub DeletedMode(szMode As String)
   If szMode = "DELETED" Then
      cmdDeleteLessee.Enabled = False
      cmdEdit.Enabled = False
      cmdCopy.Enabled = False
      cmdEditTenantAddress.Enabled = False
'      cmdNewBank.Enabled = False
'      cmdEditBank.Enabled = False
      cmdNewEvent.Enabled = False
      cmdEditEvent.Enabled = False
   End If

   If szMode = "CURRENT" Then TenantTabEnabled True

End Sub

Private Sub gridTenantLookup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gridTenantLookup_Click
    End If
End Sub
Private Function BolCompareDataGT(szData1 As String, szData2 As String, dtDataType As String) As Boolean
   BolCompareDataGT = False

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
Private Sub SortingGrid2(flxGrid As MSHFlexGrid, iSortCol As Integer, bAscDsc As Boolean, Optional dtDataType As String)
'written by anol 20160524 Email was not sorted int he grid
   Dim i As Integer, j As Integer, c As Integer
   Dim szTemp() As String
   ReDim szTemp(flxGrid.Cols - 1) As String
   
   For i = 1 To flxGrid.Rows - 2
      If flxGrid.RowHeight(i) > 0 Then
         For j = i + 1 To flxGrid.Rows - 1
            If flxGrid.RowHeight(j) > 0 Then
            
               'If Not bAscDsc Then                 'Sorting ascending order
'                  If flxGrid.TextMatrix(i, iSortCol) > flxGrid.TextMatrix(j, iSortCol) Then
                  If BolCompareDataGT(flxGrid.TextMatrix(i, iSortCol), flxGrid.TextMatrix(j, iSortCol), dtDataType) Then
                     For c = 0 To flxGrid.Cols - 1
                        szTemp(c) = flxGrid.TextMatrix(i, c)
                        flxGrid.TextMatrix(i, c) = flxGrid.TextMatrix(j, c)
                        flxGrid.TextMatrix(j, c) = szTemp(c)
                     Next c
                  End If
               'End If

'               If bAscDsc Then                 'Sorting decending order
''                  If flxGrid.TextMatrix(i, iSortCol) < flxGrid.TextMatrix(j, iSortCol) Then
'                  If BolCompareDataST(flxGrid.TextMatrix(i, iSortCol), flxGrid.TextMatrix(j, iSortCol), dtDataType) Then
'                     For c = 0 To flxGrid.Cols - 1
'                        szTemp(c) = flxGrid.TextMatrix(i, c)
'                        flxGrid.TextMatrix(i, c) = flxGrid.TextMatrix(j, c)
'                        flxGrid.TextMatrix(j, c) = szTemp(c)
'                     Next c
'                  End If
'               End If
            
            End If
         Next j
      End If
   Next i
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lblTenantSort_Click(Index As Integer)
   If Index = 0 Then                               ' Sort Tenant ID
      SortingGrid gridTenantLookup, Index, bSortingCol1
      bSortingCol1 = IIf(bSortingCol1, False, True)
      lblTenantSort(0).FontBold = True
      lblTenantSort(1).FontBold = False
      lblTenantSort(2).FontBold = False
   End If

   If Index = 1 Then                               ' Sort Tenant Name
      SortingGrid gridTenantLookup, Index, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      lblTenantSort(0).FontBold = False
      lblTenantSort(1).FontBold = True
      lblTenantSort(2).FontBold = False
   End If

   If Index = 2 Then                               ' Sort Unit Name
      SortingGrid gridTenantLookup, Index, bSortingCol3
      bSortingCol3 = IIf(bSortingCol3, False, True)
      lblTenantSort(0).FontBold = False
      lblTenantSort(1).FontBold = False
      lblTenantSort(2).FontBold = True
   End If
End Sub

Private Sub optAll_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget / Location"
   Label61(3).Caption = "Ref"
   Label61(1).Caption = "Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
End Sub

Private Sub optDiary_Click()
   Dim i As Integer

   Label61(10).Caption = "Location"
   Label61(3).Caption = "Diary No"
   Label61(1).Caption = "Event Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) = "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optJobs_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget"
   Label61(3).Caption = "Job No"
   Label61(1).Caption = "Maintenance Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) <> "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optExitingGroup_Click()
   cboGroup.Enabled = True
   cboGroup.Locked = False
End Sub

Private Sub optNewGroup_Click()
   cboGroup.Enabled = False
   cboGroup.Locked = True
End Sub

Private Sub tabTenant_Click(PreviousTab As Integer)
   Dim adoConn As New ADODB.Connection
'   If tabTenant.Tab = 2 Then
'        Picture3.Visible = True
'   Else
'        Picture3.Visible = False
'        txtAmountOutStanding(0).text = "0.00"
'   End If
   adoConn.Open getConnectionString

   Select Case tabTenant.Tab
   Case 2:
      If Not bDEPOSIT_HELD Then
         ConfigFlxDeposit
         LoadComboes adoConn

         bDEPOSIT_HELD = True
      End If

      If txtTenantID.text <> "" Then LoadFlxDeposit adoConn, 0, "ASC"      'Filling the deposit grid

      ButtonHanlding DefaultMode
      FocusControl cmdBank
   Case 4:
      If txtTenantID.text <> "" Then _
         Call LoadAttachmentFiles(cmbFiles, txtTenantID.text, "Tenants")

   End Select

   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Sub SquezeExpand()
    On Error GoTo Err
       Dim i As Integer, iCurRowHeight As Integer

  iCurRowHeight = 280
   

   If flxDeposit.TextMatrix(flxDeposit.row, 0) = "+" Then           'Expanding the grid 'FlxDeposit.col = 0 And
      flxDeposit.TextMatrix(flxDeposit.row, 0) = ">"
      iCurRowHeight = flxDeposit.RowHeight(flxDeposit.row)
      i = 1

      While flxDeposit.TextMatrix(flxDeposit.row + i, 0) = "-"
         flxDeposit.RowHeight(flxDeposit.row + i) = iCurRowHeight
         i = i + 1
         If (flxDeposit.row + i) = flxDeposit.Rows Then Exit Sub
      Wend
      Exit Sub
   End If

   If flxDeposit.TextMatrix(flxDeposit.row, 0) = ">" Then          'Squeezing the grid 'FlxDeposit.col = 1 And
      flxDeposit.TextMatrix(flxDeposit.row, 0) = "+"
      i = 1
      While flxDeposit.TextMatrix(flxDeposit.row + i, 0) = "-"
         flxDeposit.RowHeight(flxDeposit.row + i) = 0
         i = i + 1
         If (flxDeposit.row + i) = flxDeposit.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   Exit Sub
Err:
   'HighLightRowFlxGridA FlxDeposit, FlxDeposit.row
End Sub
Private Sub LoadFlxDeposit(ByVal adoConn As ADODB.Connection, j As Integer, strSort As String)
    Dim adoRST As New ADODB.Recordset
    Dim szSQL As String, iRow As Integer, i As Integer
    Dim szOrderby As String
    If j = 0 Then
        szOrderby = "ORDER BY DepositTypePrefix,DepositSL " & strSort
    End If
    If j = 1 Then
        szOrderby = "ORDER BY TransactionID " & strSort
    End If
    If j = 2 Then
        szOrderby = "ORDER BY DptType " & strSort
    End If
    If j = 4 Then
        szOrderby = "ORDER BY DepositDate " & strSort
    End If
    
    
   szSQL = "SELECT D.*, B.ReconNow, (Select fund.FundCODE from Fund where Fund.FundID=D.FundID) as FundCODE  " & _
           "FROM TenantDeposit AS D LEFT JOIN tlbBankPayment AS B ON D.DepositID = B.TenantDeposit " & _
           "Where " & _
               "D.TenantID = '" & txtTenantID.text & "' AND " & _
               "D.Deleted = False " & szOrderby

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With adoRST
      iRow = 1
      flxDeposit.Clear
      Call ConfigFlxDeposit
      flxDeposit.Rows = 1
      While Not .EOF
         If Not .EOF Then flxDeposit.AddItem ""
         flxDeposit.TextMatrix(iRow, 1) = .Fields.Item("DepositID").Value
         flxDeposit.TextMatrix(iRow, 2) = .Fields.Item("DepositTypePrefix").Value & .Fields.Item("DepositSL").Value
         flxDeposit.ColAlignment(7) = vbLeftJustify
         flxDeposit.TextMatrix(iRow, 7) = .Fields.Item("BANKCODE").Value ' .Fields.Item("BCName").Value
         flxDeposit.TextMatrix(iRow, 4) = IIf(IsNull(.Fields.Item("TransactionID").Value), "", .Fields.Item("TransactionID").Value)
         flxDeposit.ColAlignment(6) = vbLeftJustify
         flxDeposit.TextMatrix(iRow, 6) = Format(.Fields.Item("DepositDate").Value, "dd/mm/yyyy")
         If Left(.Fields.Item("TransactionID").Value, 1) = "D" Then _
            flxDeposit.TextMatrix(iRow, 3) = "Deposit"
         If Left(.Fields.Item("TransactionID").Value, 1) = "R" Then _
            flxDeposit.TextMatrix(iRow, 3) = "Refund"
         If Left(.Fields.Item("TransactionID").Value, 1) = "E" Then _
            flxDeposit.TextMatrix(iRow, 3) = "Expense"

         flxDeposit.TextMatrix(iRow, 8) = .Fields.Item("DptDetails").Value '
         If flxDeposit.TextMatrix(iRow, 8) <> "" Then flxDeposit.ColAlignment(8) = vbLeftJustify

         If Not .Fields.Item("DptRefund").Value Then
            flxDeposit.TextMatrix(iRow, 5) = Value_SecondaryCode("DPTYP", .Fields.Item("DptType").Value, adoConn)   'Deposit Type code name
            flxDeposit.TextMatrix(iRow, 9) = Format(Val(.Fields.Item("DptAmount").Value), "0.00")                   'amount
            flxDeposit.TextMatrix(iRow, 10) = Format(Val(.Fields.Item("OSRefund").Value), "0.00") 'COL_OUT_STANDING_REFUND=9
            flxDeposit.TextMatrix(iRow, 11) = ""
         Else
            If .Fields.Item("TransactionID").Value = "RF" Then
               flxDeposit.TextMatrix(iRow, 5) = Value_SecondaryCode("RTYP", .Fields.Item("DptType").Value, adoConn)   'Full refund
            Else
               flxDeposit.TextMatrix(iRow, 5) = Value_SecondaryCode("EXPTYP", .Fields.Item("DptType").Value, adoConn)     'Deposit Type code name
            End If
            flxDeposit.TextMatrix(iRow, 11) = Format(Val(.Fields.Item("DptAmount").Value), "0.00")
            If flxDeposit.TextMatrix(iRow, 10) = "" Then
                    flxDeposit.TextMatrix(iRow, 10) = Format(Val(.Fields.Item("OSRefund").Value), "0.00")
             End If
         End If

         flxDeposit.TextMatrix(iRow, 12) = IIf(IsNull(.Fields.Item("GroupNo").Value), "", .Fields.Item("GroupNo").Value)
         flxDeposit.TextMatrix(iRow, 13) = .Fields.Item("BankCode").Value
         flxDeposit.TextMatrix(iRow, 14) = IIf(IsNull(.Fields.Item("NominalCode")), "", .Fields.Item("NominalCode").Value)
         flxDeposit.TextMatrix(iRow, 17) = IIf(IsNull(.Fields.Item("NCName")), "", .Fields.Item("NCName").Value)
         flxDeposit.TextMatrix(iRow, 15) = .Fields.Item("DptAmtType").Value
         flxDeposit.TextMatrix(iRow, 16) = .Fields.Item("DptType").Value ''
         flxDeposit.TextMatrix(iRow, 18) = .Fields.Item("FundID").Value
         flxDeposit.TextMatrix(iRow, 19) = IIf(IsNull(.Fields.Item("ReconNow").Value), "", .Fields.Item("ReconNow").Value)
         flxDeposit.TextMatrix(iRow, 20) = .Fields.Item("BCName").Value
         flxDeposit.TextMatrix(iRow, 21) = .Fields.Item("FundCODE").Value
         flxDeposit.TextMatrix(iRow, 22) = .Fields.Item("DptType").Value
         iRow = iRow + 1
         .MoveNext
      Wend
      .Close
   End With
   Set adoRST = Nothing
   flxDeposit.row = 0
End Sub

Private Sub LoadComboes(ByVal adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim Data() As String, i As Integer

'   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
'               "NominalLedger.Name AS BNN " & _
'           "FROM tlbClientBanks, NominalLedger " & _
'           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code " & _
'           "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Please setup bank account for the client."
'   Else
'      ReDim Data(1, adoRst.RecordCount - 1) As String
'      i = 0
'      While Not adoRst.EOF
'         Data(0, i) = adoRst.Fields.Item("BNC").Value
'         Data(1, i) = adoRst.Fields.Item("BNN").Value
'         i = i + 1
'         adoRst.MoveNext
'      Wend
'      cmbBank.Clear
'      cmbBank.Column() = Data()
'   End If
'
'   adoRst.Close

'   szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
'             "FROM PrimaryCode, SecondaryCode " & _
'             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
'                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
'             "ORDER BY SecondaryCode.Value;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      ReDim Data(1, adoRst.RecordCount - 1) As String
'
'      i = 0
'      While Not adoRst.EOF
'         Data(0, i) = adoRst!SC
'         Data(1, i) = adoRst!V
'         adoRst.MoveNext
'         i = i + 1
'      Wend
'
'      cmbDptAmtType.Clear
'      cmbDptAmtType.Column() = Data()
'   End If

'   adoRst.Close
'////////////////////////////////// DEPOSIT TYPE //////////////////////////////////////////////////////
'   szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
'             "FROM PrimaryCode, SecondaryCode " & _
'             "WHERE (PrimaryCode.Value = 'DEPOSIT TYPE' OR " & _
'                  "PrimaryCode.Value = 'REFUND TYPE' OR " & _
'                  "PrimaryCode.Value = 'EXPENSES TYPE') AND " & _
'                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
'             "ORDER BY SecondaryCode.Value;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      ReDim Data(1, adoRst.RecordCount - 1) As String
'
'      i = 0
'      While Not adoRst.EOF
'         Data(0, i) = adoRst!SC
'         Data(1, i) = adoRst!V
'         adoRst.MoveNext
'         i = i + 1
'      Wend
'
'      cboDepositType.Clear
'      cboDepositType.Column() = Data()
'   End If

'   adoRst.Close
'////////////////////////////////// DEPOSIT TYPE //////////////////////////////////////////////////////
'   szSQL = "SELECT FundID, FundName FROM FUND;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim szaData(1, adoRst.RecordCount) As String
'
'   i = 0
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst.Fields.Item("FundID").Value
'      szaData(1, i) = adoRst.Fields.Item("FundName").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cboFund.Clear
'   cboFund.Column() = szaData()
'
'   adoRst.Close
'   Set adoRst = Nothing
End Sub

Private Sub ConfigFlxACHistory()
   Dim szHeader As String, iCol As Integer

   flxACHistory.Clear
   flxACHistory.Cols = 11
   flxACHistory.Rows = 2
   flxACHistory.RowHeight(0) = 0
'   szHeader$ = "|<REF|<TYPE|<UnitNumber|<IssueDate|>AMOUNT|>OS|>DEBIT|>CREDIT"
'   flxACHistory.FormatString = szHeader$

   flxACHistory.ColWidth(0) = Label1(1).Left - flxACHistory.Left
   For iCol = 1 To flxACHistory.Cols - 4
      flxACHistory.ColWidth(iCol) = Label1(iCol + 1).Left - Label1(iCol).Left
   Next iCol

   flxACHistory.ColAlignment(4) = vbLeftJustify
   flxACHistory.ColWidth(8) = flxACHistory.Width + flxACHistory.Left - Label1(8).Left - 360
   flxACHistory.ColWidth(9) = 0
   flxACHistory.ColWidth(10) = 0
End Sub

'  Build up lessee's Account History
Private Sub LoadFlxACHistory_old(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoRpt As New ADODB.Recordset, adoRptDtl As New ADODB.Recordset
   Dim sqlcount As Integer

   ConfigFlxACHistory
   ConfigFlxACHistorySplit

   szSQL = "SELECT Rpt.*, TT.DESCRIPTION AS TT_DES, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
           "FROM tlbReceipt AS Rpt, tlbTransactionTypes AS TT " & _
           "WHERE Rpt.SageAccountNumber = '" & txtTenantID.text & "' And " & _
               "Rpt.Type = TT.TYPE_ID " & _
           "ORDER BY Rpt.RDate;"
           sqlcount = 1
'Debug.Print szSQL
   adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoRpt.EOF                                           '//1
         If adoRpt!Type = 1 Or adoRpt!Type = 23 Then                                                             '//2
            szSQL = "SELECT MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, SQ.* " & _
                    "FROM (tlbReceipt AS R INNER JOIN " & _
                    "(" & _
                        "SELECT RT.SlNumber, RT.FromTran, RT.ReceiptAmount, R.DemandRef " & _
                        "FROM (RptTransactions AS RT INNER JOIN " & _
                              "tlbReceipt AS R ON RT.ToTran = R.TransactionID) INNER JOIN " & _
                              "tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                        "Where RT.ToTran = " & adoRpt.Fields.Item("TransactionID").Value & " " & _
                    ") AS SQ ON R.TransactionID = SQ.FromTran) " & _
                        "INNER JOIN tlbTransactionTypes AS T ON R.Type = T.TYPE_ID;"
                         sqlcount = sqlcount + 1
         Else
            szSQL = "SELECT RT.*, R.SlNumber AS RefID, " & _
                        "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
                    "FROM (RptTransactions AS RT INNER JOIN tlbReceipt AS R ON RT.ToTran = R.TransactionID) " & _
                        "INNER JOIN tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                    "WHERE RT.FromTran = " & adoRpt.Fields.Item("TransactionID").Value & ";"
                      sqlcount = sqlcount + 1
         End If
'Debug.Print szSQL
         adoRptDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         iChild = 0
         If adoRptDtl.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(iKount, 0) = "+"
            iChild = iKount + 1
            While Not adoRptDtl.EOF
               .TextMatrix(iChild, 0) = "-"
               If adoRpt!Type = 1 Or adoRpt!Type = 23 Then
                  .TextMatrix(iChild, 4) = "Receipt from " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 4) = "Receipt to " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("RefID").Value
               End If
               .TextMatrix(iChild, 5) = Format(adoRptDtl.Fields.Item("ReceiptAmount").Value, "0.00")
               .RowHeight(iChild) = 0
               iChild = iChild + 1
               adoRptDtl.MoveNext
               If Not adoRptDtl.EOF Then .AddItem ""
            Wend
         Else
            .TextMatrix(iKount, 0) = ""
         End If
'1:DemandRef, 2:Invoice, 3:Date, 4:Details, 5:Amount, 6:Amount (OS), 7:Amount (Dr)
         adoRptDtl.Close
'*************
         .TextMatrix(iKount, 1) = adoRpt.Fields.Item("PF").Value & adoRpt.Fields.Item("SlNumber").Value
         .TextMatrix(iKount, 2) = IIf(UCase(Left(adoRpt.Fields.Item("TT_DES").Value, 5)) = "SALES", Mid(adoRpt.Fields.Item("TT_DES").Value, 7), adoRpt.Fields.Item("TT_DES").Value)
         .TextMatrix(iKount, 3) = IIf(IsNull(adoRpt.Fields.Item("RDate").Value), "", _
                                             adoRpt.Fields.Item("RDate").Value)
'Resolved by BOSL
'Modified by anol 20 Apr 2015
'Issue 0000530: Batch receipts not working correctly
'Note 1014When the user processes a multiple batch receipt, the reference shown should be the reference entered by the user in batch receipts with multiple.
'This should be displayed in
'4/ Lessee account history
'I have changed Extref to Ref for tlbReceipt
'Reveresed back on 29 Apr 2015
         If adoRpt.Fields.Item("Type").Value = 3 Or _
               adoRpt.Fields.Item("Type").Value = 4 Or _
               adoRpt.Fields.Item("Type").Value = 23 Then
            .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item("Extref").Value), "", _
                                                adoRpt.Fields.Item("Extref").Value)
         Else
            .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item("Details").Value), "", _
                                                adoRpt.Fields.Item("Details").Value)
         End If

         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item("OSAmount").Value, "0.00")
         If adoRpt!Type = 1 Or adoRpt!Type = 23 Then
            .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 8) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Credit
         End If
         adoRpt.MoveNext
         iKount = IIf(iChild = 0, iKount + 1, iChild)
         If Not adoRpt.EOF Then .AddItem ""
      Wend

      adoRpt.Close
'############################## if there any unposted demand in the demand then its picking up here ##################
      szSQL = "SELECT D.DmdSlNo, D.IssueDate, D.Details, SUM(DS.TotalAmount),D.TransactionType " & _
              "FROM DemandRecords AS D, DemandSplitRecords AS DS " & _
              "WHERE D.DemandID = DS.DemandID AND DS.TrfReceipt = FALSE AND " & _
                  "D.SageAccountNumber = '" & txtTenantID.text & "' " & _
              "GROUP BY D.DmdSlNo, D.Details, D.IssueDate,TransactionType;"
              sqlcount = sqlcount + 1
      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = IIf((adoRpt.Fields("TransactionType").Value) = 1, "SI", "SC") & adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
      'issue 520 added by anol 20180209
      '############################## if there any  demand it hasbeen zero rised then its picking up here ##################
      szSQL = "SELECT D.DmdSlNo, D.IssueDate, D.Details, '0',D.TransactionType " & _
              "FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS DS " & _
              "ON D.DemandID = DS.DemandID where  " & _
                  "D.SageAccountNumber = '" & txtTenantID.text & "' AND isnull(DS.TotalAmount)" & _
              ""
    sqlcount = sqlcount + 1
      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = IIf((adoRpt.Fields("TransactionType").Value) = 1, "SI", "SC") & adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
   End With

   Set adoRpt = Nothing
   Set adoRptDtl = Nothing
   flxACHistory.row = 0
   flxACHistory.row = 0
   Debug.Print sqlcount
   'MsgBox flxACHistory.Rows
End Sub
Private Sub LoadFlxACHistory(adoConn As ADODB.Connection, Filter As String)
   On Error GoTo Err
   Dim tempstr As String
   Dim szSQL As String, iKount  As Integer
   Dim rsReportLAChistory As New ADODB.Recordset, adoRptDtl As New ADODB.Recordset
   Dim adoRpt As New ADODB.Recordset
   Dim sqlcount As Integer
   Dim strWhere As String

   ConfigFlxACHistory
   ConfigFlxACHistorySplit
   
   szSQL = "delete from  ReportLAChistory WHERE SessionID = '" & sessionID & "';"
   adoConn.Execute szSQL
   
   szSQL = "delete from  ReportLAChistory WHERE ReportingDate < #" & reportingDate & "# ;"
   adoConn.Execute szSQL
   If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            strWhere = " AND PF Like '%" & tempstr & "%'" 'PF=Inovice number/Receipt number
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            strWhere = " AND Extref Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "3" Or Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND RDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND RDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# "
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
   End If
    
    szSQL = "insert into ReportLAChistory(reportingDate,SessionID,SIGN,transactionID,Type,Type_desc,PF,slnumber,Rdate,Details ,extref ,amount,Osamount, Balance ,isMaster)  SELECT '" & reportingDate & _
            "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
            "'',TransactionID,Type,TT.Description,(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& slnumber) AS INVno , slnumber as No,Rcpt.RDate ,  Details ,extref,amount,osamount,0,0  FROM " & _
            "tlbReceipt AS Rcpt, tlbTransactionTypes AS TT " & _
            "WHERE Rcpt.SageAccountNumber = '" & txtTenantID.text & "' And Rcpt.Type = TT.TYPE_ID ORDER BY Rcpt.RDate "
    adoConn.Execute szSQL
   'Update allocation sign
    adoConn.Execute "Update ReportLAChistory A INNER JOIN RptTransactions R ON A.transactionID=R.FromTran SET SIGN='+' where Type in (2,3,4)"
    adoConn.Execute "Update ReportLAChistory A INNER JOIN RptTransactions R ON A.transactionID=R.ToTran SET SIGN='+' where Type in (1,23)"
    
    'allocation slave for invoices
    szSQL = "insert into ReportLAChistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,slnumber,RDate,amount,isMaster) " & _
            "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
            "'-', SQ.ToTran, R.Type, (Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & SQ.SlNumber) AS INVNO , SQ.SlNumber as NO,SQ.AllocDate,SQ.ReceiptAmount,'1'" & _
            "FROM (ReportLAChistory AS R INNER JOIN [SELECT RT.SlNumber, RT.Totran, RT.FromTran,RT.AllocDate, RT.ReceiptAmount, R.DemandRef " & _
                            "FROM (RptTransactions AS RT INNER JOIN " & _
                                  "tlbReceipt AS R ON RT.ToTran = R.TransactionID) INNER JOIN " & _
                                  "tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                        "]. AS SQ ON R.transactionID = SQ.FromTran) INNER JOIN tlbTransactionTypes AS T ON R.Type = T.TYPE_ID where R.Type In(2,3,4)  ;"
    adoConn.Execute szSQL
  'allocation slave for Receipts
    szSQL = "insert into ReportLAChistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,slnumber,RDate,amount,isMaster) " & _
            "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
            "'-', SQ.FromTran, SQ.Type, (SQ.PF & SQ.SlNumber) as INVNO, SQ.SlNumber as NO,RT.AllocDate,SQ.ReceiptAmount,'1'" & _
            "FROM (ReportLAChistory AS R INNER JOIN [SELECT R.SlNumber, RT.Totran,R.Type, RT.FromTran,Mid(TT.CONSTANT,4,Len(TT.CONSTANT)-3) AS PF,RT.AllocDate, RT.ReceiptAmount," & _
            "R.DemandRef " & _
                        "FROM (RptTransactions AS RT INNER JOIN " & _
                              "tlbReceipt AS R ON RT.ToTran = R.TransactionID) INNER JOIN " & _
                              "tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                    "]. AS SQ ON R.transactionID = SQ.FromTran) INNER JOIN tlbTransactionTypes AS T ON R.Type = T.TYPE_ID where R.Type In (2,3,4) ;"
                        
    adoConn.Execute szSQL
    'trick to select only OSamount:
    'as master record and its slave record contains the same transactionId column so  SQL shall update balance field with 1 , which is now working as a flag
    '2nd one set to 1 where osamount>0
    
    If chkShowOutstanding.Value = 0 Then
        'rsReportLAChistory.Open "Select * from ReportLAChistory where 1=1 " & strWhere & "  order by transactionID,ismaster", adoconn, adOpenStatic, adLockReadOnly
        adoConn.Execute "Update  ReportLAChistory A, (Select transactionID,sessionID from ReportLAChistory where  SessionID= '" & sessionID & "'" & _
                         strWhere & " order by transactionID,ismaster) As B Set balance=1 where A.transactionID=B.transactionID AND A.sessionID=B.sessionID"
    Else
        adoConn.Execute "Update  ReportLAChistory A, (Select transactionID,sessionID from ReportLAChistory where SessionID= '" & sessionID & "' AND osamount>0 " & _
                         strWhere & " order by transactionID,ismaster) As B Set balance=1 where A.transactionID=B.transactionID AND  A.sessionID=B.sessionID"
        
    End If
    rsReportLAChistory.Open "Select * from ReportLAChistory where SessionID= '" & sessionID & "' AND balance>0 order by transactionID,ismaster", adoConn, adOpenStatic, adLockReadOnly
    If rsReportLAChistory.RecordCount = 0 Then
        flxACHistory.Rows = 2
    Else
        flxACHistory.Rows = rsReportLAChistory.RecordCount + 1
    End If
    iKount = 1
    With flxACHistory
    While Not rsReportLAChistory.EOF
        .TextMatrix(iKount, 0) = rsReportLAChistory("SIGN").Value
        .TextMatrix(iKount, 1) = rsReportLAChistory("PF").Value
        .TextMatrix(iKount, 2) = IIf(IsNull(rsReportLAChistory("Type_desc").Value), "", rsReportLAChistory("Type_desc").Value)
        .TextMatrix(iKount, 2) = IIf(UCase(Left(.TextMatrix(iKount, 2), 5)) = "SALES", Mid(.TextMatrix(iKount, 2), 7), .TextMatrix(iKount, 2))
        .TextMatrix(iKount, 3) = IIf(IsNull(rsReportLAChistory("RDate").Value), "", rsReportLAChistory("RDate").Value)
        If rsReportLAChistory("SIGN").Value = "-" Then
             .RowHeight(iKount) = 0
            If rsReportLAChistory("type").Value = 1 Or rsReportLAChistory("type").Value = 23 Then
                .TextMatrix(iKount, 4) = "Receipt to " & rsReportLAChistory("PF").Value
            Else
                .TextMatrix(iKount, 4) = "Receipt From " & rsReportLAChistory("PF").Value
            End If
        Else
            .RowHeight(iKount) = 260
            If rsReportLAChistory("Type").Value = 3 Or rsReportLAChistory("Type").Value = 4 Or rsReportLAChistory("Type").Value = 23 Then
                    .TextMatrix(iKount, 4) = rsReportLAChistory("Extref").Value
            Else
                    .TextMatrix(iKount, 4) = rsReportLAChistory("Details").Value
            End If
        End If
        .TextMatrix(iKount, 5) = Format(rsReportLAChistory("Amount").Value, "0.00")
        .TextMatrix(iKount, 6) = IIf(rsReportLAChistory("OSAmount").Value = 0, "", Format(rsReportLAChistory("OSAmount").Value, "0.00")) 'balance
        If rsReportLAChistory("SIGN").Value <> "-" Then 'for the allocated amount you don't need debit or credit row
            If rsReportLAChistory("Type").Value = 1 Or rsReportLAChistory("Type").Value = 23 Then
                .TextMatrix(iKount, 7) = Format(rsReportLAChistory("Amount").Value, "0.00")
            Else
                .TextMatrix(iKount, 8) = Format(rsReportLAChistory("Amount").Value, "0.00")
            End If
        End If
        iKount = iKount + 1
        rsReportLAChistory.MoveNext
   Wend
    adoConn.Execute "Delete from ReportLAChistory where SessionID= '" & sessionID & "'"
 
'############################## if there any unposted demand in the demand then its picking up here ##################
      szSQL = "SELECT D.DmdSlNo, D.IssueDate, D.Details, SUM(DS.TotalAmount),D.TransactionType " & _
              "FROM DemandRecords AS D, DemandSplitRecords AS DS " & _
              "WHERE D.DemandID = DS.DemandID AND DS.TrfReceipt = FALSE AND " & _
                  "D.SageAccountNumber = '" & txtTenantID.text & "' " & _
              "GROUP BY D.DmdSlNo, D.Details, D.IssueDate,TransactionType;"
              sqlcount = sqlcount + 1
      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = IIf((adoRpt.Fields("TransactionType").Value) = 1, "SI", "SC") & adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
      'issue 520 added by anol 20180209
      '############################## if there any  demand it hasbeen zero rised then its picking up here ##################
      szSQL = "SELECT D.DmdSlNo, D.IssueDate, D.Details, '0',D.TransactionType " & _
              "FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS DS " & _
              "ON D.DemandID = DS.DemandID where  " & _
                  "D.SageAccountNumber = '" & txtTenantID.text & "' AND isnull(DS.TotalAmount)" & _
              ""
    sqlcount = sqlcount + 1
      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = IIf((adoRpt.Fields("TransactionType").Value) = 1, "SI", "SC") & adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
   End With

   Set adoRpt = Nothing
   Set adoRptDtl = Nothing
   flxACHistory.row = 0
   flxACHistory.row = 0
   Exit Sub
Err:
   MsgBox Err.description
'   Debug.Print sqlcount
   'MsgBox flxACHistory.Rows
End Sub

'  Build up lessee's Account Balance
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
'      'szaTenantBalance(1, iIndex) = RoundingNumber(adoRptDr.Fields.Item("Dr").Value, 2)
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
    Debug.Print time
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
'will it collect distict or not?ans : this is distinct because you are using group by
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

   szSQL = "SELECT SageAccountNumber, Type, Round(SUM(Amount),2) AS Amt " & _
           "FROM tlbReceipt " & _
           "GROUP BY SageAccountNumber,Type;"
           'added round function because for 'CRAU01A'  it was showing balance 919.99=>920 issue 859 fixed by anol 2020-07-08
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
   Debug.Print time
   'it took 9 second to build balance 2023-07-05
End Sub
Private Sub tabTenant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Me.MousePointer = vbArrow
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
       ' gridTenantLookup.SetFocus
        FocusControl gridTenantLookup
        If gridTenantLookup.Rows > 1 Then
            gridTenantLookup.row = 1
             gridTenantLookup.TopRow = 1
        End If
    End If
End Sub

Private Sub txtBank_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdBank.SetFocus
    End If
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

Private Sub txtClientList_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdClientList.SetFocus
    End If
End Sub

Private Sub txtCompanyName_GotFocus()
   If txtCompanyName.text = "" Then Exit Sub

   SelTxtInCtrl txtCompanyName
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
End Sub

Private Sub txtDate_GotFocus()
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        cmdDepositType.SetFocus
        FocusControl cmdBank
    End If
   TextBoxKeyPrsDate txtDate, KeyAscii

End Sub

Private Sub txtDate_LostFocus()
   TextBoxFormatDate txtDate
End Sub

Private Sub txtDNC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 And Index = 1 Then
        cmdNCList.SetFocus
    End If
End Sub

Private Sub txtDptAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdDptAmtType
    End If
   DigitTextKeyPress txtDptAmount, KeyAscii
End Sub

Private Sub txtDptAmount_LostFocus()
   If txtDptAmount.text = "" Then Exit Sub
'    If yDEPOSIT = 4 Then Exit Sub
   
   If yDEPOSIT = 1 Then
        If bEdit = True Then  'When you edit  deposit
               txtOSDpt.text = Format(txtDptAmount.text, "0.00")
        Else    'When you create new deposit
               txtOSDpt.text = Format(txtDptAmount.text, "0.00")
        End If
   End If
'  if user changes the deposit amount then system should change the O/S amount as the same time
   If yDEPOSIT = 2 And CCur(txtDptAmount.text) - cCurDepAmt <> 0 And txtOSDpt.text <> "" Then
      txtOSDpt.text = Format(CCur(txtOSDpt.text) + CCur(txtDptAmount.text) - cCurDepAmt, "0.00")
   End If

'  system should check if the refund amount is greater than the O/S amount then system will warn user
   If (yDEPOSIT = 3) And Val(txtOSDpt.text) < Val(txtDptAmount.text) And bEdit = False Then               'Refund
      MsgBox "Maximum refund amount cannot be greater than £" & Format(txtOSDpt.text, "0.00"), vbCritical + vbOKOnly, "Refund - Deposit"
      txtDptAmount.text = Format(txtOSDpt.text, "0.00")
      SelTxtInCtrl txtDptAmount
      txtDptAmount.SetFocus
   End If
   If (yDEPOSIT = 4) And Val(txtOSDpt.text) < Val(txtDptAmount.text) And bEdit = False Then                'expense
      MsgBox "Maximum expense amount cannot be greater than £" & Format(txtOSDpt.text, "0.00"), vbCritical + vbOKOnly, "Expense - Deposit"
      txtDptAmount.text = Format(txtOSDpt.text, "0.00")
      SelTxtInCtrl txtDptAmount
      txtDptAmount.SetFocus
   End If
   If txtDptAmount.text <> "" Then txtDptAmount.text = Format(txtDptAmount.text, "0.00")
'   txtOSDpt = txtDptAmount.text 'If txtOSDpt.text = "" Then
End Sub

Private Sub txtDptDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDptAmount
    End If
End Sub

Private Sub txtEmail1_LostFocus()
'   Dim szErrMsg As String
'
'   If Trim(txtEmail1.text) <> "" Then
'      If Not ValidateEmail(txtEmail1.text, szErrMsg) Then
'         MsgBox szErrMsg, vbCritical + vbOKOnly, "Lessee Email"
'         SelTxtInCtrl txtEmail1
'         txtEmail1.SetFocus
'      End If
'   End If
End Sub

Private Sub txtEmail2_LostFocus()
   Dim szErrMsg As String

   If Trim(txtEmail2.text) <> "" Then
      If Not ValidateEmail(txtEmail2.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Supplier Email"
         SelTxtInCtrl txtEmail2
         txtEmail2.SetFocus
      End If
   End If
End Sub



Private Sub txtFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdFund.SetFocus
    End If
End Sub

Private Sub txtName_LostFocus()
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   If UBound(szaChoice) > 0 Then
      If szaChoice(2) <> "" Then
         If InStr(szaChoice(2), "L") > 0 Then
            If NEWMODE_ And txtTenantID.text = "" And Trim(txtName.text) <> "" Then
'               txtTenantID.text = CreateTenantId(txtName.text)
               cboSageAccountNumber.text = txtTenantID.text
            End If
         End If
      End If
   End If
   txtCompanyName.text = txtName.text
End Sub

Private Function CreateTenantId(szName As String) As String
   Dim szSQL As String, i As Integer, szChar As String, j As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   For i = 1 To Len(szName) - 1
      szChar = UCase(Mid(szName, i, 1))
      If (szChar >= "A" And szChar <= "Z") Then
         CreateTenantId = CreateTenantId & szChar
         j = j + 1
      End If
      If j = 8 Then Exit For
   Next i

   If j < 8 Then CreateTenantId = Left(CreateTenantId & "01234567", 8)

   adoConn.Open getConnectionString
   
   szSQL = "SELECT SageAccountNumber " & _
           "FROM Tenants " & _
           "WHERE Tenants.SageAccountNumber = '" & CreateTenantId & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   j = 1
   Do
      If adoRST.EOF Then Exit Do
      adoRST.Close
      CreateTenantId = Left(CreateTenantId & "01234567", 6) & Format(j, "00")
      szSQL = "SELECT SageAccountNumber " & _
              "FROM Tenants " & _
              "WHERE Tenants.SageAccountNumber = '" & CreateTenantId & "';"
   
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      
      j = j + 1
   Loop

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function

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
'   Dim j As Integer
'   If szSel = "Client" Or szSel = "Property" Then
'        j = 1
'   Else
'        j = 0
'   End If
   If Len(txtSearch2.text) > 0 Then
        txtSearch1.text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
        flxSupplier(0).RowHeight(i) = 240
        If InStr(1, UCase(flxSupplier(0).TextMatrix(i, 1)), UCase(txtSearch2.text), vbTextCompare) = 0 Then
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

Private Sub txtSearchCompany_Change()
    Dim tempstr As String
    Dim Filter As String
'     If Len(txtSearchCompany.text) > 0 Then
'           txtSearchUnitName.text = ""
'           txtSearchTenant.text = ""
'           txtSearchName.text = ""
'           txtSearchUnitName.text = ""
'     End If
   If Len(txtSearchCompany.text) > 0 Then
        txtSearchTenant.text = ""
        txtSearchName.text = ""
        txtSearchUnitName.text = ""
        tempstr = Replace(UCase(txtSearchCompany.text), "'", "''")
        Filter = " UnitNumber LIKE '%" + tempstr + "*'"
   Else
        Filter = ""
   End If
   If sText = "UnitNumber" Then
        Call FilterTenantsList(Filter)
   End If
   
End Sub

Private Sub txtSearchCompany_GotFocus()
    sText = "UnitNumber"
End Sub

Private Sub txtSearchCompany_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     'FilterTenantsList
End Sub

Private Sub txtSearchName_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     'FilterTenantsList
End Sub

Private Sub txtSearchTenant_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     'FilterTenantsList
    
End Sub

Private Sub txtSearchUnitName_Change()
    Dim tempstr As String
   Dim Filter As String
'   Dim i As Integer
'
'   If Len(txtSearchUnitName.text) > 0 Then
'      txtSearchTenant.text = ""
'      txtSearchName.text = ""
'   End If
'
'   For i = 1 To gridTenantLookup.Rows - 1
'      gridTenantLookup.RowHeight(i) = 240
'      If UCase(Left(gridTenantLookup.TextMatrix(i, 2), Len(txtSearchUnitName.text))) <> UCase(txtSearchUnitName.text) Then
'         gridTenantLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 26 Jul 2014
     
'    If Len(txtSearchUnitName.text) > 0 Then
'           txtSearchCompany.text = ""
'           txtSearchTenant.text = ""
'           txtSearchName.text = ""
'           txtSearchTenant.text = ""
'     End If

   If Len(txtSearchUnitName.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
      txtSearchName.text = ""
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
   SelTxtInCtrl txtSearchUnitName
End Sub

Private Sub txtSearchUnitName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
        gridTenantLookup.SetFocus
    End If
End Sub

Private Sub txtSearchUnitName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            gridTenantLookup.SetFocus
    End If
End Sub

Private Sub txtSearchName_Change()
    Dim tempstr As String
   Dim Filter As String
'   Dim i As Integer
'
'   If Len(txtSearchName.text) > 0 Then
'      txtSearchTenant.text = ""
'      txtSearchUnitName.text = ""
'   End If
'
'   For i = 1 To gridTenantLookup.Rows - 1
'      gridTenantLookup.RowHeight(i) = 240
'      If UCase(Left(gridTenantLookup.TextMatrix(i, 1), Len(txtSearchName.text))) <> UCase(txtSearchName.text) Then
'         gridTenantLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 26 Jul 2014
'  If Len(txtSearchName.text) > 0 Then
'           txtSearchUnitName.text = ""
'           txtSearchTenant.text = ""
'           txtSearchCompany.text = ""
'           txtSearchUnitName.text = ""
'     End If
  If Len(txtSearchName.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchCompany.text = ""
      txtSearchUnitName.text = ""
      tempstr = Replace(UCase(txtSearchName.text), "'", "''")
      Filter = " Name LIKE '%" + tempstr + "*'"
   Else
      Filter = ""
   End If
    If sText = "TenantName" Then
       Call FilterTenantsList(Filter)
    End If
End Sub

Private Sub txtSearchName_GotFocus()
   sText = "TenantName"
   SelTxtInCtrl txtSearchName
End Sub

Private Sub txtSearchName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
        gridTenantLookup.SetFocus
    End If
    If KeyCode = 13 Then
            If Len(Trim(txtSearchCompany.text)) > 0 Then
                     FocusControl gridTenantLookup
                     If gridTenantLookup.Rows > 1 Then
                        gridTenantLookup.row = 1
                         gridTenantLookup.TopRow = 1
                    End If
            Else
                    txtSearchCompany.SetFocus
            End If
     End If
End Sub

Private Sub txtSearchTenant_Change()
   Dim tempstr As String
   Dim Filter As String
'   Dim i As Integer
'
'   If Len(txtSearchTenant.text) > 0 Then
'      txtSearchName.text = ""
'      txtSearchUnitName.text = ""
'   End If
'
'   For i = 1 To gridTenantLookup.Rows - 1
'      gridTenantLookup.RowHeight(i) = 240
'      If UCase(Left(gridTenantLookup.TextMatrix(i, 0), Len(txtSearchTenant.text))) <> UCase(txtSearchTenant.text) Then
'         gridTenantLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445.
'Modified By: Asif. 26 Jul 2014
'    If Len(txtSearchTenant.text) > 0 Then
'           txtSearchUnitName.text = ""
'           txtSearchName.text = ""
'           txtSearchCompany.text = ""
'           txtSearchUnitName.text = ""
'     End If


  If Len(Trim(txtSearchTenant.text)) > 0 Then
      txtSearchName.text = ""
      txtSearchUnitName.text = ""
      tempstr = Replace(UCase(txtSearchTenant.text), "'", "''")
      Filter = " SageAccountNumber LIKE '%" + tempstr + "*'"
   Else
      Filter = ""
   End If
   If sText = "TenantID" Then
        Call FilterTenantsList(Filter)
   End If
End Sub

'Resolved by BOSL
'Issue No: 0000445.
'The function generates the expression of matching string pattern by using SQL LIKE operation and
'uses the in-built Filter function of the ADODB recordset to filter the records that match with the
'expression and finally bind the filtered records to the grid.
'Modified By: Asif. 26 Jul 2014
Private Function FilterTenantsList(Filter As String) As String
'   Debug.Print 1
   
   Dim tempstr As String
   'Wild card search has been implemented by anol
   'issue 0000445: Searching issues found through out Prestige
   'Date 22 Feb 2015
   
'   Dim szOrderBy As String
'   Dim szWhere As String
'   szOrderBy = "LeaseDetails.SageAccountNumber ASC"
   
   
  
   
'   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
'      szWhere = ""
'
'   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
'      szWhere = "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
'
'   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
'      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' "
'
'   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
'      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
'                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
                         
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString


   
   Dim adoRST As New ADODB.Recordset
'   adoRst.Close
    'Debug.Print time
    'adoRst.Close
   adoRST.Open SQLTenantList, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        adoRST.Filter = Filter
   End If
   'Debug.Print time
'   MsgBox adoRst.RecordCount
      
  
   'issue 556 Col was coming 7 but in record field was 8
   ConfigGridTenantLookup
   lblLoading.Caption = "Please wait while loading..."
   fmeLoading.Visible = True
   fmeLoading.Top = 2715
   fmeLoading.Left = 3560
   fmeLoading.ZOrder 0
   lblLoading.ZOrder 0
   
   fmeLoading.Refresh
   gridTenantLookup.Rows = adoRST.RecordCount + 1
   
   Dim i As Integer, j As Integer
   Dim iRow As Integer
   iRow = 1
   While Not adoRST.EOF
       gridTenantLookup.TextMatrix(iRow, 0) = adoRST!SageAccountNumber
       gridTenantLookup.TextMatrix(iRow, 1) = adoRST!Name
       gridTenantLookup.TextMatrix(iRow, 2) = IIf(IsNull(adoRST!UnitNumber), "", adoRST!UnitNumber) 'adoRst!CompanyName
       gridTenantLookup.TextMatrix(iRow, 3) = IIf(IsNull(adoRST!UnitName), "", adoRST!UnitName)
       gridTenantLookup.TextMatrix(iRow, 4) = IIf(IsNull(adoRST!Notes), "", adoRST!Notes)
       gridTenantLookup.TextMatrix(iRow, 5) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)
       gridTenantLookup.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!ClientName), "", adoRST!ClientName)
       gridTenantLookup.TextMatrix(iRow, 8) = IIf(IsNull(adoRST!LeaseID), "0", adoRST!LeaseID) '7 is updating with balance
       gridTenantLookup.TextMatrix(iRow, 7) = Format(IIf(IsNull(adoRST!amt), "0", adoRST!amt), "0.00")
   '    Debug.Print gridTenantLookup.TextMatrix(iRow, 7)
       If iRow = 11 Then
            gridTenantLookup.Refresh
            fmeTenantLookup.Refresh
       End If
       adoRST.MoveNext
       iRow = iRow + 1
    Wend
   
   
   fmeLoading.Visible = False
   SetControlStyle gridTenantLookup
   If gridTenantLookup.Rows > 1 Then
        gridTenantLookup.row = 1
         gridTenantLookup.TopRow = 1
   End If
   adoRST.Close
   Set adoRST = Nothing
    'this is buiding column 7 . rem by anol 2023-07-05
   'UpdateBalance ' this is updating col 7 . updating balance

   adoConn.Close
   Set adoConn = Nothing

End Function
Public Function PopulateLeaseInformation(ByVal adoConn As ADODB.Connection, ByVal sTenantSageAC As String) As Boolean
   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT TENANTS.*, LEASEINFO.* " _
       & " FROM TENANTS LEFT JOIN (" _
       & "  SELECT " _
       & "     CLIENT.CLIENTNAME AS CLIENT, CLIENT.ClientID, " _
       & "     PROPERTY.PROPERTYID + '-' + PROPERTY.PROPERTYNAME AS PROPERTY, " _
       & "     UNITS.UNITNUMBER + '-'+ UNITS.UNITNAME AS UNIT, " _
       & "     LeaseDetails.LeaseID,LeaseDetails.SageAccountNumber as LeaseSAGEAC, " _
       & "     LeaseDetails.HoldingOver, UNITS.UNITNUMBER, " _
       & "     LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.RentReviewDate " _
       & "  From " _
       & "     LEASEDETAILS, UNITS , CLIENT, TENANTS, PROPERTY " _
       & "  Where " _
       & "     LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " _
       & "     UNITS.PROPERTYID = PROPERTY.PROPERTYID AND " _
       & "     Property.CLIENTID = CLIENT.CLIENTID AND " _
       & "     LEASEDETAILS.SageAccountNumber=TENANTS.SageAccountNumber " _
       & " )AS LEASEINFO ON TENANTS.SAGEACCOUNTNUMBER = LEASEINFO.LeaseSAGEAC " _
       & " WHERE LEASEDETAILS.leaseID = '" & txtLeaseId.text & "'"
'Debug.Print sSQLQuery_
   If Not FillFormUsingADODB(Me, adoConn, sSQLQuery_) Then
      
      MsgBox "WARNING !! No information found for the specified Lessee.", vbExclamation
   End If



   bLeaseSetup = True
   If txtLeaseId.text = "" Then
       fmeLoading.Visible = False
       lblLoading.Visible = False
       MsgBox "WARNING !! There is no Lease setup for this Lessee.", vbExclamation
       bLeaseSetup = False
   End If
End Function
Public Function PopulateTenantInformation(ByVal adoConn As ADODB.Connection, ByVal sTenantSageAC As String) As Boolean
   Dim sSQLQuery_ As String
   Dim statusSQL As String
   If optCurrentTenant.Value = True Then
        statusSQL = " and isCurrent"
   ElseIf optExTenant.Value = True Then
         statusSQL = " and isCurrent=0"
   ElseIf optBoth.Value = True Then
   End If
   
  
   sSQLQuery_ = "SELECT TENANTS.*, LEASEINFO.* " _
       & " FROM TENANTS LEFT JOIN (" _
       & "  SELECT " _
       & "     CLIENT.CLIENTNAME AS CLIENT, CLIENT.ClientID, " _
       & "     PROPERTY.PROPERTYID + '-' + PROPERTY.PROPERTYNAME AS PROPERTY, " _
       & "     UNITS.UNITNUMBER + '-'+ UNITS.UNITNAME AS UNIT, " _
       & "     LeaseDetails.LeaseID,LeaseDetails.SageAccountNumber as LeaseSAGEAC, " _
       & "     LeaseDetails.HoldingOver, UNITS.UNITNUMBER, " _
       & "     LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.RentReviewDate " _
       & "  From " _
       & "     LEASEDETAILS, UNITS , CLIENT, TENANTS, PROPERTY " _
       & "  Where " _
       & "     LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " _
       & "     UNITS.PROPERTYID = PROPERTY.PROPERTYID AND " _
       & "     Property.CLIENTID = CLIENT.CLIENTID AND " _
       & "     LEASEDETAILS.SageAccountNumber=TENANTS.SageAccountNumber " _
       & " )AS LEASEINFO ON TENANTS.SAGEACCOUNTNUMBER = LEASEINFO.LeaseSAGEAC " _
       & " WHERE SageAccountNumber = '" & sTenantSageAC & "'" & statusSQL
'Debug.Print sSQLQuery_
   If Not FillFormUsingADODB(Me, adoConn, sSQLQuery_) Then

      MsgBox "WARNING !! No information found for the specified Lessee.", vbExclamation
   End If
'    Dim rsTenants As New ADODB.Recordset
'    rsTenants.Open sSQLQuery_, adoConn, adOpenKeyset, adLockOptimistic
'    If Not rsTenants.EOF Then
'        With rsTenants
'            '!SageAccountNumber = txtTenantID.text
'            '!TenantID = txtTenantID.text
'            txtName.text = !Name
'            txtCompanyName.text = IIf(IsNull(!CompanyName), "", !CompanyName)
'            txtContact1.text = IIf(IsNull(!Contact1), "", !Contact1)
'            txtEmail1.text = IIf(IsNull(!Email1), "", !Email1)
'            txtDirectLine1.text = IIf(IsNull(!DirectLine1), "", !DirectLine1)
'            txtContact2.text = IIf(IsNull(!Contact2), "", !Contact2)
'            txtEmail2.text = IIf(IsNull(!Email2), "", !Email2)
'            txtDirectLine2.text = IIf(IsNull(!DirectLine2), "", !DirectLine2)
'            txtHOAddressLine1.text = IIf(IsNull(!HOAddressLine1), "", !HOAddressLine1)
'            txtHOAddressLine2.text = IIf(IsNull(!HOAddressLine2), "", !HOAddressLine2)
'            txtHOAddressLine3.text = IIf(IsNull(!HOAddressLine3), "", !HOAddressLine3)
'            txtHOAddressLine4.text = IIf(IsNull(!HOAddressLine4), "", !HOAddressLine4)
'            txtHOPostCode.text = IIf(IsNull(!HOPostCode), "", !HOPostCode)
'            txtHOTelephone.text = IIf(IsNull(!HOTelephone), "", !HOTelephone)
'            txtHOFax.text = IIf(IsNull(!HOFax), "", !HOFax)
'            txtBillAddressLine1.text = IIf(IsNull(!BillAddressLine1), "", !BillAddressLine1)
'            txtBillAddressLine2.text = IIf(IsNull(!BillAddressLine2), "", !BillAddressLine2)
'            txtBillAddressLine3.text = IIf(IsNull(!BillAddressLine3), "", !BillAddressLine3)
'            txtBillAddressLine4.text = IIf(IsNull(!BillAddressLine4), "", !BillAddressLine4)
'            txtBillPostCode.text = IIf(IsNull(!BillPostCode), "", !BillPostCode)
'            txtBillTelephone.text = IIf(IsNull(!BillTelephone), "", !BillTelephone)
'            txtBillFax.text = IIf(IsNull(!BillFax), "", !BillFax)
'            cboInvoiceTo.text = IIf(IsNull(!InvoiceTo), "", !InvoiceTo)
''            txtInvoiceTo.text = IIf(IsNull(!TenantMemo, "", !TenantMemo)
'            '!TenantMemo = txtInvoiceTo.text
'            '!balance = txtBalance.text
'            txtDeposit.text = IIf(IsNull(!Deposite), "", !Deposite)
'            txtBank.text = IIf(IsNull(!BankCode), "", !BankCode)
'            '!spare1=txtInvoiceTo.text
'           ' IIf(chkEmailDmd.Value = 0, False, True) = IIf(IsNull(!EmailDmd), "", !EmailDmd)
'
'        End With
'     End If



   bLeaseSetup = True
   If txtLeaseId.text = "" Then
       fmeLoading.Visible = False
       lblLoading.Visible = False
       MsgBox "WARNING !! There is no Lease setup for this Lessee.", vbExclamation
       bLeaseSetup = False
   End If
End Function
'Private Function FillFormInformation() As Boolean
'    Dim adoRst As ADODB.Recordset
'   Set adoRst = New ADODB.Recordset
'
'   adoRst.Open sSQLQuery, adoConnector, adOpenStatic, adLockOptimistic
'
'    If adoRst.EOF Then
'       adoRst.Close
'       Set adoRst = Nothing
'       FillFormInformation = False
'       Exit Function
'    End If
'    cboSageAccountNumber.text = adoRst!SageAccountNumber
'    txtTenantID.text = adoRst!TenantID
'    txtName.text = adoRst!Name
'    txtCompanyName.text = adoRst!CompanyName
'    txtContact1.text = adoRst!Contact1
'    txtEmail1.text = adoRst!Email1
'    txtDirectLine1.text = adoRst!DirectLine1
'    txtContact2.text = adoRst!Contact2
'    txtEmail2.text = adoRst!Email2
'    txtDirectLine2.text = adoRst!DirectLine2
'    txtHOAddressLine1.text = adoRst!HOAddressLine1
'    txtHOAddressLine2.text = adoRst!HOAddressLine2
'    txtHOAddressLine3.text = adoRst!HOAddressLine3
'    txtHOAddressLine4.text = adoRst!HOAddressLine4
'    txtHOPostCode.text = adoRst!HOPostCode
'    txtHOTelephone.text = adoRst!HOTelephone
'    txtHOFax.text = adoRst!HOFax
'    txtBillAddressLine1.text = adoRst!BillAddressLine1
'    txtBillAddressLine2.text = adoRst!BillAddressLine2
'    txtBillAddressLine3.text = adoRst!BillAddressLine3
'    txtBillAddressLine4.text = adoRst!BillAddressLine4
'    txtBillPostCode.text = adoRst!BillPostCode
'    txtBillTelephone.text = adoRst!BillTelephone
'    txtBillFax.text = adoRst!BillFax
'    cboInvoiceTo.text = adoRst!InvoiceTo
'    txtBalance.text = adoRst!balance
'    txtRCCComments.text = adoRst!RCCComments
'    chkEmailDmd.text = adoRst!EmailDmd
'    chkEmailSt.text = adoRst!EmailSt
'    frmLeasee1.cmdSentStByEmail.Enabled = adoRst!chkEmailSt
'    chkCombEmail.text = adoRst!CombEmail
'    chkEmailSC.text = adoRst!EmailSC
'    txtSLControl.text = adoRst!SLControl
'    txtClient.text = adoRst!CLIENT
'    txtClientID.text = adoRst!clientID
'    txtProperty.text = adoRst!Property
'    txtUnit.text = adoRst!unit
'    txtLeaseId.text = adoRst!LeaseID
'    chkHoldingOver.Value = adoRst!HoldingOver
'    txtUnitNumber.text = adoRst!UNITNUMBER
'    txtStartDate .text = adoRst!StartDate
'    txtEndDate.text = adoRst!EndDate
'    txtRentReviewDate.text = adoRst!RentReviewDate
'
'End Function
Public Sub PopulateCodes(adoConn As ADODB.Connection)
   Dim sSQLQuery As String

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
''   sSQLQuery = "SELECT BANK_ID, BANK_NAME, SORT_CODE, BANK_BRANCH, BANK_ADDRESS1, BANK_ADDRESS2, BANK_ADDRESS3, BANK_POST_CODE " & _
''                 "FROM tlbBank "
''
''   populateCombo adoConn, sSQLQuery, cboBankId
   
''   ' Payment Method
''   sSQLQuery = "SELECT CODE, VALUE " & _
''                 "FROM SECONDARYCODE " & _
''                 "WHERE PRIMARYCODE = 'PM'"
''
''   populateCombo adoConn, sSQLQuery, cboPaymentMethod
End Sub

Public Sub populateGroupCombo(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRST As ADODB.Recordset
   Set adoRST = New ADODB.Recordset
   
   szSQL = "SELECT Distinct  TenantDeposit.GroupNo " & _
            "FROM TenantDeposit " & _
            "WHERE TenantDeposit.TenantID = '" & txtTenantID.text & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   Dim TotalRow As Long, TotalCol As Long

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   Dim Data() As String
   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   Dim i As Integer, j As Integer

   For i = 0 To adoRST.RecordCount - 1
      For j = 0 To adoRST.Fields.Count - 1
         Data(j, i) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
      Next j
      adoRST.MoveNext
   Next i

   cboGroup.Clear
   cboGroup.ColumnCount = TotalCol
   cboGroup.Column() = Data()
   cboGroup.BoundColumn = 1
   cboGroup.TextColumn = 1

   adoRST.Close
   Set adoRST = Nothing

   Exit Sub
'     Error Handling Code
Error_Handler:
   MsgBox "An Error occurred while populating the Group ID"
End Sub

Public Sub EventHistoryButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control
   Select Case mode

   Case ComponentMode.DefaultMode
      cmdNewEvent.Enabled = True
      cmdEditEvent.Enabled = False
      cmdSaveEvent.Enabled = False
      cmdCancelEvent.Enabled = False
'      gridEventHistory.Enabled = True

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
'      gridEventHistory.Enabled = True

   Case ComponentMode.NewEntryMode
      cmdNewEvent.Enabled = False
      cmdEditEvent.Enabled = False
      cmdSaveEvent.Enabled = True
      cmdCancelEvent.Enabled = True
'      gridEventHistory.Enabled = False

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
'      gridEventHistory.Enabled = False

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

'Public Sub ConfigureFlxBank()
'   With gridBank
'      .Clear
'      .Rows = 2
'      .Cols = 7
'
'      .ColWidth(0) = 0
'      .ColWidth(1) = 0
'      .ColWidth(2) = Label31(1).Left - Label31(0).Left
'      .ColWidth(3) = Label31(2).Left - Label31(1).Left
'      .ColWidth(4) = Label31(3).Left - Label31(2).Left
'      .ColWidth(5) = Label31(4).Left - Label31(3).Left
'      .ColWidth(6) = Label31(5).Left - Label31(4).Left
'      .ColWidth(7) = .Width + .Left - Label31(5).Left - 200
'   End With
'End Sub

Public Function SaveTenantInformation() As Boolean
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rstTenants As New ADODB.Recordset
   
   ' Event Type
   Dim sSQLQuery As String

   If Not NEWMODE_ Then
      sSQLQuery = "SELECT  SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                "TenantMemo, Balance, Deposite, BankCode, spare1, " & _
                "EmailDmd, EmailSt, CombEmail, EmailSC " & _
             "FROM TENANTS " & _
             "Where TenantID = '" & lblLeaseChanged.Caption & "'"
        'This is edit mode
'        rstTenants.Open sSQLQuery, adoConn, adOpenStatic, adLockOptimistic
'        If rstTenants.EOF Then
'                 MsgBox "This record does not exist.", vbInformation, "WARNING !!"
'        End If
      If PostToDBUsingADODB(Me, fmeTenant, adoConn, sSQLQuery, False) Then
         SaveTenantInformation = True

         UpdateLesseeRec lblLeaseChanged.Caption, txtTenantID.text, adoConn
      Else
         SaveTenantInformation = False
      End If
   Else
      sSQLQuery = "SELECT  SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                  "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                  "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                  "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                  "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                  "TenantMemo, Balance, Deposite, BankCode, spare1, " & _
                "EmailDmd, EmailSt, CombEmail, EmailSC,isCurrent,CreatedBy,CreatedDate " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & txtTenantID.text & "' AND Comments<>'DELETED'"
       If SaveNewLessee(Me, fmeTenant, adoConn, sSQLQuery, True) Then
           SaveTenantInformation = True
       Else
           SaveTenantInformation = False
       End If
       'Updating User name and update time
       sSQLQuery = "SELECT  LastModifiedBy,LastModifiedDate " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & txtTenantID.text & "' AND Comments<>'DELETED'"
               
       Dim rs  As New ADODB.Recordset
       rs.Open sSQLQuery, adoConn, adOpenKeyset, adLockOptimistic
       If Not rs.EOF Then
            rs!LastModifiedBy = User
            rs!LastModifiedDate = Now
            rs.Update
       End If
        rs.Close
       
               
   End If

   adoConn.Close
   Set adoConn = Nothing
End Function
Public Function SaveNewLessee(frmCurrent As Form, ByVal oContainer As Control, ByVal oConnector As ADODB.Connection, ByVal sSQLQuery As String, ByVal IsNewRecord As Boolean) As Boolean
   Dim iFieldsCount, iControlCount, i, j As Integer
   Dim sNextField As String
   Dim oControl As Control

   Dim oResultSet As New ADODB.Recordset

'   On Error GoTo Exception
'Debug.Print sSQLQuery
   oResultSet.Open sSQLQuery, oConnector, adOpenStatic, adLockOptimistic

   If IsNewRecord Then
      If oResultSet.EOF Or oResultSet.BOF Then
          oResultSet.AddNew
          oResultSet("isCurrent") = True
          oResultSet("CreatedBy") = User
          oResultSet("CreatedDate") = Now
      Else
          MsgBox "WARNING !! This reference already exists. Please enter a unique reference.", vbInformation
          SaveNewLessee = False
          Exit Function
      End If
   Else
      If oResultSet.EOF Or oResultSet.BOF Then
          MsgBox "WARNING !! This record does not exist.", vbInformation
          SaveNewLessee = False
          Exit Function
      End If
   End If

   iFieldsCount = oResultSet.Fields.Count
   iControlCount = frmCurrent.Controls.Count

   For i = 0 To iFieldsCount - 1
      sNextField = oResultSet.Fields(i).Name
      Debug.Print sNextField

      For Each oControl In frmCurrent.Controls
         If UCase(sNextField) = UCase(Mid(CStr(oControl.Name), 4)) Then
               If sNextField = UCase("SageAccountNumber") Then
                            Debug.Print "1"
               End If
            Select Case TypeName(oControl)
'               If sNextField = UCase("SageAccountNumber") Then
'                            Debug.Print "1"
'               End If
               Case "TextBox"
                 If oControl.Container Is oContainer Then
                    'issue 506 note 1031
                    'Resolved BOSL
                    'Modified by anol 17 Apr 2015
                    'Edit lessee account
                    'It should not be possible to edit a lessee account ID if
                    '1/ It is linked to a lease
                    '2/ It has been used for any transactions
                    '3/ It is connected to any other record
                    
                    If frmCurrent.Name = "frmLeasee1" Then
                        If (sNextField = "SageAccountNumber" Or sNextField = "TenantID") And IsNewRecord = False Then
                            If isHaveTrans(oConnector) = False Then
                                oResultSet.Fields(i).Value = IIf(oControl.text <> "", oControl.text, "")
                            Else
                                'ShowMsgInTaskBar "The lessee ID cannot be changed, it is used in some transactions ", "Y", "N"
                                frmLeasee1.txtTenantID.text = frmLeasee1.lblLeaseChanged.Caption
                            End If
                        Else
                            oResultSet.Fields(i).Value = IIf(oControl.text <> "", oControl.text, "")
                        End If
                    Else
                         oResultSet.Fields(i).Value = IIf(oControl.text <> "", oControl.text, "")
                    End If
                    'End of modification
                    Exit For
                 End If

               Case "CheckBox"
                 If oControl.Container Is oContainer Then
                    If oControl.Value Or oControl.Value = 1 Then
                      oResultSet.Fields(i).Value = True
                    Else
                      oResultSet.Fields(i).Value = False
                    End If
                    Exit For
                 End If

               Case "ComboBox"
                 If oControl.Container Is oContainer Then
                    oResultSet.Fields(i).Value = IIf(oControl.Value <> "", oControl.Value, "")
                    Exit For
                 End If

            End Select
         End If
      Next oControl
   Next i

'  Added by Samrat 19.04.2007. Bug fix by Samrat 19/06/2012
   If frmCurrent.Name = "frmLeasee1" Then
      For Each oControl In frmCurrent.Controls
         If oControl.Name = "chkPrintDmd" Then
            If InStr(oResultSet.Fields.Item("spare1").Value, "NotD") > 0 Then
               If InStr(InStr(oResultSet.Fields.Item("spare1").Value, "NotD") + 1, oResultSet.Fields.Item("spare1").Value, "NotD") > 0 Then
                  oResultSet.Fields.Item("spare1").Value = Replace(oResultSet.Fields.Item("spare1").Value, "NotD", "")
                  oResultSet.Fields.Item("spare1").Value = oResultSet.Fields.Item("spare1").Value & "NotD"
               End If
            End If
            If InStr(oResultSet.Fields.Item("spare1").Value, "NotS") > 0 Then
               If InStr(InStr(oResultSet.Fields.Item("spare1").Value, "NotS") + 1, oResultSet.Fields.Item("spare1").Value, "NotS") > 0 Then
                  oResultSet.Fields.Item("spare1").Value = Replace(oResultSet.Fields.Item("spare1").Value, "NotS", "")
                  oResultSet.Fields.Item("spare1").Value = oResultSet.Fields.Item("spare1").Value & "NotS"
               End If
            End If

            If Not oControl.Value Then
               If IsNull(oResultSet.Fields.Item("spare1").Value) Then
                  oResultSet.Fields.Item("spare1").Value = "NotD"
               Else
                  If InStr(oResultSet.Fields.Item("spare1").Value, "NotD") <= 0 Then
                     oResultSet.Fields.Item("spare1").Value = oResultSet.Fields.Item("spare1").Value & "NotD"
                  End If
               End If
            Else
               If InStr(oResultSet.Fields.Item("spare1").Value, "NotD") > 0 Then
                  oResultSet.Fields.Item("spare1").Value = Replace(oResultSet.Fields.Item("spare1").Value, "NotD", "")
               End If
            End If
         End If
         If oControl.Name = "chkPrintSt" Then
            If Not oControl.Value Then
               If IsNull(oResultSet.Fields.Item("spare1").Value) Then
                  oResultSet.Fields.Item("spare1").Value = "NotS"
               Else
                  If InStr(oResultSet.Fields.Item("spare1").Value, "NotS") <= 0 Then
                     oResultSet.Fields.Item("spare1").Value = oResultSet.Fields.Item("spare1").Value & "NotS"
                  End If
               End If
            Else
               If InStr(oResultSet.Fields.Item("spare1").Value, "NotS") > 0 Then
                  oResultSet.Fields.Item("spare1").Value = Replace(oResultSet.Fields.Item("spare1").Value, "NotS", "")
               End If
            End If
         End If
         If oControl.Name = "chkEmailDmd" Then
            If oControl.Value Then
               oResultSet.Fields("EmailDmd").Value = 1
            Else
               oResultSet.Fields("EmailDmd").Value = 0
            End If
         End If
      Next oControl
   End If
   'added by anol 20170322 issue 343
   If Len(oResultSet.Fields(0).Value) > 30 Then
        MsgBox "Lessee ID cannot be more than 30 character", vbInformation, "Failed to save"
        'oResultSet.Close
        'Set oResultSet = Nothing
        Exit Function
   End If
   oResultSet.Update
   SaveNewLessee = True
   oResultSet.Close
   Set oResultSet = Nothing
   Exit Function

Exception:
   SaveNewLessee = False
   MsgBox Err.description
   Set oResultSet = Nothing
End Function
Private Function isHaveTrans(oConnector As ADODB.Connection) As Boolean
    Dim rsNLposting As New ADODB.Recordset
    rsNLposting.Open "Select ACCOUNT_NUMBER from NLposting where ACCOUNT_NUMBER='" & frmLeasee1.lblLeaseChanged.Caption & "'", oConnector, adOpenStatic, adLockReadOnly
    If rsNLposting.EOF = False Then
        isHaveTrans = True
    End If
End Function
Private Sub UpdateLesseeRec(szOldID As String, szNewID As String, adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szTemp  As String

   szTemp = Replace(txtCompanyName.text, "'", "''")

   szSQL = "UPDATE DemandRecords " & _
           "SET TenantCompanyName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "';"
'Debug.Print szSQL
   adoConn.Execute szSQL

   szSQL = "UPDATE LeaseDetails " & _
           "SET CompanyName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "' AND status=true;"
   adoConn.Execute szSQL

   szTemp = Replace(txtName.text, "'", "''")
   szSQL = "UPDATE tlbLetterReports " & _
           "SET LesseeName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "';"
   adoConn.Execute szSQL
End Sub

Private Sub TenantTabEnabled(ByVal IsEnabled As Boolean)
   tabTenant.Enabled = IsEnabled

   If IsEnabled Then
       ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
'       ComponentInFrameEnableMode Me, fmeTenancyDetails, EditMode
'       ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
       ComponentInFrameEnableMode Me, fmeEventHistory, DefaultMode
   End If
End Sub

Private Sub LoadNominalCode()
     flxSupplier(0).Clear
    flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
    flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

'   txtSearch1.Width = 1400
'   txtSearch1.Left = 40
'
'   txtSearch2.Width = 2600
'   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "N/C"
   lblSearch1(0).Caption = "Name"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger " & _
           "WHERE ClientID = '" & txtClientID.text & "' " & _
           "ORDER BY Code;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("Code").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("Name").Value
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadFundGrid()
    flxSupplier(0).Clear
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   'txtSearch1.Width = 1400
  ' txtSearch1.Left = 40

   'txtSearch2.Width = 2600
  ' txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Fund Code"
   lblSearch1(0).Caption = "Fund Name"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code;"
   szSQL = "SELECT FundID,FundCode, FundName FROM FUND;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   


   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("FundCode").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("FundName").Value
       flxSupplier(0).TextMatrix(iRows, 2) = adoRST.Fields.Item("FundID").Value
       flxSupplier(0).RowHeight(iRows) = 280
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
       flxSupplier(0).row = 1
      iRows = iRows + 1
      adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadBankGrid()
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   'txtSearch1.Width = 1400
  ' txtSearch1.Left = 40

   'txtSearch2.Width = 2600
  ' txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Bank Code"
   lblSearch1(0).Caption = "Name"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code;"
szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND tlbClientBanks.CLient_ID = NominalLedger.CLientID AND tlbClientBanks.CLient_ID='" & txtClientID.text & "' " & _
           "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
'AND  tlbClientBanks.CLIENT_ID='" & txtClientID.text & "'
'   If adoRst.EOF Then
'      MsgBox "Please setup bank account for the client."
'   Else
'      ReDim Data(1, adoRst.RecordCount - 1) As String
'      i = 0
'      While Not adoRst.EOF
'         Data(0, i) = adoRst.Fields.Item("BNC").Value
'         Data(1, i) = adoRst.Fields.Item("BNN").Value
'         i = i + 1
'         adoRst.MoveNext
'      Wend
'      cmbBank.Clear
'      cmbBank.Column() = Data()
'   End If
'
'   adoRst.Close
'
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("BNC").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("BNN").Value
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub
Private Sub ButtonHanlding(ByVal mode As ComponentMode)
   Dim X As String

   Select Case mode

   Case ComponentMode.DefaultMode
      cmdDptNew.Enabled = True
      cmdDptRefund.Enabled = True
      cmdDptExpenses.Enabled = True
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = False
'      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = False
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Deposit"
      cmdDptCancel.Caption = "Cancel Deposit"
'      cmdDptEdit.Caption = "Edit Deposit"
'      cmdDptDelete.Caption = "Delete Deposit"
      txtDptAmtType.text = ""
      txtDptAmtType.Tag = ""
      txtDepositType.text = ""
      txtDepositType.Tag = ""
      txtFund.text = ""
      txtFund.Tag = ""
      flxDeposit.Enabled = True

      'cmbBank.Locked = True
      txtDate.Locked = True
'      cmbDptAmtType.Locked = True
'      cboDepositType.Locked = True
      cmdDptAmtType.Enabled = False
      cmdSetAmtType.Enabled = False
      cmdSetDptType.Enabled = False
      txtDptDetails.Locked = True
      txtDptAmount.Locked = True
      txtOSDpt.Locked = True

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = False
      cboGroup.Enabled = False
      cboGroup.Locked = True

      txtBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = ""
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

'      yDEPOSIT = 0

   Case ComponentMode.NewEntryMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
'      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False
      cmdDptExpenses.Enabled = False

      cmdDptSave.Caption = "Save Deposit"
      cmdDptCancel.Caption = "Cancel Deposit"
'      cmdDptEdit.Caption = "Edit Deposit"
'      cmdDptDelete.Caption = "Delete Deposit"

      flxDeposit.Enabled = False

      'cmbBank.Locked = False
      txtDate.Locked = False

'      cmbDptAmtType.Locked = False
'      cboDepositType.Locked = False
      cmdDptAmtType.Enabled = True
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False

      optNewGroup.Enabled = True
      optExitingGroup.Enabled = True
      cboGroup.Enabled = False
      cboGroup.Locked = True

      txtBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = Format(Now, "dd/mm/yyyy")
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

   Case ComponentMode.EditMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
'      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save"
      cmdDptCancel.Caption = "Cancel"
'      cmdDptEdit.Caption = "Edit"
'      cmdDptDelete.Caption = "Delete"

      flxDeposit.Enabled = False

      'cmbBank.Locked = False
      txtDate.Locked = False
'      cmbDptAmtType.Locked = False
      cmdDptAmtType.Enabled = True
'      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = True
      optExitingGroup.Value = True
      cboGroup.Enabled = True
      cboGroup.Locked = False

   Case ComponentMode.GridRowOnSelection
      cmdDptNew.Enabled = True
      cmdDptRefund.Enabled = True
      cmdDptExpenses.Enabled = True
      cmdDptEdit.Enabled = True
      cmdDptSave.Enabled = False
'      cmdDptDelete.Enabled = True
      cmdDptCancel.Enabled = False
      cmdDptPrint.Enabled = True

      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "D" Then X = "Deposit"
      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "R" Then X = "Refund"
      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "E" Then X = "Expenses"
      cmdDptSave.Caption = "Save " & X
      cmdDptEdit.Caption = "Edit " & X
'      cmdDptDelete.Caption = "Delete " & X
      cmdDptCancel.Caption = "Cancel " & X

      'cmbBank.Locked = True
      txtDate.Locked = True
      'cmbDptAmtType.Locked = True
      cmdDptAmtType.Enabled = False
'      cboDepositType.Locked = True
      cmdSetAmtType.Enabled = False
      cmdSetDptType.Enabled = False
      txtDptDetails.Locked = True
      txtDptAmount.Locked = True
      txtOSDpt.Locked = True

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = False
      optExitingGroup.Value = True
      cboGroup.Enabled = False
      cboGroup.Locked = True

      'yDEPOSIT = 0
      txtBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = ""
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

   Case ComponentMode.RefundMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
'      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Refund"
      cmdDptCancel.Caption = "Cancel Refund"
      txtDate.text = ""

      flxDeposit.Enabled = False

      'cmbBank.Locked = False
      txtDate.Locked = False
'      cmbDptAmtType.Locked = False
      cmdDptAmtType.Enabled = True
'      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False
   
   Case ComponentMode.ExpensesMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
'      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Expenses"
      cmdDptCancel.Caption = "Cancel Expenses"

      flxDeposit.Enabled = False
      txtDptAmount.text = txtOSDpt.text
      txtDate.text = ""

      'cmbBank.Locked = False
      txtDate.Locked = False
'      cmbDptAmtType.Locked = False
      cmdDptAmtType.Enabled = True
'      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False
   End Select
End Sub

Private Sub ConfigFlxLetter()
   Dim szHeader As String, iCol As Integer

   flxLetters.Clear
   flxLetters.Cols = 6
   flxLetters.Rows = 2
   flxLetters.RowHeight(0) = 260
   szHeader$ = "|<ID|<TemplateName|<Subject|<Unit No|<Print Date"
   '           0  1     2           3         4        5
   flxLetters.FormatString = szHeader$
   flxLetters.ColWidth(0) = 200
   flxLetters.ColWidth(1) = 0
   flxLetters.ColWidth(2) = 3000
   flxLetters.ColWidth(3) = 5000
   flxLetters.ColWidth(4) = 0
   flxLetters.ColWidth(5) = 1600
End Sub

Private Sub ConfigFlxEmails()
   Dim szHeader As String, iCol As Integer

   flxEmails.Clear
   flxEmails.Cols = 8
   flxEmails.Rows = 2
   flxEmails.RowHeight(0) = 260
   szHeader$ = "|<Date|<Time|<Email|<Subject|<Body|ID"
   '           0    1    2      3       4       5   6
   flxEmails.FormatString = szHeader$
   flxEmails.ColWidth(0) = 200
   flxEmails.ColWidth(1) = 1000
   flxEmails.ColWidth(2) = 700
   flxEmails.ColWidth(3) = 2600
   flxEmails.ColWidth(4) = 3450
   flxEmails.ColWidth(5) = 3100
   flxEmails.ColWidth(6) = 0
   flxEmails.ColWidth(7) = 0 'temp date
End Sub

Public Sub LoadFlxEmails()
   Dim bLesseeEmail  As Boolean
   Dim bFlag         As Boolean
   Dim szPath        As String
   Dim szLine        As String
   Dim szaDateTime() As String
   Dim iRow          As Integer

   ConfigFlxEmails

   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"

   bLesseeEmail = FileExists(szPath)

   If bLesseeEmail Then
      Open szPath For Input As #2

      iRow = 1
      While Not EOF(2)
         Line Input #2, szLine
         bFlag = False
'                                szHeader$ = "|<Date|<Time|<Email|<Subject|<Body|ID"
         If InStr(szLine, "Email sent on:") > 0 Then
            szLine = Mid(szLine, 15)
            szaDateTime = Split(szLine, "#")
            flxEmails.TextMatrix(iRow, 1) = szaDateTime(0)
            If UBound(szaDateTime) > 0 Then
               flxEmails.TextMatrix(iRow, 2) = szaDateTime(1)
            End If
            If UBound(szaDateTime) > 1 Then
               flxEmails.TextMatrix(iRow, 6) = szaDateTime(2)           'ID
            End If
            bFlag = True
         End If
         If InStr(szLine, "Email Address:") > 0 Then
            szLine = Mid(szLine, 15)
            flxEmails.TextMatrix(iRow, 3) = szLine
            bFlag = True
         End If
         If InStr(szLine, "Email Subject:") > 0 Then
            szLine = Mid(szLine, 15)
            flxEmails.TextMatrix(iRow, 4) = szLine
            bFlag = True
         End If
         If Not bFlag Then
            If InStr(szLine, "*****") = 0 Then
               flxEmails.TextMatrix(iRow, 5) = szLine
               While Not EOF(2) And InStr(szLine, "*****") = 0
                  Line Input #2, szLine
               Wend
            Else
               flxEmails.TextMatrix(iRow, 5) = Mid(szLine, 1, Len(szLine) - 5)
            End If
            iRow = iRow + 1
            If Not EOF(2) Then flxEmails.AddItem ""
         End If
      Wend
      Close #2
   End If
   For iRow = 1 To flxEmails.Rows - 1
         flxEmails.TextMatrix(iRow, 7) = flxEmails.TextMatrix(iRow, 1)
         flxEmails.TextMatrix(iRow, 1) = Format(flxEmails.TextMatrix(iRow, 7), "yyyymmdd")
'         flxEmails.TextMatrix(iRow, 5) = Replace(flxEmails.TextMatrix(iRow, 5), vbCrLf, " ")
'          flxEmails.TextMatrix(iRow, 4) = Replace(flxEmails.TextMatrix(iRow, 4), vbCrLf, " ")
   Next
   SortingGrid flxEmails, 1, True, "Integer"
   For iRow = 1 To flxEmails.Rows - 1
         flxEmails.TextMatrix(iRow, 1) = flxEmails.TextMatrix(iRow, 7)
   Next
   flxEmails.col = 1
   flxEmails.ColSel = 2
   'flxEmails.Sort = flexSortGenericDescending
   flxEmails.row = 0
End Sub

Private Sub LoadFlxLetter(conUnit_ As ADODB.Connection)
   Dim rstReport As New ADODB.Recordset
   Dim sSQLQuery_ As String, iRow As Integer

   ConfigFlxLetter

   sSQLQuery_ = "SELECT L.Id, T.TemplateName, " & _
                     "T.Description AS Subject, L.UnitNo, L.printDate " & _
                "FROM tlbLetterReports AS L, Template AS T " & _
                "WHERE L.sageAccountNumber = '" & txtTenantID.text & "' AND " & _
                     " L.TemplateID = T.TemplateID;"

   rstReport.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

'   szHeader$ = "|<ID|<TemplateName|<Subject|<Unit No|<Print Date"
   iRow = 1
   With flxLetters
      While Not rstReport.EOF
         .TextMatrix(iRow, 1) = IIf(IsNull(rstReport!Id), "", rstReport!Id)
         .TextMatrix(iRow, 2) = IIf(IsNull(rstReport!TemplateName), "", rstReport!TemplateName)
         .TextMatrix(iRow, 3) = IIf(IsNull(rstReport!Subject), "", rstReport!Subject)
         .TextMatrix(iRow, 4) = IIf(IsNull(rstReport!UnitNo), "", rstReport!UnitNo)
         .TextMatrix(iRow, 5) = IIf(IsNull(rstReport!PrintDate), "", rstReport!PrintDate)
         iRow = iRow + 1
         rstReport.MoveNext
         If Not rstReport.EOF Then .AddItem ""
      Wend
   End With
   rstReport.Close
   Set rstReport = Nothing
End Sub

Private Sub txtSearchTenant_GotFocus()
    sText = "TenantID"
   SelTxtInCtrl txtSearchTenant
End Sub

Private Sub txtSearchTenant_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown And Len(txtSearchTenant.text) > 0 Then
        gridTenantLookup.SetFocus
    End If
    If KeyCode = 13 Then
            If Len(txtSearchTenant.text) > 0 Then
                    If gridTenantLookup.Rows > 1 Then
                        gridTenantLookup.row = 1
                        gridTenantLookup.TopRow = 1
                    End If
                   gridTenantLookup.SetFocus
            Else
                   txtSearchName.SetFocus
            End If
   End If
End Sub

Private Sub txtSearchTenant_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            If Len(Trim(txtSearchTenant.text)) > 0 Then
                     FocusControl gridTenantLookup
                     If gridTenantLookup.Rows > 1 Then
                        gridTenantLookup.row = 1
                         gridTenantLookup.TopRow = 1
                    End If
            Else
                   txtSearchName.SetFocus
            End If
    End If
End Sub
 

Private Sub txtSearchUnitName_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    FilterTenantsList
End Sub

Private Sub txtTenantID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
   If (KeyAscii >= 65 And KeyAscii <= 90) Or _
         (KeyAscii >= 97 And KeyAscii <= 122) Or _
         (KeyAscii >= 48 And KeyAscii <= 57) Then
      If (KeyAscii >= 97 And KeyAscii <= 122) Then
         KeyAscii = KeyAscii - 32
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txtTenantID_LostFocus()
   If txtTenantID.Locked Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim szID As String

   adoConn.Open getConnectionString

   szID = txtTenantID.text

   If (IsAccountExist(szID, adoConn)) Then
      If (Not (txtTenantID.text = szID)) Then
         MsgBox "This ID is already in use. A possible suggestion is '" & szID & "' but you may choose a different ID"
         txtTenantID.text = szID
         SelTxtInCtrl txtTenantID
      End If
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub ConfigGridMaintenanceHistory(ByVal rstMHistory_ As ADODB.Recordset)
   Dim iColumn As Integer
   Dim oColumn As ADODB.Field

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = rstMHistory_.Fields.Count + 1

   For iColumn = 1 To 10
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn - 1).Left - 70
   gridMaintenanceHistory.ColWidth(11) = 900
   For iColumn = 12 To rstMHistory_.Fields.Count
      gridMaintenanceHistory.ColWidth(iColumn) = 0
   Next iColumn

   iColumn = 0
   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.RowHeight(0) = 0
   For Each oColumn In rstMHistory_.Fields
      gridMaintenanceHistory.TextMatrix(0, iColumn) = oColumn.Name
      gridMaintenanceHistory.col = iColumn
      gridMaintenanceHistory.CellFontBold = True
      iColumn = iColumn + 1
   Next oColumn
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
'''  All tenants
''   If optBoth.Value Then
''      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
''                   "'' AS Balance, " & _
''                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
''                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
''              "FROM Tenants AS T LEFT JOIN " & _
''                   "[" & _
''                   "SELECT U.UnitName, L.SageAccountNumber, " & _
''                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
''                   "FROM Units AS U, LeaseDetails AS L, " & _
''                        "Property AS P " & _
''                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
''                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "
''
''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
''                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE ORDER BY T.SageAccountNumber;" 'AND OCCUPIDE_ = FALSE
''
''   End If
'''AND IsNull(TerminateDate) has been added by anol 14 May 2015
'''The bug was that it was showing termindated lease
'''Changed join type LEFT to inner
''
''
''
'''  Current tenants Only
''
'''Changed join type LEFT to inner  by anol 26 Jun 2015
'''Changed join type inner to left  by anol 23 aug 2015
'''Changed query   by anol 21 01 2017
'''PRESTIGE VALIDATION ALL MODULES AND FORMS Note 1146
'''L.Status = TRUE AND ommited AND..... AND IsNull(IQ.TerminateDate) added
''   If optCurrentTenant.Value Then
'''      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'''                   "'' AS Balance, " & _
'''                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'''                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber,IQ.Status  " & _
'''              "FROM Tenants AS T LEFT JOIN " & _
'''                   "[" & _
'''                   "SELECT U.UnitName, L.SageAccountNumber,L.TerminateDate, " & _
'''                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'''                   "FROM Units AS U, LeaseDetails AS L, " & _
'''                        "Property AS P " & _
'''                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
'''                      "  U.PropertyID = P.PropertyID AND L.Status = TRUE"
'''
'''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'''                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE ORDER BY T.SageAccountNumber; "
''
''                 'Below code has been implemented on 20180324
'''            szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'''                   "'' AS Balance, " & _
'''                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'''                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber,IQ.Status  " & _
'''                   "FROM Tenants AS T LEFT JOIN " & _
'''                   "[" & _
'''                   "Select SELECT U.UnitName, S.SageAccountNumber,S.TerminateDate,P.PropertyID, P.ClientID, U.UnitNumber,L.STATUS from LeaseDetails as S INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  " & _
'''                          "MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails" & _
'''                   "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber INNER JOIN  Units AS U ON U.UnitNumber = S.UnitNumber  " & _
'''                        "INNER JOIN Property AS P  U.PropertyID = P.PropertyID  where " & _
'''                        "IQ.MaxOfStartDate=S.StartDate AND S.status=false "
'''            szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'''                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE ORDER BY T.SageAccountNumber; "
'''
''''Below code has been implemented on 20170121
'''           szSQL = "SELECT  LeaseDetails.SageAccountNumber, " & _
'''         "Tenants.CompanyName, UnitName, '' AS Balance, IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes," & _
'''               "Property.PropertyID, Property.ClientID " & _
'''           "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
'''           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
'''               "LeaseDetails.Status = True And " & _
'''               "Units.PropertyId = Property.PropertyID And " & _
'''               "Property.ClientID = Client.ClientID AND " & _
'''               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
'''               "" & szWhere & " " & _
'''           "ORDER BY " & szOrderBy & ";"
'''Implemented by anol 20180325, issse 556
''    szSQL = "Select LA.SageAccountNumber, LA.CompanyName,LA.UnitName,  Balance,Notes,LA.PropertyID,LA.ClientID, LA.UnitNumber,LA.Status  FROM (SELECT T.SageAccountNumber, IQ.UnitName, '' AS Balance, IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
''    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber,IQ.Status,CompanyName  FROM Tenants AS T LEFT JOIN (SELECT U.UnitName, L.SageAccountNumber,L.TerminateDate, " & _
''    "P.PropertyID, P.ClientID, U.UnitNumber,L.Status FROM Units AS U, LeaseDetails AS L, Property AS P WHERE U.UnitNumber = L.UnitNumber AND " & _
''    "U.PropertyID = P.PropertyID AND L.Status = TRUE) AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber WHERE ((T.Comments) IS NULL OR T.Comments = '') AND  " & _
''    "OCCUPIDE_ = FALSE ORDER BY T.SageAccountNumber ) as LA LEFT JOIN (  " & _
''    "Select  S.SageAccountNumber from LeaseDetails as S INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM " & _
''    "LeaseDetails GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where   IQ.MaxOfStartDate=S.StartDate AND S.status=false) " & _
''    "as IM ON IM.Sageaccountnumber=LA.sageaccountnumber where IM.Sageaccountnumber IS null order by LA.sageaccountnumber; "
''
''           '
''           Debug.Print szSQL
''
''   End If
''
'''  Deleted tenants only
''   If optExTenant.Value Then
'''   "SELECT SageAccountNumber, CompanyName " & _
'''             "FROM Tenants " & _
'''             "WHERE Tenants.SageAccountNumber NOT IN " & _
'''                 "(SELECT LeaseDetails.SageAccountNumber " & _
'''                 "FROM LeaseDetails " & _
'''                 "WHERE Status=True) AND " & _
'''                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
'''             "ORDER BY SageAccountNumber"
'''rem by anol 20160121
'''       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
'''                   "'' AS Balance, " & _
'''                    "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
'''                    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
'''               "FROM Tenants AS T INNER JOIN " & _
'''                    "[" & _
'''                    "SELECT U.UnitName, L.SageAccountNumber, " & _
'''                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'''                    "FROM Units AS U, LeaseDetails AS L, " & _
'''                        "Property AS P " & _
'''                    "WHERE U.UnitNumber = L.UnitNumber AND " & _
'''                        "L.Status = FALSE AND U.PropertyID = P.PropertyID "
'''
'''      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber ORDER BY T.SageAccountNumber;"
''      'added by anol 21070121
'''      szSQL = "SELECT  LeaseDetails.SageAccountNumber, " & _
'''         "Tenants.CompanyName, UnitName, '' AS Balance, IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes," & _
'''               "Property.PropertyID, Property.ClientID " & _
'''           "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
'''           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
'''               "LeaseDetails.Status = false And " & _
'''               "Units.PropertyId = Property.PropertyID And " & _
'''               "Property.ClientID = Client.ClientID AND " & _
'''               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
'''               "" & szWhere & " " & _
'''           "ORDER BY " & szOrderBy & ";"
'''            szSQL = "Select SageAccountNumber, CompanyName, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) as UnitName ," & _
'''                        " '' AS Balance,IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes,Property.PropertyID, Property.ClientID " & _
'''             "FROM (  Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
'''             "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
'''             "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false) as IM ORDER BY SageAccountNumber"
''            szSQL = "SELECT IM.SageAccountNumber, IM.CompanyName, (Select UnitName from units where IM.Unitnumber=units.Unitnumber) AS UnitName, IIF(ISNULL(Tenants.Comments),'CURRENT','DELETED') AS Notes, " & _
''             "Property.propertyID , Property.clientID FROM (Property INNER JOIN ([Select  S.SageAccountNumber,S.CompanyName,S.UnitNumber  from LeaseDetails as S " & _
''           "INNER JOIN (SELECT Max(LeaseDetails.StartDate) AS  MaxOfStartDate, LeaseDetails.SageAccountNumber FROM LeaseDetails " & _
''             "GROUP BY LeaseDetails.SageAccountNumber) as IQ ON IQ.Sageaccountnumber=S.sageaccountnumber  where  IQ.MaxOfStartDate=S.StartDate AND S.status=false]. AS IM INNER JOIN Units " & _
''             "ON IM.UnitNumber = Units.UnitNumber) ON Property.PropertyID = Units.PropertyID) INNER JOIN Tenants ON IM.SageAccountNumber = Tenants.SageAccountNumber " & _
''            "ORDER BY IM.SageAccountNumber;"
''
''   End If
''
'''   Debug.Print szSQL

