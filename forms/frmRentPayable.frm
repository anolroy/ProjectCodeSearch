VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRentPayable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Rent Payable"
   ClientHeight    =   13020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   22245
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
   Icon            =   "frmRentPayable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13020
   ScaleWidth      =   22245
   Begin VB.PictureBox picClientList 
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
      Height          =   4140
      Left            =   8640
      ScaleHeight     =   4110
      ScaleWidth      =   5565
      TabIndex        =   1
      Top             =   8325
      Visible         =   0   'False
      Width           =   5595
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
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   3480
         Left            =   45
         TabIndex        =   6
         Top             =   585
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   6138
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   8
         Top             =   315
         Width           =   1755
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3096;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1845
         TabIndex        =   7
         Top             =   315
         Width           =   3645
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6429;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   90
         Width           =   180
         VariousPropertyBits=   276824083
         Caption         =   "ID"
         Size            =   "317;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   585
         Index           =   1
         Left            =   1860
         TabIndex        =   4
         Top             =   90
         Width           =   1545
         VariousPropertyBits=   276824083
         Caption         =   "Client Name"
         Size            =   "2725;1032"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   795
         VariousPropertyBits=   276824083
         Caption         =   "Post Code"
         Size            =   "1402;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   240
         Index           =   0
         Left            =   5400
         Top             =   1530
         Width           =   5175
      End
   End
   Begin TabDlg.SSTab tabFees 
      Height          =   11325
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   22080
      _ExtentX        =   38947
      _ExtentY        =   19976
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Rent Payable"
      TabPicture(0)   =   "frmRentPayable.frx":1202
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Rent Payable History"
      TabPicture(1)   =   "frmRentPayable.frx":121E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   11040
         Left            =   -74955
         TabIndex        =   139
         Top             =   315
         Width           =   21885
         Begin VB.CommandButton cmdReverseHistory 
            Caption         =   "Reverse History"
            Height          =   495
            Left            =   225
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   10170
            Width           =   1485
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayFeesHistory 
            Height          =   9825
            Left            =   90
            TabIndex        =   140
            Top             =   225
            Width           =   21615
            _ExtentX        =   38126
            _ExtentY        =   17330
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
            BackColorUnpopulated=   -2147483643
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
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   10995
         Left            =   0
         TabIndex        =   9
         Top             =   270
         Width           =   21930
         Begin VB.Frame Frame1 
            Caption         =   "Produce Rent Summary Statement"
            Height          =   10770
            Index           =   6
            Left            =   3735
            TabIndex        =   65
            Top             =   3330
            Visible         =   0   'False
            Width           =   14055
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   0
               Left            =   12510
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   10125
               Width           =   1200
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "&Save"
               Height          =   465
               Left            =   12555
               TabIndex        =   103
               Top             =   9495
               Width           =   1170
            End
            Begin VB.Frame Frame2 
               Caption         =   "Retention Details"
               Height          =   5100
               Left            =   3150
               TabIndex        =   96
               Top             =   7695
               Visible         =   0   'False
               Width           =   6270
               Begin VB.Frame Frame6 
                  Height          =   1230
                  Left            =   180
                  TabIndex        =   129
                  Top             =   1890
                  Width           =   5910
                  Begin VB.CommandButton cmdClose12 
                     Caption         =   "X"
                     Height          =   240
                     Left            =   5535
                     TabIndex        =   138
                     Top             =   90
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdAddToGrid 
                     Caption         =   "Add"
                     Height          =   375
                     Left            =   4500
                     TabIndex        =   137
                     Top             =   810
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option2 
                     Caption         =   "-"
                     Height          =   210
                     Left            =   4590
                     TabIndex        =   136
                     Top             =   450
                     Width           =   780
                  End
                  Begin VB.OptionButton Option1 
                     Caption         =   "+"
                     Height          =   210
                     Left            =   3960
                     TabIndex        =   135
                     Top             =   450
                     Value           =   -1  'True
                     Width           =   780
                  End
                  Begin VB.TextBox txtRetentionDescriptions 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000014&
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "dd/MM/yyyy"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2057
                        SubFormatType   =   3
                     EndProperty
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
                     Height          =   285
                     Left            =   1440
                     MaxLength       =   50
                     TabIndex        =   133
                     Top             =   405
                     Width           =   2115
                  End
                  Begin VB.TextBox txtRetensionAmount1 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000014&
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "dd/MM/yyyy"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2057
                        SubFormatType   =   3
                     EndProperty
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
                     Height          =   285
                     Left            =   315
                     MaxLength       =   10
                     TabIndex        =   130
                     Text            =   "0.00"
                     Top             =   405
                     Width           =   945
                  End
                  Begin VB.Label Label16 
                     AutoSize        =   -1  'True
                     Caption         =   "Sign"
                     Height          =   210
                     Left            =   3870
                     TabIndex        =   134
                     Top             =   135
                     Width           =   345
                  End
                  Begin VB.Label Label15 
                     AutoSize        =   -1  'True
                     Caption         =   "Descriptions"
                     Height          =   210
                     Left            =   1485
                     TabIndex        =   132
                     Top             =   135
                     Width           =   975
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "Amount"
                     Height          =   210
                     Left            =   405
                     TabIndex        =   131
                     Top             =   135
                     Width           =   660
                  End
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "X"
                  Height          =   240
                  Left            =   5850
                  TabIndex        =   97
                  Top             =   90
                  Width           =   375
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetails 
                  Height          =   4575
                  Left            =   45
                  TabIndex        =   98
                  Top             =   450
                  Width           =   6150
                  _ExtentX        =   10848
                  _ExtentY        =   8070
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483630
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
                  _Band(0).Cols   =   3
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.CommandButton cmdCalculateAvailableFund 
               Caption         =   "Calculate Available Fund"
               Height          =   465
               Left            =   8730
               TabIndex        =   95
               Top             =   9495
               Width           =   2385
            End
            Begin VB.Frame Frame1 
               Caption         =   "Rent Payable"
               Height          =   690
               Index           =   14
               Left            =   6885
               TabIndex        =   91
               Top             =   9405
               Width           =   1770
               Begin VB.TextBox txtRentPayable 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   3
                  EndProperty
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
                  Height          =   285
                  Left            =   45
                  MaxLength       =   10
                  TabIndex        =   92
                  Text            =   "0.00"
                  Top             =   225
                  Width           =   1485
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Available  Fund"
               Height          =   690
               Index           =   13
               Left            =   5220
               TabIndex        =   89
               Top             =   9405
               Width           =   1635
               Begin VB.TextBox txtAvailableFunds 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   3
                  EndProperty
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
                  Height          =   285
                  Left            =   45
                  MaxLength       =   10
                  TabIndex        =   90
                  Text            =   "0.00"
                  Top             =   225
                  Width           =   1485
               End
            End
            Begin VB.CommandButton cmdTestReport 
               Caption         =   "Report For Test Purpose"
               Height          =   420
               Left            =   8730
               TabIndex        =   88
               Top             =   10125
               Visible         =   0   'False
               Width           =   3705
            End
            Begin VB.CommandButton cmdClose1 
               Caption         =   "X"
               Height          =   240
               Left            =   13500
               TabIndex        =   87
               Top             =   270
               Width           =   375
            End
            Begin VB.Frame Frame1 
               Caption         =   "Clients:"
               Height          =   3780
               Index           =   12
               Left            =   135
               TabIndex        =   85
               Top             =   1170
               Width           =   6690
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
                  Height          =   3495
                  Left            =   135
                  TabIndex        =   86
                  Top             =   225
                  Width           =   6375
                  _ExtentX        =   11245
                  _ExtentY        =   6165
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483630
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
                  _Band(0).Cols   =   3
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.TextBox txtStatementDate1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   1485
               MaxLength       =   10
               TabIndex        =   84
               Text            =   "01/01/2000"
               Top             =   450
               Width           =   1575
            End
            Begin VB.TextBox txtLastStatementDate1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   4905
               MaxLength       =   10
               TabIndex        =   82
               Text            =   "01/01/2000"
               Top             =   450
               Width           =   1575
            End
            Begin VB.TextBox txtClientSearch 
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
               Left            =   1485
               MaxLength       =   10
               TabIndex        =   79
               Top             =   810
               Width           =   1575
            End
            Begin VB.CommandButton cmdOKInouts 
               Caption         =   "&Preview"
               Height          =   465
               Left            =   0
               TabIndex        =   78
               Top             =   0
               Width           =   1170
            End
            Begin VB.Frame Frame1 
               Caption         =   "Bank Accounts:"
               Height          =   4185
               Index           =   8
               Left            =   6930
               TabIndex        =   76
               Top             =   765
               Width           =   6690
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
                  Height          =   3810
                  Left            =   90
                  TabIndex        =   77
                  Top             =   300
                  Width           =   6450
                  _ExtentX        =   11377
                  _ExtentY        =   6720
                  _Version        =   393216
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   -2147483640
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  BackColorUnpopulated=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   8421504
                  WordWrap        =   -1  'True
                  GridLinesFixed  =   1
                  Appearance      =   0
                  BandDisplay     =   1
                  RowSizingMode   =   1
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
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Payable Types:"
               Height          =   2970
               Index           =   11
               Left            =   13635
               TabIndex        =   74
               Top             =   6525
               Visible         =   0   'False
               Width           =   6465
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayableTypes 
                  Height          =   2370
                  Left            =   90
                  TabIndex        =   75
                  Top             =   540
                  Width           =   6225
                  _ExtentX        =   10980
                  _ExtentY        =   4180
                  _Version        =   393216
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   -2147483640
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  BackColorUnpopulated=   -2147483643
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
                  _Band(0).Cols   =   2
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Funds:"
               Height          =   4455
               Index           =   10
               Left            =   6930
               TabIndex        =   71
               Top             =   4950
               Width           =   6465
               Begin VB.CheckBox chkInFunds 
                  Caption         =   "All Funds"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   72
                  Top             =   270
                  Width           =   1095
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInFunds 
                  Height          =   3765
                  Left            =   120
                  TabIndex        =   73
                  Top             =   570
                  Width           =   6225
                  _ExtentX        =   10980
                  _ExtentY        =   6641
                  _Version        =   393216
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   -2147483640
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  BackColorUnpopulated=   -2147483643
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
                  _Band(0).Cols   =   2
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Retention"
               Height          =   690
               Index           =   7
               Left            =   135
               TabIndex        =   69
               Top             =   9405
               Width           =   4785
               Begin VB.CommandButton Command3 
                  Caption         =   "-"
                  Height          =   375
                  Left            =   4185
                  TabIndex        =   94
                  Top             =   225
                  Width           =   510
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Retentions"
                  Height          =   375
                  Left            =   90
                  TabIndex        =   93
                  Top             =   225
                  Width           =   2310
               End
               Begin VB.TextBox txtRetention 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   3
                  EndProperty
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
                  Height          =   285
                  Left            =   2610
                  MaxLength       =   10
                  TabIndex        =   70
                  Text            =   "0.00"
                  Top             =   225
                  Width           =   1485
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Properties:"
               Height          =   4455
               Index           =   9
               Left            =   135
               TabIndex        =   66
               Top             =   4950
               Width           =   6690
               Begin VB.CheckBox chkAllProperties 
                  Caption         =   "All Properties"
                  Height          =   255
                  Left            =   135
                  TabIndex        =   67
                  Top             =   240
                  Width           =   2025
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
                  Height          =   3810
                  Left            =   135
                  TabIndex        =   68
                  Top             =   495
                  Width           =   6405
                  _ExtentX        =   11298
                  _ExtentY        =   6720
                  _Version        =   393216
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   -2147483640
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  BackColorUnpopulated=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   8421504
                  WordWrap        =   -1  'True
                  GridLinesFixed  =   1
                  Appearance      =   0
                  BandDisplay     =   1
                  RowSizingMode   =   1
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
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.Shape Shape2 
               Height          =   10500
               Left            =   45
               Top             =   225
               Width           =   13920
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Statement Date"
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
               Left            =   225
               TabIndex        =   83
               Top             =   450
               Width           =   1110
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last Statement Date"
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
               Left            =   3285
               TabIndex        =   81
               Top             =   450
               Width           =   1425
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Client (Search)"
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
               Left            =   225
               TabIndex        =   80
               Top             =   810
               Width           =   1035
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Create PI from Statement"
            Height          =   10590
            Left            =   7245
            TabIndex        =   99
            Top             =   8415
            Visible         =   0   'False
            Width           =   13020
            Begin VB.Frame Frame5 
               Caption         =   "Fund List"
               Height          =   3615
               Left            =   450
               TabIndex        =   125
               Top             =   5400
               Visible         =   0   'False
               Width           =   3885
               Begin VB.CommandButton cmdFrameFundClose 
                  Caption         =   "X"
                  Height          =   240
                  Left            =   3510
                  TabIndex        =   126
                  Top             =   135
                  Width           =   375
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFundList 
                  Height          =   3225
                  Left            =   45
                  TabIndex        =   127
                  Top             =   315
                  Width           =   3720
                  _ExtentX        =   6562
                  _ExtentY        =   5689
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  BackColorSel    =   15329508
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   16777215
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483630
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
                  _Band(0).Cols   =   3
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
            End
            Begin VB.CommandButton cmdFundListForCreatePI 
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
               Left            =   7245
               TabIndex        =   124
               Top             =   1575
               Width           =   315
            End
            Begin VB.TextBox txtFundForPI 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   5715
               MaxLength       =   10
               TabIndex        =   123
               Top             =   1575
               Width           =   1485
            End
            Begin VB.TextBox txtStatementNumber2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   2025
               MaxLength       =   10
               TabIndex        =   121
               Top             =   630
               Width           =   1485
            End
            Begin VB.TextBox txtStatementDate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   5715
               MaxLength       =   10
               TabIndex        =   120
               Text            =   "01/01/2000"
               Top             =   720
               Width           =   1485
            End
            Begin VB.TextBox txtStatementBalance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   9450
               MaxLength       =   10
               TabIndex        =   119
               Text            =   "0.00"
               Top             =   720
               Width           =   1485
            End
            Begin VB.TextBox txtRetentions2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   2025
               MaxLength       =   10
               TabIndex        =   118
               Text            =   "0.00"
               Top             =   1080
               Width           =   1485
            End
            Begin VB.TextBox txtPayableDate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   2025
               MaxLength       =   10
               TabIndex        =   117
               Text            =   "01/01/2000"
               Top             =   1575
               Width           =   1485
            End
            Begin VB.TextBox txtAvailableFund1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   5715
               MaxLength       =   10
               TabIndex        =   111
               Text            =   "0.00"
               Top             =   1125
               Width           =   1485
            End
            Begin VB.TextBox txtRentPayable1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
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
               Height          =   285
               Left            =   9450
               MaxLength       =   10
               TabIndex        =   110
               Text            =   "0.00"
               Top             =   1170
               Width           =   1485
            End
            Begin VB.CommandButton cmdCalculateavailableFund1 
               Caption         =   "Recalculate Statement"
               Height          =   735
               Left            =   4545
               TabIndex        =   102
               Top             =   6795
               Width           =   2745
            End
            Begin VB.CommandButton cmdCreatePI 
               Caption         =   "Generate Rent Payable"
               Height          =   735
               Left            =   8325
               TabIndex        =   101
               Top             =   6705
               Width           =   2250
            End
            Begin VB.CommandButton Command5 
               Caption         =   "X"
               Height          =   240
               Left            =   12510
               TabIndex        =   100
               Top             =   180
               Width           =   375
            End
            Begin MSForms.Label Label13 
               Height          =   210
               Left            =   3960
               TabIndex        =   122
               Top             =   1620
               Width           =   420
               VariousPropertyBits=   276824091
               Caption         =   "Fund"
               Size            =   "741;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label12 
               Height          =   210
               Left            =   225
               TabIndex        =   116
               Top             =   1575
               Width           =   1020
               VariousPropertyBits=   276824091
               Caption         =   "Payable date"
               Size            =   "1799;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label11 
               Height          =   210
               Left            =   225
               TabIndex        =   115
               Top             =   1125
               Width           =   780
               VariousPropertyBits=   276824091
               Caption         =   "Retention"
               Size            =   "1376;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label10 
               Height          =   210
               Left            =   7650
               TabIndex        =   114
               Top             =   720
               Width           =   1485
               VariousPropertyBits=   276824091
               Caption         =   "Statement Balance"
               Size            =   "2619;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label9 
               Height          =   210
               Left            =   3915
               TabIndex        =   113
               Top             =   720
               Width           =   1230
               VariousPropertyBits=   276824091
               Caption         =   "Statement Date"
               Size            =   "2170;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label8 
               Height          =   210
               Left            =   225
               TabIndex        =   112
               Top             =   630
               Width           =   1500
               VariousPropertyBits=   276824091
               Caption         =   "Statement number "
               Size            =   "2646;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label6 
               Height          =   210
               Left            =   7605
               TabIndex        =   109
               Top             =   1170
               Width           =   1110
               VariousPropertyBits=   276824091
               Caption         =   "Rent Payables"
               Size            =   "1958;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label4 
               Height          =   210
               Left            =   3915
               TabIndex        =   108
               Top             =   1170
               Width           =   1260
               VariousPropertyBits=   276824091
               Caption         =   "Available Funds"
               Size            =   "2222;370"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   20340
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   10440
            Width           =   1200
         End
         Begin VB.Frame Frame1 
            Height          =   3120
            Index           =   5
            Left            =   2025
            TabIndex        =   49
            Top             =   7785
            Width           =   19815
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetailsTransaction 
               Height          =   2355
               Left            =   135
               TabIndex        =   62
               Top             =   270
               Width           =   19545
               _ExtentX        =   34475
               _ExtentY        =   4154
               _Version        =   393216
               FixedCols       =   0
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483640
               BackColorSel    =   12648447
               ForeColorSel    =   -2147483640
               BackColorBkg    =   16777215
               BackColorUnpopulated=   -2147483643
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
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Paid/Received Credit"
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
               Left            =   14040
               TabIndex        =   60
               Top             =   405
               Width           =   1500
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Paid/Received Debit"
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
               Left            =   12285
               TabIndex        =   59
               Top             =   405
               Width           =   1455
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Balance (OS)"
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
               Left            =   10440
               TabIndex        =   58
               Top             =   405
               Width           =   870
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Original Amount"
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
               Left            =   9000
               TabIndex        =   57
               Top             =   405
               Width           =   1155
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
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
               Index           =   17
               Left            =   225
               TabIndex        =   56
               Top             =   405
               Width           =   780
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Property Name"
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
               Left            =   1350
               TabIndex        =   55
               Top             =   405
               Width           =   1050
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Code"
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
               Left            =   2790
               TabIndex        =   54
               Top             =   405
               Width           =   735
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Transaction Date"
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
               Left            =   4995
               TabIndex        =   53
               Top             =   405
               Width           =   1425
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Reference"
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
               Left            =   6480
               TabIndex        =   52
               Top             =   405
               Width           =   720
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
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
               Index           =   12
               Left            =   7560
               TabIndex        =   51
               Top             =   405
               Width           =   840
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Transaction No"
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
               Left            =   3780
               TabIndex        =   50
               Top             =   405
               Width           =   1065
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000000&
               Height          =   435
               Index           =   3
               Left            =   90
               Top             =   270
               Width           =   15855
            End
         End
         Begin VB.Frame Frame1 
            Height          =   7035
            Index           =   4
            Left            =   2025
            TabIndex        =   31
            Top             =   765
            Width           =   19815
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayFees 
               Height          =   6810
               Left            =   90
               TabIndex        =   61
               Top             =   135
               Width           =   19635
               _ExtentX        =   34634
               _ExtentY        =   12012
               _Version        =   393216
               FixedCols       =   0
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483640
               BackColorSel    =   12648447
               ForeColorSel    =   -2147483640
               BackColorBkg    =   16777215
               BackColorUnpopulated=   -2147483643
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
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.CommandButton cmdRecharge 
               Caption         =   "&Recharge"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   12375
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   4455
               Width           =   1335
            End
            Begin VB.CommandButton cmdPostRP 
               Caption         =   "&Post"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   9630
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   4680
               Width           =   1335
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Edit"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   11115
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   4995
               Width           =   1335
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   10890
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   4545
               Width           =   1335
            End
            Begin VB.CommandButton cmdCreate 
               Caption         =   "C&reate"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   12375
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   4950
               Width           =   1335
            End
            Begin VB.CommandButton cmdRechargeGenerate 
               Caption         =   "&Generate Rent Payable"
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
               Left            =   7605
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   4140
               Width           =   2415
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Prev. Statement Date"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   92
               Left            =   3645
               TabIndex        =   64
               Top             =   540
               Width           =   1200
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice No."
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   98
               Left            =   12420
               TabIndex        =   63
               Top             =   540
               Width           =   885
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Invoiced"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   101
               Left            =   14805
               TabIndex        =   48
               Top             =   540
               Width           =   885
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Statement Date"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   93
               Left            =   4950
               TabIndex        =   46
               Top             =   540
               Width           =   1200
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Emailed"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   100
               Left            =   14130
               TabIndex        =   45
               Top             =   540
               Width           =   615
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Payable Amount"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   97
               Left            =   11205
               TabIndex        =   44
               Top             =   540
               Width           =   1155
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Printed"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   99
               Left            =   13410
               TabIndex        =   43
               Top             =   540
               Width           =   690
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Available Funds"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   96
               Left            =   9630
               TabIndex        =   42
               Top             =   540
               Width           =   1560
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Retentions"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   95
               Left            =   8550
               TabIndex        =   41
               Top             =   540
               Width           =   960
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Statement Closing Balance"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   94
               Left            =   6345
               TabIndex        =   40
               Top             =   540
               Width           =   2595
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Code"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   91
               Left            =   2700
               TabIndex        =   39
               Top             =   540
               Width           =   1110
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Client/Landlord"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   90
               Left            =   1485
               TabIndex        =   38
               Top             =   540
               Width           =   1305
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Statement ID"
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
               Index           =   89
               Left            =   270
               TabIndex        =   37
               Top             =   540
               Width           =   930
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000000&
               Height          =   435
               Index           =   1
               Left            =   45
               Top             =   360
               Width           =   15945
            End
         End
         Begin VB.CommandButton cmdClient 
            Caption         =   " .."
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
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   435
            Width           =   345
         End
         Begin VB.TextBox txtClientID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3285
            TabIndex        =   25
            Top             =   435
            Width           =   3255
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
            Height          =   2925
            Index           =   1
            Left            =   45
            TabIndex        =   22
            Top             =   3375
            Width           =   1935
            Begin VB.CommandButton cmdGenerateRentPayable 
               Caption         =   "Generate Rent Payable"
               Height          =   855
               Left            =   200
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   450
               Width           =   1485
            End
            Begin VB.CommandButton cmdPosttoHistory 
               Caption         =   "Post to Hist."
               Height          =   495
               Left            =   200
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1395
               Width           =   1485
            End
            Begin MSForms.CommandButton cmdReverseRentPayable 
               Height          =   765
               Left            =   225
               TabIndex        =   142
               Top             =   2070
               Width           =   1470
               VariousPropertyBits=   8388635
               Caption         =   "Reverse Rent Payable"
               Size            =   "2593;1349"
               FontName        =   "Myriad Web"
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label lblEditDemand 
               Alignment       =   2  'Center
               BackColor       =   &H00E5E5E5&
               BackStyle       =   0  'Transparent
               Caption         =   "Generate Rent"
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
               Height          =   255
               Left            =   135
               MousePointer    =   99  'Custom
               TabIndex        =   24
               Top             =   165
               Width           =   1155
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
            Height          =   2685
            Index           =   2
            Left            =   45
            TabIndex        =   18
            Top             =   6255
            Width           =   1935
            Begin VB.CheckBox chkShowDue 
               Caption         =   "Incl. Mngt Fees Due"
               Enabled         =   0   'False
               Height          =   210
               Left            =   135
               TabIndex        =   146
               Top             =   765
               UseMaskColor    =   -1  'True
               Width           =   1995
            End
            Begin VB.CheckBox chkExcludeSupOS 
               Caption         =   "Incl. Supplier OS"
               Enabled         =   0   'False
               Height          =   210
               Left            =   135
               TabIndex        =   145
               Top             =   450
               UseMaskColor    =   -1  'True
               Width           =   1590
            End
            Begin VB.CommandButton cmdPrintClientStatement 
               Caption         =   "Selected Statement"
               Height          =   495
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   144
               Top             =   990
               Width           =   1485
            End
            Begin VB.CommandButton cmdFix 
               Caption         =   "Fix"
               Height          =   405
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   2160
               Width           =   1485
            End
            Begin VB.CommandButton cmdPrintAll 
               Caption         =   "All statement"
               Height          =   540
               Left            =   165
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1560
               Width           =   1485
            End
            Begin VB.CommandButton cmdPrintThis 
               Caption         =   "Selected Statement"
               Height          =   495
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   2160
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label Label19 
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
               Height          =   255
               Index           =   35
               Left            =   225
               MousePointer    =   99  'Custom
               TabIndex        =   21
               Top             =   135
               Width           =   1155
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1980
            Index           =   3
            Left            =   45
            TabIndex        =   13
            Top             =   8865
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "delete All Statements"
               Height          =   735
               Left            =   225
               TabIndex        =   128
               Top             =   2025
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "Sea&rch"
               Height          =   375
               Left            =   225
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   1530
               Width           =   1485
            End
            Begin VB.CommandButton cmdEmailDmds 
               Caption         =   "Statement"
               Height          =   405
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   540
               Width           =   1485
            End
            Begin VB.CommandButton cmdArchive 
               Caption         =   "Archive"
               Height          =   360
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H00E5E5E5&
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
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
               Height          =   210
               Index           =   5
               Left            =   270
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   225
               Width           =   1155
            End
         End
         Begin VB.Frame Frame1 
            Height          =   3210
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   1950
            Begin VB.CommandButton cmdPreViewGenDmds 
               Caption         =   "Finalise/Modify Selected Statement"
               Height          =   810
               Left            =   270
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   2115
               Width           =   1350
            End
            Begin VB.CommandButton cmdProduceClientSummaryStatement 
               Caption         =   "Produce New Client Statement"
               Height          =   990
               Left            =   270
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   855
               Width           =   1350
            End
            Begin VB.Label lblGenerate 
               Alignment       =   2  'Center
               BackColor       =   &H00E5E5E5&
               BackStyle       =   0  'Transparent
               Caption         =   "Produce Rent Summary Statement"
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
               Left            =   90
               MousePointer    =   99  'Custom
               TabIndex        =   12
               Top             =   270
               Width           =   1590
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Client ID:"
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
            Left            =   2115
            TabIndex        =   30
            Top             =   495
            Width           =   660
         End
         Begin MSForms.Label Label3 
            Height          =   195
            Left            =   6960
            TabIndex        =   29
            Top             =   450
            Width           =   795
            VariousPropertyBits=   276824083
            Caption         =   "Properties:"
            Size            =   "1402;344"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboProperty 
            Height          =   285
            Left            =   7995
            TabIndex        =   28
            Top             =   405
            Width           =   3495
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6165;503"
            TextColumn      =   2
            ColumnCount     =   2
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1234"
         End
         Begin MSForms.CheckBox chkONDD 
            Height          =   285
            Left            =   11640
            TabIndex        =   27
            ToolTipText     =   "Override Next Due Date"
            Top             =   450
            Width           =   855
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1508;503"
            Value           =   "0"
            Caption         =   "ONDD"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   330
      Left            =   13230
      Top             =   9135
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc adoProperty 
      Height          =   330
      Left            =   11745
      Top             =   9990
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Property"
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
End
Attribute VB_Name = "frmRentPayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NO_AGENT_INFO As Boolean

Private szProertyID As String
Private bData As Boolean
Private szSupplierAccount As String
'Private lLastID As Long
Private szaFreq() As String
Private szaTrans(1) As Integer
Private szAgentVATCode As String
Dim szPayableTypes As String
Dim szSelectedClient As String
Dim szSelectedBankAccount As String
Dim bPreviewMode As Boolean
Dim szSelectedStatement As String
Dim szSelectedFund As String
Dim szCurrentStatementID As String
Dim szAvailableFund1 As Double
Dim whichFieldToCheck As String
Dim hasSelProperty As Boolean
Dim hasSelBankAccounts As Boolean
Dim szCurrentStatementHistoryID As String
'Dim szSelectedPayableTypeID As String
'Dim szCurrentRentsummarySTID As String
Private Function ListOfFundsForDBSave() As String
    Dim rsStatement As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    rsStatement.Open "Select * from RentSummaryStatement R,Client C where R.ClientIDLandlordID=C.ClientID AND StatementID=" & _
    Replace(szCurrentStatementID, "SS", "") & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsStatement.EOF Then
           ListOfFundsForDBSave = rsStatement("ListOfFundID").Value
    End If
    rsStatement.Close
    Set rsStatement = Nothing
    adoConn.Close
End Function
Private Sub LoadRentSummaryDetails()
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim i As Integer
    Dim rsDetailsTransaction As New ADODB.Recordset
    Call configflxDetailsTransaction
    adoConn.Open getConnectionString
    szSQL = "Select UnitID as propertyID,SageAccountNumber, P.TransactionID,SP.Type,MID(T.CONSTANT, 4,3) & P.SlNumber AS TRANID,ExtRef,PDate,Details, " & _
            "BankCode,S.Amount,S.OSAmount,SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount) as DebitCredit from " & _
            "tlbPayment P ,tlbPaymentSplit S ,tlbTransactionTypes T ,Supplier SP where S.Payheader=P.TransactionID AND Sp.SupplierID=P.ClientID AND P.Type=T.TYPE_ID AND S.ClientStatementID=" & szCurrentStatementID & ""
'PropertyID
'AccountNo
'Account Type
'Transaction No
'Transaction Type
'Transaction Date
'Reference
'description
'amount
'balance
'Paid/Received Debit
'Paid/Received Credit
'szHeader$ = "|<|<PropertyID|<AccountNo|<Account Type |<Transaction No|<Transaction|<TypeTransaction |<DateReference|<Description|<Amount|<Balance|<Paid/Received Debit|<Paid/Received Credit"
    i = 1
    rsDetailsTransaction.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsDetailsTransaction.EOF
        flxDetailsTransaction.TextMatrix(i, 1) = IIf(IsNull(rsDetailsTransaction("propertyID").Value), "", rsDetailsTransaction("propertyID").Value) 'propertyID
        flxDetailsTransaction.TextMatrix(i, 2) = rsDetailsTransaction("SageAccountNumber").Value 'AccountNo
        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TRANID").Value 'TransactionID
        flxDetailsTransaction.TextMatrix(i, 4) = rsDetailsTransaction("Type").Value 'Account Type/This is supplier type
        flxDetailsTransaction.TextMatrix(i, 5) = rsDetailsTransaction("TRANID").Value
        flxDetailsTransaction.TextMatrix(i, 6) = rsDetailsTransaction("PDate").Value
        flxDetailsTransaction.TextMatrix(i, 7) = IIf(IsNull(rsDetailsTransaction("ExtRef").Value), "", rsDetailsTransaction("ExtRef").Value)
        flxDetailsTransaction.TextMatrix(i, 8) = IIf(IsNull(rsDetailsTransaction("Details").Value), "", rsDetailsTransaction("Details").Value)
        flxDetailsTransaction.TextMatrix(i, 9) = Format(rsDetailsTransaction("Amount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 10) = Format(rsDetailsTransaction("OSAmount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 11) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") > 0, Format(rsDetailsTransaction("DebitCredit").Value, "0.00"), 0)
        flxDetailsTransaction.TextMatrix(i, 12) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") < 0, Abs(Format(rsDetailsTransaction("DebitCredit").Value, "0.00")), 0)

        
        
'        flxDetailsTransaction.TextMatrix(i, 1) = rsDetailsTransaction("TransactionID").Value
'        flxDetailsTransaction.TextMatrix(i, 2) = "Payment"
'        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TRANID").Value
'        flxDetailsTransaction.TextMatrix(i, 4) = rsDetailsTransaction("Details").Value
'        flxDetailsTransaction.TextMatrix(i, 5) = IIf(IsNull(rsDetailsTransaction("BankCode").Value), "", rsDetailsTransaction("BankCode").Value)
'        flxDetailsTransaction.TextMatrix(i, 6) = Format(rsDetailsTransaction("Amount").Value, "0.00")
'        flxDetailsTransaction.TextMatrix(i, 7) = rsDetailsTransaction("Type").Value
        flxDetailsTransaction.AddItem ""
        i = i + 1
        rsDetailsTransaction.MoveNext
    Wend
    rsDetailsTransaction.Close
    Set rsDetailsTransaction = Nothing
    
    
         szSQL = "Select U.PropertyID,R.SageAccountNumber,R.TransactionID,R.Type,RIGHT(T.CONSTANT, 2)  & R.SlNumber as TRANID,RDate ,ExtRef,Details  ,BankCode,S.Amount,S.OSAmount " & _
         ",SWITCH(R.TYPE=23,-S.Amount,R.TYPE=1,-S.Amount,R.TYPE=2,S.Amount,R.TYPE=3,S.Amount,R.TYPE=4,S.Amount) as DebitCredit from tlbReceipt R ,tlbReceiptSplit S ,tlbTransactionTypes T,UnitS U where " & _
         "U.UnitNumber=R.UnitID AND R.TransactionID=S.rptHeader AND R.Type=T.TYPE_ID AND S.ClientStatementID=" & szCurrentStatementID & ""
         
    rsDetailsTransaction.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsDetailsTransaction.EOF
        flxDetailsTransaction.TextMatrix(i, 1) = rsDetailsTransaction("propertyID").Value
        flxDetailsTransaction.TextMatrix(i, 2) = rsDetailsTransaction("SageAccountNumber").Value
        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TransactionID").Value
        flxDetailsTransaction.TextMatrix(i, 4) = "Lessee" 'rsDetailsTransaction("Type").Value
        flxDetailsTransaction.TextMatrix(i, 5) = rsDetailsTransaction("TRANID").Value
        flxDetailsTransaction.TextMatrix(i, 6) = rsDetailsTransaction("RDate").Value
         flxDetailsTransaction.TextMatrix(i, 7) = IIf(IsNull(rsDetailsTransaction("ExtRef").Value), "", rsDetailsTransaction("ExtRef").Value)
        flxDetailsTransaction.TextMatrix(i, 8) = IIf(IsNull(rsDetailsTransaction("Details").Value), "", rsDetailsTransaction("Details").Value)
        flxDetailsTransaction.TextMatrix(i, 9) = Format(rsDetailsTransaction("Amount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 10) = Format(rsDetailsTransaction("OSAmount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 11) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") > 0, Format(rsDetailsTransaction("DebitCredit").Value, "0.00"), 0)
        flxDetailsTransaction.TextMatrix(i, 12) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") < 0, Abs(Format(rsDetailsTransaction("DebitCredit").Value, "0.00")), 0)
        flxDetailsTransaction.AddItem ""
        i = i + 1
        rsDetailsTransaction.MoveNext
    Wend
    rsDetailsTransaction.Close
    Set rsDetailsTransaction = Nothing
    
    szSQL = "Select PropertyID,MY_ID as TransactionID,BANK_AC as SageAccountNumber,TRANS,TRANS & B.TRAN_ID as TRID,TRAN_DATE,PROJ_REF,DESCRIPTION as Details ,BANK_AC as BankCode, NET_AMOUNT as Amount,0 as OSAmount,SWITCH(TRANS='BP',Amount,TRANS='BR',-amount) as DebitCredit from tlbBankPayment B " & _
         "where RentSumStatement='" & szCurrentStatementID & "'"
         'Select MY_ID as TransactionID,TRANS & B.TRAN_ID,DESCRIPTION as Details ,BANK_AC , NET_AMOUNT as Amount from tlbBankPayment B where RentSumStatement='1'
    rsDetailsTransaction.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsDetailsTransaction.EOF
'        flxDetailsTransaction.TextMatrix(i, 1) = rsDetailsTransaction("TransactionID").Value
'        flxDetailsTransaction.TextMatrix(i, 2) = "Bank Payment"
'        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TRID").Value
'        flxDetailsTransaction.TextMatrix(i, 4) = rsDetailsTransaction("Details").Value
'        flxDetailsTransaction.TextMatrix(i, 5) = IIf(IsNull(rsDetailsTransaction("BankCode").Value), "", rsDetailsTransaction("BankCode").Value)
'        flxDetailsTransaction.TextMatrix(i, 6) = Format(rsDetailsTransaction("Amount").Value, "0.00")
'        flxDetailsTransaction.TextMatrix(i, 7) = IIf(rsDetailsTransaction("TRANS").Value = "BP", 11, 12)
        flxDetailsTransaction.TextMatrix(i, 1) = rsDetailsTransaction("propertyID").Value
        flxDetailsTransaction.TextMatrix(i, 2) = rsDetailsTransaction("SageAccountNumber").Value
        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TransactionID").Value
        flxDetailsTransaction.TextMatrix(i, 4) = IIf(rsDetailsTransaction("TRANS").Value = "BP", "Bank Payment", "Bank Receipt")
        flxDetailsTransaction.TextMatrix(i, 5) = rsDetailsTransaction("TRAN_DATE").Value
        flxDetailsTransaction.TextMatrix(i, 6) = rsDetailsTransaction("TRID").Value
        flxDetailsTransaction.TextMatrix(i, 7) = IIf(IsNull(rsDetailsTransaction("PROJ_REF").Value), "", rsDetailsTransaction("PROJ_REF").Value)
        flxDetailsTransaction.TextMatrix(i, 8) = IIf(IsNull(rsDetailsTransaction("Details").Value), "", rsDetailsTransaction("Details").Value)
        flxDetailsTransaction.TextMatrix(i, 9) = Format(rsDetailsTransaction("Amount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 10) = Format(rsDetailsTransaction("OSAmount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 11) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") > 0, Format(rsDetailsTransaction("DebitCredit").Value, "0.00"), 0)
        flxDetailsTransaction.TextMatrix(i, 12) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") < 0, Abs(Format(rsDetailsTransaction("DebitCredit").Value, "0.00")), 0)
        flxDetailsTransaction.AddItem ""
        i = i + 1
        rsDetailsTransaction.MoveNext
    Wend
    rsDetailsTransaction.Close
    Set rsDetailsTransaction = Nothing
    
    
    
    If chkExcludeSupOS.Value = 1 Then
    '********wrting code for showing in grid// additional data from snapshot table**************
        szSQL = "select StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID, " & _
                "NOMINAL_CODE,NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef from ClientStatementPurchasesSnapshot  where StatementID=" & _
                szCurrentStatementID & " and isManagementFee=false"

'szHeader$ = "|<|<PropertyID|<AccountNo|<Account Type |<Transaction No|<Transaction|<TypeTransaction |<DateReference|<Description|<Amount|<Balance|<Paid/Received Debit|<Paid/Received Credit"
 
    rsDetailsTransaction.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsDetailsTransaction.EOF
        flxDetailsTransaction.TextMatrix(i, 2) = IIf(IsNull(rsDetailsTransaction("propertyID").Value), "", rsDetailsTransaction("propertyID").Value) 'propertyID
        flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("NOMINAL_CODE").Value 'AccountNo
       ' flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TransactionID").Value 'TransactionID
        flxDetailsTransaction.TextMatrix(i, 4) = "PI Invoice" 'rsDetailsTransaction("Type").Value 'Account Type/This is supplier type
        flxDetailsTransaction.TextMatrix(i, 7) = "snapshot" 'rsDetailsTransaction("TransactionID").Value
        flxDetailsTransaction.TextMatrix(i, 6) = rsDetailsTransaction("TranDate").Value
        flxDetailsTransaction.TextMatrix(i, 5) = IIf(IsNull(rsDetailsTransaction("PaymentRef").Value), "", rsDetailsTransaction("PaymentRef").Value)
        'flxDetailsTransaction.TextMatrix(i, 8) = IIf(IsNull(rsDetailsTransaction("Details").Value), "", rsDetailsTransaction("Details").Value)
        flxDetailsTransaction.TextMatrix(i, 8) = Format(rsDetailsTransaction("PaymentAmount").Value, "0.00")
        flxDetailsTransaction.TextMatrix(i, 9) = Format(rsDetailsTransaction("osAmount").Value, "0.00")
'        flxDetailsTransaction.TextMatrix(i, 11) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") > 0, Format(rsDetailsTransaction("DebitCredit").Value, "0.00"), 0)
'        flxDetailsTransaction.TextMatrix(i, 12) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") < 0, Abs(Format(rsDetailsTransaction("DebitCredit").Value, "0.00")), 0)
        flxDetailsTransaction.AddItem ""
        i = i + 1
        rsDetailsTransaction.MoveNext
    Wend
    rsDetailsTransaction.Close
    Set rsDetailsTransaction = Nothing
    End If
    
    If chkShowDue.Value = 1 Then
    '********wrting code for showing in grid// additional data from snapshot table**************
        szSQL = "select StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID, " & _
                            "NOMINAL_CODE,NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef from ClientStatementPurchasesSnapshot  where StatementID=" & _
                            szCurrentStatementID & " and isManagementFee=true"
            
            'szHeader$ = "|<|<PropertyID|<AccountNo|<Account Type |<Transaction No|<Transaction|<TypeTransaction |<DateReference|<Description|<Amount|<Balance|<Paid/Received Debit|<Paid/Received Credit"
             
                rsDetailsTransaction.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                While Not rsDetailsTransaction.EOF
                    flxDetailsTransaction.TextMatrix(i, 2) = IIf(IsNull(rsDetailsTransaction("propertyID").Value), "", rsDetailsTransaction("propertyID").Value) 'propertyID
                    flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("NOMINAL_CODE").Value 'AccountNo
                   ' flxDetailsTransaction.TextMatrix(i, 3) = rsDetailsTransaction("TransactionID").Value 'TransactionID
                    flxDetailsTransaction.TextMatrix(i, 4) = "PI Invoice" 'rsDetailsTransaction("Type").Value 'Account Type/This is supplier type
                    flxDetailsTransaction.TextMatrix(i, 7) = "snapshot" 'rsDetailsTransaction("TransactionID").Value
                    flxDetailsTransaction.TextMatrix(i, 6) = rsDetailsTransaction("TranDate").Value
                    flxDetailsTransaction.TextMatrix(i, 5) = IIf(IsNull(rsDetailsTransaction("PaymentRef").Value), "", rsDetailsTransaction("PaymentRef").Value)
                    'flxDetailsTransaction.TextMatrix(i, 8) = IIf(IsNull(rsDetailsTransaction("Details").Value), "", rsDetailsTransaction("Details").Value)
                    flxDetailsTransaction.TextMatrix(i, 8) = Format(rsDetailsTransaction("PaymentAmount").Value, "0.00")
                    flxDetailsTransaction.TextMatrix(i, 9) = Format(rsDetailsTransaction("osAmount").Value, "0.00")
            '        flxDetailsTransaction.TextMatrix(i, 11) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") > 0, Format(rsDetailsTransaction("DebitCredit").Value, "0.00"), 0)
            '        flxDetailsTransaction.TextMatrix(i, 12) = IIf(Format(rsDetailsTransaction("DebitCredit").Value, "0.00") < 0, Abs(Format(rsDetailsTransaction("DebitCredit").Value, "0.00")), 0)
                    flxDetailsTransaction.AddItem ""
                    i = i + 1
                    rsDetailsTransaction.MoveNext
                Wend
                rsDetailsTransaction.Close
                Set rsDetailsTransaction = Nothing
    End If
          
    adoConn.Close
    Set adoConn = Nothing
    

End Sub
Private Sub configflxDetailsTransaction()
'    Dim szHeader As String
'    flxDetailsTransaction.Clear
'    szHeader$ = "|<|<Type|<Transaction ID|<Details|<BankCode|<Amount"
'    flxDetailsTransaction.FormatString = szHeader$
'    flxDetailsTransaction.Rows = 2
'    flxDetailsTransaction.Cols = 8
'    flxDetailsTransaction.ColWidth(0) = 350
'    flxDetailsTransaction.ColWidth(1) = 0
'    flxDetailsTransaction.ColWidth(2) = 2000 'Table Type
'    flxDetailsTransaction.ColWidth(3) = 3000
'    flxDetailsTransaction.ColWidth(4) = 2000
'    flxDetailsTransaction.ColWidth(5) = 3000
'    flxDetailsTransaction.ColWidth(6) = 2000
'    flxDetailsTransaction.ColWidth(7) = 0 '"transacation Type

    Dim szHeader As String
    flxDetailsTransaction.Clear
    szHeader$ = "|<|<PropertyID|<AccountNo|<Account Type |<Transaction No|<Transaction Date|<Reference|<Description|<Amount|<Balance|<Paid/Received Debit|<Paid/Received Credit"
    flxDetailsTransaction.FormatString = szHeader$
    flxDetailsTransaction.Rows = 2
    flxDetailsTransaction.Cols = 13
    flxDetailsTransaction.ColWidth(0) = 350
    flxDetailsTransaction.ColWidth(1) = 2000
    flxDetailsTransaction.ColWidth(2) = 2000 'Table Type
    flxDetailsTransaction.ColWidth(3) = 3000
    flxDetailsTransaction.ColWidth(4) = 2000
    flxDetailsTransaction.ColWidth(5) = 3000
    flxDetailsTransaction.ColWidth(6) = 2000
    flxDetailsTransaction.ColWidth(7) = 2000
    flxDetailsTransaction.ColWidth(8) = 2000
    flxDetailsTransaction.ColWidth(9) = 2000
    flxDetailsTransaction.ColWidth(10) = 2000
    flxDetailsTransaction.ColWidth(11) = 2000
    flxDetailsTransaction.ColWidth(12) = 2000
    
    
    
    

End Sub

Private Function getClosingBalance(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************


    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 255
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getClosingBalance = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "' "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -50
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getClosingBalance = getClosingBalance + dblAmt

    'd)  Add (+): Sum of Bank payments and receipts
   
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,B.NET_AMOUNT,TransactionType=12 ,-B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=cstr(F.FundID) " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
 End Function
Private Function getAvailablefunds(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
    'AND tlbReceipt.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND tlbReceipt.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and S.Amount>S.OSAmount and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getAvailablefunds = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "S.FundID=F.FundID and  S.Amount>S.OSAmount and P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -15
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt

    'd)  Add (+): Sum of Bank payments and receipts
   
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=cstr(F.FundID) " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    dblAmt = GetSupplierOSAmount
    'If negative then ignore this
    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
    'it should be -40

    Dim rsNLposting As New ADODB.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefunds = getAvailablefunds + GetClientACBalance + GetLandLordACBalance
        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
        dblAmt = GetLandLordACBalance
          'if dblAmt is negative then ignore
        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
     
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    dblAmt = GetAgentBalance
    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
    
    Debug.Print getAvailablefunds
    
    
    Dim ManagementFeeControl As String
    AccrualsControl = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    'Dim rsNLPosting As New ADODB.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & _
                    " NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
                    szSelectedClient & "' AND DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
    'j)  Less (-): Tenant Deposits received for the client selected
    'REMOVE THIS AS PER SPEC
    'getAvailablefunds = getAvailablefunds - GetRentDeposit
    
    getAvailablefunds = getAvailablefunds - Val(txtRetention.text)
    MsgBox "Available fund is: " & getAvailablefunds
    txtAvailableFunds.text = getAvailablefunds
    txtRentPayable.text = txtAvailableFunds.text
     
End Function
Private Sub GenerateSummaryStatement(szStatmentID As String)
    Frame1(6).Visible = True
    Frame1(6).Top = 135
    Frame1(6).Left = 2025
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & _
    szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
   
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    
    'szStatementID
    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where 1=2", adoConn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !statementID = szStatmentID 'we are setting this column atutomatically
             szCurrentStatementID = szStatmentID
            !statementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient = "", Null, GetLastStatementDateByClient) 'This is Fromdate
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
            !StatementOpBal = dblLasClosingBalance
            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
            !Clearretentions = False 'Will need to come again
            
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
'            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit()
            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
            !PaymentsonAccount = -GetPaymentsonAccount
            'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceipts
            !SupplierPayments = GetSupplierPayment
            !BankPaymentReceipts = GetBankPaymentReceipts
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            
            
            !PayableAmount = txtRentPayable.text
            !StatementClosingBal = getClosingBalance(dblLasClosingBalance)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostToHistory = False
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    Call SaveRetentionDetails(adoConn)
    Call loadflxPayFees("")
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub SaveRetentionDetails(adoConn As ADODB.Connection)
        Dim rsRetensionDetails As New ADODB.Recordset
        If szCurrentStatementID = "" Then Exit Sub
        adoConn.Execute "Delete from RetentionDetails where statementID =" & szCurrentStatementID & ""
        'Enter data into grid only memory version
        Dim iRow As Integer
        rsRetensionDetails.Open "Select * from RetentionDetails where 1=2", adoConn, adOpenDynamic, adLockOptimistic
        For iRow = 1 To flxRetensionDetails.Rows - 1
            With rsRetensionDetails
                    .AddNew
                    !statementID = szCurrentStatementID
                    !SlNumber = iRow
                    !description = flxRetensionDetails.TextMatrix(iRow, 3)
                    !amount = Val(flxRetensionDetails.TextMatrix(iRow, 4))
                    .Update
            End With
       Next iRow
        'then load everything to the grid
        Call loadflxRetensionDetails
End Sub
'Public Sub loadflxPayFees(strFilter As String)
'       Dim szSQL As String
'        Call configflxPayFees
'        Dim adoconn As New ADODB.Connection
'        Dim rsRentSummaryStatement As New ADODB.Recordset
'        Dim rsRentSummaryStatementSplit1 As New ADODB.Recordset
'        Dim dblPayableAmount As Double
'        Dim a As String
'        Dim b As String
'        Dim c As String
'        Dim d As String
'        Dim e As String
'        Dim f As String
'        Dim g As String
'        Dim h As String
'        Dim k As String
'        Dim l As String
'        Dim m As String
'        Dim n As String
'        Dim j, o, p, q, r, s, t, u As String
'        'Exit Sub
'
'        Dim i As Long
'        adoconn.Open getConnectionString
'            If strFilter = "" Then
'                    rsRentSummaryStatement.Open "Select '+' as SIGN,R.* from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
'                    "R.ClientIDLandlordID=S.SupplierID  Order by statementID DESC", adoconn, adOpenDynamic, adLockOptimistic
'            Else
'                    rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
'                    "R.ClientIDLandlordID=S.SupplierID " & strFilter & " Order by statementID DESC", adoconn, adOpenDynamic, adLockOptimistic
'            End If
'
'
''            If strFilter = "" Then
''                    rsRentSummaryStatement.Open "Select '+' as SIGN,R.* from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
''                    "R.ClientIDLandlordID=S.SupplierID " & _
''                    "UNION ALL" & _
''                    "SELECT * from RentSummaryStatementdetails"
''            Else
''                    rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
''                    "R.ClientIDLandlordID=S.SupplierID " & strFilter & " Order by statementID DESC", adoconn, adOpenDynamic, adLockOptimistic
''            End If
'
'            i = 1
'            With rsRentSummaryStatement
'            While Not rsRentSummaryStatement.EOF
'                flxPayFees.AddItem ""
'                flxPayFees.TextMatrix(i, 1) = "+" ' Expansion/collapse
'                flxPayFees.TextMatrix(i, 2) = "CS" & !statementID
'                flxPayFees.TextMatrix(i, 3) = !StatementNo 'statement no by client ID
'                flxPayFees.TextMatrix(i, 4) = !ClientIDLandlordID 'This naming should be only clientID its  mistake on spec
'                flxPayFees.TextMatrix(i, 5) = !BankCode
'                a = !BankCode
'                flxPayFees.TextMatrix(i, 6) = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
'                b = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
'                flxPayFees.TextMatrix(i, 7) = !StatementDate
'                c = !StatementDate
'                flxPayFees.TextMatrix(i, 8) = Format(!Availablefunds, "0.00")
'                d = !StatementDate
'                flxPayFees.TextMatrix(i, 9) = Format(!Retentions, "0.00")
'                e = Format(!Retentions, "0.00")
'                flxPayFees.TextMatrix(i, 10) = "" 'Keep it blank for master 'apportionment
'                flxPayFees.TextMatrix(i, 11) = "£" & Format(!PayableAmount, "0.00") 'Keep it blank for master 'Payable Paid amount
'                f = "£" & Format(!PayableAmount, "0.00")
'                dblPayableAmount = !PayableAmount
'                flxPayFees.TextMatrix(i, 12) = Format(!StatementClosingBal, "0.00")
'                g = Format(!StatementClosingBal, "0.00")
'                 flxPayFees.TextMatrix(i, 13) = Format((IIf(IsNull(!TenantReceipts), 0, !TenantReceipts)), "0.00")
'                h = Format((IIf(IsNull(!TenantReceipts), 0, !TenantReceipts)), "0.00")
''                flxPayFees.TextMatrix(i, 13) = Format((IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)), "0.00")
''                h = Format((IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)), "0.00")
'                flxPayFees.TextMatrix(i, 14) = Format(IIf(IsNull(!SupplierPayments), 0, !SupplierPayments), "0.00")
'                k = Format(IIf(IsNull(!SupplierPayments), 0, !SupplierPayments), "0.00")
'                flxPayFees.TextMatrix(i, 15) = !BankACBalance 'Format(!PaymentsonAccount, "0.00")
'                 l = Format(!PaymentsonAccount, "0.00")
'
'                flxPayFees.TextMatrix(i, 16) = IIf(IsNull(!ClientPayments), "0.00", !ClientPayments)
'                m = IIf(IsNull(!ClientPayments), "0.00", !ClientPayments)
'                flxPayFees.TextMatrix(i, 17) = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '0 '!LandlordPayments
'                n = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '
'                flxPayFees.TextMatrix(i, 18) = IIf(IsNull(!ManagingAgentPayments), "0.00", !ManagingAgentPayments) '!ManagingAgentPayments
'                o = IIf(IsNull(!ManagingAgentPayments), "0.00", !ManagingAgentPayments) '
'                flxPayFees.TextMatrix(i, 19) = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '!BankPaymentReceipts
'                p = !BankPaymentReceipts
'                  flxPayFees.TextMatrix(i, 20) = Format(!ClientAcBalance, "0.00")
'                r = Format(!ManagingAgentAcBalance, "0.00")
'                flxPayFees.TextMatrix(i, 21) = Format(!SupplierAcBalance, "0.00")
'                q = Format(!SupplierAcBalance, "0.00")
'                flxPayFees.TextMatrix(i, 22) = Format(!ManagingAgentAcBalance, "0.00")
'                r = Format(!ManagingAgentAcBalance, "0.00")
'
'                flxPayFees.TextMatrix(i, 23) = Format(!LandlordACBalance, "0.00")
'                s = Format(!ClientLandlordBalance, "0.00")
'                flxPayFees.TextMatrix(i, 24) = Format(!AccrualsAcBalance, "0.00") ' IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date
'                t = Format(!AccrualsAcBalance, "0.00")
'                flxPayFees.TextMatrix(i, 25) = Format(!TenantDepositsReceived, "0.00")
'                u = Format(!TenantDepositsReceived, "0.00")
'                flxPayFees.TextMatrix(i, 26) = IIf(!Printed = True, "Yes", "No")
'                flxPayFees.TextMatrix(i, 27) = IIf(!Emailed = True, "Yes", "No") 'IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date'!Emailed
'                flxPayFees.TextMatrix(i, 28) = IIf(!Invoiced = True, "Yes", "No")
'                flxPayFees.TextMatrix(i, 29) = IIf(IsNull(!PINumber), "", !PINumber)
'                If Not IsNull(!isfinalized) Then
'                    If !isfinalized = "1" Then
'                        flxPayFees.TextMatrix(i, 30) = "Finalised"
'                    Else
'                         flxPayFees.TextMatrix(i, 30) = "Pending"
'                    End If
'                End If
'                flxPayFees.RowHeight(i) = 280
'
'                    szSQL = "SELECT * from RentSummaryStatementdetails where StatementID=" & rsRentSummaryStatement!statementID & ""
'                     rsRentSummaryStatementSplit1.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'                     While Not rsRentSummaryStatementSplit1.EOF
'                            i = i + 1
'                            flxPayFees.AddItem ""
'                            flxPayFees.TextMatrix(i, 1) = "-"
'                            flxPayFees.TextMatrix(i, 2) = !statementID
'                            flxPayFees.TextMatrix(i, 3) = rsRentSummaryStatementSplit1!splitID 'This naming should be only clientID its  mistak on spec
'                            flxPayFees.TextMatrix(i, 4) = rsRentSummaryStatementSplit1!ClientID
'                            flxPayFees.TextMatrix(i, 5) = rsRentSummaryStatementSplit1!PINumber
'                            flxPayFees.TextMatrix(i, 6) = rsRentSummaryStatementSplit1!SageAccountNumber
'                            flxPayFees.TextMatrix(i, 7) = "Percentage" 'Payable _BASIS_
'                            flxPayFees.TextMatrix(i, 8) = "" 'Payable _BASIS_
'                            flxPayFees.TextMatrix(i, 10) = IIf(IsNull(rsRentSummaryStatementSplit1!PercentageLL), 100, rsRentSummaryStatementSplit1!PercentageLL) 'IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value), "0.00", Format(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value, "0.00")) & "%" 'strPercentage 'Percentage
'                            flxPayFees.TextMatrix(i, 11) = rsRentSummaryStatementSplit1!amount '"£" & dblPayableAmount * IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value), "0.00", Format(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value, "0.00")) / 100
'                            flxPayFees.TextMatrix(i, 12) = ""
'                            flxPayFees.TextMatrix(i, 13) = ""
'                            flxPayFees.TextMatrix(i, 14) = ""
'                            flxPayFees.TextMatrix(i, 15) = ""
'                            flxPayFees.TextMatrix(i, 16) = ""
'                            flxPayFees.TextMatrix(i, 17) = ""
'                            flxPayFees.TextMatrix(i, 18) = ""
'                            'flxPayFees.TextMatrix(i, 19) = q
'                            flxPayFees.TextMatrix(i, 20) = ""
'                            flxPayFees.TextMatrix(i, 21) = ""
'                            flxPayFees.TextMatrix(i, 22) = ""
'                            flxPayFees.TextMatrix(i, 23) = ""
'                            flxPayFees.TextMatrix(i, 24) = ""
'                            flxPayFees.TextMatrix(i, 25) = ""
'                            flxPayFees.TextMatrix(i, 26) = "" 'IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date'!Emailed
'                            flxPayFees.TextMatrix(i, 27) = ""
'                            flxPayFees.TextMatrix(i, 28) = ""
'                            flxPayFees.RowHeight(i) = 0
'                        rsRentSummaryStatementSplit1.MoveNext
'                     Wend
'                     rsRentSummaryStatementSplit1.Close
'                     Set rsRentSummaryStatementSplit1 = Nothing
'
'
'                rsRentSummaryStatement.MoveNext
'                i = i + 1
'            Wend
'            End With
'            adoconn.Close
'            Set adoconn = Nothing
'End Sub
Public Sub loadflxPayFees(strFilter As String)
       Dim szSQL As String
        Call configflxPayFees
        Dim adoConn As New ADODB.Connection
        Dim rsRentSummaryStatement As New ADODB.Recordset
        Dim rsRentSummaryStatementSplit1 As New ADODB.Recordset
        Dim dblPayableAmount As Double
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        Dim e As String
        Dim f As String
        Dim g As String
        Dim h As String
        Dim K As String
        Dim l As String
        Dim m As String
        Dim n As String
        Dim j, o, p, q, r, s, t, u As String
        'Exit Sub
        
        Dim i As Long
        adoConn.Open getConnectionString
'            If strFilter = "" Then
'                    rsRentSummaryStatement.Open "Select '+' as SIGN,R.* from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
'                    "R.ClientIDLandlordID=S.SupplierID  Order by statementID DESC", adoConn, adOpenDynamic, adLockOptimistic
'            Else
'                    rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
'                    "R.ClientIDLandlordID=S.SupplierID " & strFilter & " Order by statementID DESC", adoConn, adOpenDynamic, adLockOptimistic
'            End If
            
            
            If strFilter = "" Then
                    szSQL = "Select SIGN,StatementID,statementNo,ClientIDLandlordID,BankCode,PreviousStatementDate, " & _
                    "StatementDate,AvailableFunds, Retentions,0 as amount1,PayableAmount, StatementClosingBal,TenantReceipts,SupplierPayments,  " & _
                    "BankAcBalance, ClientPayments, BankAcBalance, ClientPayments, LandlordPayments, ManagingAgentPayments,LandlordPayments,  " & _
                    "ClientAcBalance, SupplierAcBalance ,ManagingAgentACBalance,LandlordACBalance, AccrualsAcBalance,TenantDepositsReceived ,  " & _
                    "Printed,emailed,invoiced, PInumber, DateFinalized,isfinalized,InclSupplierOS,InclMngtFeesDue From (Select '+' as SIGN,StatementID,statementNo,ClientIDLandlordID,BankCode,PreviousStatementDate, " & _
                    "StatementDate,AvailableFunds, Retentions,0 as amount1,PayableAmount, StatementClosingBal,TenantReceipts,SupplierPayments,  " & _
                    "BankAcBalance, ClientPayments, LandlordPayments, ManagingAgentPayments,  " & _
                    "ClientAcBalance, SupplierAcBalance ,ManagingAgentACBalance,LandlordACBalance, AccrualsACBalance,TenantDepositsReceived ,  " & _
                    "Printed,emailed,invoiced, PInumber, DateFinalized,isfinalized,InclSupplierOS,InclMngtFeesDue from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
                    "R.ClientIDLandlordID=S.SupplierID " & _
                    "UNION ALL " & _
                    "SELECT '-' as SIGN,StatementID,SplitID  as statementNo,ClientID as ClientIDLandlordID,PInumber as BankCode,SageAccountNumber as PreviousStatementDate, " & _
                    " 'Percentage' AS StatementDate,'' as AvailableFunds,PercentageLL as Retentions,Amount as Amount1,Amount as PayableAmount,'' as StatementClosingBal,'' as TenantReceipts,'' as SupplierPayments,  " & _
                    "'' as BankAcBalance, '' as ClientPayments,'' as LandlordPayments,'' as ManagingAgentPayments, " & _
                    "'' as ClientAcBalance,'' as SupplierAcBalance ,'' as ManagingAgentACBalance,'' as LandlordACBalance,'' as AccrualsAcBalance,'' as TenantDepositsReceived ,  " & _
                    "'' as Printed,'' as emailed,'' as invoiced, '' as PInumber,'' as DateFinalized,'' as isfinalized,'' as InclSupplierOS,'' as InclMngtFeesDue  from RentSummaryStatementdetails) order by StatementID DESC,SIGN DESC"
                    'Exit Sub
                    'rsRentSummaryStatement.Close
                    rsRentSummaryStatement.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
            Else
                     szSQL = "Select SIGN,StatementID,statementNo,ClientIDLandlordID,BankCode,PreviousStatementDate, " & _
                    "StatementDate,AvailableFunds, Retentions,0 as amount1,PayableAmount, StatementClosingBal,TenantReceipts,SupplierPayments,  " & _
                    "BankAcBalance, ClientPayments, BankAcBalance, ClientPayments, LandlordPayments, ManagingAgentPayments,LandlordPayments,  " & _
                    "ClientAcBalance, SupplierAcBalance ,ManagingAgentACBalance,LandlordACBalance, AccrualsAcBalance,TenantDepositsReceived ,  " & _
                    "Printed,emailed,invoiced, PInumber, DateFinalized,isfinalized,InclSupplierOS,InclMngtFeesDue From (Select '+' as SIGN,StatementID,statementNo,ClientIDLandlordID,BankCode,PreviousStatementDate, " & _
                    "StatementDate,AvailableFunds, Retentions,0 as amount1,PayableAmount, StatementClosingBal,TenantReceipts,SupplierPayments,  " & _
                    "BankAcBalance, ClientPayments, LandlordPayments, ManagingAgentPayments,  " & _
                    "ClientAcBalance, SupplierAcBalance ,ManagingAgentACBalance,LandlordACBalance, AccrualsACBalance,TenantDepositsReceived ,  " & _
                    "Printed,emailed,invoiced, PInumber, DateFinalized,isfinalized,InclSupplierOS,InclMngtFeesDue from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
                    "R.ClientIDLandlordID=S.SupplierID " & _
                    "UNION ALL " & _
                    "SELECT '-' as SIGN,StatementID,SplitID  as statementNo,ClientID as ClientIDLandlordID,PInumber as BankCode,SageAccountNumber as PreviousStatementDate, " & _
                    " 'Percentage' AS StatementDate,'' as AvailableFunds,PercentageLL as Retentions ,Amount as Amount1,Amount as PayableAmount,'' as StatementClosingBal,'' as TenantReceipts,'' as SupplierPayments,  " & _
                    "'' as BankAcBalance, '' as ClientPayments,'' as LandlordPayments,'' as ManagingAgentPayments, " & _
                    "'' as ClientAcBalance,'' as SupplierAcBalance ,'' as ManagingAgentACBalance,'' as LandlordACBalance,'' as AccrualsAcBalance,'' as TenantDepositsReceived ,  " & _
                    "'' as Printed,'' as emailed,'' as invoiced, '' as PInumber,'' as DateFinalized,'' as isfinalized,'' as InclSupplierOS,'' as InclMngtFeesDue  from RentSummaryStatementdetails) U " & strFilter & " order by StatementID DESC,SIGN DESC"

                     rsRentSummaryStatement.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'                    rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where PostToHistory=false AND " & _
'                    "R.ClientIDLandlordID=S.SupplierID " & strFilter & " Order by statementID DESC", adoConn, adOpenDynamic, adLockOptimistic
            End If
            flxPayFees.Rows = RecordCount(rsRentSummaryStatement) + 1
            i = 1
'            Exit Sub
            With rsRentSummaryStatement
            While Not rsRentSummaryStatement.EOF
               ' flxPayFees.AddItem ""
                If !SIGN = "+" Then
                    flxPayFees.RowHeight(i) = 280
                Else
                    flxPayFees.RowHeight(i) = 0
                End If
                flxPayFees.TextMatrix(i, 1) = !SIGN ' Expansion/collapse
                flxPayFees.TextMatrix(i, 2) = "CS" & !statementID
                flxPayFees.TextMatrix(i, 3) = !statementNo 'statement no by client ID
                flxPayFees.TextMatrix(i, 4) = !ClientIDLandlordID 'This naming should be only clientID its  mistake on spec
                flxPayFees.TextMatrix(i, 5) = !BankCode
                a = !BankCode
                flxPayFees.TextMatrix(i, 6) = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
                b = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
                flxPayFees.TextMatrix(i, 7) = !StatementDate
                c = !StatementDate
                flxPayFees.TextMatrix(i, 8) = Format(!Availablefunds, "0.00")
                d = !StatementDate
                If !SIGN = "+" Then
                    flxPayFees.TextMatrix(i, 9) = Format(!Retentions, "0.00")
                Else
                    flxPayFees.TextMatrix(i, 9) = CStr(!Retentions) & "%"
                End If
                'flxPayFees.TextMatrix(i, 9) = Format(!Retentions, "0.00")
                e = Format(!Retentions, "0.00")
                flxPayFees.TextMatrix(i, 10) = "" 'Keep it blank for master 'apportionment
                If Val(!PayableAmount) > 0 Then
                    flxPayFees.TextMatrix(i, 11) = Format(!PayableAmount, "0.00")  'Keep it blank for master 'Payable Paid amount
                Else
                    flxPayFees.TextMatrix(i, 11) = ""
                End If
                f = "£" & Format(!PayableAmount, "0.00")
                'dblPayableAmount = !PayableAmount
                flxPayFees.TextMatrix(i, 12) = Format(!StatementClosingBal, "0.00")
                g = Format(!StatementClosingBal, "0.00")
                 flxPayFees.TextMatrix(i, 13) = Format((IIf(IsNull(!TenantReceipts), 0, !TenantReceipts)), "0.00")
                h = Format((IIf(IsNull(!TenantReceipts), 0, !TenantReceipts)), "0.00")
'                flxPayFees.TextMatrix(i, 13) = Format((IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)), "0.00")
'                h = Format((IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)), "0.00")
                flxPayFees.TextMatrix(i, 14) = Format(IIf(IsNull(!SupplierPayments), 0, !SupplierPayments), "0.00")
                K = Format(IIf(IsNull(!SupplierPayments), 0, !SupplierPayments), "0.00")
                flxPayFees.TextMatrix(i, 15) = !BankACBalance 'Format(!PaymentsonAccount, "0.00")
                 'l = Format(!PaymentsonAccount, "0.00")
                  
                flxPayFees.TextMatrix(i, 16) = IIf(IsNull(!ClientPayments), "0.00", !ClientPayments)
                m = IIf(IsNull(!ClientPayments), "0.00", !ClientPayments)
                flxPayFees.TextMatrix(i, 17) = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '0 '!LandlordPayments
                n = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '
                flxPayFees.TextMatrix(i, 18) = IIf(IsNull(!ManagingAgentPayments), "0.00", !ManagingAgentPayments) '!ManagingAgentPayments
                o = IIf(IsNull(!ManagingAgentPayments), "0.00", !ManagingAgentPayments) '
                flxPayFees.TextMatrix(i, 19) = IIf(IsNull(!LandlordPayments), "0.00", !LandlordPayments) '!BankPaymentReceipts
               ' p = !BankPaymentReceipts
                  flxPayFees.TextMatrix(i, 20) = Format(!ClientACBalance, "0.00")
                r = Format(!ManagingAgentAcBalance, "0.00")
                flxPayFees.TextMatrix(i, 21) = Format(!SupplierAcBalance, "0.00")
                q = Format(!SupplierAcBalance, "0.00")
                flxPayFees.TextMatrix(i, 22) = Format(!ManagingAgentAcBalance, "0.00")
                r = Format(!ManagingAgentAcBalance, "0.00")
                
                flxPayFees.TextMatrix(i, 23) = Format(!LandlordACBalance, "0.00")
                's = Format(!ClientLandlordBalance, "0.00")
                flxPayFees.TextMatrix(i, 24) = Format(!AccrualsAcBalance, "0.00") ' IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date
                t = Format(!AccrualsAcBalance, "0.00")
                flxPayFees.TextMatrix(i, 25) = Format(!TenantDepositsReceived, "0.00")
                u = Format(!TenantDepositsReceived, "0.00")
                flxPayFees.TextMatrix(i, 26) = IIf(!Printed = True, "Yes", "No")
                flxPayFees.TextMatrix(i, 27) = IIf(!Emailed = True, "Yes", "No") 'IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date'!Emailed
                flxPayFees.TextMatrix(i, 28) = IIf(!Invoiced = True, "Yes", "No")
                flxPayFees.TextMatrix(i, 29) = IIf(IsNull(!PINumber), "", !PINumber)
                If Not IsNull(!isfinalized) Then
                    If !isfinalized = "1" Then
                        flxPayFees.TextMatrix(i, 30) = Format(IIf(IsNull(!DateFinalized), "", !DateFinalized), "dd/MM/yyyy")
                     End If
                    If flxPayFees.TextMatrix(i, 30) = "" Then
                        If !isfinalized = "1" Then
                            flxPayFees.TextMatrix(i, 30) = "Finalised"
                        Else
                             flxPayFees.TextMatrix(i, 30) = "Pending"
                        End If
                     End If
                End If
                flxPayFees.TextMatrix(i, 31) = IIf(!InclSupplierOS = True, 1, 0)
                flxPayFees.TextMatrix(i, 32) = IIf(!InclMngtFeesDue = True, 1, 0)
               
                
'                    szSQL = "SELECT * from RentSummaryStatementdetails where StatementID=" & rsRentSummaryStatement!statementID & ""
'                    rsRentSummaryStatementSplit1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                    While Not rsRentSummaryStatementSplit1.EOF
'                           i = i + 1
'                           flxPayFees.AddItem ""
'                           flxPayFees.TextMatrix(i, 1) = "-"
'                           flxPayFees.TextMatrix(i, 2) = !statementID
'                           flxPayFees.TextMatrix(i, 3) = rsRentSummaryStatementSplit1!splitID 'This naming should be only clientID its  mistak on spec
'                           flxPayFees.TextMatrix(i, 4) = rsRentSummaryStatementSplit1!ClientID
'                           flxPayFees.TextMatrix(i, 5) = rsRentSummaryStatementSplit1!PINumber
'                           flxPayFees.TextMatrix(i, 6) = rsRentSummaryStatementSplit1!SageAccountNumber
'                           flxPayFees.TextMatrix(i, 7) = "Percentage" 'Payable _BASIS_
'                           flxPayFees.TextMatrix(i, 8) = "" 'Payable _BASIS_
'                           flxPayFees.TextMatrix(i, 10) = IIf(IsNull(rsRentSummaryStatementSplit1!PercentageLL), 100, rsRentSummaryStatementSplit1!PercentageLL) 'IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value), "0.00", Format(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value, "0.00")) & "%" 'strPercentage 'Percentage
'                           flxPayFees.TextMatrix(i, 11) = rsRentSummaryStatementSplit1!amount '"£" & dblPayableAmount * IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value), "0.00", Format(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value, "0.00")) / 100
'                           flxPayFees.TextMatrix(i, 12) = ""
'                           flxPayFees.TextMatrix(i, 13) = ""
'                           flxPayFees.TextMatrix(i, 14) = ""
'                           flxPayFees.TextMatrix(i, 15) = ""
'                           flxPayFees.TextMatrix(i, 16) = ""
'                           flxPayFees.TextMatrix(i, 17) = ""
'                           flxPayFees.TextMatrix(i, 18) = ""
'                           'flxPayFees.TextMatrix(i, 19) = q
'                           flxPayFees.TextMatrix(i, 20) = ""
'                           flxPayFees.TextMatrix(i, 21) = ""
'                           flxPayFees.TextMatrix(i, 22) = ""
'                           flxPayFees.TextMatrix(i, 23) = ""
'                           flxPayFees.TextMatrix(i, 24) = ""
'                           flxPayFees.TextMatrix(i, 25) = ""
'                           flxPayFees.TextMatrix(i, 26) = "" 'IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date'!Emailed
'                           flxPayFees.TextMatrix(i, 27) = ""
'                           flxPayFees.TextMatrix(i, 28) = ""
'                           flxPayFees.RowHeight(i) = 0
'                       rsRentSummaryStatementSplit1.MoveNext
'                    Wend
'                    rsRentSummaryStatementSplit1.Close
'                    Set rsRentSummaryStatementSplit1 = Nothing
'
'
                    rsRentSummaryStatement.MoveNext
                    i = i + 1
                    Wend
            End With
            adoConn.Close
            Set adoConn = Nothing
End Sub
Public Sub loadflxPayFeesHistory()
      Dim szSQL As String
        Call configflxPayFeesHistory
        Dim adoConn As New ADODB.Connection
        Dim rsRentSummaryStatement As New ADODB.Recordset
        Dim rsRentSummaryStatementSplit1 As New ADODB.Recordset
        Dim i As Long
        adoConn.Open getConnectionString
            rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where PostToHistory=true AND " & _
                    "R.ClientIDLandlordID=S.SupplierID Order by StatementID desc", adoConn, adOpenDynamic, adLockOptimistic
            i = 1
            With rsRentSummaryStatement
            While Not rsRentSummaryStatement.EOF
                flxPayFeesHistory.AddItem ""
                flxPayFeesHistory.TextMatrix(i, 1) = "+" ' Expansion/collapse
                flxPayFeesHistory.TextMatrix(i, 2) = "CS" & !statementID
                flxPayFeesHistory.TextMatrix(i, 3) = !statementNo
                flxPayFeesHistory.TextMatrix(i, 4) = !ClientIDLandlordID 'This naming should be only clientID its  mistak on spec
                flxPayFeesHistory.TextMatrix(i, 5) = !BankCode
                flxPayFeesHistory.TextMatrix(i, 6) = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
                flxPayFeesHistory.TextMatrix(i, 7) = !StatementDate
                flxPayFeesHistory.TextMatrix(i, 8) = !StatementOpBal
                flxPayFeesHistory.TextMatrix(i, 9) = !Retentions
                flxPayFeesHistory.TextMatrix(i, 10) = !AccrualsAcBalance
                flxPayFeesHistory.TextMatrix(i, 11) = IIf(IsNull(!SupplierAcBalance), 0, !SupplierAcBalance)
                flxPayFeesHistory.TextMatrix(i, 12) = IIf(IsNull(!ManagingAgentAcBalance), 0, !ManagingAgentAcBalance)
                flxPayFeesHistory.TextMatrix(i, 13) = IIf(IsNull(!ClientACBalance), 0, !ClientACBalance)
                flxPayFeesHistory.TextMatrix(i, 14) = IIf(IsNull(!LandlordACBalance), 0, !LandlordACBalance)
                flxPayFeesHistory.TextMatrix(i, 15) = IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)
                flxPayFeesHistory.TextMatrix(i, 16) = !Availablefunds
                flxPayFeesHistory.TextMatrix(i, 17) = !PaymentsonAccount
                flxPayFeesHistory.TextMatrix(i, 18) = !PayableAmount
                flxPayFeesHistory.TextMatrix(i, 19) = !StatementClosingBal
                flxPayFeesHistory.TextMatrix(i, 20) = IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date
                flxPayFeesHistory.TextMatrix(i, 21) = !Printed
                flxPayFeesHistory.TextMatrix(i, 22) = !Emailed
                flxPayFeesHistory.TextMatrix(i, 23) = !Invoiced
                flxPayFeesHistory.RowHeight(i) = 280
                
                szSQL = "SELECT S.SupplierName as CID, T.ID as PID,P.PAYABLE_ID,P.CPA_ID, T.PayType , PAYABLE_TYPE,F.FundID," & _
                    "F.FundName,clientLandlordID, PAY_START_DATE, PAY_END_DATE,   P.ONDD,P.PAYABLE_BASIS_,PAY_NtDueDate,Percentage,StopDate,PAY_END_DATE " & _
                    "FROM tlbPayable AS P, ClientProAgr AS C,  PayableTypes AS T, FUND as F,Supplier S " & _
                    "WHERE   F.FundID=P.PAY_Fund AND P.CPA_ID = C.CPA_ID And S.SupplierID=P.ClientLandlordID and C.ClientID = '" & !ClientIDLandlordID & "' And " & _
                    "T.ID = P.PAYABLE_TYPE And C.PropertyID in (" & !ListOfinputProperties & ")"
                     rsRentSummaryStatementSplit1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                     While Not rsRentSummaryStatementSplit1.EOF
                            i = i + 1
                            flxPayFeesHistory.AddItem ""
                            flxPayFeesHistory.TextMatrix(i, 1) = "-"
                            flxPayFeesHistory.TextMatrix(i, 2) = rsRentSummaryStatementSplit1!clientLandlordID 'This naming should be only clientID its  mistak on spec
                            flxPayFeesHistory.TextMatrix(i, 3) = rsRentSummaryStatementSplit1!PayType
                            flxPayFeesHistory.TextMatrix(i, 4) = rsRentSummaryStatementSplit1!FundName
                            flxPayFeesHistory.TextMatrix(i, 5) = rsRentSummaryStatementSplit1!PAYABLE_BASIS_
                            flxPayFeesHistory.TextMatrix(i, 6) = IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value), "0.00", Format(rsRentSummaryStatementSplit1.Fields.Item("Percentage").Value, "0.00")) & "%" 'strPercentage 'Percentage
                            flxPayFeesHistory.TextMatrix(i, 7) = IIf(IsNull(rsRentSummaryStatementSplit1.Fields.Item("StopDate").Value) = True, "", rsRentSummaryStatementSplit1.Fields.Item("StopDate").Value) 'rsRentSummaryStatementSplit1!Percentage
                            'flxPayFeesHistory.TextMatrix(i, 8) = !ClientIDLandlordID
                            flxPayFeesHistory.RowHeight(i) = 0
                        rsRentSummaryStatementSplit1.MoveNext
                     Wend
                     rsRentSummaryStatementSplit1.Close
                     Set rsRentSummaryStatementSplit1 = Nothing


                rsRentSummaryStatement.MoveNext
                i = i + 1
            Wend
            End With
            adoConn.Close
            Set adoConn = Nothing
End Sub
Private Sub configflxPayFees()
        Dim szHeader As String
        flxPayFees.Clear
        szHeader$ = "|<Exp-Collapse|<StatementID|<Statement No|<Client/Landlord|<Bank Code|<Previous Statement Date|<Statement Date|<Available Funds|<Retentions" & _
            "|>Apportionment (%)|<Payable/Paid Amount|>Statement Closing Balance|>Tenant receipts|<Supplier Payments|<Bank Balance |<Client Payments|<Landlord Payments" & _
            " |<Managing Agent Payments|<Bank Payment/Receipts ADD|<Client AC Balance|<Suppliers Balance|<Managing Agents Balance|<Landlord Balance|<Accruals Balance|<Tenant Deposits Received|<Printed|<Emailed|<Invoiced|<PI Number|<Finalised |<Finalised   "
        flxPayFees.FormatString = szHeader$
        'flxPayFees.Clear
        flxPayFees.Cols = 33
        flxPayFees.Rows = 2
        'flxPayFees.RowHeight(0) = 0
        flxPayFees.ColWidth(0) = 350
        flxPayFees.ColWidth(1) = 420
        flxPayFees.ColWidth(2) = 2000 'Label5(101).Left - Label5(78).Left'StatementID
        flxPayFees.ColWidth(3) = 2000 'Label5(101).Left - Label5(79).Left StatementNo
        flxPayFees.ColWidth(4) = 2500 'Label5(101).Left - Label5(80).Left ClientIDLandlordID
        flxPayFees.ColWidth(5) = 1500 'Label5(101).Left - Label5(81).Left BankCode
        flxPayFees.ColWidth(6) = 1500 'Label5(101).Left - Label5(82).Left PreviousStatementDate
        flxPayFees.ColAlignment(6) = vbRightJustify
        flxPayFees.ColWidth(7) = 1500 'Label5(101).Left - Label5(83).Left StatementDate
        flxPayFees.ColWidth(8) = 1500 'Label5(101).Left - Label5(84).Left StatementOpBal
        flxPayFees.ColWidth(9) = 1500 'Label5(101).Left - Label5(85).Left Retentions
        flxPayFees.ColWidth(10) = 1500 ' Label5(101).Left - Label5(86).Left AccrualsACBalance
        flxPayFees.ColWidth(11) = 2300 'Label5(101).Left - Label5(87).Left SupplierACBalance
        flxPayFees.ColWidth(12) = 2000 'Label5(101).Left - Label5(88).Left ManagingAgentACBalance
        flxPayFees.ColWidth(13) = 2000  'Label5(101).Left - Label5(89).Left ClientACBalance
        flxPayFees.ColWidth(14) = 2000  'Label5(101).Left - Label5(90).Left LandLordACBalance
        flxPayFees.ColWidth(15) = 2000  'Label5(101).Left - Label5(91).Left TenantDepositsreceived
        flxPayFees.ColWidth(16) = 2000  'Label5(101).Left - Label5(92).Left Availablefunds
        flxPayFees.ColWidth(17) = 2000  'Label5(101).Left - Label5(93).Left PaymentsonAccount
        flxPayFees.ColWidth(18) = 2600  'Label5(101).Left - Label5(94).Left PayableAmount
        flxPayFees.ColWidth(19) = 2200  'Label5(101).Left - Label5(95).Left StatementClosingBal
        flxPayFees.ColWidth(20) = 2200  'Label5(101).Left - Label5(96).Left Generated_Date
        flxPayFees.ColWidth(21) = 2200  'Label5(101).Left - Label5(97).Left !Printed
        flxPayFees.ColWidth(22) = 2200  'Label5(101).Left - Label5(98).Left  !Emailed 22 !Invoiced
        flxPayFees.ColWidth(23) = 2200  'Label5(101).Left - Label5(99).Left
        flxPayFees.ColWidth(24) = 2200  'Label5(101).Left - Label5(100).Left
        flxPayFees.ColWidth(25) = 2200  '
        flxPayFees.ColWidth(26) = 2200  '
        flxPayFees.ColWidth(27) = 1200  '
        flxPayFees.ColWidth(28) = 1200
        flxPayFees.ColWidth(29) = 1600
        flxPayFees.ColWidth(30) = 2200
        flxPayFees.ColWidth(31) = 0
        flxPayFees.ColWidth(32) = 0
        
        
    
End Sub
Private Sub configflxPayFeesHistory()
        Dim szHeader As String
        flxPayFeesHistory.Clear
        szHeader$ = "|<Exp-Collapse|<StatementID|<StatementNo|<ClientIDLandlordID|<BankCode|<PreviousStatementDate|<StatementDate|<StatementOpBal|<Retentions" & _
            "|>AccrualsACBalance|<SupplierACBalance|>ManagingAgentACBalance|>ClientACBalance|<LandLordACBalance|<TenantDepositsreceived|<Availablefunds|<PaymentsonAccount" & _
            " |<PayableAmount|<StatementClosingBal|<Generated_Date|<Printed|<Emailed|<Invoiced"
        flxPayFeesHistory.FormatString = szHeader$
        'flxPayFeesHistory.Clear
        flxPayFeesHistory.Cols = 24
        flxPayFeesHistory.Rows = 2
        'flxPayFeesHistory.RowHeight(0) = 0
        flxPayFeesHistory.ColWidth(0) = 350
        flxPayFeesHistory.ColWidth(1) = 420
        flxPayFeesHistory.ColWidth(2) = 2000 'Label5(101).Left - Label5(78).Left'StatementID
        flxPayFeesHistory.ColWidth(3) = 3000 'Label5(101).Left - Label5(79).Left StatementNo
        flxPayFeesHistory.ColWidth(4) = 2500 'Label5(101).Left - Label5(80).Left ClientIDLandlordID
        flxPayFeesHistory.ColWidth(5) = 1500 'Label5(101).Left - Label5(81).Left BankCode
        flxPayFeesHistory.ColWidth(6) = 1500 'Label5(101).Left - Label5(82).Left PreviousStatementDate
        flxPayFeesHistory.ColAlignment(6) = vbRightJustify
        flxPayFeesHistory.ColWidth(7) = 1500 'Label5(101).Left - Label5(83).Left StatementDate
        flxPayFeesHistory.ColWidth(8) = 1500 'Label5(101).Left - Label5(84).Left StatementOpBal
        flxPayFeesHistory.ColWidth(9) = 1500 'Label5(101).Left - Label5(85).Left Retentions
        flxPayFeesHistory.ColWidth(10) = 1500 ' Label5(101).Left - Label5(86).Left AccrualsACBalance
        flxPayFeesHistory.ColWidth(11) = 1500 'Label5(101).Left - Label5(87).Left SupplierACBalance
        flxPayFeesHistory.ColWidth(12) = 1500 'Label5(101).Left - Label5(88).Left ManagingAgentACBalance
        flxPayFeesHistory.ColWidth(13) = 1500  'Label5(101).Left - Label5(89).Left ClientACBalance
        flxPayFeesHistory.ColWidth(14) = 1500  'Label5(101).Left - Label5(90).Left LandLordACBalance
        flxPayFeesHistory.ColWidth(15) = 1500  'Label5(101).Left - Label5(91).Left TenantDepositsreceived
        flxPayFeesHistory.ColWidth(16) = 1500  'Label5(101).Left - Label5(92).Left Availablefunds
        flxPayFeesHistory.ColWidth(17) = 1500  'Label5(101).Left - Label5(93).Left PaymentsonAccount
        flxPayFeesHistory.ColWidth(18) = 1500  'Label5(101).Left - Label5(94).Left PayableAmount
        flxPayFeesHistory.ColWidth(19) = 1800  'Label5(101).Left - Label5(95).Left StatementClosingBal
        flxPayFeesHistory.ColWidth(20) = 1700  'Label5(101).Left - Label5(96).Left Generated_Date
        flxPayFeesHistory.ColWidth(21) = 1200  'Label5(101).Left - Label5(97).Left !Printed
        flxPayFeesHistory.ColWidth(22) = 1200  'Label5(101).Left - Label5(98).Left  !Emailed 22 !Invoiced
        flxPayFeesHistory.ColWidth(23) = 1200  'Label5(101).Left - Label5(99).Left
        flxPayFeesHistory.ColWidth(24) = 1200  'Label5(101).Left - Label5(100).Left
        
    
End Sub
Private Sub GeneratePreview(szStatmentID As String)
'    Exit Sub
'    Frame1(6).Visible = True
'    Frame1(6).Top = 135
'    Frame1(6).Left = 2025
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    'Before writing this table you need to delete this table
    If szSelectedFund = "" Then
        MsgBox "Please select a fund", vbInformation, "Warning!"
        Exit Sub
    End If
    adoConn.Execute "Delete from  RentSummaryStatementPreview"
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    Dim X As String
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    rsRentSummaryStatement.Open "Select * from RentSummaryStatementPreview where 1=2", adoConn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !statementID = szStatmentID 'we are setting this column atutomatically
            !statementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient = "", Null, GetLastStatementDateByClient)
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy")
            !StatementOpBal = dblLasClosingBalance
            !Retentions = txtRetention.text 'we need to further analyse detail/add/deduct retension
            !Clearretentions = False 'Will need to come again
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave
            !ListOfPayableTypeID = ListOfFundsForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit()
            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
            !PaymentsonAccount = -GetPaymentsonAccount 'date  filter added
            'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceipts
            !SupplierPayments = GetSupplierPayment
            !BankPaymentReceipts = GetBankPaymentReceipts
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            
            
            !PayableAmount = Val(txtRentPayable.text)
            !StatementClosingBal = getClosingBalance(dblLasClosingBalance)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostToHistory = False
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub
'Private Function GetClientLandLordBalance() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    adoConn.Open getConnectionString
'
'    szSQL = "SELECT  P.SageAccountNumber,SUM(SWITCH(P.Type = 6,P.Amount, P.Type = 24,P.Amount,P.Type = 7, " & _
'            "-P.Amount,P.Type = 8,-P.Amount,P.Type = 9,-P.Amount)) AS Dr " & _
'            "FROM tlbPayment AS P , Client WHERE  P.SageAccountNumber = Client.ClientID " & _
'            "AND P.ClientID ='" & szSelectedClient & "' " & _
'            "GROUP BY P.SageAccountNumber"
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetClientLandLordBalance = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
'Private Function GetClientLandLordPA() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoconn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    adoconn.Open getConnectionString
'
'    szSQL = "SELECT  P.SageAccountNumber,SUM(SP.Amount) AS Dr " & _
'            "FROM tlbPayment AS P ,tlbPaymentSplit SP, Client,Supplier SS WHERE  SS.SupplierID=P.SageAccountNumber AND SP.Payheader=P.TransactionID AND " & _
'            "P.SageAccountNumber = Client.ClientID AND SS.Type in ('LLORD','CLIENT') " & _
'            "AND P.ClientID ='" & szSelectedClient & "' AND P.Type=9 " & _
'            "GROUP BY P.SageAccountNumber"
'
'    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetClientLandLordPA = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
'    End If
'    rsPayment.Close
'    adoconn.Close
'    Set adoconn = Nothing
'End Function

Private Function GetBankPaymentReceipts() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString

    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND " & _
            "B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=(F.FundID)  AND (B.RentSumStatement=''OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "'"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentReceipts = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLandLordACBalance() As Double   'This function return result as minus This is getting LLORD balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetClientACBalance() As Double   'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('CLIENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAgentBalance() As Double   'This function return result as minus This is getting Agent balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetSupplierOSAmount() As Double   'This function return result as minus'This is getting supplier balance
'Temporarily remming it 2023-08-11 by anol
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    'Bank code does not exits in PI,so do not put it in where clause
'    adoConn.Open getConnectionString
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
'            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
'            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')"
'            'AND   F.FundCode in (" &   ListOfFunds & ")  remmed
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetSupplierOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
    GetSupplierOSAmount = 0
End Function
Private Function GetSupplierPayment() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
            "P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' AND SP.Type='Supplier' "
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierPayment = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetTenantReceipts() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
            "R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND R.ClientID ='" & szSelectedClient & "' "
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetTenantReceipts = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Function GetPaymentsonAccount() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  SP.SupplierID=P.SAGEACCOUNTNUMBER AND " & _
            "P.TransactionID=S.PayHeader AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.TYPE  " & _
            "IN(9) AND S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & _
             ListOfFunds & ") AND SP.SupplierID ='" & szSelectedClient & "' AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) AND SP.Type in ('CLIENT','LLORD') "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetPaymentsonAccount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub cboProperty_Click()
   Me.Caption = "Rent Payable - " + cboProperty.Column(1)

   If Not IsGlobalData(cboProperty.Column(0)) Then MsgBox cboProperty.Column(1) & " does not have global data setup; no fees will be generated.", vbCritical + vbOKOnly, "Global Data"
End Sub

Private Sub chkAllPayableType_Click()
'    Dim i As Integer
'    If chkAllPayableType.Value = 0 Then
'        For i = 1 To flxPayableTypes.Rows - 1
'             flxPayableTypes.TextMatrix(i, 0) = ""
'        Next i
'    Else
'         For i = 1 To flxPayableTypes.Rows - 1
'             flxPayableTypes.TextMatrix(i, 0) = "X"
'        Next i
'    End If
End Sub

Private Sub chkAllProperties_Click()
    Dim i As Integer
    If chkAllProperties.Value = 0 Then
        For i = 1 To flxProperties.Rows - 1
             flxProperties.TextMatrix(i, 0) = ""
        Next i
    Else
         For i = 1 To flxProperties.Rows - 1
             If flxProperties.TextMatrix(i, 1) <> "" Then
                   flxProperties.TextMatrix(i, 0) = "X"
             End If
        Next i
    End If
End Sub

Private Sub chkInFunds_Click()
    Dim i As Integer
    If chkInFunds.Value = 0 Then
        For i = 1 To flxInFunds.Rows - 1
             flxInFunds.TextMatrix(i, 0) = ""
        Next i
    Else
         For i = 1 To flxInFunds.Rows - 1
             flxInFunds.TextMatrix(i, 0) = "X"
        Next i
    End If
End Sub

Private Sub cmdAddToGrid_Click()
            If Val(txtRetensionAmount1.text) = 0 Then
                    MsgBox "Please enter amount greater than zero", vbInformation, "Warning"
                    Exit Sub
            End If
            flxRetensionDetails.Enabled = True
            'Enter data into grid only memory version
            'statementId you shall generate it when you finally save the statement
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = IIf(Option1.Value = True, "+", "-")
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
            flxRetensionDetails.AddItem ""
           ' txtRetensionAmount1.Visible = False
            
            txtRetensionAmount1.text = "0.00"
            FocusControl txtRetensionAmount1
            txtRetensionAmount1.SelStart = 0
            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
            Call MakeSummaryRetention
End Sub

Private Sub cmdCalculateAvailableFund_Click()
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If ListOfFunds = "" Then
         MsgBox "Please select a fund", vbInformation, "Fund!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning "
        Exit Sub
    End If
     Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    'Before writing this table you need to delete this table
    adoConn.Execute "Delete from  RentSummaryStatementPreview"
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
    txtAvailableFunds.text = getAvailablefunds(dblLasClosingBalance)
    txtRentPayable.text = txtAvailableFunds.text
End Sub

Private Sub cmdCalculateavailableFund1_Click()
    'This is recalculate rent summary sub procedure where we are clearing all the flags when we press recalculate button
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update tlbBankPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
    adoConn.Execute "Update tlbPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
    adoConn.Execute "Update tlbReceipt Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
    adoConn.Close
    Set adoConn = Nothing
    
End Sub

Private Sub cmdClient_Click()
    txtSearchClientID.text = ""
txtSearchClientName.text = ""
   Call PrepareList
   picClientList.Top = txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = txtClientID.Left + 5
   picClientList.Visible = True
   FocusControl txtSearchClientID
   picClientList.ZOrder 0
End Sub

Private Sub cmdClient_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then picClientList.Visible = False
End Sub

Private Sub cmdClose_Click(Index As Integer)
        'If Index = 0 Then
                Frame1(6).Visible = False
                 
        'End If
        If Index = 1 Then
            Unload Me
        End If
End Sub

Private Sub cmdGenAll_Click()
    
End Sub

Private Sub cmdClose1_Click()
    Frame1(6).Visible = False
End Sub

Private Sub cmdGenReport_Click() 'produce clientsummarystatement
    'This table shall write into the table clientsummarystatement
    'In this table StatementID is the primary key , Detail shall be loaded from the tlbPayment,tlbReceipt and tlbBankPayment and receipt
    
    '******************** Inputs for Populating this table**************************
    '1.Take input form the input frame
    '2. Then read tlbPayable
    
    
    
End Sub
'Private Function GetControlAccountForPayable(adoconn As ADODB.Connection, szSelectedPayableTypeID As String) As Boolean
'    Dim rsPayableTypes As New ADODB.Recordset
'    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoconn, adOpenStatic, adLockReadOnly
'    If Not rsPayableTypes.EOF Then
'            GetControlAccountForPayable = rsPayableTypes!isUseControlAccount 'PayNCAmt
'    End If
'    rsPayableTypes.Close
'    Set rsPayableTypes = Nothing
'
'End Function

Private Function GetControlAccountForPayableString(adoConn As ADODB.Connection, szSelectedPayableTypeID As String) As Boolean
    Dim rsPayableTypes As New ADODB.Recordset
    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayableTypes.EOF Then
            GetControlAccountForPayableString = rsPayableTypes!PayNCAmt 'PayNCAmt
    End If
    rsPayableTypes.Close
    Set rsPayableTypes = Nothing

End Function



Private Sub cmdClose12_Click()
    Frame6.Visible = False
End Sub

Private Sub cmdEmailDmds_Click()
    If szFromEmail = "" Or szSMTPserver = "" Then
      MsgBox "Company email or SMTP server IP has not been setup."
      Exit Sub
   End If
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
   
'   If IsLoadedAndVisible("frmReport") Then
'      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
'      Exit Sub
'   End If
'   Dim szTemp As String
'   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
'   If FileExists(szTemp) Then
'      MsgBox "There are open reports on another computer. Please close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
'      Exit Sub
'   End If

   Dim adoConn       As New ADODB.Connection
   Dim szID          As String
   Dim bEmailResult  As Boolean
   Dim szSQL         As String
   Dim szClientID As String
 
    iIncDec = 0
    
    
    Dim isitPlus As Boolean
    For rCount = 1 To flxPayFees.Rows - 1
         If flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If flxPayFees.TextMatrix(rCount, 1) = "+" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select a statement.", vbInformation + vbOKOnly, "statement Selection"
       Exit Sub
    End If
'    If szCurrentStatementID = "" Then
'         Exit Sub
'    End If
'    'flxPayFees.TextMatrix(i, 3) is the statement ID by client
'    '66) It should only be possible to modify a statement provided a rent payable
'    ' PI has not been generated against the statement and a subsequent statement has not been produced.
    If isitPlus = True Then
        'MsgBox "This statement cannot be modified, because a Rent Payable invoice  " & flxPayFees.TextMatrix(selRow, 29) & " has been generated against it.", vbInformation + vbOKOnly, "statement Selection"
        szCurrentStatementID = Replace(flxPayFees.TextMatrix(selRow, 2), "CS", "")
        szClientID = flxPayFees.TextMatrix(selRow, 4)
        'Exit Sub
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select a statement to modify", vbInformation + vbOKOnly, "statement Selection"
        Exit Sub
    End If
    'Dim adoConn As New ADODB.Connection
    Dim rsClient As New ADODB.Recordset
    Dim rsSupplier As New ADODB.Recordset
    Dim EmailsTo As String
    adoConn.Open getConnectionString
    
    
'use this 4 field frod land lord type
'SupplierOfficeEmail
'StLandlordHomeEmail
'StLandlordStatementEmail
'StLandlordOfficeEmail
    EmailsTo = "anolcse@gmail.com"
'    szSQL = "Select * from Supplier where SupplierID='" & szClientID & "' AND SupplierType='LL'"
'    rsSupplier.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsSupplier.EOF Then
'        If IIf(IsNull(rsSupplier!StToLandlordAddress), 0, rsSupplier!StToLandlordAddress) = 1 Then
'              EmailsTo = EmailsTo & rsSupplier!SupplierOfficeEmail & ";"
'              EmailsTo = EmailsTo & rsSupplier!StLandlordHomeEmail & ";"
'        End If
'        If IIf(IsNull(rsSupplier!StToStatementAddress), 0, rsSupplier!StToStatementAddress) = 1 Then
'              EmailsTo = EmailsTo & rsSupplier!StLandlordStatementEmail & ";"
'              EmailsTo = EmailsTo & rsSupplier!StLandlordOfficeEmail & ";"
'        End If
'    End If
'    rsSupplier.Close
'    szSQL = "Select * from client where clientId='" & szClientID & "'"
'    rsClient.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsClient.EOF Then
'        If IIf(IsNull(rsClient!StToClientAddress), 0, rsClient!StToClientAddress) = 1 Then
'              EmailsTo = EmailsTo & rsClient!ClientPersonalEmail & ";"
'              EmailsTo = EmailsTo & rsClient!ClientOfficeEmail & ";"
'        End If
'        If IIf(IsNull(rsClient!StToStatementAddress), 0, rsClient!StToStatementAddress) = 1 Then
'              EmailsTo = EmailsTo & rsClient!StClientPersonalEmail & ";"
'              EmailsTo = EmailsTo & rsClient!StClientOfficeEmail & ";"
'        End If
'    End If
'    rsClient.Close
    Dim szSub As String
    Dim szBody As String
    Dim adoRstEmailDetails As New ADODB.Recordset
    EmailsTo = Left(EmailsTo, Len(EmailsTo) - 1)
'    szSub = "Test"
    EmailsTo = "anolcse@gmail.com"
    '  Get the subject and body of the email from template
   szSQL = "SELECT * FROM Template WHERE TemplateName = 'Client Statement Email Template';"
   adoRstEmailDetails.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRstEmailDetails.EOF Then
      szSub = adoRstEmailDetails.Fields.Item("Description").Value
      szBody = adoRstEmailDetails.Fields.Item("Body").Value
   End If
   adoRstEmailDetails.Close
   
    Dim szColl As New Collection
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatement.rpt")
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'   Report.ParameterFields(1).AddCurrentValue CInt(Replace(szCurrentStatementID, "CS", ""))
'
'   Report.EnableParameterPrompting = False
'   If Report.HasSavedData Then Report.DiscardSavedData
'    'Dim szSQL As String
'   szSQL = txtTenantID.text & "_" & UniqueID() & ".pdf"
   szSQL = UniqueID() & ".pdf"
'
'   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
'   Report.ExportOptions.DestinationType = crEDTDiskFile
'   Report.ExportOptions.FormatType = crEFTPortableDocFormat
'   Report.ExportOptions.PDFExportAllPages = True
'   Report.Export False
'   Set Report = Nothing
   szColl.Add DB_PATH & "\AllStuff\Temp\" & szSQL
   Call Export2PdfCSlinebyLine(szSQL)    'Creating PDF for attachements main procudure is here
   Call SendEmail(szFromEmail, EmailsTo, _
                                     szSub, _
                                     szBody, , , _
                                     szColl, "", "", , "") '
                                     

'                    Attach the PDF in the email
'   SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL

   'bEmailResult = SendDemandByE_Mail("General Letter", "Please find the letter in the attachment.", "General Letter")
    MsgBox "Email sent."
'   If bEmailResult Then
'      ShowMsgInTaskBar "Email sent.", "Y", "P"
'   Else
'      ShowMsgInTaskBar "No email sent.", "Y", "N"
'   End If
End Sub

Private Sub cmdFix_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update  RentSummaryStatement set AvailableFunds=0 where statementID in( 175,176,177,178) AND isFinalized=1"
    adoConn.Close
'    Sleep (100)
'    Call UpdateRentPayableOnCSDetails
'    Call loadflxPayFees("")
  
    MsgBox "AvailableFunds fix process is complete"
End Sub

Private Sub cmdFrameFundClose_Click()
    Frame5.Visible = False
End Sub

Private Sub cmdFundListForCreatePI_Click()
    Frame5.Visible = True
    Call LoadFlxFundList
    FocusControl flxFundList
End Sub

Private Sub cmdGenerateRentPayable_Click()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select a statement.", vbInformation + vbOKOnly, "Statement Selection"
'      chkSelectAllDemands.Value = 0
      'ClearGridSelection
      Exit Sub
   End If
   
    iIncDec = 0
    
    
    Dim isitPlus As Boolean
    For rCount = 1 To flxPayFees.Rows - 1
         If flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If flxPayFees.TextMatrix(rCount, 1) = "+" Or flxPayFees.TextMatrix(rCount, 1) = ">" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Or isitPlus = False Then
       MsgBox "Please select a statement at header level.", vbInformation + vbOKOnly, "Statement Selection"
       Exit Sub
    End If
    If szCurrentStatementID = "" Then
         Exit Sub
    End If
    'flxPayFees.TextMatrix(i, 3) is the statement ID by client
    '66) It should only be possible to modify a statement provided a rent payable
    ' PI has not been generated against the statement and a subsequent statement has not been produced.
    If isitPlus = True And flxPayFees.TextMatrix(selRow, 29) <> "" And flxPayFees.TextMatrix(selRow, 3) <= GetLastStatementNoByClient + 1 Then
        MsgBox "A Rent Payable invoice  " & flxPayFees.TextMatrix(selRow, 29) & " has already been generated against this statement.", vbInformation + vbOKOnly, "statement Selection"
        Exit Sub
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select a statement to modify.", vbInformation + vbOKOnly, "Statement Selection"
        Exit Sub
    End If
    
   If szCurrentStatementID <> "" Then
            Frame1(6).Visible = False
            Frame4.Caption = "Create PI from Statement: SS" & szCurrentStatementID
            frmGenaratePayable.strRef = "CS" & szCurrentStatementID & "/" & flxPayFees.TextMatrix(selRow, 3)
            frmGenaratePayable.szCurrentStatementID = szCurrentStatementID
            frmGenaratePayable.szClientID = flxPayFees.TextMatrix(selRow, 4)
            'frmGenaratePayable.txtClientAccount.text = XX
'            frmGenaratePayable.Show
'            frmGenaratePayable.ZOrder 0
            LoadForm frmGenaratePayable
'            Frame4.Left = 2070
'            Frame4.Top = 180
'            Frame4.Visible = True
'            txtAvailableFund1.text = szAvailableFund1
'            FocusControl txtRentPayable1
'            txtRentPayable1.SelStart = 0
'            txtRentPayable1.SelLength = Len(txtRentPayable1.text)
    End If
    
End Sub

Private Sub cmdGridUnitLookup_Click()
   picClientList.Visible = False
End Sub

Private Sub DrawPayFeesGrid()
   Dim iCol As Integer, szHeader As String, szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rstMngFees As New ADODB.Recordset
   Dim iRow As Integer, iRecCol As Integer

   MousePointer = vbHourglass

   adoConn.Open getConnectionString

   flxPayFees.Clear
   flxPayFees.Cols = 15
   flxPayFees.Rows = 2
   flxPayFees.RowHeight(0) = 0
   
   szHeader$ = "<TRAN_DATE|<CATEGORY_CODE|<SUPP_AC|<UNIT_ID|<NOMINAL_CODE|<DEPT_ID|<PROJ_REF|<COST_CODE|<DESCRIPTION|>NET_AMOUNT|<TAX_CODE|>VAT|>TotalAmt|<UPDATE_SAGE|<MY_ID"
   flxPayFees.FormatString = szHeader$

   For iCol = 1 To flxPayFees.Cols - 2
      flxPayFees.ColWidth(iCol - 1) = Label5(iCol + 1).Left - Label5(iCol).Left
   Next iCol
   flxPayFees.ColWidth(13) = flxPayFees.Width + flxPayFees.Left - Label5(14).Left - 40
   flxPayFees.ColWidth(14) = 0         'ID field

   szSQL = "SELECT P.TRAN_DATE, S.MY_ID, P.SUPP_AC, S.CATEGORY_CODE, " & _
               "S.NOMINAL_CODE, S.DEPT_ID, S.PROJ_REF, S.COST_CODE, " & _
               "S.DESCRIPTION, S.NET_AMOUNT, S.TAX_CODE, S.VAT, P.UPDATE_SAGE, " & _
               "S.VAT + S.NET_AMOUNT AS TotalAmt " & _
            "FROM tblPurInv AS P, tblPurInvSRec AS S " & _
            "WHERE P.MY_ID = S.ParentID AND " & _
               "P.TTP = " & CByte(TransactionTakePlace("TTP", "RENT PAYABLE", adoConn)) & " AND " & _
               "P.HISTORY = FALSE;"

   rstMngFees.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstMngFees.EOF Then
      For iRow = 1 To rstMngFees.RecordCount
         For iRecCol = 0 To flxPayFees.Cols - 1
            For iCol = 0 To flxPayFees.Cols - 1
               If flxPayFees.TextMatrix(0, iCol) = rstMngFees.Fields.Item(iRecCol).Name Then Exit For
            Next iCol
            flxPayFees.TextMatrix(iRow, iCol) = IIf(IsNull(rstMngFees.Fields.Item(iRecCol).Value), "", rstMngFees.Fields.Item(iRecCol).Value)
            If (UCase(flxPayFees.TextMatrix(iRow, iCol)) = "TRUE") Then flxPayFees.TextMatrix(iRow, iCol) = "YES"
            If (UCase(flxPayFees.TextMatrix(iRow, iCol)) = "FALSE") Then flxPayFees.TextMatrix(iRow, iCol) = "NO"
         Next iRecCol
         rstMngFees.MoveNext
         If Not rstMngFees.EOF Then flxPayFees.AddItem ""
      Next iRow

      flxPayFees.row = 0
      flxPayFees.col = 0
   End If

   rstMngFees.Close

'Agent id, which is supplier id for all transactions of the Management fees
   If szSupplierAccount = "" Then
      szSQL = "SELECT AgentID " & _
                  "FROM Agent;"
      rstMngFees.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      iCol = rstMngFees.RecordCount
      If iCol = 0 Then
         'ShowMsgInTaskBar "Error in the Agent information, please contact with PCM Consulting.", , "N"
         MsgBox "No Agent information has been entered, please enter the Agent data.", vbCritical + vbOKOnly, iCol & " Agent"
      Else
         szSupplierAccount = rstMngFees!AgentID
      End If
      rstMngFees.Close
   End If
   adoConn.Close
   Set rstMngFees = Nothing
   Set adoConn = Nothing

   MousePointer = vbDefault
End Sub



Private Function ListOfProperties() As String
   Dim i As Integer
   ListOfProperties = "''," ' This shall always include No Property
   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" Then
         ListOfProperties = ListOfProperties & " '" & flxProperties.TextMatrix(i, 1) & "', "
      End If
   Next i
   If Len(ListOfProperties) > 0 Then ListOfProperties = Left(ListOfProperties, Len(ListOfProperties) - 2)
End Function


Private Function ListOfFunds() As String
    Dim i As Integer
    For i = 1 To flxInFunds.Rows - 1
         If flxInFunds.TextMatrix(i, 0) = "X" Then
                ListOfFunds = ListOfFunds & "'" & flxInFunds.TextMatrix(i, 3) & "', "
         End If
    Next
    If Len(ListOfFunds) > 0 Then ListOfFunds = Left(ListOfFunds, Len(ListOfFunds) - 2)
End Function

Private Function GetBalanceSupplier() As Double
    Dim PurchaseLedgerControl As String
    Dim dblAmt As Double
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn, "Purchase Ledger Control", szSelectedClient)
    Dim rsNLposting As New ADODB.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & PurchaseLedgerControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoConn.Close
    Set adoConn = Nothing
    GetBalanceSupplier = dblAmt
End Function
Private Function GetBalanceAgent() As Double
    Dim ManagementFeesControl As String
    Dim dblAmt As Double
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    ManagementFeesControl = GetNominalCodeForControlAccount(adoConn, "Managing Agents control Account (B/S)", szSelectedClient)
    Dim rsNLposting As New ADODB.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & ManagementFeesControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoConn.Close
    Set adoConn = Nothing
    GetBalanceAgent = dblAmt
End Function

Private Function GetRentDeposit() As Double
    Dim szSQL As String
'    Dim szSQL1 As String
    Dim szSQL2 As String
'    Dim szSQL3 As String
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt1 As New ADODB.Recordset
    Dim rsReceipt2 As New ADODB.Recordset
    Dim rsReceipt3 As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim dblAmt, dblamt1, dblamt2, dblamt3 As Double
    adoConn.Open getConnectionString
'tlbBankPayment
'BANK_AC
'TRAN_TYPE
'DEPT_ID
'propertyID
'clientID
'NET_AMOUNT
    'From unit Id i Ned to build a relation with the selected properties
    szSQL = "Select  SUM(SWITCH(TYPE=1,R.Amount,TYPE=2,R.Amount,TYPE=3,-R.Amount,TYPE=4,-R.Amount,TYPE=23,-R.Amount)) as DR from tlbReceipt R,Fund F where TYPE IN(1,2,3,4,23) AND R.FundID=F.FundID and F.FundCode='TENANTDEPOSIT' AND ClientID ='" & _
             szSelectedClient & "'"
'    szSQL1 = "Select  SUM(R.Amount)  as DR from tlbReceipt R,Fund F where TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
'              szSelectedClient & "'"
    
    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND B.DEPT_ID=F.FundID and " & _
            "F.FundCode='TENANTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "'"
'    szSQL3 = "Select  SUM(B.NET_AMOUNT)  as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND B.DEPT_ID= cstr(F.FundID) and " & _
'            "F.FundCode='RENTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "'"

    
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
        dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
    End If
    rsReceipt.Close
'    rsReceipt1.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt1.EOF Then
'         dblamt1 = IIf(IsNull(rsReceipt1.Fields.Item("Dr").Value), 0, rsReceipt1.Fields.Item("Dr").Value)
'    End If
'    rsReceipt1.Close
    
    rsReceipt2.Open szSQL2, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt2.EOF Then
        dblamt2 = IIf(IsNull(rsReceipt2.Fields.Item("Dr").Value), 0, rsReceipt2.Fields.Item("Dr").Value)
    End If
    rsReceipt2.Close
    
'    rsReceipt3.Open szSQL3, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt3.EOF Then
'        dblamt3 = IIf(IsNull(rsReceipt3.Fields.Item("Dr").Value), 0, rsReceipt3.Fields.Item("Dr").Value)
'    End If
'    rsReceipt3.Close
'    dblamt2 = 0
'    dblamt1 = 0
'    dblAmt = 0
    GetRentDeposit = dblAmt + dblamt2
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBalance(szType As String) As Double
    'szType is the suppplier type from supplier table
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset
   adoConn.Open getConnectionString
   Dim szaClientBal(1, 1) As String

   szSQL = "SELECT  SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Client C, Supplier S " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND C.ClientID=S.SupplierID AND P.SageAccountNumber = C.ClientID " & _
           "and  C.ClientID='" & szSelectedClient & "' AND S.Type='" & szType & "'  "

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaClientBal(1, iIndex) = IIf(IsNull(adoPayDr.Fields.Item("Dr").Value), 0, adoPayDr.Fields.Item("Dr").Value)
      adoPayDr.MoveNext
   Wend
   adoPayDr.Close

   szSQL = "SELECT  SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Client C,  Supplier S " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber = C.ClientID and  C.ClientID='" & szSelectedClient & "' " & _
           "AND C.ClientID=S.SupplierID  AND S.Type='" & szType & "'"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
         szaClientBal(1, iIndex) = IIf(IsNull(adoPayCr.Fields.Item("Cr").Value), 0, adoPayCr.Fields.Item("Cr").Value) 'adoPayCr.Fields.Item("Cr").Value
         adoPayCr.MoveNext
   Wend

   adoPayCr.Close
   GetBalance = szaClientBal(1, 0)
   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function
Private Function GetManagingAgentACBalance() As Double
    Dim adoConn As New ADODB.Connection
    Dim rsAccrualsControlBalance As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
'    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & _
'            NominalCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetManagingAgentACBalance = rsAccrualsControlBalance("SumAmount").Value
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function
Private Function GetAccrualsControlBalance() As Double
    'include no property when  calculating accruals
    Dim adoConn As New ADODB.Connection
    Dim rsAccrualsControlBalance As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsCode As String
    If ListOfProperties = "'" Then Exit Function
    AccrualsCode = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    
    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & AccrualsCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetAccrualsControlBalance = IIf(IsNull(rsAccrualsControlBalance("SumAmount").Value), 0, rsAccrualsControlBalance("SumAmount").Value)
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function

'Private Function getAvailablefunds() As Double
'    Dim intmaxStatementNo As Integer
'    Dim adoconn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoconn.Open getConnectionString
'    Dim szSQL As String
'    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
'    'This is by client
'    'Get ID by Client max ID from RentSummaryStatement
'    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientID='" & szSelectedClient & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        getAvailablefunds = rsRentSummaryStatement!IDbyCL
'    End If
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    adoconn.Close
'    Set adoconn = Nothing
'End Function
Private Function GetLastStatementNoByClient() As Integer
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementNoByClient = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLastStatementID() As Long 'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementID = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLastStatementDateByClient() As String
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select StatementDate from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "' order by StatementNo Desc"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementDateByClient = rsRentSummaryStatement!StatementDate
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function



Private Sub cmdPostRP_Click()
   If MsgBox("Do you want to post to the history?", vbQuestion + vbYesNo, "Post") = vbNo Then Exit Sub
   Dim adoConn As New ADODB.Connection
   Dim adoRST As ADODB.Recordset
   Dim sSQLQuery As String

   adoConn.Open getConnectionString
   Set adoRST = New ADODB.Recordset

   sSQLQuery = "UPDATE tblPurInv " & _
               "SET tblPurInv.HISTORY = TRUE " & _
               "WHERE tblPurInv.TTP = " & CByte(TransactionTakePlace("TTP", "RENT PAYABLE", adoConn)) & " AND " & _
                  "tblPurInv.HISTORY = FALSE AND tblPurInv.UPDATE_SAGE = TRUE"
   adoRST.Open sSQLQuery, adoConn, adOpenStatic, adLockReadOnly
   
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdProduceClientSummeryStatement_Click()
   'make visi
    
End Sub

Private Sub cmdPosttoHistory_Click()
    Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   adoConn.Execute "Update rentSummaryStatement set PostToHistory=true where statementID =" & szCurrentStatementID & ""
   Call loadflxPayFees("")
   Call loadflxPayFeesHistory
   adoConn.Close
   Set adoConn = Nothing
   MsgBox "Rent Summary Statement has been posted to history", vbInformation, "Posted to history"
End Sub

Public Sub cmdPreViewGenDmds_Click()
    Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    For rCount = 1 To flxPayFees.Rows - 1
         If flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If flxPayFees.TextMatrix(rCount, 1) = "+" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select a statement.", vbInformation + vbOKOnly, "statement Selection"
       For rCount = 1 To flxPayFees.Rows - 1
         flxPayFees.TextMatrix(rCount, 0) = ""
       Next
       Exit Sub
    End If
'    If szCurrentStatementID = "" Then
'         Exit Sub
'    End If
'    'flxPayFees.TextMatrix(i, 3) is the statement ID by client
'    '66) It should only be possible to modify a statement provided a rent payable
'    ' PI has not been generated against the statement and a subsequent statement has not been produced.
    If isitPlus = True Then
        'MsgBox "This statement cannot be modified, because a Rent Payable invoice  " & flxPayFees.TextMatrix(selRow, 29) & " has been generated against it.", vbInformation + vbOKOnly, "statement Selection"
        szCurrentStatementID = Replace(flxPayFees.TextMatrix(selRow, 2), "CS", "")
        'Exit Sub
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select a statement to modify", vbInformation + vbOKOnly, "statement Selection"
        Exit Sub
    End If
    'flxPayFees.TextMatrix(i, 29)
    'PI has not been generated against the statement OR a subsequent statement has not been produced.
    Call configflxDetailsTransaction
    LoadForm frmRentPayableModification
    frmRentPayableModification.Caption = "Finalise/Modify Client Statement"
    frmRentPayableModification.bEditMode = True
    frmRentPayableModification.Top = Me.Top + 600
    
    frmRentPayableModification.cmdSave.Caption = "Modify Statement"
    frmRentPayableModification.cmdFinalizeStatement.Visible = True
    frmRentPayableModification.szCurrentStatementID = szCurrentStatementID
    frmRentPayableModification.flxBankAccounts.Clear
    frmRentPayableModification.flxProperties.Clear
    frmRentPayableModification.flxInFunds.Clear

End Sub


Private Sub PrintCSlineByLineNew()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim CSID As String
   On Error GoTo Err
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
    szCurrentStatementID = flxPayFees.TextMatrix(selRow, 2)
    CSID = Replace(szCurrentStatementID, "CS", "")
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblCrReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    
    Dim DateTO As String
    If szCurrentStatementID = "" Then
        MsgBox "Please select a statement", vbInformation, "Warning"
        Exit Sub
    End If
'    adoConn.Execute "Update DemandSplitRecords DS,DemandRecords D,Units U,Property P  set  ReportCsShowFlag= '',ReportNetAmountS=0,ReportVATAmountS=0,ReportReceivedAmountS= 0,ReportDateFromS=Null," & _
'         " ReportCreditAmountS=0,reportOSAmountS=0,ReportDateTOS =null,ReportDemandTypeDescS= '' where D.DemandID=DS.DemandID  and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
'         flxPayFees.TextMatrix(selRow, 4) & "'"
    
       
    Dim szPreviousStatementDate As Date
    Dim szStatementDate As Date
    szCurrentStatementID = Replace(szCurrentStatementID, "CS", "")
    rsRentSummaryStatement.Open "Select ClientIDLandlordID,ListOfFundId,PreviousStatementDate,StatementDate  from RentSummaryStatement where statementID=" & _
                szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly ' group by D.DemandId", adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
            szListofFunds = rsRentSummaryStatement("ListOfFundId").Value
            szPreviousStatementDate = rsRentSummaryStatement("PreviousStatementDate").Value
            szStatementDate = rsRentSummaryStatement("StatementDate").Value
            szSelectedClient = rsRentSummaryStatement("ClientIDLandlordID").Value
    End If
    rsRentSummaryStatement.Close
    
   Dim strDueDate As String
   Dim rsDemandSplitAmt As New ADODB.Recordset
   Dim iCount As Integer
   Dim dblDemandSplitamt As Double
   Dim szSQL As String
   Dim SQLforInsert As String
   Dim adoOsamount As New ADODB.Recordset
    
   adoConn.Execute "Delete from ReportClientStatementDemands"
   adoConn.Execute "Delete from ReportClientStatementPurchases"
'Type 1 SI

   SQLforInsert = " Select " & szCurrentStatementID & " as StatementID,TransactionID,ClientID,PropertyID,DemandID,SplitID,Sageaccountnumber,'',Type as DemandTypeDesc,TypeOfDemand,DueDate,D.DateFrom,D.DateTo,switch(Type=1,D.Amount,Type=2,-D.Amount)as NETAmount," & _
                    "switch(Type=1,D.VATAmount,Type=2,-D.VATAmount),switch(Type=1,(NetAmounts+VATAmounts))as ReceivedAmountS,switch(Type=2,-NetAmounts-VATAmounts) as CreditAmount,T.OSAmount from " & _
                    "(Select D.DemandID,T.TransactionID,T.sageaccountnumber,U.Unitnumber,U.PropertyID,T.ClientID,D.TotalAmount,D.SplitID,T.Type,D.TypeOfDemand,D.DueDate,D.DateFrom,D.DateTO,D.Amount,T.OSAmount,D.VATAmount" & _
                    " from DemandSplitRecords D INNER JOIN  ((tlbReceipt T INNER JOIN ( SELECT Distinct A.ToTran as TrxId FROM" & _
                    " RptTransactionsSplit A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND" & _
                    " RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false  AND " & _
                    " RS.ClientStatementID=" & szCurrentStatementID & "   group by  A.ToTran ) B ON   B.trxID=T.TransactionID) INNER JOIN units U ON T.UnitID=U.UnitNumber)" & _
                    " on T.Demandref=D.DemandID) X LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofSI," & _
                    " A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit  A" & _
                    " , tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false  AND  RS.ClientStatementID=" & szCurrentStatementID & "" & _
                    "   group by  A.ToTran, A.SplitIDofSI, A.fundID  Union SELECT" & _
                    " A.FromTran, A.SplitIDofSI, A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit" & _
                    " A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND A.Deleteflag=false    and" & _
                    " A.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " # And A.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & " #  group" & _
                    " by  A.FromTran, A.SplitIDofSI, A.FundID) Y ON X.TransactionID=Y.FromTran  AND X.splitID=Y.SplitIDofSI"
             adoConn.Execute "Insert into ReportClientStatementDemands(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber" & _
            ",TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,ReceivedAmountS,CreditAmount,OSAmount)" & _
            SQLforInsert
 ' putis code in middle  of A.ToTran ) B
'            UNION SELECT  Distinct" & _
'                    " A.FromTran as TrxId FROM RptTransactionsSplit A, tlbReceipt R,  tlbReceiptSplit RS" & _
'                    "  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND" & _
'                    " A.Deleteflag=false and A.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And A.Allocdate > # " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#" & _
'                    "  group by  A.FromTran
'
'
            'There is theory behind allocating SC I had note/found before now I dont know. Because Receipt and payment are marking  basis. but we are not marking SC. Same for payment.answer is date Range
           'Finding OSamount column here. which is independant of CSID ( or can I use the SI osamount without being calculative answer is: no)
             'Exit Sub
           SQLforInsert = " Select OSAmount,M.Netamount,vatamount, ReceiptAMOUNTS,SITrxID,SplitID from  ReportClientStatementDemands M INNER JOIN  (SELECT A.ToTran as FromTran," & _
           "A.SplitIDofSI,Sum(A.NetAmount+A.VATAMOUNT) AS ReceiptAMOUNTS FROM RptTransactionsSplit  A , tlbReceipt R,  tlbReceiptSplit RS  where " & _
           "R.RDate <=#" & Format(szStatementDate, "DD MMM yyyy") & "# AND " & _
           "A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false " & _
           " group by  A.ToTran, A.SplitIDofSI )  AS X ON M.SITrxID = X.FromTran ANd X.SplitIDofSI=M.SplitID"
            adoOsamount.Open SQLforInsert, adoConn, adOpenStatic, adLockReadOnly
            If Not adoOsamount.EOF Then
                adoOsamount.MoveFirst
            End If
            While Not adoOsamount.EOF
                adoConn.Execute "Update  ReportClientStatementDemands SET Osamount =" & adoOsamount("NetAmount").Value + adoOsamount("vatamount").Value - adoOsamount("ReceiptAMOUNTS").Value & "" & _
                                " where SITrxID=" & adoOsamount("SITrxID").Value & " and  SplitID=" & adoOsamount("SplitID").Value & ""
                adoOsamount.MoveNext
            Wend
            'adoOsamount.UpdateBatch
            adoOsamount.Close
           adoConn.Execute "update ReportClientStatementDemands set CreditAmount=0  where CreditAmount is null"
           adoConn.Execute "update ReportClientStatementDemands set OSAmount=0  where OSAmount is null"
           adoConn.Execute "update ReportClientStatementDemands set ReceivedAmountS=0  where ReceivedAmountS is null"
           
           'Sleep (100)
           'Type 2 SC
           'Here D.[FromTran])=[R].[TransactionID]  R tlbReceipt represents Credit side
           'Taking there some by Sum(D.ReceiptAmount) AS Amt where D is allocation table with date range criteria
           'So now  have collected all allocated amount for credit notes
           'Now I inner join them with DemandsplitRecords DS to get split level Credit note details
           'Finally inserting them into a report table
''            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,R.TransactionID,ClientID,U.PropertyID,D.DemandID,SplitID,D.Sageaccountnumber,'',TransactionType,TypeOfDemand,DueDate, DS.DateFrom,DS.DateTo,  " & _
''               "switch(D.transactionType=1,X.Amt,D.transactionType=2,x.amt)as NETAmount,switch(D.transactionType=1,DS.VATAmount,D.transactionType=2,DS.VATAmount) as VAT, " & _
''                "switch(D.transactionType=2,DS.Amount+DS.VATAmount) as CreditAmount,0 as DS1  from (DemandsplitRecords DS Inner join  " & _
''               "(SELECT R.DemandRef,RC.SlNumber, Sum(D.ReceiptAmount) AS Amt, R.TransactionID, D.SPlitIDofSi FROM RptTransactionsSplit AS D, tlbReceiptSplit AS RS, tlbReceipt AS RC, tlbReceipt AS R  " & _
''               "WHERE (((RC.TransactionID)=[RS].[RptHeader]) AND R.Amount>R.OsAmount AND ((RS.RptTransactionsIDSplit)=[D].[TransactionID]) AND ((D.[FromTran])=[R].[TransactionID]) AND ((D.[deleteflag])=False) AND ((R.Type)=2)  " & _
''               "AND ((R.ClientID)='" & flxPayFees.TextMatrix(selRow, 4) & "')) and D.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And D.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#GROUP BY  R.DemandRef,  " & _
''                "RC.SlNumber, R.TransactionID, D.SPlitIDofSi)X ON X.DemandRef=DS.DemandID and DS.SplitID=X.SPlitIDofSi), DemandRecords D,UNITS U,Property P where  P.PropertyID=U.PropertyID AND  " & _
''                "D.DemandID=DS.DemandID AND D.UnitNumber=U.UnitNumber and D.exclCRNtoCS=false; " '
''            adoConn.Execute "Insert into ReportClientStatementDemands(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber," & _
''            "TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,CreditAmount,osAmount)" & _
''            SQLforInsert
''
''            adoConn.Execute "update ReportClientStatementDemands D set ReceivedAmountS=0,D.OSAmount =-(NetAmount+VATAmount)+CreditAmount where TransactionType=2"
''            'Neet to check OS amounts here for acceptance
''            'Credit note needs to show relavent SI. I shall be using UnitNumber field from ReportClientStatementDemandsPreview for showing SI number
''             adoConn.Execute "update ReportClientStatementDemands D,RptTransactionsSplit T, tlbReceipt AS R,DemandRecords DR set D.UnitNumber ='/ SI'& R.slnumber where D.TransactionType=2 " & _
''                    " AND T.ToTran=R.TransactionID and R.DemandRef=Dr.DemandID and D.SITrxID=T.FromTran"
           'writing code for inserting SRR  23    by anol 2023-08-20
            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,TransactionID, R.clientID,PropertyID,R.RDate,SageAccountNumber," & _
                        "23,1,-X.Amt,0,-X.Amt,0,0  from ( (SELECT R.TransactionID,R.UNITID as PropertyID,R.clientID,R.RDate,SageAccountNumber,slnumber,R.NominalCode,Sum(RS.Amount) AS Amt FROM " & _
                        "tlbReceiptSplit AS RS,  tlbReceipt AS R  WHERE R.Amount>R.OsAmount AND ((R.Type)=23)  AND  " & _
                        "(R.ClientID)='" & szSelectedClient & "' AND RS.rptHeader=R.TransactionID AND RS.ClientStatementID=" & szCurrentStatementID & " " & _
                        " GROUP BY R.TransactionID,R.clientID,R.UNITID,R.RDate,slnumber,SageAccountNumber,R.NominalCode)X )"
                        
               adoConn.Execute "Insert into ReportClientStatementDemands(StatementID,SITrxID,ClientID,PropertyID,DueDate,SageAccountNumber," & _
            "TransactionType,SplitID,NetAmount,VATAmount,ReceivedAmountS,CreditAmount,osAmount)" & _
            SQLforInsert
            
'                adoconn.Execute "update ReportClientStatementDemands D,RptTransactionsSplit T, tlbReceipt AS R, DemandSplitRecords AS RS set D.UnitNumber ='SRR'& R.slnumber, " & _
'                    " D.DateFrom=RS.DateFrom, D.DateTo=RS.DateTo  where D.TransactionType=23 AND T.SplitIDofSI=RS.SplitID AND R.Demandref=RS.DemandID " & _
'                    " AND RS.rptHeader=R.TransactionID AND T.ToTran=R.TransactionID  and D.SITrxID=T.toTran and T.DeleteFlag=false "

  adoConn.Execute "update ReportClientStatementDemands D,RptTransactionsSplit T, tlbReceipt AS R, DemandSplitRecords AS RS set D.UnitNumber ='SRR'& R.slnumber,  " & _
                    "      D.DateFrom=RS.DateFrom, D.DateTo=RS.DateTo  where D.TransactionType=23  AND R.Demandref=RS.DemandID " & _
                    "     AND T.fromTran=R.TransactionID  and D.SITrxID=T.toTran and T.DeleteFlag=false   "
                    
                    

            'Insert expense for PI Type 6
            SQLforInsert = "  Select StatementID, '6' as type,D.ParentID as MY_ID,SplitID,TransactionID,T.ClientID, PropertID," & _
                            " T.Pdate,SageAccountNumber,NOMINAL_CODE,D.Net_Amount,D.VAT,PaidAmounts from (Select  " & szCurrentStatementID & _
                             " as StatementID, D.ParentID as MY_ID,D.TRAN_ID as SplitID, D.ParentID,T.ClientID, D.TRANS as PropertID" & _
                            " , T.Pdate,T.SageAccountNumber,NOMINAL_CODE,T.TransactionID,D.Net_Amount,D.VAT from tblPurInvSRec D INNER JOIN  (Select" & _
                            " * from tlbPayment T INNER JOIN (SELECT Distinct A.ToTran as TrxId" & _
                            " FROM PayTransactionsSplit A, tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID" & _
                            " AND RS.PayHeader=R.TransactionID and RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "'  AND Deleteflag=false  AND" & _
                            "  RS.ClientStatementID=" & szCurrentStatementID & " group by  A.ToTran)X ON T.TransactionID=X.TrxId)N ON" & _
                            " N.PI=D.ParentID ) Y LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofPI," & _
                            "  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit  A" & _
                            " , tlbPayment R, tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "'  AND Deleteflag=false  AND  RS.ClientStatementID=" & szCurrentStatementID & "" & _
                            "  group by  A.ToTran, A.SplitIDofPI, A.fundID  Union SELECT A.FromTran," & _
                            " A.SplitIDofPI,  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit A," & _
                            " tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND A.Deleteflag=false and A.Allocdate <=#" & Format(szStatementDate, "DD MMM yyyy") & "#" & _
                            " And A.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#  group by  A.FromTran, A.SplitIDofPI," & _
                            " A.FundID )Z  ON Y.TransactionID=Z.FromTran AND Y.SplitID=cstr(Z.SplitIDofPI)"
           adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,Netamount,VatAmount,PaymentAmount)" & _
                    SQLforInsert
                    
                    'Rent payable should be removed from the records
                    'Manaagement Fee we are inserting from above
                    Debug.Print "start1" & time
                     adoConn.Execute "DELETE A.* FROM ReportClientStatementPurchases  AS A " & _
                                "LEFT JOIN tblPurInv   AS B ON A.MY_ID = B.MY_ID" & _
                                 " WHERE B.isRentPayable = True"
                                 Debug.Print "end1" & time
           adoConn.Execute "update ReportClientStatementPurchases SET CreditAmount=0  where CreditAmount is null"
            adoConn.Execute "update ReportClientStatementPurchases SET OSAmount=0  where OSAmount is null"
            '2023-/08/16 by anol
            Debug.Print "start2" & time
             adoConn.Execute "update ReportClientStatementPurchases R,tlbPayment P SET R.OSAmount=P.OSAmount where P.PI=R.MY_ID "
           Debug.Print "end2" & time
           
           'adoconn.Execute "update ReportClientStatementPurchases SET OSAmount=Netamount+vatamount-paymentAmount" ' I cant fully remember were it date based?
           'adoConn.Execute "update ReportClientStatementPurchasesPreview R,tlbPayment P SET R.OSAmount=P.OSAmount where R.MY_ID=R.MY_ID" ' I cant fully remember were it date based?
           adoConn.Execute "UPDATE ReportClientStatementPurchasesPreview AS R INNER JOIN tlbPayment AS P ON R.MY_ID = P.PI SET R.OSAmount = P.OSAmount"
            adoConn.Execute "update ReportClientStatementPurchases a,tblPurInv b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where b.MY_ID=a.MY_ID and C.TYPE_ID=a.Type"
                    
           'Insert code for type 7 PI Credit note
'''Debug.Print "mw1" & time
        SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'7' as type,INV.MY_ID,TRAN_ID,TransactionID,ClientID,P.PropertyID,INV.TRAN_DATE,INV.SUPP_AC,DS.NOMINAL_CODE," & _
                "-DS.Net_amount, -DS.VAT,0, -X.Amt,0  from (tblPurInvSRec DS Inner join " & _
                "(SELECT P.PI,P.TransactionID,Sum(D.PaymentAmount) AS Amt, D.SPlitIDofPI FROM PayTransactionsSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((D.[FromTran])=[P].[TransactionID]) AND " & _
                "((D.[deleteflag])=False) AND ((P.Type)=7)  AND ((P.ClientID)='" & flxPayFees.TextMatrix(selRow, 4) & "') and D.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & _
                " #  And D.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "# GROUP BY  P.PI,   P.TransactionID,  P.TransactionID,D.SPlitIDofPI)X " & _
                "ON X.PI=DS.ParentID and DS.TRAN_ID=cstr(X.SPlitIDofPI)), tblPurInv INV,Property P where  P.PropertyID=INV.PropertyID AND  INV.My_ID=DS.ParentID "
                    adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                       Debug.Print "mw2" & time
'    'Insert code for type 24 PI Purchase Payment Refund
'             SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,TransactionID, 1 as TRAN_ID,TransactionID,P.clientID,PropertyID,P.PDate,SageAccountNumber,P.NominalCode," & _
'                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.clientID,P.PDate,SageAccountNumber,slnumber,P.NominalCode,Sum(D.Amount) AS Amt FROM " & _
'                        "tlbPaymentSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
'                        "(P.ClientID)='" & szSelectedClient & "' AND D.PayHeader=P.TransactionID AND D.ClientStatementID=" & szCurrentStatementID & " " & _
'                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,slnumber,SageAccountNumber,P.NominalCode)X )"'
'                    adoconn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
'                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
'                            SQLforInsert

         adoConn.Execute "Update   tlbPaymentSplit PS,tlbPayment P,ReportClientStatementPurchases R  set R.PaymentDescription=PS.Description where   Ps.PayHeader=p.transactionID and  P.type=R.Type and P.transactionID=R.TransactionID "
         adoConn.Execute "Update  Supplier S, ReportClientStatementPurchases R,GlobalData G set R.VATAmount= Round((R.NetAmount * 20/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTax=true "
         adoConn.Execute "Update  Supplier S, ReportClientStatementPurchases R,GlobalData G set R.NetAmount= Round((R.NetAmount * 100/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTAx=true "
                            Debug.Print "mw3" & time

 
     
     
     

          ' Code for transfer data from snapshot table into expenditure table
             If chkExcludeSupOS.Value = 1 Then
                                Debug.Print "start3" & time
                            'Select P.MY_ID from ReportClientStatementPurchases P, ClientStatementPurchasesnapdhot R where P.MY_ID=R.MYID
'                            SQLforInsert = "select StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID, " & _
'                            "NOMINAL_CODE,NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef from ClientStatementPurchasesSnapshot CS " & _
'                            "where isManagementFee=false AND StatementID=" & szCurrentStatementID & "  AND CS.MY_ID not in (" & _
'                            "Select P.MY_ID from ReportClientStatementPurchases P, ClientStatementPurchasesSnapshot R where P.MY_ID=R.MY_ID) "
                             SQLforInsert = "SELECT P.StatementID, P.Type, P.MY_ID, P.SplitID, P.TransactionID, P.ClientID, P.PropertyID, P.TranDate, P.SupplierID, " & _
                                            "P.NOMINAL_CODE, P.NetAmount, P.VATAmount, P.PaymentAmount, P.CreditAmount, P.osAmount, P.PaymentRef " & _
                                            "FROM ReportClientStatementPurchases P " & _
                                            "LEFT JOIN ClientStatementPurchasesSnapshot CS ON P.MY_ID = CS.MY_ID " & _
                                            "WHERE CS.isManagementFee = False " & _
                                            "AND CS.StatementID = " & szCurrentStatementID & " " & _
                                            "AND CS.MY_ID IS NULL;"
            

                         adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef)" & _
                            SQLforInsert
                            Debug.Print "end 3" & time
             End If

           If chkShowDue.Value = 1 Then
                            Debug.Print "start4" & time
                            SQLforInsert = "select StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID, " & _
                            "NOMINAL_CODE,NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef from ClientStatementPurchasesSnapshot CS " & _
                            "where isManagementFee=true AND StatementID=" & szCurrentStatementID & "  AND CS.MY_ID not in (" & _
                            "Select P.MY_ID from ReportClientStatementPurchases P, ClientStatementPurchasesSnapshot R where P.MY_ID=R.MY_ID) "
                            
                            SQLforInsert = "SELECT P.StatementID, P.Type, P.MY_ID, P.SplitID, P.TransactionID, P.ClientID, P.PropertyID, P.TranDate, P.SupplierID, " & _
                                            "P.NOMINAL_CODE, P.NetAmount, P.VATAmount, P.PaymentAmount, P.CreditAmount, P.osAmount, P.PaymentRef " & _
                                            "FROM ReportClientStatementPurchases P " & _
                                            "LEFT JOIN ClientStatementPurchasesSnapshot CS ON P.MY_ID = CS.MY_ID " & _
                                            "WHERE CS.isManagementFee = true " & _
                                            "AND CS.StatementID = " & szCurrentStatementID & " " & _
                                            "AND CS.MY_ID IS NULL;"

                         adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,PaymentRef)" & _
                            SQLforInsert
                            
                            Debug.Print "end4" & time
             End If

             
'        adoconn.Execute "update ReportClientStatementPurchases a,tlbPayment b,tlbTransactionTypes C  " & _
'        "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber,a.PaymentDescription=b.Details  where b.TransactionID=a.TransactionID and C.TYPE_ID=a.Type and a.Type=24 "
         'Insert code for type 24 PI Purchase Payment Refund
        'for type 24 there is no corresponding entry in tlbPurInv. so need to remove that relationship
        'TRAN_ID is the split ID in tblPurInvSRec table
'         SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,TransactionID, 1 as TRAN_ID,slnumber,P.clientID,PropertyID,P.PDate,SageAccountNumber,P.NominalCode," & _
'                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.clientID,P.PDate,SageAccountNumber,slnumber,P.NominalCode,Sum(D.Amount) AS Amt FROM " & _
'                        "tlbPaymentSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
'                        "(P.ClientID)='" & szSelectedClient & "' AND D.PayHeader=P.TransactionID AND D.ClientStatementPrevID=" & szCurrentStatementID & " " & _
'                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,slnumber,SageAccountNumber,P.NominalCode)X )"

         SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,P.TransactionID, 1 as TRAN_ID,P.slnumber,P.Details,P.clientID,PropertyID,P.PDate,P.SageAccountNumber,PD.NOMINAL_CODE," & _
                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.clientID,P.Details,P.PDate,P.SageAccountNumber,P.slnumber,PD.NOMINAL_CODE,Sum(D.Amount) AS Amt FROM " & _
                        "tlbPaymentSplit AS D,  tlbPayment AS P,  tlbPayment AS Q , PayTransactions PS,tblPurInvSRec PD  WHERE PD.ParentID=Q.PI and P.TransactionID=PS.totran and ps.fromtran=Q.transactionID and P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
                        "(P.ClientID)='" & szSelectedClient & "' AND ps.DeleteFlag=false AND D.PayHeader=P.TransactionID AND D.ClientStatementID=" & szCurrentStatementID & " " & _
                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,P.slnumber,P.SageAccountNumber,PD.NOMINAL_CODE,P.Details)X )"
                        

                    adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,PaymentDescription,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                            
                            
'                                           adoconn.Execute "update ReportClientStatementPurchases a,tblPurInv b,tlbTransactionTypes C  " & _
'                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where b.MY_ID=a.MY_ID and C.TYPE_ID=a.Type"

 adoConn.Execute "update ReportClientStatementPurchases a,tlbPayment b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where cstr(b.transactionID)=a.MY_ID and C.TYPE_ID=a.Type"
     '*************************Now insert Management fee when OSAmount >0 here  date 13/08/2023****************
     
     
    'Exit Sub
    Dim rsStatementTemplate As New ADODB.Recordset
    Dim strReportName As String
    rsStatementTemplate.Open "Select CSTemplate from client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("CSTemplate").Value), "", rsStatementTemplate("CSTemplate").Value)
    End If
    If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
    Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
    End If
   rsStatementTemplate.Close
   
    Dim rsReportName As New ADODB.Recordset
'    rsReportName.Open "Select LesseeTemplate from Client where clientID='" & flxPayFees.TextMatrix(selRow, 4) & "'", adoConn, adOpenStatic, adLockReadOnly
'    If rsReportName.EOF Then
         'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'    Else
'        If IsNull(rsReportName!LesseeTemplate) Then
'            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'        Else
'            If rsReportName!LesseeTemplate = "" Then
'                Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'            Else
'                Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & rsReportName.Fields.Item("LesseeTemplate").Value)
'            End If
'
'        End If
'    End If
'    rsReportName.Close
    adoConn.Close
    Sleep (500)
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    Report.ParameterFields(1).AddCurrentValue CLng(StrDigitVal(szCurrentStatementID))
    

     Dim selRowTemp As Integer
     selRowTemp = selRow
     If flxPayFees.TextMatrix(selRow, 1) = "+" Or flxPayFees.TextMatrix(selRow, 1) = ">" Then
            'Report as header
            szSelectedClient = flxPayFees.TextMatrix(selRow, 4)
            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRow, 4) 'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRow, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
            Report.ParameterFields(6).AddCurrentValue "0" '0 Means header
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, flxPayFees.TextMatrix(selRow, 4))
            Report.ParameterFields(8).AddCurrentValue 0 '((flxPayFees.TextMatrix(selRow, 10))) 'amount paid to LL no use of this param
            Report.ParameterFields(9).AddCurrentValue -CDbl(GetTotalExpenditure)
            'Report.ParameterFields(10).AddCurrentValue CDbl(GetTotalExpenditure) 'Send supplier OS Amount at parameter 10
            adoConn.Close
      Else
            'Report as Split in LL
             szSelectedClient = flxPayFees.TextMatrix(selRow, 4)
            dblPercentage = Replace(flxPayFees.TextMatrix(selRow, 9), "%", "") 'Take Percenatge from Grid
            Do
                selRowTemp = selRowTemp - 1
            Loop Until (flxPayFees.TextMatrix(selRowTemp, 1) = "+" Or flxPayFees.TextMatrix(selRowTemp, 1) = ">")
            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRowTemp, 4)  'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRowTemp, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRowTemp, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRowTemp, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue dblPercentage
            Report.ParameterFields(6).AddCurrentValue "1" '1 Means Split
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findLandlordAddress(adoConn, flxPayFees.TextMatrix(selRow, 6)) 'take clientlnadlord ID from grid
            Report.ParameterFields(8).AddCurrentValue Val(flxPayFees.TextMatrix(selRow, 11)) 'amount paid to LL' I cannot replace poound sign . So I am not going to display pound sign in the grid
            Report.ParameterFields(9).AddCurrentValue CDbl(GetTotalExpenditure)
            'Report.ParameterFields(10).AddCurrentValue CDbl(GetTotalExpenditure) 'Send supplier OS Amount at parameter 10
            adoConn.Close
      End If
        
    Load frmReport
    frmReport.LoadReportViewer Report
    Exit Sub
Err:
    MsgBox Err.description
    
End Sub
Private Function GetTotalExpenditure() As Currency  'This function works on OS column on the expediture @ CS
            Dim rsPayment As New ADODB.Recordset
            Dim szSQL As String
            Dim adoConn As New ADODB.Connection

            adoConn.Open getConnectionString
            Dim whereProperty As String

            szSQL = "Select  SUM(P.OSAmount) as AMT from ReportClientStatementPurchases P,tblPurInv S where " & _
                    " S.MY_ID=P.MY_ID and isManagementFee=false"
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF And chkExcludeSupOS.Value = 1 Then
                GetTotalExpenditure = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close
            
            szSQL = "Select  SUM(P.OSAmount) as AMT from ReportClientStatementPurchases P,tblPurInv S where " & _
                    " S.MY_ID=P.MY_ID and isManagementFee=true"
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF And chkShowDue.Value = 1 Then
                GetTotalExpenditure = GetTotalExpenditure + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            adoConn.Close
            Set adoConn = Nothing
End Function

Private Sub Export2PdfCSlinebyLine(szOutputpdfto As String)
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim CSID As String
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
    szCurrentStatementID = flxPayFees.TextMatrix(selRow, 2)
    CSID = Replace(szCurrentStatementID, "CS", "")
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblCrReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    
    Dim DateTO As String
    If szCurrentStatementID = "" Then
        MsgBox "Please select a statement", vbInformation, "Warning"
        Exit Sub
    End If
'    adoConn.Execute "Update DemandSplitRecords DS,DemandRecords D,Units U,Property P  set  ReportCsShowFlag= '',ReportNetAmountS=0,ReportVATAmountS=0,ReportReceivedAmountS= 0,ReportDateFromS=Null," & _
'         " ReportCreditAmountS=0,reportOSAmountS=0,ReportDateTOS =null,ReportDemandTypeDescS= '' where D.DemandID=DS.DemandID  and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
'         flxPayFees.TextMatrix(selRow, 4) & "'"
    
       
    Dim szPreviousStatementDate As Date
    Dim szStatementDate As Date
    szCurrentStatementID = Replace(szCurrentStatementID, "CS", "")
    rsRentSummaryStatement.Open "Select ListOfFundId,PreviousStatementDate,StatementDate  from RentSummaryStatement where statementID=" & _
                szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly ' group by D.DemandId", adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
            szListofFunds = rsRentSummaryStatement("ListOfFundId").Value
            szPreviousStatementDate = rsRentSummaryStatement("PreviousStatementDate").Value
            szStatementDate = rsRentSummaryStatement("StatementDate").Value
    End If
    
   Dim strDueDate As String
   Dim rsDemandSplitAmt As New ADODB.Recordset
   Dim iCount As Integer
   Dim dblDemandSplitamt As Double
   Dim szSQL As String
   Dim SQLforInsert As String
   Dim adoOsamount As New ADODB.Recordset
    
   adoConn.Execute "Delete from ReportClientStatementDemands"
   adoConn.Execute "Delete from ReportClientStatementPurchases"
'Type 1 SI

   SQLforInsert = " Select " & szCurrentStatementID & " as StatementID,TransactionID,ClientID,PropertyID,DemandID,SplitID,Sageaccountnumber,'',Type as DemandTypeDesc,TypeOfDemand,DueDate,D.DateFrom,D.DateTo,switch(Type=1,D.Amount,Type=2,-D.Amount)as NETAmount," & _
                    "switch(Type=1,D.VATAmount,Type=2,-D.VATAmount),switch(Type=1,(NetAmounts+VATAmounts))as ReceivedAmountS,switch(Type=2,-NetAmounts-VATAmounts) as CreditAmount,T.OSAmount from " & _
                    "(Select D.DemandID,T.TransactionID,T.sageaccountnumber,U.Unitnumber,U.PropertyID,T.ClientID,D.TotalAmount,D.SplitID,T.Type,D.TypeOfDemand,D.DueDate,D.DateFrom,D.DateTO,D.Amount,T.OSAmount,D.VATAmount" & _
                    " from DemandSplitRecords D INNER JOIN  ((tlbReceipt T INNER JOIN ( SELECT Distinct A.ToTran as TrxId FROM" & _
                    " RptTransactionsSplit A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND" & _
                    " RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false  AND " & _
                    " RS.ClientStatementID=" & szCurrentStatementID & "   group by  A.ToTran ) B ON   B.trxID=T.TransactionID) INNER JOIN units U ON T.UnitID=U.UnitNumber)" & _
                    " on T.Demandref=D.DemandID) X LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofSI," & _
                    " A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit  A" & _
                    " , tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false  AND  RS.ClientStatementID=" & szCurrentStatementID & "" & _
                    "   group by  A.ToTran, A.SplitIDofSI, A.fundID  Union SELECT" & _
                    " A.FromTran, A.SplitIDofSI, A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit" & _
                    " A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND A.Deleteflag=false    and" & _
                    " A.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " # And A.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & " #  group" & _
                    " by  A.FromTran, A.SplitIDofSI, A.FundID) Y ON X.TransactionID=Y.FromTran  AND X.splitID=Y.SplitIDofSI"
             adoConn.Execute "Insert into ReportClientStatementDemands(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber" & _
            ",TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,ReceivedAmountS,CreditAmount,OSAmount)" & _
            SQLforInsert
 ' putis code in middle  of A.ToTran ) B
'            UNION SELECT  Distinct" & _
'                    " A.FromTran as TrxId FROM RptTransactionsSplit A, tlbReceipt R,  tlbReceiptSplit RS" & _
'                    "  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND" & _
'                    " A.Deleteflag=false and A.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And A.Allocdate > # " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#" & _
'                    "  group by  A.FromTran
'
'
            'There is theory behind allocating SC I had note/found before now I dont know. Because Receipt and payment are marking  basis. but we are not marking SC. Same for payment.answer is date Range
           'Finding OSamount column here. which is independant of CSID ( or can I use the SI osamount without being calculative answer is: no)
             'Exit Sub
           SQLforInsert = " Select OSAmount,M.Netamount,vatamount, ReceiptAMOUNTS,SITrxID,SplitID from  ReportClientStatementDemands M INNER JOIN  (SELECT A.ToTran as FromTran," & _
           "A.SplitIDofSI,Sum(A.NetAmount+A.VATAMOUNT) AS ReceiptAMOUNTS FROM RptTransactionsSplit  A , tlbReceipt R,  tlbReceiptSplit RS  where " & _
           "R.RDate <=#" & Format(szStatementDate, "DD MMM yyyy") & "# AND " & _
           "A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND Deleteflag=false " & _
           " group by  A.ToTran, A.SplitIDofSI )  AS X ON M.SITrxID = X.FromTran ANd X.SplitIDofSI=M.SplitID"
            adoOsamount.Open SQLforInsert, adoConn, adOpenStatic, adLockReadOnly
            If Not adoOsamount.EOF Then
                adoOsamount.MoveFirst
            End If
            While Not adoOsamount.EOF
                adoConn.Execute "Update  ReportClientStatementDemands SET Osamount =" & adoOsamount("NetAmount").Value + adoOsamount("vatamount").Value - adoOsamount("ReceiptAMOUNTS").Value & "" & _
                                " where SITrxID=" & adoOsamount("SITrxID").Value & " and  SplitID=" & adoOsamount("SplitID").Value & ""
                adoOsamount.MoveNext
            Wend
            'adoOsamount.UpdateBatch
            adoOsamount.Close
           adoConn.Execute "update ReportClientStatementDemands set CreditAmount=0  where CreditAmount is null"
           adoConn.Execute "update ReportClientStatementDemands set OSAmount=0  where OSAmount is null"
           adoConn.Execute "update ReportClientStatementDemands set ReceivedAmountS=0  where ReceivedAmountS is null"
           
           'Sleep (100)
           'Type 2 SC
           'Here D.[FromTran])=[R].[TransactionID]  R tlbReceipt represents Credit side
           'Taking there some by Sum(D.ReceiptAmount) AS Amt where D is allocation table with date range criteria
           'So now  have collected all allocated amount for credit notes
           'Now I inner join them with DemandsplitRecords DS to get split level Credit note details
           'Finally inserting them into a report table
            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,R.TransactionID,ClientID,U.PropertyID,D.DemandID,SplitID,D.Sageaccountnumber,'',TransactionType,TypeOfDemand,DueDate, DS.DateFrom,DS.DateTo,  " & _
               "switch(D.transactionType=1,X.Amt,D.transactionType=2,x.amt)as NETAmount,switch(D.transactionType=1,DS.VATAmount,D.transactionType=2,DS.VATAmount) as VAT, " & _
                "switch(D.transactionType=2,DS.Amount+DS.VATAmount) as CreditAmount,0 as DS1  from (DemandsplitRecords DS Inner join  " & _
               "(SELECT R.DemandRef,RC.SlNumber, Sum(D.ReceiptAmount) AS Amt, R.TransactionID, D.SPlitIDofSi FROM RptTransactionsSplit AS D, tlbReceiptSplit AS RS, tlbReceipt AS RC, tlbReceipt AS R  " & _
               "WHERE (((RC.TransactionID)=[RS].[RptHeader]) AND R.Amount>R.OsAmount AND ((RS.RptTransactionsIDSplit)=[D].[TransactionID]) AND ((D.[FromTran])=[R].[TransactionID]) AND ((D.[deleteflag])=False) AND ((R.Type)=2)  " & _
               "AND ((R.ClientID)='" & flxPayFees.TextMatrix(selRow, 4) & "')) and D.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And D.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#GROUP BY  R.DemandRef,  " & _
                "RC.SlNumber, R.TransactionID, D.SPlitIDofSi)X ON X.DemandRef=DS.DemandID and DS.SplitID=X.SPlitIDofSi), DemandRecords D,UNITS U,Property P where  P.PropertyID=U.PropertyID AND  " & _
                "D.DemandID=DS.DemandID AND D.UnitNumber=U.UnitNumber and D.exclCRNtoCS=false; " '
            adoConn.Execute "Insert into ReportClientStatementDemands(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber," & _
            "TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,CreditAmount,osAmount)" & _
            SQLforInsert

            adoConn.Execute "update ReportClientStatementDemands D set ReceivedAmountS=0,D.OSAmount =-(NetAmount+VATAmount)+CreditAmount where TransactionType=2"
            'Neet to check OS amounts here for acceptance
            'Credit note needs to show relavent SI. I shall be using UnitNumber field from ReportClientStatementDemandsPreview for showing SI number
             adoConn.Execute "update ReportClientStatementDemands D,RptTransactionsSplit T, tlbReceipt AS R,DemandRecords DR set D.UnitNumber ='/ SI'& R.slnumber where D.TransactionType=2 " & _
                    " AND T.ToTran=R.TransactionID and R.DemandRef=Dr.DemandID and D.SITrxID=T.FromTran"
                    

            'Insert expense for PI Type 6
            SQLforInsert = "  Select StatementID, '6' as type,D.ParentID as MY_ID,SplitID,TransactionID,T.ClientID, PropertID," & _
                            " T.Pdate,SageAccountNumber,NOMINAL_CODE,D.Net_Amount,D.VAT,PaidAmounts from (Select  " & szCurrentStatementID & _
                             " as StatementID, D.ParentID as MY_ID,D.TRAN_ID as SplitID, D.ParentID,T.ClientID, D.TRANS as PropertID" & _
                            " , T.Pdate,T.SageAccountNumber,NOMINAL_CODE,T.TransactionID,D.Net_Amount,D.VAT from tblPurInvSRec D INNER JOIN  (Select" & _
                            " * from tlbPayment T INNER JOIN (SELECT Distinct A.ToTran as TrxId" & _
                            " FROM PayTransactionsSplit A, tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID" & _
                            " AND RS.PayHeader=R.TransactionID and RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "'  AND Deleteflag=false  AND" & _
                            "  RS.ClientStatementID=" & szCurrentStatementID & " group by  A.ToTran)X ON T.TransactionID=X.TrxId)N ON" & _
                            " N.PI=D.ParentID ) Y LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofPI," & _
                            "  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit  A" & _
                            " , tlbPayment R, tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "'  AND Deleteflag=false  AND  RS.ClientStatementID=" & szCurrentStatementID & "" & _
                            "  group by  A.ToTran, A.SplitIDofPI, A.fundID  Union SELECT A.FromTran," & _
                            " A.SplitIDofPI,  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit A," & _
                            " tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' AND A.Deleteflag=false and A.Allocdate <=#" & Format(szStatementDate, "DD MMM yyyy") & "#" & _
                            " And A.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "#  group by  A.FromTran, A.SplitIDofPI," & _
                            " A.FundID )Z  ON Y.TransactionID=Z.FromTran AND Y.SplitID=cstr(Z.SplitIDofPI)"
           adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,Netamount,VatAmount,PaymentAmount)" & _
                    SQLforInsert
                    'Payment description and  payment ref needs to be included
                    
           adoConn.Execute "update ReportClientStatementPurchases SET CreditAmount=0  where CreditAmount is null"
           adoConn.Execute "update ReportClientStatementPurchases SET OSAmount=0  where OSAmount is null"
           adoConn.Execute "update ReportClientStatementPurchases SET OSAmount=Netamount+vatamount-paymentAmount" ' I cant fully remember were it date based?
            adoConn.Execute "update ReportClientStatementPurchases a,tblPurInv b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where b.MY_ID=a.MY_ID and C.TYPE_ID=a.Type"
                            
                    'There is  no os amount calculation for expense side calculations
                    'Also no credit note is coming in this section
                    
           'Insert code for type 7 PI Credit note


'        SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'7' as type,INV.MY_ID,TRAN_ID,TransactionID,ClientID,P.PropertyID,INV.TRAN_DATE,INV.SUPP_AC,DS.NOMINAL_CODE," & _
'                "-DS.Net_amount, -DS.VAT,0, -X.Amt,0  from (tblPurInvSRec DS Inner join " & _
'                "(SELECT P.PI,P.TransactionID,Sum(D.PaymentAmount) AS Amt, D.SPlitIDofPI FROM PayTransactionsSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((D.[FromTran])=[P].[TransactionID]) AND " & _
'                "((D.[deleteflag])=False) AND ((P.Type)=7)  AND ((P.ClientID)='" & flxPayFees.TextMatrix(selRow, 4) & "') and D.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And D.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "# GROUP BY  P.PI,   P.TransactionID,  P.TransactionID,D.SPlitIDofPI)X " & _
'                "ON X.PI=DS.ParentID and DS.TRAN_ID=cstr(X.SPlitIDofPI)), tblPurInv INV,Property P where  P.PropertyID=INV.PropertyID AND  INV.My_ID=DS.ParentID "
                
                
        SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'7' as type,INV.MY_ID,TRAN_ID,TransactionID,ClientID,P.PropertyID,INV.TRAN_DATE,INV.SUPP_AC,DS.NOMINAL_CODE," & _
                "-DS.Net_amount, -DS.VAT,0, -X.Amt,0  from (tblPurInvSRec DS Inner join " & _
                "(SELECT P.PI,P.TransactionID,Sum(D.PaymentAmount) AS Amt, D.SPlitIDofPI FROM PayTransactionsSplit AS D,  tlbPayment AS P , tlbPayment AS Q  WHERE P.Amount>P.OsAmount AND ((D.[FromTran])=[P].[TransactionID]) AND " & _
                "((D.[deleteflag])=False) AND ((D.[TOTran])=[Q].[TransactionID]) AND Q.NominalCODE='" & szSelectedBankAccount & "'  AND ((P.Type)=7)  AND ((P.ClientID)='" & flxPayFees.TextMatrix(selRow, 4) & "') and D.Allocdate <=# " & Format(szStatementDate, "DD MMM yyyy") & " #  And D.Allocdate ># " & Format(szPreviousStatementDate, "DD MMM yyyy") & "# GROUP BY  P.PI,   P.TransactionID,  P.TransactionID,D.SPlitIDofPI)X " & _
                "ON X.PI=DS.ParentID and DS.TRAN_ID=cstr(X.SPlitIDofPI)), tblPurInv INV,Property P where  P.PropertyID=INV.PropertyID AND  INV.My_ID=DS.ParentID "
                    
                    
                    adoConn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                            
'    'Insert code for type 24 PI Purchase Payment Refund
'             SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,TransactionID, 1 as TRAN_ID,TransactionID,P.clientID,PropertyID,P.PDate,SageAccountNumber,P.NominalCode," & _
'                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.clientID,P.PDate,SageAccountNumber,slnumber,P.NominalCode,Sum(D.Amount) AS Amt FROM " & _
'                        "tlbPaymentSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
'                        "(P.ClientID)='" & szSelectedClient & "' AND D.PayHeader=P.TransactionID AND D.ClientStatementID=" & szCurrentStatementID & " " & _
'                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,slnumber,SageAccountNumber,P.NominalCode)X )"
'
'                    adoconn.Execute "Insert into ReportClientStatementPurchases(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
'                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
'                            SQLforInsert
                            
                            
    
         adoConn.Execute "Update   tlbPaymentSplit PS,tlbPayment P,ReportClientStatementPurchases R  set R.PaymentDescription=PS.Description where   Ps.PayHeader=p.transactionID and  P.type=R.Type and P.transactionID=R.TransactionID "
         adoConn.Execute "Update  Supplier S, ReportClientStatementPurchases R,GlobalData G set R.VATAmount= Round((R.NetAmount * 20/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTax=true "
         adoConn.Execute "Update  Supplier S, ReportClientStatementPurchases R,GlobalData G set R.NetAmount= Round((R.NetAmount * 100/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTAx=true "
'        adoconn.Execute "update ReportClientStatementPurchases a,tlbPayment b,tlbTransactionTypes C  " & _
'        "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber,a.PaymentDescription=b.Details  where b.TransactionID=a.TransactionID and C.TYPE_ID=a.Type and a.Type=24 "
         
    'Exit Sub
    Dim rsStatementTemplate As New ADODB.Recordset
    Dim strReportName As String
    rsStatementTemplate.Open "Select CSTemplate from client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("CSTemplate").Value), "", rsStatementTemplate("CSTemplate").Value)
    End If
    If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
    Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
    End If
   rsStatementTemplate.Close
   
''''    Dim rsReportName As New ADODB.Recordset
'''''    rsReportName.Open "Select LesseeTemplate from Client where clientID='" & flxPayFees.TextMatrix(selRow, 4) & "'", adoConn, adOpenStatic, adLockReadOnly
'''''    If rsReportName.EOF Then
''''         'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'''''    Else
'''''        If IsNull(rsReportName!LesseeTemplate) Then
'''''            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'''''        Else
'''''            If rsReportName!LesseeTemplate = "" Then
'''''                Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplitNew.rpt")
'''''            Else
'''''                Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & rsReportName.Fields.Item("LesseeTemplate").Value)
'''''            End If
'''''
'''''        End If
'''''    End If
'''''    rsReportName.Close
''''    adoconn.Close
''''    Dim dblPercentage As Double
''''    Report.EnableParameterPrompting = False
''''    Report.DiscardSavedData
''''    Report.ParameterFields(1).AddCurrentValue CLng(StrDigitVal(szCurrentStatementID))
''''
''''
''''     Dim selRowTemp As Integer
''''     selRowTemp = selRow
''''     If flxPayFees.TextMatrix(selRow, 1) = "+" Or flxPayFees.TextMatrix(selRow, 1) = ">" Then
''''            'Report as header
''''            szSelectedClient = flxPayFees.TextMatrix(selRow, 4)
''''            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRow, 4) 'client ID
''''            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRow, 7)) 'statement date
''''            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
''''            Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
''''            Report.ParameterFields(6).AddCurrentValue "0" '0 Means header
''''            adoconn.Open getConnectionString
''''            Report.ParameterFields(7).AddCurrentValue findClientaddress(adoconn, flxPayFees.TextMatrix(selRow, 4))
''''            Report.ParameterFields(8).AddCurrentValue 0 '((flxPayFees.TextMatrix(selRow, 10))) 'amount paid to LL no use of this param
''''            Report.ParameterFields(9).AddCurrentValue -GetSupplierOSAmount
''''            adoconn.Close
''''      Else
''''            'Report as Split in LL
''''             szSelectedClient = flxPayFees.TextMatrix(selRow, 4)
''''            dblPercentage = Replace(flxPayFees.TextMatrix(selRow, 9), "%", "") 'Take Percenatge from Grid
''''            Do
''''                selRowTemp = selRowTemp - 1
''''            Loop Until (flxPayFees.TextMatrix(selRowTemp, 1) = "+" Or flxPayFees.TextMatrix(selRowTemp, 1) = ">")
''''            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRowTemp, 4)  'client ID
''''            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRowTemp, 7)) 'statement date
''''            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRowTemp, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRowTemp, 6))) 'Previuos statement date
''''            Report.ParameterFields(5).AddCurrentValue dblPercentage
''''            Report.ParameterFields(6).AddCurrentValue "1" '1 Means Split
''''            adoconn.Open getConnectionString
''''            Report.ParameterFields(7).AddCurrentValue findLandlordAddress(adoconn, flxPayFees.TextMatrix(selRow, 6)) 'take clientlnadlord ID from grid
''''            Report.ParameterFields(8).AddCurrentValue Val(flxPayFees.TextMatrix(selRow, 11)) 'amount paid to LL' I cannot replace poound sign . So I am not going to display pound sign in the grid
''''            Report.ParameterFields(9).AddCurrentValue -GetSupplierOSAmount
''''            adoconn.Close
''''      End If
''''
''''    Load frmReport
''''    frmReport.LoadReportViewer Report
    

    adoConn.Close
    'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementSplit.rpt")
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    Report.ParameterFields(1).AddCurrentValue CInt(Trim(Replace(szCurrentStatementID, "CS", "")))


     Dim selRowTemp As Integer
     selRowTemp = selRow
     If flxPayFees.TextMatrix(selRow, 1) = "+" Or flxPayFees.TextMatrix(selRow, 1) = ">" Then
            'Report as header
            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRow, 4) 'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRow, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
            Report.ParameterFields(6).AddCurrentValue "0" '0 Means header
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, flxPayFees.TextMatrix(selRow, 4))
            Report.ParameterFields(8).AddCurrentValue 0 '((flxPayFees.TextMatrix(selRow, 10))) 'amount paid to LL no use of this param
            adoConn.Close
      Else
            'Report as Split in LL
            dblPercentage = Replace(flxPayFees.TextMatrix(selRow, 9), "%", "") 'Take Percenatge from Grid
            Do
                selRowTemp = selRowTemp - 1
            Loop Until (flxPayFees.TextMatrix(selRowTemp, 1) = "+" Or flxPayFees.TextMatrix(selRowTemp, 1) = ">")
            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRowTemp, 4)  'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRowTemp, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRowTemp, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRowTemp, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue dblPercentage
            Report.ParameterFields(6).AddCurrentValue "1" '1 Means Split
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findLandlordAddress(adoConn, flxPayFees.TextMatrix(selRow, 6)) 'take clientlnadlord ID from grid
            Report.ParameterFields(8).AddCurrentValue Val(flxPayFees.TextMatrix(selRow, 11)) 'amount paid to LL' I cannor replace poound sign . SO I am not going to dipaly pound sign in the grid
            adoConn.Close
      End If

'    Load frmReport
'    frmReport.LoadReportViewer Report

       ' szSQL = "test.pdf"
        Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szOutputpdfto 'DB_PATH & "\AllStuff\Temp\" & szSQL
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.FormatType = crEFTPortableDocFormat
        Report.ExportOptions.PDFExportAllPages = True
        Report.Export False
        Set Report = Nothing
End Sub
Public Function StrDigitVal1(szStr As String) As Long
   Dim i As Long, j As Long, K As String
   Dim strText As String
   
   For i = 1 To Len(szStr)
      K = Mid$(szStr, i, 1)
      If IsNumeric(K) Or K = "." Then
         strText = strText & K
      End If
   Next
   StrDigitVal1 = CLng(strText)
End Function
Private Sub PrintCsConsolidated()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim CSID As String
   For rCount = 1 To flxPayFees.Rows - 1
        If flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
    szCurrentStatementID = flxPayFees.TextMatrix(selRow, 2)
    
    CSID = Replace(szCurrentStatementID, "CS", "")
    'run TestReportForRentSummary.rpt
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    'Dim adoconn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblCrReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    
    Dim DateTO As String
    If szCurrentStatementID = "" Then
        MsgBox "Please select a statement", vbInformation, "Warning"
        Exit Sub
    End If
    adoConn.Execute "Update DemandSplitRecords DS,DemandRecords D,Units U,Property P  set ReportNetAmount=0,ReportVATAmount=0,ReportReceivedAmount= 0,ReportDateFrom=Null," & _
                " ReportCreditAmount=0,reportOSAmount=0,ReportDateTO =null,ReportDemandTypeDesc= '' where D.DemandID=DS.DemandID  and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
                flxPayFees.TextMatrix(selRow, 4) & "'"
    
       
                
    szCurrentStatementID = Replace(szCurrentStatementID, "CS", "")
     Dim szPreviousStatementDate As Date
    Dim szStatementDate As Date
     rsRentSummaryStatement.Open "Select ListOfFundId,PreviousStatementDate,StatementDate  from RentSummaryStatement where statementID=" & _
                szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly ' group by D.DemandId", adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
            szListofFunds = rsRentSummaryStatement("ListOfFundId").Value
            szPreviousStatementDate = rsRentSummaryStatement("PreviousStatementDate").Value
             szStatementDate = rsRentSummaryStatement("StatementDate").Value
    End If
    'Exit Sub
    Dim strDueDate As String
   
    
    rsDemandSplit.Open "Select D.DemandId,D.TransactionType,sum(Amount) as NAmt,sum(VATAmount)as TVAT from  DemandSplitRecords DS,DemandRecords D,Units U,Property P where D.DemandID=DS.DemandID " & _
                " and D.TransactionType=1 AND U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID   AND P.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' group by  " & _
                " D.DemandId,D.TransactionType", adoConn, adOpenStatic, adLockReadOnly
'
    Dim rsDemandSplit1 As New ADODB.Recordset
    Dim dblCrSumAmt As Double
    Dim rsDemandSplitCredit As New ADODB.Recordset
'Update credit receipt amount
'    rsDemandSplitCredit.Open "Select D.DemandId,sum(RS.Amount) as NAmt from  tlbReceipt R,tlbReceiptSplit RS,DemandRecords D where  RS.ClientStatementID=" & CSID & " " & _
'                "AND R.DemandRef=D.DemandID AND rptHeader=R.TransactionID AND Type=2 AND R.OSAmount=0 AND R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & _
'                "' group by D.DemandId", adoConn, adOpenStatic, adLockReadOnly
'    While Not rsDemandSplitCredit.EOF
'                 dateFrom = ""
'                 DateTO = ""
'                 strDueDate = ""
'                 szTypeOfDemanddesc = ""
'                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & rsDemandSplitCredit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
'                 While Not rsDemandSplit1.EOF
'                        dateFrom = rsDemandSplit1("DateFrom").Value
'                        DateTO = rsDemandSplit1("DateTo").Value
'                        szTypeOfDemanddesc = szTypeOfDemanddesc + " " + rsDemandSplit1("Type").Value
'                        strDueDate = rsDemandSplit1("DueDate").Value
'                        rsDemandSplit1.MoveNext
'                Wend
'                 rsDemandSplit1.Close
'
'                adoConn.Execute "Update DemandRecords DS  set ReportNetAmount=" & -rsDemandSplitCredit("NAmt").Value & ",ReportVATAmount=0,ReportCreditAmount=" & -rsDemandSplitCredit("NAmt").Value & ",ReportReceivedAmount= 0,ReportDateFrom=#" & dateFrom & "#," & _
'                " reportOSAmount=0,ReportDateTO =#" & DateTO & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where TransactionType=2 AND DemandId=" & rsDemandSplitCredit("DEMANDID").Value & " "
'            rsDemandSplitCredit.MoveNext
'    Wend
'    rsDemandSplitCredit.Close
    Dim rsDemandSplitAmt As New ADODB.Recordset
    Dim iCount As Integer
    Dim dblDemandSplitamt As Double
     rsDemandSplitCredit.Open "Select D.DemandId,D.SplitID, Sum(RS.Amount) as NAmt,R.TransactionID from  tlbReceipt R,tlbReceiptSplit RS, DemandSplitRecords D where D.SplitID=Rs.SplitID AND RS.rptHeader=R.transactionID " & _
                " AND R.DemandRef=D.DemandID " & _
                "AND Type=2 AND R.OSAmount=0 AND R.ClientID='" & flxPayFees.TextMatrix(selRow, 4) & "' group by R.TransactionID,D.DemandId,D.SplitID", adoConn, adOpenStatic, adLockReadOnly
    While Not rsDemandSplitCredit.EOF

                dateFrom = ""
                 DateTO = ""
                 strDueDate = ""
'                 FromTran = 1461
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                 rsDemandSplitCredit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                        strDueDate = rsDemandSplit1("DueDate").Value
                 End If
                 rsDemandSplit1.Close
                 '*************************
                 rsDemandSplitAmt.Open "Select Count(amount) as CNT from DemandSplitRecords DS where  DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             iCount = rsDemandSplitAmt("CNT").Value
                 End If
                 rsDemandSplitAmt.Close
                 
                 
                 rsDemandSplitAmt.Open "Select amount from DemandSplitRecords DS where  DS.SPLITID=" & rsDemandSplitCredit("SplitID").Value & "  and DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             dblDemandSplitamt = rsDemandSplitAmt("amount").Value
                 End If
                 dblOSAmount = dblDemandSplitamt
                 rsDemandSplitAmt.Close
                  
                        
                If iCount = 1 Then
                    rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                    "RS.RptTransactionsIDSplit=D.TransactionID  AND D.Allocdate <=#" & Format(szStatementDate, "dd MMM yyyy") & "#  AND D.Allocdate > #" & Format(szPreviousStatementDate, "dd MMM yyyy") & "#  AND FromTran=" & _
                    rsDemandSplitCredit("TransactionID").Value & " and deleteflag=false  ", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsDemandSplit1.EOF Then
                           dblCrSumAmt = IIf(IsNull(rsDemandSplit1("Amt").Value), 0, rsDemandSplit1("Amt").Value)
                           dblOSAmount = dblOSAmount - dblCrSumAmt
                    End If
                    rsDemandSplit1.Close
                    If dblCrSumAmt > 0 Then
                        adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID  AND DemandId=" & _
                        rsDemandSplitCredit("DEMANDID").Value & " "
                     End If
                 Else
                        rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                        "RS.RptTransactionsIDSplit=D.TransactionID  AND FromTran=" & _
                        rsDemandSplitCredit("TransactionID").Value & " and SPlitIDofSi=" & rsDemandSplitCredit("SplitID").Value & " and deleteflag=false  group by D.SPlitIDofSi ", adoConn, adOpenStatic, adLockReadOnly
                        If Not rsDemandSplit1.EOF Then
                               dblCrSumAmt = rsDemandSplit1("Amt").Value
                               dblOSAmount = dblOSAmount - dblCrSumAmt
                        End If
                        rsDemandSplit1.Close
                        If dblCrSumAmt > 0 Then
                            adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID  AND DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & " "
                        End If
                 End If

                adoConn.Execute "Update DemandSplitRecords DS  set ReportNetAmountS=" & -dblDemandSplitamt & ",ReportVATAmountS=0,ReportCreditAmountS=" & _
                                -dblCrSumAmt & ",ReportReceivedAmountS= 0,ReportDateFromS=#" & dateFrom & "#," & _
                                " reportOSAmountS=" & -dblOSAmount & ",ReportDateTOS=#" & DateTO & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                                rsDemandSplitCredit("DEMANDID").Value & " AND DS.SplitID =" & rsDemandSplitCredit("SplitID").Value & ""
            rsDemandSplitCredit.MoveNext
            
            
    Wend
    rsDemandSplitCredit.Close
    While Not rsDemandSplit.EOF
                'Income Side
                 dblReceivedAmt = 0
                 dblCrReceivedAmt = 0
                 dblOSAmount = 0
                 If rsDemandSplit("DemandId").Value = 128 Then
                    Debug.Print rsDemandSplit("transactionType").Value
                 End If
                 If rsDemandSplit("DEMANDID").Value = 204 Then
                    Debug.Print ""
                 End If
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID in ( " & szListofFunds & ") " & _
                 "AND  R.ClientStatementID=" & CSID & " AND RL.Type in(3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " " & _
                 " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 dblReceivedAmt = 0 'AND R.FundID in ( " & szListofFunds & ")'AND RC.Type in(3,4,23)'sum(switch(RC.Type=3,R.Amount,RC.Type=4,R.Amount,RC.Type=23,-R.Amount))
                 If Not rsReceived.EOF Then
                        dblReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If

                 rsReceived.Close
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID in ( " & szListofFunds & ") " & _
                 "AND  R.ClientStatementID=" & CSID & " AND RL.Type in(2) and RC.DemandRef=D.DemandID " & _
                 "AND D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 'dblReceivedAmt = 0 'AND R.FundID in ( " & szListofFunds & ")'AND RC.Type in(3,4,23)'sum(switch(RC.Type=3,R.Amount,RC.Type=4,R.Amount,RC.Type=23,-R.Amount))
                 If Not rsReceived.EOF Then
                         dblCrReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                   'writing code for collectiong the osamount (osamount report field we dont have to filter by csID here
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID AND " & _
                 "RL.RDate <=#" & Format(rsRentSummaryStatement("StatementDate").Value, "dd MMM yyyy") & "# " & _
                 "AND RL.Type in(2,3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 
                 If Not rsReceived.EOF Then
                             dblOSAmount = rsDemandSplit("NAmt").Value + rsDemandSplit("TVAT").Value - IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 
                 
                 dateFrom = ""
                 DateTO = ""
                 szTypeOfDemanddesc = ""
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 While Not rsDemandSplit1.EOF
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = szTypeOfDemanddesc + " " + rsDemandSplit1("Type").Value
                        rsDemandSplit1.MoveNext
                 Wend
                 rsDemandSplit1.Close
                 If dateFrom = "" Then
                        MsgBox "Date from in the demand split is empty", vbInformation, "Warning"
                        Exit Sub
                 End If

                  If dblReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandRecords set ReportNetAmount= " & rsDemandSplit("NAmt").Value & ",ReportVATAmount= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmount =  " & dblOSAmount & ",ReportReceivedAmount= " & dblReceivedAmt & ",ReportDateFrom=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTO =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where DemandId=" & rsDemandSplit("DEMANDID").Value & ""
                        dblReceivedAmt = 0
                  End If
                  
                  If dblCrReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandRecords set ReportNetAmount= " & rsDemandSplit("NAmt").Value & ",ReportVATAmount= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmount =  " & dblOSAmount & ", ReportCreditAmount= " & dblCrReceivedAmt & ",ReportDateFrom=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTO =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where DemandId=" & rsDemandSplit("DEMANDID").Value & ""
                        dblCrReceivedAmt = 0
                  End If
                    
                    
            rsDemandSplit.MoveNext
    Wend
    rsDemandSplit.Close
    ''' Expense side ***********************************************************************
    Dim rstblPurInv As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim dblPayAmount As Double
    Dim dblinvNetamount As Double
    Dim dblinvVATamount As Double
    Dim szPDesc As String
    Dim szNominalCode As String
    rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where PI.MY_ID=P.PI AND CL_ID='" & flxPayFees.TextMatrix(selRow, 4) & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S where S.payHeader=P.TransactionID " & _
                        "AND S.ClientStatementID=" & CSID & " AND S.PayTransactionIDSplit=T.TransactionID AND S.FundID in ( " & szListofFunds & ")" & _
                        "AND T.FromTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.ToTran=" & _
                        rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
         End If
         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & dblinvNetamount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"
                
         rstblPurInv.MoveNext
    Wend
    rstblPurInv.Close
    rsRentSummaryStatement.Close
    adoConn.Close
  
    
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatement.rpt")
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
     Report.ParameterFields(1).AddCurrentValue CInt(Trim(Replace(szCurrentStatementID, "CS", "")))
    

     Dim selRowTemp As Integer
     selRowTemp = selRow
     If flxPayFees.TextMatrix(selRow, 1) = "+" Or flxPayFees.TextMatrix(selRow, 1) = ">" Then
             Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRow, 4) 'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRow, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
            Report.ParameterFields(6).AddCurrentValue "0" '0 is for detail record so print address for passing client ID in parameter 2
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, flxPayFees.TextMatrix(selRow, 4))
            adoConn.Close
      Else
            dblPercentage = Replace(flxPayFees.TextMatrix(selRow, 9), "%", "") 'Take Percenatge from Grid
            Do
                selRowTemp = selRowTemp - 1
            Loop Until (flxPayFees.TextMatrix(selRowTemp, 1) = "+" Or flxPayFees.TextMatrix(selRowTemp, 1) = ">")
            Report.ParameterFields(2).AddCurrentValue flxPayFees.TextMatrix(selRowTemp, 4)  'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(flxPayFees.TextMatrix(selRowTemp, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(flxPayFees.TextMatrix(selRowTemp, 6) = "", "01-01-1900", flxPayFees.TextMatrix(selRowTemp, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue dblPercentage
            Report.ParameterFields(6).AddCurrentValue "1" '1 is for detail record so print address for passing clientlnadlord ID
            
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findLandlordAddress(adoConn, flxPayFees.TextMatrix(selRow, 3)) 'take clientlnadlord ID from grid
            adoConn.Close
      End If
    Load frmReport
    frmReport.LoadReportViewer Report
End Sub
Private Function findClientaddress(adoConn As ADODB.Connection, ClientID As String) As String
    Dim rsClient As New ADODB.Recordset
    rsClient.Open "Select ClientAddressLine1,ClientAddressLine2,ClientAddressLine3,ClientAddressLine4,Client.ClientPostCode from Client where ClientID='" & ClientID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsClient.EOF Then
            findClientaddress = rsClient("ClientAddressLine1").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine2").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine3").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine4").Value + Chr(10) + Chr(13) + rsClient("ClientPostCode").Value + Chr(10) + Chr(13)
    End If
    rsClient.Close
End Function
Private Function findLandlordAddress(adoConn As ADODB.Connection, landLordID As String) As String
    Dim rsClient As New ADODB.Recordset
    rsClient.Open "Select LandlordAddressLine1,LandlordAddressLine2,LandlordAddressLine3,LandlordAddressLine4,LandlordPostCode from Landlord where LandlordID='" & landLordID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsClient.EOF Then
            findLandlordAddress = rsClient("LandlordAddressLine1").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine2").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine3").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine4").Value + Chr(10) + Chr(13) + rsClient("LandlordPostCode").Value + Chr(10) + Chr(13)
    End If
    rsClient.Close
End Function



Private Sub cmdPrintClientStatement_Click()
    Call PrintCSlineByLineNew
End Sub

'Private Sub cmdPrintThis_Click()
''    If chkConsolidatedCS.Value = 0 Then
''        Call PrintCsConsolidated
''    Else
'
'    Debug.Print time & "-01"
'
'        Call printCSlinebyLine
'        Debug.Print time & "-02"
''    End If
'End Sub

Private Sub cmdProduceClientSummaryStatement_Click()
'    Frame1(6).Visible = True
'    Frame1(6).Top = 135
'    Frame1(6).Left = 2025
'    bPreviewMode = False
'    Call LoadFlxClients
'    ConfigFlxProperties
'    ConfigFlxBankAccounts
''    ConfigflxFundiList
'    ConfigFlxInFunds
'    frmNewRentSummaryStatement.Top = Me.Top + 500
    frmRentPayableNew.bEditMode = False
'    frmNewRentSummaryStatement.Show
   ' frmRentPayableNew.Caption = "New Rent Summary Statement"
    'frmRentPayableNew.cmdFinalizeStatement.Visible = False
    frmRentPayableNew.cmdSave.Caption = "Produce Statement"
'    frmNewRentSummaryStatement.ZOrder 0
    LoadForm frmRentPayableNew
End Sub

Private Sub cmdRecharge_Click()
   Load frmRecharge
   frmRecharge.Show
End Sub

Private Sub cmdRechargeGenerate_Click()
   If MsgBox("Do you wish to generate rent payable?", vbQuestion + vbYesNo, "Generate Fees") = vbNo Then Exit Sub

   szaTrans(1) = 0
   GenerateRentPayable
   ShowMsgInTaskBar "Total " & szaTrans(1) & " rents have been generated to payable."
End Sub

Private Sub AgentVATCode(conADO As ADODB.Connection)
   Dim rstAgn As New ADODB.Recordset
   Dim szSQL As String
   
   szSQL = "SELECT VATReg " & _
           "FROM AGENT " & _
           "WHERE InactiveAgent = FALSE;"
   rstAgn.Open szSQL, conADO, adOpenStatic, adLockReadOnly
   
   If Not rstAgn.EOF Then
      If IsNull(rstAgn!VATReg) Or rstAgn!VATReg = "" Then
         szAgentVATCode = "T0"
      Else
         szAgentVATCode = "T1"
      End If
   Else
      szAgentVATCode = "T0"
   End If
End Sub

Private Sub GenerateRentPayable()
   Dim szSQL As String
   Dim conRent As New ADODB.Connection
   Dim rstRent As ADODB.Recordset, rstPI As ADODB.Recordset
   Dim szDemandCategoryID As String
   If szSupplierAccount = "" Then
        MsgBox "Please select the supplier account", vbInformation, "Warning"
        Exit Sub
   End If
'   MousePointer = vbHourglass

   conRent.Open getConnectionString
   Set rstRent = New ADODB.Recordset

'*****                           RENT PAYABLE: Received AMOUNT               *************************
   szSQL = "SELECT PA.PAYABLE_ID, " & _
               "PA.PAY_AnnualCharge , PA.PAY_START_DATE, " & _
               "PA.PAY_END_DATE , PA.PAY_FUND, PT.CategoryCode, " & _
               "PA.PAY_FREQUENCY , PA.PAY_NtDueDate, " & _
               "PT.PAYTYPE, PT.PAYNCAmt, PA.PAYABLE_TYPE, VC.VAT_CODE, VC.VAT_RATE " & _
           "FROM tlbPayable AS PA, ClientProAgr AS CPA, ClientGlobalData AS GD, " & _
               "tlbVatCode AS VC, PayableTypes AS PT " & _
           "WHERE  And " & _
               "DATEDIFF('d', DATEVALUE('" & Format(Date, "dd/mmm/yyyy") & "'), PA.PAY_NtDueDate) <= CPA.NOTICE_DAYS And " & _
               "PA.CPA_ID=CPA.CPA_ID And CPA.PropertyID=GD.PropertyID And " & _
               "GD.PropertyID='" & cboProperty.Column(0) & "' And PT.ID = PA.PAYABLE_TYPE;"
               
               'cboProperty.Column(0) will be array of properties as you need to generate PI for multiple properties

   rstRent.Open szSQL, conRent, adOpenStatic, adLockReadOnly

'   lLastID = CLng(GetLastID("tblPurInv", conRent)) + 1

   Set rstPI = New ADODB.Recordset
   szSQL = "SELECT * FROM tblPurInv;"
   rstPI.Open szSQL, conRent, adOpenDynamic, adLockPessimistic

   While Not rstRent.EOF
      rstPI.AddNew
      rstPI!My_ID = UniqueID()
      rstPI!TRAN_DATE = Format(Now, "DD MMMM YYYY")
      rstPI!CATEGORY_CODE = rstRent!CategoryCode
      rstPI!Nominal_code = rstRent!PayNCAmt
      rstPI!DEPT_ID = rstRent!PAY_FUND
      rstPI!description = rstRent!PayType
      rstPI!NET_AMOUNT = Format(rstRent!PAY_AnnualCharge, "0.00")
      'rstPI!TAX_CODE = rstRent!VAT_CODE 'pick up the Vat rate form the global data , if it isnt sett up then create a warning
      rstPI!vat = Format((rstRent!VAT_RATE * rstRent!PAY_AnnualCharge) / 100, "0.00")
      rstPI!UPDATE_SAGE = False
      rstPI!SUPP_AC = szSupplierAccount
      rstPI!TRANS = "Prop"
      rstPI!TRAN_TYPE = "PI"
      rstPI!UNIT_ID = cboProperty.Column(0)
      rstPI!TTP = CByte(TransactionTakePlace("TTP", "RENT PAYABLE", conRent))
'      lLastID = lLastID + 1
      rstPI!SlNumber = SlNumber("PI", "tblPurInv", conRent)
      rstPI!INV_NO = NextMFid_ADO(conRent)

      SetPayNtDueDt conRent, CInt(rstRent!PAYABLE_ID), CInt(rstRent!PAY_FREQUENCY), _
               rstRent!PAY_END_DATE, rstRent!PAY_NtDueDate, rstRent!PAYABLE_TYPE

      rstPI.Update
      szaTrans(1) = szaTrans(1) + 1
      rstRent.MoveNext
   Wend
   rstRent.Close

''***********                     RENT PAYABLE: RECEIVED AMOUNT               *************************
'      szSQL = "SELECT PB.PAYABLE_ID, RPT.UnitID, " & _
'                  "PB.PAY_START_DATE, PB.PAY_END_DATE, " & _
'                  "PB.PAY_FUND, PB.PAY_FREQUENCY, PT.CategoryCode, " & _
'                  "PB.PAY_NtDueDate, PT.PAYTYPE, PT.PAYNCAmt, " & _
'                  "PB.PAYABLE_TYPE, VC.VAT_CODE, VC.VAT_RATE, " & _
'                  "SUM(RPT.ReceiptAmount) AS TTR "
'      szSQL = szSQL + _
'               "FROM tlbPayable AS PB, PayableTypes AS PT, " & _
'                  "GlobalData AS GD, tlbVatCode AS VC, " & _
'                  "tlbReceipt AS RPT, Units as UT, " & _
'                  "DemandSplitRecords AS DS, ClientProAgr AS CPA "
'      szSQL = szSQL + _
'               "WHERE PB.PAY_Handling = 'Automatic' AND PB.PAYABLE_METHOD = 'RECEIVED' AND " & _
'                  "DATEDIFF('d',DATEVALUE('" & Format(Now, "DD MMMM YYYY") & "'), PB.PAY_NtDueDate) <= CPA.NOTICE_DAYS AND " & _
'                  "DATEDIFF('d', RPT.RDate, PB.PAY_NtDueDate) >= 0 AND " & _
'                  "PB.CPA_ID = CPA.CPA_ID AND " & _
'                  "CPA.PropertyID = UT.PropertyID AND " & _
'                  "GD.PropertyID = UT.PropertyID AND " & _
'                  "GD.VATRate = VC.VAT_ID AND " & _
'                  "UT.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
'                  "UT.UnitNumber = RPT.UnitID AND " & _
'                  "RPT.DemandRef = DS.DemandID AND " & _
'                  "DS.TypeOfDemand = PT.ID AND " & _
'                  "PT.CategoryCode = 1 "
'      szSQL = szSQL + _
'               "GROUP BY PB.PAYABLE_ID, RPT.UnitID, " & _
'                  "PB.PAY_START_DATE, PB.PAY_END_DATE, " & _
'                  "PB.PAY_FUND, PB.PAY_FREQUENCY, PT.CategoryCode, " & _
'                  "PB.PAY_NtDueDate, PT.PAYTYPE, PT.PAYNCAmt, " & _
'                  "PB.PAYABLE_TYPE, VC.VAT_CODE, VC.VAT_RATE;"
'Debug.Print szSQL
'   rstRent.Open szSQL, conRent, adOpenStatic, adLockReadOnly
'
'   While Not rstRent.EOF
'      rstPI.AddNew
'      rstPI!My_ID = UniqueID()
'      rstPI!TRAN_DATE = Format(Now, "DD MMMM YYYY")
'      rstPI!CATEGORY_CODE = rstRent!CategoryCode
'      rstPI!Nominal_code = rstRent!PayNCAmt
'      rstPI!DEPT_ID = rstRent!PAY_FUND
'      rstPI!description = rstRent!PayType & " - Unit ID: " & rstRent!unitid
'      rstPI!NET_AMOUNT = Format(rstRent!TTR, "0.00")
'      rstPI!TAX_CODE = rstRent!VAT_CODE
'      rstPI!vat = Format((rstRent!VAT_RATE * rstRent!TTR) / 100, "0.00")
'      rstPI!UPDATE_SAGE = False
'      rstPI!SUPP_AC = szSupplierAccount
'      rstPI!TRANS = "Prop"
'      rstPI!TRAN_TYPE = "PI"
'      rstPI!UNIT_ID = rstRent!unitid
'      rstPI!TTP = CByte(TransactionTakePlace("TTP", "RENT PAYABLE", conRent))
''      lLastID = lLastID + 1
''      rstPI!TRAN_ID = lLastID
'      rstPI!SlNumber = SlNumber("PI", "tblPurInv", conRent)
'
'      rstPI!INV_NO = NextMFid_ADO(conRent)
'
'      SetPayNtDueDt conRent, CInt(rstRent!PAYABLE_ID), CInt(rstRent!PAY_FREQUENCY), _
'            CDate(rstRent!PAY_END_DATE), CDate(rstRent!PAY_NtDueDate), CInt(rstRent!PAYABLE_TYPE)
'
'      rstPI.Update
'      szaTrans(1) = szaTrans(1) + 1
'      rstRent.MoveNext
'   Wend
'   rstRent.Close
'
''***********                     RENT PAYABLE: RECEIVABLE AMOUNT               *************************
'      szSQL = "SELECT PB.PAYABLE_ID, DR.UnitNumber, " & _
'                  "PB.PAY_START_DATE, PB.PAY_END_DATE, " & _
'                  "PB.PAY_FUND, PB.PAY_FREQUENCY, PT.CategoryCode, " & _
'                  "PB.PAY_NtDueDate, PT.PAYTYPE, PT.PAYNCAmt, " & _
'                  "PB.PAYABLE_TYPE, VC.VAT_CODE, VC.VAT_RATE, " & _
'                  "SUM(DS.TotalAmount) AS TTR "
'      szSQL = szSQL + _
'               "FROM tlbPayable AS PB, PayableTypes AS PT, " & _
'                  "GlobalData AS GD, tlbVatCode AS VC, " & _
'                  "DemandRecords AS DR, Units as UT, " & _
'                  "DemandSplitRecords AS DS, ClientProAgr AS CPA "
'      szSQL = szSQL + _
'               "WHERE PB.PAY_Handling = 'Automatic' AND PB.PAYABLE_METHOD = 'RECEIVABLE' AND " & _
'                  "DATEDIFF('d',DATEVALUE('" & Format(Now, "DD/MMMM/YYYY") & "'), PB.PAY_NtDueDate) <= CPA.NOTICE_DAYS AND " & _
'                  "DATEDIFF('d', DS.DueDate, PB.PAY_NtDueDate) >= 0 AND " & _
'                  "PB.CPA_ID = CPA.CPA_ID AND " & _
'                  "CPA.PropertyID = UT.PropertyID AND " & _
'                  "GD.PropertyID = UT.PropertyID AND " & _
'                  "GD.VATRate = VC.VAT_ID AND " & _
'                  "UT.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
'                  "UT.UnitNumber = DR.UnitNumber AND " & _
'                  "DR.DemandID = DS.DemandID AND " & _
'                  "DS.TypeOfDemand = PT.ID AND " & _
'                  "PT.CategoryCode = 1 "
'      szSQL = szSQL + _
'               "GROUP BY PB.PAYABLE_ID, DR.UnitNumber, " & _
'                  "PB.PAY_START_DATE, PB.PAY_END_DATE, " & _
'                  "PB.PAY_FUND, PB.PAY_FREQUENCY, PT.CategoryCode, " & _
'                  "PB.PAY_NtDueDate, PT.PAYTYPE, PT.PAYNCAmt, " & _
'                  "PB.PAYABLE_TYPE, VC.VAT_CODE, VC.VAT_RATE"
''Debug.Print szSQL
'   rstRent.Open szSQL, conRent, adOpenStatic, adLockReadOnly
'
'   While Not rstRent.EOF
'      rstPI.AddNew
'      rstPI!My_ID = UniqueID()
'      rstPI!TRAN_DATE = Format(Now, "DD MMMM YYYY")
'      rstPI!CATEGORY_CODE = rstRent!CategoryCode
'      rstPI!Nominal_code = rstRent!PayNCAmt
'      rstPI!DEPT_ID = rstRent!PAY_FUND
'      rstPI!description = rstRent!PayType & " - Unit ID: " & rstRent!unitid
'      rstPI!NET_AMOUNT = Format(rstRent!TTR, "0.00")
'      rstPI!TAX_CODE = rstRent!VAT_CODE
'      rstPI!vat = Format((rstRent!VAT_RATE * rstRent!TTR) / 100, "0.00")
'      rstPI!UPDATE_SAGE = False
'      rstPI!SUPP_AC = szSupplierAccount
'      rstPI!TRANS = "Prop"
'      rstPI!TRAN_TYPE = "PI"
'      rstPI!UNIT_ID = rstRent!unitid
'      rstPI!TTP = CByte(TransactionTakePlace("TTP", "RENT PAYABLE", conRent))
''      lLastID = lLastID + 1
'      rstPI!SlNumber = SlNumber("PI", "tblPurInv", conRent)
'      rstPI!INV_NO = NextMFid_ADO(conRent)
'
'      SetPayNtDueDt conRent, CInt(rstRent!PAYABLE_ID), CInt(rstRent!PAY_FREQUENCY), _
'         rstRent!PAY_END_DATE, rstRent!PAY_NtDueDate, rstRent!PAYABLE_TYPE
'
'      rstPI.Update
'      szaTrans(1) = szaTrans(1) + 1
'      rstRent.MoveNext
'   Wend
'   rstRent.Close

'   SetLastIDado "tblPurInv", lLastID, conRent

   rstPI.Close
   Set rstPI = Nothing
   Set rstRent = Nothing

   conRent.Close
   Set conRent = Nothing

   MousePointer = vbDefault
End Sub

Private Sub SetAgrNtDueDt(conMng As ADODB.Connection, szAgreementID As String, iFreq As Integer, dtEndDate As Date, dtNtDueDate As Date, iCTId As Integer)
   Dim szSQL As String
   Dim rstMng As ADODB.Recordset

   Set rstMng = New ADODB.Recordset
   szSQL = "SELECT NtDueDate " & _
           "FROM tlbAgreement " & _
           "WHERE AGREEMENT_ID =" & szAgreementID & ";"
   rstMng.Open szSQL, conMng, adOpenDynamic, adLockPessimistic

   rstMng!NtDueDate = FindNextDueDate(dtEndDate, dtNtDueDate, iFreq, conMng, iCTId, "ChargeTypes")

   rstMng.Update

   rstMng.Close
   Set rstMng = Nothing
End Sub

Private Sub SetPayNtDueDt(conMng As ADODB.Connection, iPayableID As Integer, iFreq As Integer, dtEndDate As Date, dtNtDueDate As Date, iPTId As Integer)
   Dim szSQL As String
   Dim rstPay As ADODB.Recordset

   Set rstPay = New ADODB.Recordset
   szSQL = "SELECT PAY_NtDueDate " & _
           "FROM tlbPayable " & _
           "WHERE PAYABLE_ID =" & iPayableID & ";"
   rstPay.Open szSQL, conMng, adOpenDynamic, adLockPessimistic

   rstPay!PAY_NtDueDate = FindNextDueDate(dtEndDate, dtNtDueDate, iFreq, conMng, iPTId, "PayableTypes")

   rstPay.Update

   rstPay.Close
   Set rstPay = Nothing
End Sub

Private Function FindNextDueDate(dtEndDate As Date, dtNtDueDate As Date, iFreq As Integer, ByVal adoConn As ADODB.Connection, iCTId As Integer, szChrPayTypes As String) As Date
   GetClientGlobalData txtClientID, adoConn, iCTId, szChrPayTypes

   Select Case iFreq
      Case 1:                               'Weekly in advance
         FindNextDueDate = dtNtDueDate
      Case 2:                               'Weekly in arrears
         FindNextDueDate = DateAdd("d", 7, dtNtDueDate)
      Case 3:                               'Fortnightly in advance
         FindNextDueDate = dtNtDueDate
      Case 4:                               'Fortnightly in arrears
         FindNextDueDate = DateAdd("d", 14, dtNtDueDate)
      Case 5:                               'Monthly in advance
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Monthly)
      Case 6:                               'Monthly in arrears
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Monthly)
      Case 7:                               'Quarterly in advance
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Quarterly)
      Case 8:                               'Quarterly in arrears
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Quarterly)
      Case 9:                               'Half yearly in advance
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Half_Yearly)
      Case 10:                               'Half yearly in arrears
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Half_Yearly)
      Case 11:                              'yearly in advance
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Yearly)
      Case 12:                              'yearly in arrears
         FindNextDueDate = ClNextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Yearly)
   End Select

   If DateDiff("d", dtEndDate, FindNextDueDate) > 0 Then
      FindNextDueDate = dtEndDate
   End If
End Function

Private Function DemandCategoryID(conADO As ADODB.Connection, szFeesOption As String) As String
   Dim szSQL As String
   Dim rstMng As New ADODB.Recordset
   
   If szFeesOption = "ALL" Then
      DemandCategoryID = ""
   Else
      szSQL = "SELECT CODE " & _
              "FROM SECONDARYCODE " & _
              "WHERE VALUE = '" & szFeesOption & "' AND " & _
                  "PRIMARYCODE = 'DCTG';"
      rstMng.Open szSQL, conADO, adOpenStatic, adLockReadOnly

      If Not rstMng.EOF Then
         DemandCategoryID = rstMng!Code
      Else
         DemandCategoryID = ""
      End If

      rstMng.Close
      Set rstMng = Nothing
   End If
End Function

Private Sub cmdSavePI_Click()
      If txtStatementDate1.text = "" Then
            MsgBox "Please enter statement date", vbInformation, "Warning"
            Exit Sub
      End If
      Dim lSlNumber As Long
      Dim adoConn As New ADODB.Connection
      Dim adoPIHeader As New ADODB.Recordset
      Dim szSQL As String
      adoConn.Open getConnectionString
      szSQL = "SELECT * FROM tblPurInv"
      lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
      With adoPIHeader
                .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
                .AddNew
                .Fields.Item("SlNumber").Value = lSlNumber
                .Fields.Item("SUPP_AC").Value = szSelectedClient
                .Fields.Item("TRAN_DATE").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
                .Fields.Item("TransactionType").Value = 6
                .Fields.Item("INV_NO").Value = szSelectedStatement 'txtInv(0).text
                .Fields.Item("TOTAL_AMOUNT").Value = CCur(txtRentPayable.text)
                .Fields.Item("TTP").Value = "PURCHASE INVOICE" 'CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoconn))
                .Fields.Item("History").Value = False
                .Fields.Item("TrfPayment").Value = False
                .Fields.Item("PropertyID").Value = ""
                .Fields.Item("CL_ID").Value = szSelectedClient
                .Fields.Item("NLPost").Value = False
                .Fields.Item("DueDate").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
                .Fields.Item("PostingDate").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
                .Update
      End With
      adoConn.Close
      Set adoConn = Nothing
End Sub

Private Sub cmdReverseHistory_Click()
    Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   For rCount = 1 To flxPayFeesHistory.Rows - 1
        If flxPayFeesHistory.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   adoConn.Execute "Update rentSummaryStatement set PostToHistory=false where statementID =" & szCurrentStatementHistoryID & ""
   Call loadflxPayFees("")
   Call loadflxPayFeesHistory
   adoConn.Close
   Set adoConn = Nothing
   MsgBox "Rent Summary Statement has been reversed from history", vbInformation, "Reversed"
End Sub

Private Sub cmdReverseRentPayable_Click()
    Dim adoConn As New ADODB.Connection
    Dim rsFinalized As New ADODB.Recordset
    Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    'Dim adoConn As New ADODB.Connection
    'Dim rsFinalized As New ADODB.Recordset
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = "+" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
       Exit Sub
    End If
    
    adoConn.Open getConnectionString
    rsFinalized.Open "Select * from RentSummaryStatement where StatementID=" & Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 2), "CS", "") & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsFinalized.EOF Then
        If rsFinalized("isfinalized").Value = "1" Then
                rsFinalized.Close
                adoConn.Close
                MsgBox "This statement has already been finalised.", vbInformation + vbOKOnly, "Statement Selection"

                Exit Sub
        End If
    End If
    If MsgBox("Do you want to clear Rent Payable?.", vbYesNo, "clear Rent Payable") = vbYes Then
        adoConn.Execute "Update RentSummaryStatement set PINumber=null where StatementID=" & Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 2), "CS", "") & ""
        adoConn.Execute "Delete From RentSummaryStatementdetails  where StatementID=" & Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 2), "CS", "") & ""
         Call loadflxPayFees("")
    End If
    
    adoConn.Close
End Sub

Private Sub Command1_Click()
        'This is recalculate rent summary sub procedure where we are clearing all the flags when we press recalculate button
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        adoConn.Execute "Delete from RentSummaryStatement"
        adoConn.Execute "Delete from RentSummaryStatementPreview"
        adoConn.Execute "Update tlbBankPayment Set RentSumStatement=''"
        adoConn.Execute "Update tlbPayment Set RentSumStatement=''"
        adoConn.Execute "Update tlbReceipt Set RentSumStatement=''"
        Call loadflxPayFees("")


    adoConn.Close
    Set adoConn = Nothing
    MsgBox "Flag has been cleared"
End Sub

Private Sub Command2_Click()
    Frame2.Visible = True
    Frame2.Left = 315
    Frame2.Top = 4320
    txtRetensionAmount1.Visible = True
    txtRetensionAmount1.SelStart = 0
    txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
    flxRetensionDetails.Enabled = False
    FocusControl txtRetensionAmount1
End Sub

Private Sub Command3_Click()
        Frame2.Visible = True
        flxRetensionDetails.Enabled = True
        txtRetensionAmount1.Visible = False
        FocusControl flxRetensionDetails
        Dim iRow As Integer
        For iRow = 1 To flxRetensionDetails.Rows - 1
            If flxRetensionDetails.TextMatrix(iRow, 2) <> "" Then
                    flxRetensionDetails.TextMatrix(iRow, 0) = "-"
            End If
        Next
End Sub

Private Sub Command4_Click()
    Frame2.Visible = False
End Sub

Private Sub Command5_Click()
    Frame4.Visible = False
End Sub

Private Sub CommandButton1_Click()
   
End Sub

Private Sub flxBankAccounts_Click()
    Dim iRow As Integer
    If flxBankAccounts.TextMatrix(flxBankAccounts.row, 1) = "" Then Exit Sub
    SelectOnly1RowFlxGrid flxBankAccounts, flxBankAccounts.row, 0
    If flxBankAccounts.TextMatrix(flxBankAccounts.row, 0) = "X" Then
            szSelectedBankAccount = flxBankAccounts.TextMatrix(flxBankAccounts.row, 2)
            addPropertiesTowizard szSelectedClient
'            Call LoadflxInFunds
    End If
    hasSelBankAccounts = False
    For iRow = 1 To flxBankAccounts.Rows - 1
            If flxBankAccounts.TextMatrix(iRow, 0) = "X" Then
                hasSelBankAccounts = True
                Exit For
            End If
    Next
    If hasSelBankAccounts = False Then
        Call ConfigFlxProperties
    End If
End Sub

Private Sub flxClientList_Click()
'   Dim sSQLQuery_ As String, sFilter As String
'
'   txtClientID.text = flxClientList.TextMatrix(flxClientList.row, 1)
'
'   MousePointer = vbHourglass
'
'   adoMain.ConnectionString = getConnectionString
'   sSQLQuery_ = "SELECT * " & _
'                "FROM CLIENT " & _
'                "WHERE CLIENT.ClientID = '" & flxClientList.TextMatrix(flxClientList.row, 1) & "';"
'
'   adoMain.RecordSource = sSQLQuery_
'   adoMain.CommandType = adCmdText
'   adoMain.Refresh
'
'   If Not Fill_Form(Me, adoMain) Then
'      ShowMsgInTaskBar "Error in Database.", , "N"
'   Else
'      LoadProperty
'   End If
'
'   MousePointer = vbDefault
    If flxClientList.TextMatrix(flxClientList.row, 1) = "ALL" Then
         Call loadflxPayFees("")
    End If
    If flxClientList.TextMatrix(flxClientList.row, 1) <> "" And flxClientList.TextMatrix(flxClientList.row, 1) <> "ALL" Then
        Call loadflxPayFees(" Where U.ClientIDLandlordID ='" & flxClientList.TextMatrix(flxClientList.row, 1) & "'")
    End If
    txtClientID.text = flxClientList.TextMatrix(flxClientList.row, 1)
   picClientList.Visible = False
End Sub

Private Sub flxClients_Click()
     SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
     szSelectedClient = flxClients.TextMatrix(flxClients.row, 1)
     'Auto select Properties bases on whatever you select at client
     Dim iRow As Integer
     For iRow = 1 To flxClients.Rows - 1
                If flxClients.TextMatrix(iRow, 0) = "X" Then
                    szSelectedClient = flxClients.TextMatrix(iRow, 1)
'                    addPropertiesTowizard flxClients.TextMatrix(iRow, 1)
                    Call LoadFlxBankAccounts(szSelectedClient)
'                    Call LoadflxPayableTypes 'szSelectedClient is a glbal variable and this function is loadiing the values according to client
'                    Call LoadflxInFunds
'                    Call LoadFlxBankAccounts(szSelectedClient)
                    Call LoadLaststatementdate
                    Exit For
                Else
                    Call ConfigFlxBankAccounts
                    Call ConfigFlxProperties
                    Call ConfigFlxInFunds
                    txtLastStatementDate1.text = ""
                End If
      Next
    
'     chkAllProperties.Value = 1
'     chkInFunds.Value = 1
End Sub
Private Sub addPropertiesTowizard(strClientID As String)
        Call ConfigFlxProperties
        Dim iRow As Integer
        Dim szSQL As String
        Dim adoConn As New ADODB.Connection
        Dim adoRST As New ADODB.Recordset
        adoConn.Open getConnectionString
        szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
               "FROM  PROPERTY where clientID = '" & strClientID & "'" & _
               "ORDER BY ClientID,PROPERTYID;"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        iRow = 1

        While Not adoRST.EOF 'While Not adoRst.EOF
'            flxProperties.AddItem ""
            flxProperties.TextMatrix(iRow, 0) = ""
            flxProperties.TextMatrix(iRow, 1) = adoRST("PROPERTYID").Value
            flxProperties.TextMatrix(iRow, 2) = adoRST("PROPERTYNAME").Value
            flxProperties.TextMatrix(iRow, 3) = adoRST("ClientID").Value
           ' flxProperties.RowHeight(iRow) = 280
            If iRow > 1 Then
                flxProperties.AddItem ""
            End If
            iRow = iRow + 1
    
            adoRST.MoveNext
        'End If
        Wend
       adoConn.Close
       Set adoConn = Nothing
End Sub

Private Sub flxFundList_Click()
    Frame5.Visible = False
    txtFundForPI.Tag = flxFundList.TextMatrix(flxFundList.row, 1)
    txtFundForPI.text = flxFundList.TextMatrix(flxFundList.row, 2)
    FocusControl cmdCreatePI
End Sub

Private Sub flxFundList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxFundList_Click
    End If
End Sub

Private Sub flxInFunds_Click()
    SelectFlxGridRow 0, flxInFunds, flxInFunds.row
    
    '  SelectOnly1RowFlxGrid flxInFunds, flxInFunds.row, 0
      szSelectedFund = flxInFunds.TextMatrix(flxInFunds.row, 1)
End Sub

Private Sub flxPayableTypes_Click()
     'SelectOnly1RowFlxGrid flxPayableTypes, flxPayableTypes.row, 0
      SelectFlxGridRow 0, flxPayableTypes, flxPayableTypes.row
      'SelectOnly1RowFlxGrid flxPayableTypes, flxPayableTypes.row, 0
      'szSelectedPayableTypeID = flxPayableTypes.TextMatrix(flxPayableTypes.row, 1)
End Sub

Private Sub flxPayFees_Click()
        Dim szSlNo As String
        Dim iIncDec As Integer
        If flxPayFees.TextMatrix(flxPayFees.row, 1) = "" Then Exit Sub
'        If flxPayFees.col = 0 Then
'            iIncDec = iIncDec + SelectFlxGridRow(0, flxPayFees, flxPayFees.row) 'Returns 1 or -1 depends on selection
'        End If
            SelectOnly1RowFlxGrid flxPayFees, flxPayFees.row, 0
'
'        'SelectOnly1RowFlxGrid flxPayFees, flxPayFees.row, 0
        If flxPayFees.TextMatrix(flxPayFees.row, 0) = "X" Then
            If flxPayFees.TextMatrix(flxPayFees.row, 1) = "+" Then
                szCurrentStatementID = Replace(flxPayFees.TextMatrix(flxPayFees.row, 2), "CS", "")
               chkExcludeSupOS.Value = flxPayFees.TextMatrix(flxPayFees.row, 31)
               chkShowDue.Value = flxPayFees.TextMatrix(flxPayFees.row, 32)
               ' szAvailableFund1 = flxPayFees.TextMatrix(flxPayFees.row, 8)
                Call LoadRentSummaryDetails
            End If
        End If
        
        Call SquezeExpand
End Sub
Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
   Dim iRow       As Integer
   Dim iCol       As Integer
   Dim iColPaint  As Integer

   iColPaint = IIf(iColID = 0, 1, 0)
   
   For iRow = 1 To conFlxGrid.Rows - 1
      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then Exit Sub
         conFlxGrid.TextMatrix(iRow, iColID) = ""
         conFlxGrid.row = iRow
         For iCol = iColPaint To conFlxGrid.Cols - 1
            conFlxGrid.col = iCol
            conFlxGrid.CellBackColor = vbWhite
         Next iCol
      End If
   Next iRow

   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
   conFlxGrid.row = iNewRow

   For iCol = iColPaint To conFlxGrid.Cols - 1
      conFlxGrid.col = iCol
      conFlxGrid.CellBackColor = RGB(174, 179, 233)
   Next iCol
End Sub
Private Sub SquezeExpandHistory()
   Dim i As Integer, iCurRowHeight As Integer

   iCurRowHeight = 280
   

   If flxPayFeesHistory.col = 1 And flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = "+" Then          'Expanding the grid
      flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = ">"
      iCurRowHeight = flxPayFeesHistory.RowHeight(flxPayFeesHistory.row)
      i = 1

      While flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row + i, 1) = "-"
         flxPayFeesHistory.RowHeight(flxPayFeesHistory.row + i) = iCurRowHeight
         i = i + 1
         If (flxPayFeesHistory.row + i) = flxPayFeesHistory.Rows Then Exit Sub
      Wend
      Exit Sub
   End If

   If flxPayFeesHistory.col = 1 And flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = ">" Then          'Squeezing the grid
      flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = "+"
      i = 1
      While flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row + i, 1) = "-"
         flxPayFeesHistory.RowHeight(flxPayFeesHistory.row + i) = 0
         i = i + 1
         If (flxPayFeesHistory.row + i) = flxPayFeesHistory.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   'HighLightRowFlxGridA flxPayFeesHistory, flxPayFeesHistory.row
End Sub
Private Sub SquezeExpand()
       Dim i As Integer, iCurRowHeight As Integer

  iCurRowHeight = 280
   

   If flxPayFees.col = 1 And flxPayFees.TextMatrix(flxPayFees.row, 1) = "+" Then          'Expanding the grid
      flxPayFees.TextMatrix(flxPayFees.row, 1) = ">"
      iCurRowHeight = flxPayFees.RowHeight(flxPayFees.row)
      i = 1

      While flxPayFees.TextMatrix(flxPayFees.row + i, 1) = "-"
         flxPayFees.RowHeight(flxPayFees.row + i) = iCurRowHeight
         i = i + 1
         If (flxPayFees.row + i) = flxPayFees.Rows Then Exit Sub
      Wend
      Exit Sub
   End If

   If flxPayFees.col = 1 And flxPayFees.TextMatrix(flxPayFees.row, 1) = ">" Then          'Squeezing the grid
      flxPayFees.TextMatrix(flxPayFees.row, 1) = "+"
      i = 1
      While flxPayFees.TextMatrix(flxPayFees.row + i, 1) = "-"
         flxPayFees.RowHeight(flxPayFees.row + i) = 0
         i = i + 1
         If (flxPayFees.row + i) = flxPayFees.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   'HighLightRowFlxGridA flxPayFees, flxPayFees.row
End Sub

Private Sub flxPayFees_RowColChange()
'        Dim szSlNo As String
'        Dim iIncDec As Integer
'        If flxPayFees.TextMatrix(flxPayFees.row, 1) = "" Then Exit Sub
'        iIncDec = iIncDec + SelectFlxGridRow(0, flxPayFees, flxPayFees.row) 'Returns 1 or -1 depends on selection
'
'        'SelectOnly1RowFlxGrid flxPayFees, flxPayFees.row, 0
'        If flxPayFees.TextMatrix(flxPayFees.row, 0) = "X" Then
'            szCurrentStatementID = Replace(flxPayFees.TextMatrix(flxPayFees.row, 1), "CS", "")
'            szAvailableFund1 = flxPayFees.TextMatrix(flxPayFees.row, 15)
'        End If
End Sub
Public Function SelectFlxGridRow(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      For iRow = conFlxGrid.Cols - 1 To 1
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iRow
      SelectFlxGridRow = -1
   Else
        'Here I have Implemented if no value in the grid row then do not select anol 2020-11-04
      If conFlxGrid.TextMatrix(iSelRow, iColID + 1) <> "" Then
            conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
            conFlxGrid.row = iSelRow
            For iRow = conFlxGrid.Cols - 1 To 1
               conFlxGrid.col = iRow
               conFlxGrid.CellBackColor = RGB(174, 179, 233)
            Next iRow
            SelectFlxGridRow = 1
      Else
            SelectFlxGridRow = -1
      End If
   End If
End Function
'THIS FUNCTIOn IS REMMED BEACAUSE i HAVE UPDATEd THIS FUNCTION
'Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
'   Dim iRow       As Integer
'   Dim iCol       As Integer
'   Dim iColPaint  As Integer
'
'   iColPaint = IIf(iColID = 0, 1, 0)
'
'   For iRow = conFlxGrid.Rows - 1 To 1 Step -1
'      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
'         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
'                conFlxGrid.TextMatrix(iRow, iColID) = ""
'                conFlxGrid.TextMatrix(iRow, iColID) = ""
'                conFlxGrid.row = iRow
'                For iCol = iColPaint To conFlxGrid.Cols - 1
'                   conFlxGrid.col = iCol
'                   conFlxGrid.CellBackColor = vbWhite
'                Next iCol
'                Exit Sub
'         End If
''         conFlxGrid.TextMatrix(iRow, iColID) = ""
''         conFlxGrid.row = iRow
''         For iCol = iColPaint To conFlxGrid.Cols - 1
''            conFlxGrid.col = iCol
''            conFlxGrid.CellBackColor = vbWhite
''         Next iCol
'      End If
'   Next iRow
'
'   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
'   conFlxGrid.row = iNewRow
'
'   For iCol = conFlxGrid.Cols - 1 To iColPaint Step -1
'      conFlxGrid.col = iCol
'      conFlxGrid.CellBackColor = RGB(174, 179, 233)
'   Next iCol
'End Sub

Private Sub flxPayFeesHistory_Click()
        Dim szSlNo As String
        Dim iIncDec As Integer
        If flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = "" Then Exit Sub
        If flxPayFeesHistory.col = 0 Then
            iIncDec = iIncDec + SelectFlxGridRow(0, flxPayFeesHistory, flxPayFeesHistory.row) 'Returns 1 or -1 depends on selection
        End If
        

        'SelectOnly1RowFlxGrid flxPayFeesHistory, flxPayFeesHistory.row, 0
        If flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 0) = "X" Then
            If flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 1) = "+" Then
                szCurrentStatementHistoryID = Replace(flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 2), "CS", "")
'                szAvailableFund1 = flxPayFeesHistory.TextMatrix(flxPayFeesHistory.row, 16)
'                Call LoadRentSummaryDetails
            End If
        End If
        
        Call SquezeExpandHistory
End Sub

Private Sub flxProperties_Click()
'    SelectFlxGridRow 0, flxProperties, flxProperties.row
    Dim iIncDec As Integer
    Dim iRow As Integer
    iIncDec = iIncDec + SelectFlxGridRow(0, flxProperties, flxProperties.row) 'Returns 1 or -1 depends on selection
'    Call LoadflxInFunds
    hasSelProperty = False
    For iRow = 1 To flxProperties.Rows - 1
            If flxProperties.TextMatrix(iRow, 0) = "X" Then
                hasSelProperty = True
                Exit For
            End If
    Next
    
    If hasSelProperty Then
          Call LoadflxInFunds
          'Call LoadFlxBankAccounts(szSelectedClient)
    Else
          Call ConfigFlxInFunds
    End If
End Sub

Private Sub flxRetensionDetails_Click()
    If flxRetensionDetails.TextMatrix(flxRetensionDetails.row, 0) = "-" Then
        flxRetensionDetails.RemoveItem flxRetensionDetails.row
        Call MakeSummaryRetention
    End If
End Sub

Private Sub Form_Load()

    Me.Height = 12015
    Me.Width = 22335
    Me.BackColor = MODULEBACKCOLOR
    tabFees.BackColor = MODULEBACKCOLOR
    Frame4.BackColor = MODULEBACKCOLOR
    Frame1(6).BackColor = MODULEBACKCOLOR
    Frame1(12).BackColor = MODULEBACKCOLOR
    Frame1(8).BackColor = MODULEBACKCOLOR
    Frame1(9).BackColor = MODULEBACKCOLOR
    Frame1(10).BackColor = MODULEBACKCOLOR
    Frame1(11).BackColor = MODULEBACKCOLOR
    chkAllProperties.BackColor = MODULEBACKCOLOR
    Frame1(13).BackColor = MODULEBACKCOLOR
    Frame1(7).BackColor = MODULEBACKCOLOR
    Frame1(14).BackColor = MODULEBACKCOLOR
    Label8.BackColor = MODULEBACKCOLOR
    Label9.BackColor = MODULEBACKCOLOR
    Label10.BackColor = MODULEBACKCOLOR
    Label11.BackColor = MODULEBACKCOLOR
    Label4.BackColor = MODULEBACKCOLOR
    Label6.BackColor = MODULEBACKCOLOR
    Label12.BackColor = MODULEBACKCOLOR
    Label13.BackColor = MODULEBACKCOLOR
    chkInFunds.BackColor = MODULEBACKCOLOR
    tabFees.Tab = 0
    Call ConfigflxRetensionDetails
    txtStatementDate1.text = Format(Now, "dd/mm/yyyy")

    Call LoadFlxClients
    Call LoadflxInFunds
    Call loadflxPayFees("")
    chkExcludeSupOS.Value = 1
    chkShowDue.Value = 1
    'Call loadflxPayFeesHistory
    'Call UpdateRentPayableOnCSDetails
    Call WheelHook(Me.hWnd)
End Sub
Private Sub UpdateRentPayableOnCSDetails()
'RentPayableOnCSDetails was not wrinting . beacuse we thought we dont need it.
    'now wer are populating those records
    Dim szSQL As String
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim rsRentSummaryStatementMAXID As New ADODB.Recordset
    Dim rsRentSummaryStatementDetailsNew As New ADODB.Recordset
    Dim onePI As String
    Dim adoConn As New ADODB.Connection
    Dim rsPIAdding As New ADODB.Recordset
    Dim rsPI As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim PIs
    Dim maxID As Long
    Dim splitID As Integer
    Dim iCount As Integer
    szSQL = "Select * from RentSummaryStatement  order by statementID desc"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    While Not rsRentSummaryStatement.EOF
        If IsNull(rsRentSummaryStatement!PINumber) Then
            rsRentSummaryStatement.MoveNext
         Else
            If Len(rsRentSummaryStatement!PINumber) > 0 Then
                If rsRentSummaryStatement!statementID = 90 Then
                    Debug.Print ""
                End If
                'Exit Sub
                splitID = 1
                Sleep (100)
                If rsRentSummaryStatement!statementID = 52 Then
                       Debug.Print ""
                End If
                If InStr(1, LTrim(rsRentSummaryStatement!PINumber), " ") > 0 Then
                        PIs = Split(LTrim(rsRentSummaryStatement!PINumber), " ")
                        rsRentSummaryStatementDetailsNew.Open "Select * from RentSummaryStatementDetails where statementID=" & _
                                rsRentSummaryStatement!statementID & "", adoConn, adOpenStatic, adLockReadOnly
                                
                                rsRentSummaryStatementMAXID.Open "Select (Max(ID)) AS DID from RentSummaryStatementDetails", adoConn, adOpenStatic, adLockReadOnly
                                If Not rsRentSummaryStatementMAXID.EOF Then
                                    maxID = IIf(IsNull(rsRentSummaryStatementMAXID("DID").Value), 0, rsRentSummaryStatementMAXID("DID").Value) + 1
                                End If
                                rsRentSummaryStatementMAXID.Close
                                
                        If rsRentSummaryStatementDetailsNew.EOF Then
                                  rsPIAdding.Open "Select * from RentSummaryStatementDetails  where statementID=" & _
                                                            rsRentSummaryStatement!statementID & "", adoConn, adOpenKeyset, adLockOptimistic
                                                If rsPIAdding.EOF Then
                                                    For iCount = 0 To UBound(PIs)
                                                        onePI = PIs(iCount)
                                                        'only add into the table  rsRentSummaryStatementDetails when statement ID is not there
                                                            If onePI <> "" Then
                                                                    rsPIAdding.AddNew
                                                                    rsPIAdding!Id = maxID
                                                                    rsPIAdding!splitID = splitID
                                                                    rsPIAdding!statementID = rsRentSummaryStatement!statementID
                                                                    rsPIAdding!PINumber = onePI
                                                                    rsPI.Open "Select * from tlbPayment P,tlbPayable B where P.SageAccountNumber=B.clientLandlordID AND " & _
                                                                        " B.clientID=P.ClientID AND slnumber=" & StrDigitVal(onePI) & " and type=6", adoConn, adOpenKeyset, adLockReadOnly
                                                                    If Not rsPI.EOF Then
                                                                            rsPIAdding!amount = rsPI!amount
                                                                            rsPIAdding!OSAmount = rsPI!OSAmount
                                                                            rsPIAdding!SageAccountNumber = rsPI!SageAccountNumber
                                                                            rsPIAdding!ClientID = rsPI!ClientID
                                                                            rsPIAdding!PercentageLL = rsPI!Percentage
                                                                    End If
                                                                    rsPI.Close
                                                                    rsPIAdding.Update
                                                                    splitID = splitID + 1
                                                                    maxID = maxID + 1
                                                                    
                                                             End If
                                                        'rsRentSummaryStatementDetailsNew.Close
                                                    Next
                                                 End If
                                                 rsPIAdding.Close
                          End If
                          rsRentSummaryStatementDetailsNew.Close
                  Else
                        splitID = 1
                        rsPIAdding.Open "Select * from RentSummaryStatementDetails where statementID=" & _
                                                    rsRentSummaryStatement!statementID & "", adoConn, adOpenKeyset, adLockOptimistic
                        rsRentSummaryStatementMAXID.Open "Select (Max(ID)) AS DID from RentSummaryStatementDetails", adoConn, adOpenStatic, adLockReadOnly
                        If Not rsRentSummaryStatementMAXID.EOF Then
                                    maxID = IIf(IsNull(rsRentSummaryStatementMAXID("DID").Value), 0, rsRentSummaryStatementMAXID("DID").Value) + 1
                        End If
                        rsRentSummaryStatementMAXID.Close
                        If rsPIAdding.EOF Then
                            rsPIAdding.AddNew
                            rsPIAdding!statementID = rsRentSummaryStatement!statementID
                            rsPIAdding!PINumber = LTrim(rsRentSummaryStatement!PINumber)
                            rsPI.Open "Select * from tlbPayment P,tlbPayable B where P.SageAccountNumber=B.clientLandlordID AND " & _
                                  " B.clientID=P.ClientID AND slnumber=" & StrDigitVal(LTrim(rsRentSummaryStatement!PINumber)) & " and type=6", adoConn, adOpenKeyset, adLockReadOnly
                            If Not rsPI.EOF Then
                                rsPIAdding!Id = maxID
                                rsPIAdding!splitID = splitID
                                rsPIAdding!amount = rsPI!amount
                                rsPIAdding!OSAmount = rsPI!OSAmount
                                rsPIAdding!SageAccountNumber = rsPI!SageAccountNumber
                                rsPIAdding!ClientID = rsPI!ClientID
                                rsPIAdding!PercentageLL = rsPI!Percentage
                            End If
                            rsPI.Close
                            rsPIAdding.Update
                            
                        End If
                        rsPIAdding.Close
                  End If
                  
            End If
            rsRentSummaryStatement.MoveNext
         End If
    Wend
    
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
    
End Sub
Private Sub LoadLaststatementdate()
    Dim szSQL As String
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    szSQL = "Select StatementDate from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        txtLastStatementDate1.text = rsRentSummaryStatement!StatementDate
    End If
    
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub txtAvailableFund1_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtAvailableFund1, KeyAscii
End Sub

Private Sub txtAvailableFunds_KeyPress(KeyAscii As Integer)
     DigitTextKeyPress txtAvailableFunds, KeyAscii
End Sub

Private Sub txtLastStatementDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientSearch
    End If
    TextBoxKeyPrsDate txtLastStatementDate1, KeyAscii
End Sub
Private Sub txtLastStatementDate1_LostFocus()
    If txtLastStatementDate1.text <> "" Then TextBoxFormatDate txtLastStatementDate1
End Sub
Private Sub txtLastStatementDate1_GotFocus()
   If Len(txtLastStatementDate1.text) < 10 Then txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtLastStatementDate1
End Sub
Private Sub txtLastStatementDate1_Change()
    TextBoxChangeDate txtLastStatementDate1
End Sub

Private Function NextID(adoConn As ADODB.Connection) As Long
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset
   szSQL = "SELECT MAX(Cint(StatementID))+1 AS Ref FROM RentSummaryStatement;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        NextID = IIf(adoRST.EOF, 1, IIf(IsNull(adoRST!ref), 1, adoRST!ref))
   adoRST.Close
   Set adoRST = Nothing
End Function

Private Sub LoadflxPayableTypes()
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New ADODB.Connection
   On Error GoTo ErrorHandler
   Call ConfigflxPayableTypes
   conClient.Open getConnectionString
   szSQL = "SELECT D.*, " & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
                "FROM (PayableTypes AS D INNER JOIN Property AS P ON " & _
                      "D.PropertyID = P.PropertyID) INNER JOIN Client AS C ON P.ClientID = C.ClientID " & _
                      "where C.ClientID='" & szSelectedClient & "' " & _
                " ORDER BY ClientName, PropertyName, D.PayType, D.ID;"


   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      flxPayableTypes.TextMatrix(iRow, 1) = rstClient!Id
      flxPayableTypes.TextMatrix(iRow, 2) = rstClient!PayType
      'flxPayableTypes.TextMatrix(iRow, 3) = rstClient!Id
      rstClient.MoveNext
      If Not rstClient.EOF Then flxPayableTypes.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub
Private Sub ConfigFlxBankAccounts()
   Dim szHeader As String

   flxBankAccounts.Cols = 5
   flxBankAccounts.Clear
   szHeader$ = "|<ID|<Account|<Name|Client_ID"
   flxBankAccounts.FormatString = szHeader$
   flxBankAccounts.ColWidth(0) = 280        'Solid column
   flxBankAccounts.ColWidth(1) = 0          'ID
   flxBankAccounts.ColWidth(2) = 900        'Account
   flxBankAccounts.ColWidth(3) = 3800       'Name
   flxBankAccounts.ColWidth(4) = 0          'ClientID
   flxBankAccounts.Rows = 2
End Sub
Private Sub LoadflxInFunds()
   Dim adoRST   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New ADODB.Connection
'   On Error GoTo ErrorHandler
    If ListOfProperties = "" Then
        Exit Sub
    End If
    If szSelectedClient = "" Then
            Exit Sub
    End If
   ConfigFlxInFunds
   conClient.Open getConnectionString
'   szSQL = "SELECT F.FundID, F.FundName, S.Value " & _
'           "FROM Fund AS F, SecondaryCode AS S " & _
'           "WHERE F.CategoryCode = CBYTE(S.Code) AND S.PrimaryCode = 'DCTG' " & _
'           "ORDER BY FundID;"
'
'   adoRst.Open szSQL, conClient, adOpenStatic, adLockReadOnly
    Dim rsFundMatrix As New ADODB.Recordset
    Dim iSel As Integer
'    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON F.FundID=agr.fund where PB.clientID='" & _
'            szSelectedClient & "' order by fundID;"
     szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=(Fund.FUNDID) where PB.clientID='" & _
            szSelectedClient & "' order by fundID;"
            
    
     
'    rsFundMatrix.Open "Select isfundAssign from shoppingcentre", conClient, adOpenStatic, adLockReadOnly
'    If rsFundMatrix("isfundAssign").Value = False Then
'        iSel = 0
'        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund order by fundID;"
'    Else
'        iSel = 1
'        szSQL = "Select F.*,M.PropertyID from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID  in (" & _
'                ListOfProperties & ") and ClientID='" & szSelectedClient & "' and isDeleted=false order by F.fundID"
'    End If
'    rsFundMatrix.Close
    adoRST.Open szSQL, conClient, adOpenStatic, adLockReadOnly
   

   iRow = 1

   While Not adoRST.EOF
      If iRow = 1 Then
'            flxInFunds.TextMatrix(iRow, 0) = "X"
'            szSelectedFund = adoRst!fundID
      End If
      flxInFunds.TextMatrix(iRow, 1) = adoRST!fundID
      flxInFunds.TextMatrix(iRow, 2) = adoRST!FundName
      flxInFunds.TextMatrix(iRow, 3) = adoRST!FundCode
      If flxInFunds.TextMatrix(iRow, 3) = "TENANTDEPOSIT" Then
                 flxInFunds.TextMatrix(iRow, 0) = "X"
                 flxInFunds.RowHeight(iRow) = 0
      End If
      flxInFunds.ColWidth(4) = 0
      If iSel = 1 Then
            flxInFunds.ColWidth(4) = 1500
            flxInFunds.TextMatrix(iRow, 4) = adoRST!propertyID
      End If
      adoRST.MoveNext
      If Not adoRST.EOF Then flxInFunds.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRST = Nothing
End Sub
Private Sub ConfigFlxGrids()
   Dim szHeader As String

'   flxClients.RowHeight(0) = 0
'   flxClients.ColWidth(0) = 300
'   flxClients.ColWidth(1) = 1000
'   flxClients.ColWidth(2) = 2350

   szHeader$ = "<|Property ID<|Property Name<|<"
   flxProperties.FormatString = szHeader
   flxProperties.Cols = 4
   flxProperties.RowHeight(0) = 0
   flxProperties.ColWidth(0) = 300                   '"X"
   flxProperties.ColWidth(1) = 1000                'Property ID
   flxProperties.ColWidth(2) = 2350                'Property Name
   flxProperties.ColWidth(3) = 0                   'Client ID

'   flxDemandTypes.Cols = 5
'   flxDemandTypes.RowHeight(0) = 0
'   flxDemandTypes.ColWidth(0) = 300                  '"X"
'   flxDemandTypes.ColWidth(1) = 900                  'Property ID
'   flxDemandTypes.ColWidth(2) = 0               'Demand Type ID
'   flxDemandTypes.ColAlignment(2) = vbRightJustify
'   flxDemandTypes.ColWidth(3) = 4000               'Demand Type Name
'   flxDemandTypes.ColWidth(4) = 0                  'Demand Category
'
'   flxCategory.RowHeight(0) = 0
'   flxCategory.ColWidth(0) = 0
'   flxCategory.ColWidth(1) = 0
'   flxCategory.ColWidth(2) = flxCategory.Width - 250
End Sub
Private Sub ConfigFlxProperties()
   Dim szHeader As String
   flxProperties.Clear
   flxProperties.Rows = 2
   szHeader$ = "<|ID<|Name<|<"
   With flxProperties
      .FormatString = szHeader
      .Cols = 4
      .RowHeight(0) = 0
      .ColWidth(0) = 200 'Label2(0).Left - .Left '200                 '"X"
      .ColWidth(1) = 2000 'Label2(1).Left - Label2(0).Left 'Property ID
      .ColWidth(2) = 2500 'Label2(2).Left - Label2(1).Left 'Property Name
      .ColWidth(3) = 0 '.Width + .Left - Label2(2).Left - 300 'Client ID
   End With
End Sub
Private Sub ConfigFlxInFunds()
   Dim szHeader As String

   flxInFunds.Cols = 5
   flxInFunds.Clear
   szHeader$ = "|<ID|<Name|<Fund Code|<Property ID"
   flxInFunds.FormatString = szHeader$
   flxInFunds.ColWidth(0) = 280        'Selection column
   flxInFunds.ColWidth(1) = 400        'fundID
   flxInFunds.ColWidth(2) = 2800       'FundName
   flxInFunds.ColWidth(3) = 1500       'FundCode
   flxInFunds.ColWidth(4) = 0       '
   flxInFunds.Rows = 2
End Sub

Private Sub ConfigflxPayableTypes()
   Dim szHeader As String

   flxPayableTypes.Cols = 4
   flxPayableTypes.Clear
   szHeader$ = "|<ID|<Payable Types|<Category"
   flxPayableTypes.FormatString = szHeader$
   flxPayableTypes.ColWidth(0) = 280        'Solid column
   flxPayableTypes.ColWidth(1) = 400        'ID
   flxPayableTypes.ColWidth(2) = 2800       'Name
   flxPayableTypes.ColWidth(3) = 1500       'empty text
   flxPayableTypes.Rows = 2

   flxPayableTypes.RowHeightMin = 255
End Sub
'Private Sub loadflxProperties()
'    Dim szSQL   As String
'    Dim r       As Integer
'    Dim adoConn As New ADODB.Connection
'    Dim adoRst As New ADODB.Recordset
'    adoConn.Open getConnectionString
'     szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
'           "FROM     PROPERTY where ClientID='" & szSelectedClient & "' " & _
'           "ORDER BY PROPERTYID;"
'           'where ClientID='" & txtClientID.text & "'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   ConfigFlxProperties
'   r = 1
'
'
'   While Not adoRst.EOF
'      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
'      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
'      flxProperties.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value
'      flxProperties.RowHeight(r) = 240
'      r = r + 1
'
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxProperties.AddItem ""
'   Wend
'    Debug.Print r
'   adoRst.Close
'   Set adoRst = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
'   flxProperties.row = 0
'End Sub
Private Sub LoadFlxBankAccounts(szClientID As String)
   Dim conClient As New ADODB.Connection
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer

   On Error GoTo ErrorHandler
   conClient.Open getConnectionString
   ConfigFlxBankAccounts
   szSQL = "SELECT MY_ID, NominalCode, Bank_AC_Name, CLIENT_ID " & _
           "FROM tlbClientBanks where CLIENT_ID='" & szClientID & "'" & _
           "ORDER BY NominalCode;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      If iRow = 1 Then
'         flxBankAccounts.TextMatrix(iRow, 0) = "X"
'         szSelectedBankAccount = rstClient!nominalCode
      End If
      flxBankAccounts.TextMatrix(iRow, 1) = rstClient!My_ID
      flxBankAccounts.TextMatrix(iRow, 2) = rstClient!nominalCode
      flxBankAccounts.TextMatrix(iRow, 3) = rstClient!Bank_AC_Name
      flxBankAccounts.TextMatrix(iRow, 4) = rstClient!CLIENT_ID
      rstClient.MoveNext
      If Not rstClient.EOF Then flxBankAccounts.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub
Private Sub ConfigFlxClients()
    Dim szHeader As String
    flxClients.Clear
    szHeader$ = "|<ClientID|<ClientName|<.."
    flxClients.FormatString = szHeader$
    flxClients.Cols = 4
    flxClients.Rows = 2
    flxClients.RowHeight(0) = 0
    flxClients.ColWidth(0) = 200
    flxClients.ColWidth(1) = 1200
    flxClients.ColWidth(2) = 4000
    flxClients.ColWidth(3) = 0
    
End Sub
Private Sub LoadFlxClients()
   Call ConfigFlxClients
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier where type in ('Client');"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   r = 1
   flxClients.Rows = 1

   While Not adoRST.EOF
      flxClients.AddItem ""
      flxClients.TextMatrix(r, 1) = adoRST.Fields.Item("SupplierID").Value
      flxClients.TextMatrix(r, 2) = adoRST.Fields.Item("SupplierNAME").Value
      r = r + 1
      adoRST.MoveNext
   Wend

'        If r > 1 Then
'                SelectOnly1RowFlxGrid flxClients, 1, 0
'                szSelectedClient = flxClients.TextMatrix(1, 1) 'saving the first propertyID in the list in a variable
'                addPropertiesTowizard szSelectedClient
'                LoadFlxBankAccounts szSelectedClient
'        End If
   adoRST.Close
End Sub
Private Sub LoadFreq()
   Dim adoRstFreq As ADODB.Recordset
   Dim adoConn As ADODB.Connection
   Dim strSQLTitles As String

   Set adoConn = New ADODB.Connection
   Set adoRstFreq = New ADODB.Recordset
   adoConn.Open getConnectionString
   strSQLTitles = "SELECT * FROM FREQUENCIES;"
   adoRstFreq.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaFreq(adoRstFreq.RecordCount) As String

   While Not adoRstFreq.EOF
      szaFreq(adoRstFreq.Fields("ID").Value) = adoRstFreq.Fields("CALDAYS").Value
      adoRstFreq.MoveNext
   Wend
   adoRstFreq.Close
   adoConn.Close
   Set adoRstFreq = Nothing
   Set adoConn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   UnLoadForm Me
   Unload Me
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxClientList
   LoadAllClientFlxGrd
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<ClientID|<ClientName|<ClientPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 0        'Solid column
   conFlxGrid.ColWidth(1) = Label2(1).Left - Label2(0).Left + 300        'Client ID
   conFlxGrid.ColWidth(2) = Label2(2).Left - Label2(1).Left          'Client Name
   conFlxGrid.ColWidth(3) = conFlxGrid.Width - Label2(2).Left - 300  'Post Code
   conFlxGrid.Rows = 2
   conFlxGrid.row = 1
   conFlxGrid.RowHeight(0) = 0
End Sub

Private Sub LoadAllClientFlxGrd()
   Dim conClient As New ADODB.Connection
   Dim rstClient As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conClient.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   If rstClient.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1
   flxClientList.AddItem ""
    flxClientList.TextMatrix(iRow, 1) = "ALL"
      flxClientList.TextMatrix(iRow, 2) = "ALL Client"
   iRow = 2

   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!ClientID
      flxClientList.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(iRow, 3) = IIf(IsNull(rstClient!ClientPostCode), "", rstClient!ClientPostCode)
      flxClientList.RowHeight(iRow) = 280
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

Public Sub LoadProperty()
   cboProperty.Clear

   Dim szSQL As String

   adoProperty.ConnectionString = getConnectionString

   szSQL = "SELECT PropertyID, PropertyName  " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyName;"

   adoProperty.RecordSource = szSQL
   adoProperty.CommandType = adCmdText
   adoProperty.Refresh

   If adoProperty.Recordset.RecordCount < 1 Then Exit Sub

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoProperty.Recordset.RecordCount
   TotalCol = adoProperty.Recordset.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoProperty.Recordset.RecordCount - 1
       For j = 0 To adoProperty.Recordset.Fields.Count - 1
           Data(j, i) = adoProperty.Recordset.Fields(j)
       Next j
       adoProperty.Recordset.MoveNext
   Next i

   cboProperty.Column() = Data()
End Sub

Private Sub tabFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbArrow
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

Private Sub txtClientSearch_Change()
     Dim i As Integer
   For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.RowHeight(i) = 240
            If InStr(1, UCase(flxClients.TextMatrix(i, 1)), UCase(txtClientSearch.text), vbTextCompare) = 0 And txtClientSearch.text <> "" Then
                flxClients.RowHeight(i) = 0
            End If
       
      If flxClients.RowHeight(i) = 240 Then
            flxClients.row = i
      End If
   Next i
End Sub

Private Sub loadflxRetensionDetails()
     Dim adoConn As New ADODB.Connection
     adoConn.Open getConnectionString
     Dim rsRetensionDetails As New ADODB.Recordset
     Dim iRow As Integer
     iRow = 1
     rsRetensionDetails.Open "Select * from RetentionDetails where statementID=" & szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly
     While Not rsRetensionDetails.EOF
            flxRetensionDetails.AddItem ""
            flxRetensionDetails.TextMatrix(iRow, 1) = rsRetensionDetails("statementID").Value
            flxRetensionDetails.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
            flxRetensionDetails.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
            flxRetensionDetails.TextMatrix(iRow, 4) = rsRetensionDetails("Amount").Value
            iRow = iRow + 1
            rsRetensionDetails.MoveNext
     Wend
     rsRetensionDetails.Close
     Set rsRetensionDetails = Nothing
     adoConn.Close
     Set adoConn = Nothing
End Sub
Private Sub ConfigflxRetensionDetails()
        flxRetensionDetails.Clear
        Dim szHeader As String
        szHeader$ = "|<StatementID|<SlNumber|<Amount"
        flxRetensionDetails.FormatString = szHeader$
    
        flxRetensionDetails.Cols = 5
        flxRetensionDetails.Rows = 2
        flxRetensionDetails.RowHeight(0) = 0
        flxRetensionDetails.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxRetensionDetails.ColWidth(1) = 0 'This is statementId
        flxRetensionDetails.ColWidth(2) = 1200 'This is slNumber
        flxRetensionDetails.ColWidth(3) = 1200  'This is Description
        flxRetensionDetails.ColWidth(4) = 1200  'This is amount
        flxRetensionDetails.ColAlignment(3) = vbLeftJustify

End Sub
Private Sub LoadFlxFundList()
        Call ConfigFlxFundList
        Dim adoConn As New ADODB.Connection
        Dim rstRec As New ADODB.Recordset
        Dim szSQL As String
        
        Dim rRow As Integer
        adoConn.Open getConnectionString
        
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
        rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

        rRow = 1
        While Not rstRec.EOF
                flxFundList.TextMatrix(rRow, 0) = ""
                flxFundList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundID").Value
                flxFundList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundCode").Value
                flxFundList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundName").Value
                flxFundList.RowHeight(rRow) = 280
                rstRec.MoveNext
                If Not rstRec.EOF Then flxFundList.AddItem ""
                rRow = rRow + 1
        Wend

        rstRec.Close
        adoConn.Close
        Set rstRec = Nothing
        Set adoConn = Nothing
End Sub
Private Sub ConfigFlxFundList()
        flxFundList.Clear
        Dim szHeader As String
        szHeader$ = "|<FundID|<FundCode|<FundName"
        flxFundList.FormatString = szHeader$
    
        flxFundList.Cols = 4
        flxFundList.Rows = 2
        flxFundList.RowHeight(0) = 0
        flxFundList.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxFundList.ColWidth(1) = 0 'FundID
        flxFundList.ColWidth(2) = 2000 ' FundCode
        flxFundList.ColWidth(3) = 2000  ' FundName
        flxFundList.ColAlignment(0) = vbLeftJustify
        flxFundList.ColAlignment(1) = vbLeftJustify
        flxFundList.ColAlignment(2) = vbLeftJustify
        flxFundList.ColAlignment(3) = vbLeftJustify
End Sub

Private Sub txtPayableDate2_GotFocus()
    txtPayableDate2.SelStart = 0
    txtPayableDate2.SelLength = Len(txtPayableDate2.text)
End Sub



Private Sub txtRentPayable_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtRentPayable, KeyAscii
End Sub

Private Sub txtRentPayable1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtPayableDate2
    End If
    DigitTextKeyPress txtRentPayable1, KeyAscii
End Sub

Private Sub txtRetensionAmount1_GotFocus()
     txtRetensionAmount1.SelStart = 0
     txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
End Sub

Private Sub txtRetensionAmount1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
'            flxRetensionDetails.Enabled = True
'            'Enter data into grid only memory version
'            'statementId you shall generate it when you finally save the statement
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = IIf(Option1.Value = True, "+", "-")
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
'            flxRetensionDetails.AddItem ""
'           ' txtRetensionAmount1.Visible = False
'
'            txtRetensionAmount1.text = "0.00"
'            txtRetensionAmount1.SelStart = 0
'            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
'            Call MakeSummaryRetention
        FocusControl txtRetentionDescriptions
     End If
'     If KeyAscii = 27 Then 'escape ascii key
'        txtRetensionAmount1.Visible = False
'     End If
     DigitTextKeyPress txtRetensionAmount1, KeyAscii
End Sub

Private Sub MakeSummaryRetention()
    Dim iRow As Long
    Dim dblAmt As Double
    For iRow = 1 To flxRetensionDetails.Rows - 1
            If flxRetensionDetails.TextMatrix(iRow, 2) <> "" Then
                    If flxRetensionDetails.TextMatrix(iRow, 0) = "+" Then
                            dblAmt = dblAmt + flxRetensionDetails.TextMatrix(iRow, 4)
                    Else
                            dblAmt = dblAmt - flxRetensionDetails.TextMatrix(iRow, 4)
                    End If
            End If
        Next
    txtRetention.text = dblAmt
End Sub

Private Sub txtRetensionAmount1_LostFocus()
    txtRetensionAmount1.text = Format(txtRetensionAmount1.text, "0.00")
End Sub

Private Sub txtRetentionDescriptions_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdAddToGrid
    End If
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

Private Sub txtSearchClientName_Change()
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

Private Sub txtStatementDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientSearch
    End If
    TextBoxKeyPrsDate txtStatementDate1, KeyAscii
End Sub
Private Sub txtStatementDate1_LostFocus()
    If txtStatementDate1.text <> "" Then TextBoxFormatDate txtStatementDate1
End Sub
Private Sub txtStatementDate1_GotFocus()
   If Len(txtStatementDate1.text) < 10 Then txtStatementDate1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtStatementDate1
End Sub
Private Sub txtStatementDate1_Change()
    TextBoxChangeDate txtStatementDate1
End Sub
Private Sub txtPayableDate2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         FocusControl cmdFundListForCreatePI
    End If
    TextBoxKeyPrsDate txtPayableDate2, KeyAscii
End Sub
Private Sub txtPayableDate2_LostFocus()
    If txtPayableDate2.text <> "" Then TextBoxFormatDate txtPayableDate2
End Sub

Private Sub txtPayableDate2_Change()
    TextBoxChangeDate txtPayableDate2
End Sub


