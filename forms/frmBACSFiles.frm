VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBACSFiles 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View BACS File"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17685
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBACSFiles.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   17685
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picEbanking 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   4770
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   84
      Top             =   2160
      Visible         =   0   'False
      Width           =   6285
      Begin VB.CommandButton cmdCloseEbanking 
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
         TabIndex        =   85
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEbanking 
         Height          =   3615
         Left            =   45
         TabIndex        =   86
         Top             =   540
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6376
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
         Index           =   1
         Left            =   2115
         TabIndex        =   89
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   1
         Left            =   1515
         TabIndex        =   88
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label Label7 
         Height          =   195
         Left            =   135
         TabIndex        =   87
         Top             =   270
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "E Banking"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdViewEBankiing 
      Caption         =   ".."
      Height          =   330
      Left            =   9180
      TabIndex        =   83
      Top             =   135
      Width           =   405
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   37
      Top             =   1035
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "View From File"
      TabPicture(0)   =   "frmBACSFiles.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPayPro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "View From Program"
      TabPicture(1)   =   "frmBACSFiles.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRunNumberSearch"
      Tab(1).Control(1)=   "cmdCreateBACSFile"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Label19(8)"
      Tab(1).Control(4)=   "Label19(7)"
      Tab(1).Control(5)=   "Label19(6)"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtRunNumberSearch 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68160
         TabIndex        =   79
         Top             =   540
         Width           =   1545
      End
      Begin VB.CommandButton cmdCreateBACSFile 
         Caption         =   "&Re-Create a BACS file "
         Height          =   375
         Left            =   -71445
         TabIndex        =   75
         Top             =   5985
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Height          =   5100
         Left            =   -74865
         TabIndex        =   57
         Top             =   855
         Width           =   13245
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBACSTab1 
            Height          =   4095
            Left            =   90
            TabIndex        =   58
            Top             =   945
            Width           =   13020
            _ExtentX        =   22966
            _ExtentY        =   7223
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
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
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Originating"
            Height          =   210
            Index           =   13
            Left            =   1425
            TabIndex        =   73
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reference"
            Height          =   210
            Index           =   12
            Left            =   9090
            TabIndex        =   72
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   210
            Index           =   11
            Left            =   7845
            TabIndex        =   71
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   10
            Left            =   6930
            TabIndex        =   70
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   9
            Left            =   5745
            TabIndex        =   69
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   8
            Left            =   3450
            TabIndex        =   68
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Originating"
            Height          =   210
            Index           =   0
            Left            =   405
            TabIndex        =   67
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   15
            Left            =   405
            TabIndex        =   66
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Number"
            Height          =   255
            Index           =   14
            Left            =   405
            TabIndex        =   65
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   13
            Left            =   1425
            TabIndex        =   64
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   12
            Left            =   1425
            TabIndex        =   63
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   11
            Left            =   3450
            TabIndex        =   62
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code"
            Height          =   255
            Index           =   10
            Left            =   5745
            TabIndex        =   61
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   9
            Left            =   6930
            TabIndex        =   60
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Number"
            Height          =   255
            Index           =   8
            Left            =   6930
            TabIndex        =   59
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label20 
            BackColor       =   &H00EAEAEA&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Index           =   0
            Left            =   90
            TabIndex        =   74
            Top             =   225
            Width           =   12660
         End
      End
      Begin VB.CommandButton cmdPayPro 
         Caption         =   "&Archive BACS File"
         Height          =   375
         Left            =   405
         TabIndex        =   56
         Top             =   5850
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         Height          =   5325
         Left            =   90
         TabIndex        =   38
         Top             =   360
         Width           =   13245
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBACS 
            Height          =   4275
            Left            =   90
            TabIndex        =   39
            Top             =   945
            Width           =   13020
            _ExtentX        =   22966
            _ExtentY        =   7541
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
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
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Number"
            Height          =   255
            Index           =   7
            Left            =   6930
            TabIndex        =   54
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   6
            Left            =   6930
            TabIndex        =   53
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code"
            Height          =   255
            Index           =   5
            Left            =   5745
            TabIndex        =   52
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   4
            Left            =   3450
            TabIndex        =   51
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   3
            Left            =   1425
            TabIndex        =   50
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   2
            Left            =   1425
            TabIndex        =   49
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Number"
            Height          =   255
            Index           =   1
            Left            =   405
            TabIndex        =   48
            Top             =   705
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   47
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Originating"
            Height          =   210
            Index           =   1
            Left            =   405
            TabIndex        =   46
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   3
            Left            =   3450
            TabIndex        =   45
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   4
            Left            =   5745
            TabIndex        =   44
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   210
            Index           =   5
            Left            =   6930
            TabIndex        =   43
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   210
            Index           =   6
            Left            =   7845
            TabIndex        =   42
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reference"
            Height          =   210
            Index           =   7
            Left            =   9090
            TabIndex        =   41
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Originating"
            Height          =   210
            Index           =   2
            Left            =   1425
            TabIndex        =   40
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label20 
            BackColor       =   &H00EAEAEA&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Index           =   18
            Left            =   90
            TabIndex        =   55
            Top             =   225
            Width           =   12660
         End
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View by Run Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   -70050
         TabIndex        =   78
         Top             =   585
         Width           =   1800
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   7
         Left            =   -73155
         TabIndex        =   77
         Top             =   495
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Run Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   6
         Left            =   -74730
         TabIndex        =   76
         Top             =   495
         Width           =   1515
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   13950
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   5
      Top             =   3600
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
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3525
         Left            =   45
         TabIndex        =   7
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
         Left            =   1845
         TabIndex        =   13
         Top             =   375
         Width           =   4320
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "7620;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   12
         Top             =   375
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1845
         TabIndex        =   11
         Top             =   120
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   8
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   13680
      ScaleHeight     =   3570
      ScaleWidth      =   9585
      TabIndex        =   28
      Top             =   5850
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton Command2 
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
         Left            =   9315
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOutputLocation 
         Height          =   3075
         Left            =   45
         TabIndex        =   30
         Top             =   450
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   5424
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
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   0
         Left            =   45
         Top             =   75
         Width           =   9180
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label Label4 
         Height          =   195
         Left            =   120
         TabIndex        =   32
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
      Begin MSForms.Label Label3 
         Height          =   195
         Left            =   1845
         TabIndex        =   31
         Top             =   120
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdFileLocation 
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
      Left            =   11700
      TabIndex        =   27
      Top             =   7515
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame2 
      Height          =   2625
      Left            =   11565
      TabIndex        =   20
      Top             =   2250
      Visible         =   0   'False
      Width           =   14370
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFileLocation 
         Height          =   1365
         Left            =   0
         TabIndex        =   21
         Top             =   675
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   2408
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output File Location:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   24
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Status:"
         Height          =   225
         Index           =   2
         Left            =   10665
         TabIndex        =   23
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   225
         Index           =   3
         Left            =   8460
         TabIndex        =   22
         Top             =   405
         Width           =   900
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   12375
      TabIndex        =   1
      Top             =   7425
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   11295
      TabIndex        =   0
      Top             =   540
      Width           =   1215
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
      Height          =   310
      Left            =   10620
      TabIndex        =   15
      Top             =   6705
      Visible         =   0   'False
      Width           =   345
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
      Left            =   11115
      TabIndex        =   17
      Top             =   7470
      Visible         =   0   'False
      Width           =   345
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
      Left            =   4905
      TabIndex        =   25
      Top             =   7470
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSForms.TextBox txtEBanking 
      Height          =   285
      Left            =   2160
      TabIndex        =   82
      Top             =   135
      Width           =   6975
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "12303;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File Location:"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   81
      Top             =   675
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-banking service:"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   80
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   9495
      TabIndex        =   36
      Top             =   225
      Width           =   45
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Index           =   5
      Left            =   11070
      TabIndex        =   35
      Top             =   135
      Width           =   45
   End
   Begin MSForms.TextBox txtOutputFileLocation 
      Height          =   285
      Left            =   2205
      TabIndex        =   26
      Top             =   675
      Width           =   8730
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "15399;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtBankACNUM 
      Height          =   285
      Left            =   9270
      TabIndex        =   19
      Top             =   7470
      Visible         =   0   'False
      Width           =   1800
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "3175;503"
      Value           =   "ALL"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtBC 
      Height          =   285
      Left            =   6120
      TabIndex        =   18
      Tag             =   "ALL"
      Top             =   7470
      Visible         =   0   'False
      Width           =   3105
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "5477;503"
      Value           =   "ALL BANKS"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   8100
      TabIndex        =   16
      Tag             =   "ALL"
      Top             =   6705
      Visible         =   0   'False
      Width           =   2385
      VariousPropertyBits=   746604571
      Size            =   "4207;556"
      Value           =   "ALL Properties"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1260
      TabIndex        =   14
      Tag             =   "ALL"
      Top             =   7470
      Visible         =   0   'False
      Width           =   3600
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6350;503"
      Value           =   "ALL Clients"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   210
      Index           =   11
      Left            =   435
      TabIndex        =   4
      Top             =   7455
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   210
      Index           =   12
      Left            =   7155
      TabIndex        =   3
      Top             =   6705
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      Height          =   210
      Index           =   0
      Left            =   5490
      TabIndex        =   2
      Top             =   7470
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmBACSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szChoice As String, szaChoice() As String
Dim szFileName As String, szFileEtn As String, szOutPutFilePath As String
Dim sTextBox As String
Dim BACS_ProcessFileLocation As String
Dim szEBSel As Integer
Private Sub CancelButton_Click()
   Unload Me
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
                flxClient.AddItem ""
                flxClient.TextMatrix(rRow, 1) = "ALL"
                flxClient.TextMatrix(rRow, 2) = "ALL Clients"
                flxClient.RowHeight(rRow) = 280
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
          
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub

'Private Sub cmdBC_Click()
'        picClient.Left = 4747.029
'        picClient.Top = 155.299
'        sTextBox = "3"
'        If txtClientList.text = "" Then
'            MsgBox "Please select a client.", vbInformation, "Warning"
'            FocusControl cmdClientList
'            Exit Sub
'        End If
'        LoadBankList
'        'SSTab1.Enabled = False
'        'Frame2.Enabled = False
'        picClient.Visible = True
'        txtSearchClientID.SetFocus
'End Sub
'Private Sub LoadBankList()
''On Error GoTo Error_Handler
'
'   Dim rRow As Integer
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, szaData() As String
'   flxClient.RowHeight(0) = 0
'   flxClient.Cols = 3
'   flxClient.ColWidth(0) = 100
'   flxClient.ColWidth(1) = 1500
'   flxClient.ColWidth(2) = 4500
'
'
'   txtSearchClientID.Width = 1530
'   txtSearchClientName.Visible = True
'   'picClient.Width = 5295
'   'cmdPicCLose.Left = 5010
'
'   flxClient.Clear
'   flxClient.Rows = 2
'   flxClient.ColAlignment(0) = vbLeftJustify
'   flxClient.ColAlignment(1) = vbLeftJustify
'   flxClient.ColAlignment(2) = vbLeftJustify
'
'   '~~~ Added by Anol Configuring width and position of labels and search boxes.
'   lblClientID.Caption = "Bank Account Name"
'   lblClientName.Caption = "Bank Account No"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
'   txtSearchClientName.Left = 1620
'   txtSearchClientName.text = ""
'   txtSearchClientID.text = ""
'   'txtSearchClientName.Width = 3240
'   txtSearchClientID.Left = 45
'   adoConn.Open getConnectionString
'   If txtClientList.Tag = "ALL" Then
'        szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef " & _
'           "FROM tlbClientBanks AS CB, Client AS C " & _
'           "WHERE CB.FileLoc <> '' AND " & _
'               "C.ClientID = CB.CLIENT_ID ;"
'   Else
'         szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef " & _
'           "FROM tlbClientBanks AS CB, Client AS C " & _
'           "WHERE CB.FileLoc <> '' AND " & _
'               "C.ClientID = CB.CLIENT_ID AND " & _
'               "C.ClientID = '" & txtClientList.Tag & "';"
'   End If
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      adoRst.Close
'
'      MsgBox "No bank account has found with electronic banking information."
'      Exit Sub
'   Else
'      ReDim szaData(5, adoRst.RecordCount - 1) As String
'
'      With adoRst.Fields
''         While Not adoRst.EOF
''            szaData(0, iRec) = .Item("MY_ID").Value
''            szaData(1, iRec) = .Item("Bank_AC_Name").Value
''            szaData(2, iRec) = .Item("ClientName").Value
''            szaData(3, iRec) = .Item("BANK_AC_NUM").Value
''            szaData(4, iRec) = .Item("BANK_SC").Value
''            szaData(5, iRec) = IIf(IsNull(.Item("BacsRef").Value), "", .Item("BacsRef").Value)
''
''            iRec = iRec + 1
''            adoRst.MoveNext
''         Wend
'                 rRow = 1
'                 flxClient.Rows = adoRst.RecordCount + 1
'                 flxClient.ColWidth(3) = 0
'                 flxClient.Clear
'                While Not adoRst.EOF
'                    flxClient.row = 1
'                    flxClient.TextMatrix(rRow, 0) = ""
'                    flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("Bank_AC_Name").Value
'                    flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("BANK_AC_NUM").Value
'                    flxClient.TextMatrix(rRow, 3) = adoRst.Fields.Item("MY_ID").Value
'                    If Len(adoRst.Fields.Item("Bank_AC_Name").Value) > 18 Then
'                        flxClient.RowHeight(rRow) = 600
'                    Else
'                        flxClient.RowHeight(rRow) = 280
'                    End If
'                        'flxClient.RowHeight(rRow) = 280
'
'                    adoRst.MoveNext
'                    If Not adoRst.EOF Then flxClient.AddItem ""
'                    rRow = rRow + 1
'                 Wend
'      End With
'
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   adoConn.Close
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub
'
'Private Sub cmdFileLocation_Click()
'        sTextBox = "5"
'        Picture1.Visible = True
'        Picture1.Top = 540
'        Picture1.Left = 2115
'        Call loadflxOutputLocation
'
'End Sub
'Private Sub loadflxOutputLocation()
'        Dim szSQLFileLoc As String
'        Call configflxOutputLocation
'        Dim adoConn As New ADODB.Connection
'        Dim rsLocation As New ADODB.Recordset
'        Dim i As Integer
'        adoConn.Open getConnectionString
'        szSQLFileLoc = "SELECT distinct FileLoc " & _
'        "FROM tlbClientBanks;"
'        rsLocation.Open szSQLFileLoc, adoConn, adOpenKeyset, adLockReadOnly
'        flxOutputLocation.Rows = rsLocation.RecordCount + 1
'        i = 1
'        While Not rsLocation.EOF
'            flxOutputLocation.TextMatrix(i, 1) = IIf(IsNull(rsLocation("FileLoc").Value), "", rsLocation("FileLoc").Value)
'            rsLocation.MoveNext
'            i = i + 1
'        Wend
'
'        rsLocation.Close
'        Set rsLocation = Nothing
'        adoConn.Close
'        Set adoConn = Nothing
'End Sub
Private Sub loadFirstOutputLocation()
        Dim szSQLFileLoc As String
        Call configflxOutputLocation
        Dim adoConn As New ADODB.Connection
        Dim rsLocation As New ADODB.Recordset
        Dim i As Integer
        adoConn.Open getConnectionString
        szSQLFileLoc = "SELECT distinct FileLoc, EB, Indentifier, FileExten, Indentifier " & _
        "FROM tlbClientBanks where EB='" & txtEBanking.Tag & "';"
        szEBSel = Val(txtEBanking.Tag)
        rsLocation.Open szSQLFileLoc, adoConn, adOpenKeyset, adLockReadOnly
        flxOutputLocation.Rows = rsLocation.RecordCount + 1
        i = 1
        If Not rsLocation.EOF Then
            txtOutputFileLocation.text = IIf(IsNull(rsLocation("FileLoc").Value), "", rsLocation("FileLoc").Value) & "\" & IIf(IsNull(rsLocation("Indentifier").Value), "", rsLocation("Indentifier").Value) & "." & Mid(rsLocation("FileExten").Value, 3)
        End If
        rsLocation.Close
        Set rsLocation = Nothing
        If FileExists(txtOutputFileLocation.text) = True Then
            Label19(5).Caption = "File Found"
        Else
            Label19(5).Caption = "File Not Found"
        End If
        adoConn.Close
        Set adoConn = Nothing
End Sub

'Private Sub cboEbanking_Change()
'    flxBACS.Clear
'    If cboEbanking.ListIndex > -1 Then
'            Call loadFirstOutputLocation
'    End If
'End Sub

Private Sub cmdCloseEbanking_Click()
    picEbanking.Visible = False
End Sub

Private Sub cmdCreateBACSFile_Click()
    Dim adoConn As New ADODB.Connection
    If Trim(txtRunNumberSearch.text) = "" Then
        MsgBox "Please enter a run number to create a BACS file", vbInformation, "Warning"
        FocusControl txtRunNumberSearch
        Exit Sub
    End If
    adoConn.Open getConnectionString
    
    If ValidBACSPaymentRun(adoConn, txtRunNumberSearch.text) = False Then
            MsgBox "Run number you are looking for does not exitsts", vbInformation, "Warning"
    Else
        If MsgBox("Are you sure you wish to recreate this BACS file?", vbYesNo, "Please confirm") = vbYes Then
            Call createBACSFilefromDB(adoConn)
        End If
    End If
   
    
    '***********************************************
    'Refresh file status
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim szEB As String
    Dim szOutPutFileLoc  As String
    Dim iFilecount As Long
    
    Dim file As file
     szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn)
    If Right(szOutPutFileLoc, 1) = "\" Then szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)

    iFilecount = -1
    Set FSfolder = FS.GetFolder(szOutPutFileLoc)
    For Each file In FSfolder.Files
            If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                   iFilecount = iFilecount + 1
            End If
    Next file
    If iFilecount = -1 Then
        Label19(5).Caption = "File not found"
        Exit Sub
    Else
         Label19(5).Caption = "File found :" & (iFilecount + 1) & IIf((iFilecount = 0), " File", " Files")
    End If
    Dim filearray() As String
    ReDim filearray(iFilecount) As String
    If iFilecount > -1 Then
        iFilecount = 0
    End If
    For Each file In FSfolder.Files
            If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                   filearray(iFilecount) = file
                   iFilecount = iFilecount + 1
            End If
    Next file
    If iFilecount = -1 Then
        Label19(5).Caption = "File not found"
        Exit Sub
    Else
        szFileName = filearray(0)
        txtOutputFileLocation.text = szFileName
        'Label19(5).Caption = "File found :" & (iFilecount + 1) & IIf((iFilecount = 0), " File", " Files")
    End If
     adoConn.Close
End Sub
Private Function createBACSFilefromDB(adoConn As ADODB.Connection)
        On Error GoTo Err
        Dim rsBACSPaymentRun As New ADODB.Recordset
        Dim szEB As String
        Dim szOutPutFileLoc As String
        rsBACSPaymentRun.Open "Select * from BACSPaymentRun where RunNo=" & txtRunNumberSearch.text & "", adoConn, adOpenKeyset, adLockReadOnly
        szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn)
        
        If Right(BACS_ProcessFileLocation, 1) = "\" Then BACS_ProcessFileLocation = Left(BACS_ProcessFileLocation, Len(BACS_ProcessFileLocation) - 1)
        If BACS_ProcessFileLocation = "" Then Exit Function
        BACS_ProcessFileLocation = BACS_ProcessFileLocation & "\" & txtRunNumberSearch.text & "-" & szFileName & "." & Mid(szFileEtn, 3)
        Open BACS_ProcessFileLocation For Output As #1
        While Not rsBACSPaymentRun.EOF
                 Print #1, rsBACSPaymentRun("description").Value
                 rsBACSPaymentRun.MoveNext
        Wend
        
        Close #1
        rsBACSPaymentRun.Close
        Set rsBACSPaymentRun = Nothing
        MsgBox "BACS file has been created succesfully", vbInformation, "BACS file has been created"
        Exit Function
Err:
        MsgBox Err.description
End Function
Private Sub cmdPicCLose_Click()
        picClient.Visible = False
        cmdClientList.SetFocus
End Sub

Private Sub cmdViewEBankiing_Click()
    Call loadflxEbanking
    picEbanking.Left = 3150
    picEbanking.Top = 135
    picEbanking.Visible = True
    
End Sub

'Private Sub cmdproperty_Click()
'        picClient.Left = 4747.029
'        picClient.Top = 155.299
'        sTextBox = "2"
'        LoadPropertyList
'        'SSTab1.Enabled = False
'        'Frame2.Enabled = False
'        picClient.Visible = True
'        txtSearchClientID.SetFocus
'End Sub

Private Sub Command1_Click()
    picEbanking.Visible = True
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
End Sub





Private Sub MSHFlexGrid1_Click()
    
End Sub

Private Sub flxEbanking_Click()
    txtOutputFileLocation.text = ""
    txtEBanking.text = flxEbanking.TextMatrix(flxEbanking.row, 2)
    txtEBanking.Tag = flxEbanking.TextMatrix(flxEbanking.row, 1)
    Call loadFirstOutputLocation
    picEbanking.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        FocusControl txtRunNumberSearch
        Call viewTab1
    Else
        Call viewTab0
    End If
End Sub

Private Sub txtRunNumberSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call viewTab1
    End If
    LicenceTextKeyPress txtRunNumberSearch, KeyAscii
End Sub

'Private Sub flxOutputLocation_Click()
'     If sTextBox = "5" Then
'             txtOutputFileLocation.text = flxOutputLocation.TextMatrix(flxOutputLocation.row, 1)
'              Label19(5).Caption = ""
'              ConfigFlxBACS
'     End If
'     Picture1.Visible = False
'End Sub

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
Private Sub configflxOutputLocation()
   flxOutputLocation.RowHeight(0) = 0
   flxOutputLocation.Cols = 3
   flxOutputLocation.ColWidth(0) = 0
   flxOutputLocation.ColWidth(1) = 9135
   flxOutputLocation.ColWidth(2) = 0
   flxOutputLocation.Clear
   flxOutputLocation.Rows = 2
End Sub
'Private Sub LoadPropertyList()
'   Dim rRow As Integer
'   Dim szSQL As String
'
'   Dim adoConn As New ADODB.Connection
'   Dim rstRec As New ADODB.Recordset
'   txtSearchClientID.text = ""
'   txtSearchClientName.text = ""
'   flxClient.RowHeight(0) = 0
'   flxClient.Cols = 3
'   flxClient.ColWidth(0) = 0
'   flxClient.ColWidth(1) = 1500
'   flxClient.ColWidth(2) = 4500
'   flxClient.Clear
'   flxClient.Rows = 2
'   flxClient.ColAlignment(0) = vbLeftJustify
'   flxClient.ColAlignment(1) = vbLeftJustify
'   flxClient.ColAlignment(2) = vbLeftJustify
'
'   txtSearchClientID.Width = 1530
'   txtSearchClientName.Visible = True
'   'picClient.Width = 5295
'   'cmdPicCLose.Left = 5010
'   txtSearchClientID.Left = 45
'   '~~~ Added by Anol Configuring width and position of labels and search boxes.
'   lblClientID.Caption = "Property ID"
'   lblClientName.Caption = "Property Name"
''   lblClientID.Width = 1400
''   lblClientID.Left = 50
''   lblClientName.Width = 2600
''   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
'
'   txtSearchClientName.Left = 1620
'   txtSearchClientName.text = ""
'   txtSearchClientID.text = ""
'   'txtSearchClientName.Width = 3240
'   txtSearchClientID.Left = 45
''   picClient.Height = 4095
''   flxClient.Height = 3345
''   flxClient.Width = 5175
'
'
'   adoConn.Open getConnectionString
'
'        szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
'
''Debug.Print szSQL
'   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            rRow = 1
'            flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = "ALL"
'           flxClient.TextMatrix(rRow, 2) = "ALL Properties"
'           flxClient.RowHeight(rRow) = 280
'           flxClient.AddItem ""
'           rRow = 2
'        While Not rstRec.EOF
'           flxClient.row = 1
'           flxClient.RowSel = 1
'               flxClient.ColSel = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
'           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
'           flxClient.RowHeight(rRow) = 280
'           rstRec.MoveNext
'           If Not rstRec.EOF Then flxClient.AddItem ""
'           rRow = rRow + 1
'        Wend
'
'   rstRec.Close
'   adoConn.Close
'   Set rstRec = Nothing
'   Set adoConn = Nothing
'End Sub
'
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
'Private Sub flxClient_Click()
'            Dim szSQL As String
'            Dim adoConn As New ADODB.Connection
'            Dim adoRst As New ADODB.Recordset
'            If sTextBox = "1" Then
'                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
'
'                    If txtClientList.Tag = "ALL" Then
'                        txtBC.Tag = "ALL"
'                        txtBC.text = "ALL BANKS"
'                        txtBankACNUM.text = "ALL"
'                        Call ConfigFlxBACS
''                        flxBACS.Clear
''                        flxBACS.Rows = 1
'                        'lblFileName.text = ""
'                        picClient.Visible = False
'                        Exit Sub
'                    End If
'                    adoConn.Open getConnectionString
'                    szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef " & _
'                           "FROM tlbClientBanks AS CB, Client AS C " & _
'                           "WHERE CB.FileLoc <> '' AND " & _
'                               "C.ClientID = CB.CLIENT_ID AND " & _
'                               "C.ClientID = '" & txtClientList.Tag & "';"
'                   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'                   If adoRst.EOF Then
'                      adoRst.Close
'                      If txtClientList.text = "" Then
'                            MsgBox "Please select a client.", vbInformation, "Warning"
'                            Exit Sub
'                      End If
'                      MsgBox "No bank account has found with electronic banking information.", vbInformation, "Information"
'                      txtBC.Tag = ""
'                      txtBC.text = ""
'                      txtBankACNUM.text = ""
'                   Else
'                      With adoRst.Fields
'                             txtBC.Tag = adoRst.Fields.Item("MY_ID").Value
'                             txtBC.text = adoRst.Fields.Item("Bank_AC_Name").Value
'                             txtBankACNUM.text = adoRst.Fields.Item("BANK_AC_NUM").Value
'                      End With
'                   End If
'                   Call loadflxFileLocation
'                   adoRst.Close
'                   adoConn.Close
'                   txtPropertyName.Tag = "ALL"
'                   txtPropertyName.text = "All Properties"
''                   cmdProperty.SetFocus
'                    FocusControl cmdBC
'            ElseIf sTextBox = "2" Then
'                    txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                    txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
'                    cmdBC.SetFocus
'            ElseIf sTextBox = "3" Then
'                    txtBC.Tag = flxClient.TextMatrix(flxClient.row, 2)
'                    txtBC.text = flxClient.TextMatrix(flxClient.row, 1)
'                    txtBankACNUM.text = flxClient.TextMatrix(flxClient.row, 2)
'                    Call loadflxFileLocation
'                    cmdView.SetFocus
'            End If
''            flxBACS.Clear
''            flxBACS.Rows = 1
'            Call ConfigFlxBACS
'            'lblFileName.text = ""
'            picClient.Visible = False
'            Label19(5).Caption = ""
'
'End Sub

'Private Sub flxClient_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        flxClient_Click
'    End If
'End Sub
Private Sub ConfigFlxBACS()
   Dim szHeader As String

   flxBACS.Clear
   szHeader$ = "|<OANo|<OAName|<DN|<DSC|<DAN|>Amt|<Ref"

   flxBACS.Rows = 2
   flxBACS.Cols = 8
   flxBACS.RowHeight(0) = 0

   flxBACS.FormatString = szHeader$

   flxBACS.ColAlignment(0) = vbCenter
   flxBACS.ColWidth(0) = Label1(1).Left - flxBACS.Left        'Sign
   flxBACS.ColWidth(1) = Label1(2).Left - Label1(1).Left      'OANo
   flxBACS.ColWidth(2) = Label1(3).Left - Label1(2).Left      'OAName
   flxBACS.ColWidth(3) = Label1(4).Left - Label1(3).Left      'DN
   flxBACS.ColWidth(4) = Label1(5).Left - Label1(4).Left      'DSC
   flxBACS.ColWidth(5) = Label1(6).Left - Label1(5).Left      'DAN
   flxBACS.ColWidth(6) = Label1(7).Left - Label1(6).Left      'Amt
   flxBACS.ColWidth(7) = flxBACS.Width + flxBACS.Left - Label1(7).Left - 300      'Ref
End Sub

Private Sub ConfigflxBACSTab1()
   Dim szHeader As String

   flxBACSTab1.Clear
   szHeader$ = "|<OANo|<OAName|<DN|<DSC|<DAN|>Amt|<Ref"

   flxBACSTab1.Rows = 2
   flxBACSTab1.Cols = 8
   flxBACSTab1.RowHeight(0) = 0

   flxBACSTab1.FormatString = szHeader$

   flxBACSTab1.ColAlignment(0) = vbCenter
   flxBACSTab1.ColWidth(0) = Label1(1).Left - flxBACSTab1.Left        'Sign
   flxBACSTab1.ColWidth(1) = Label1(2).Left - Label1(1).Left      'OANo
   flxBACSTab1.ColWidth(2) = Label1(3).Left - Label1(2).Left      'OAName
   flxBACSTab1.ColWidth(3) = Label1(4).Left - Label1(3).Left      'DN
   flxBACSTab1.ColWidth(4) = Label1(5).Left - Label1(4).Left      'DSC
   flxBACSTab1.ColWidth(5) = Label1(6).Left - Label1(5).Left      'DAN
   flxBACSTab1.ColWidth(6) = Label1(7).Left - Label1(6).Left      'Amt
   flxBACSTab1.ColWidth(7) = flxBACSTab1.Width + flxBACSTab1.Left - Label1(7).Left - 300      'Ref
End Sub
Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
    
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPayPro_Click()
   Dim iFileLoop As Integer
   Dim fso As New Scripting.FileSystemObject
   Dim adoConn As New ADODB.Connection
   Dim i As Integer
   Dim szEB As String
   Dim szOutPutFileLoc As String
   Dim iFilecount As Long
   'So the logic is you copy( actually move) only one file from source and rename it and already know there is only one file in unprocessed folder
   If Me.Caption = "Payment Processed" Then
            If MsgBox("Have you processed the payment?", vbQuestion + vbYesNo, "Payment Processed") = vbNo Then Exit Sub
            If MsgBox("System will archive the BACS file. Do you wish to proceed?", vbQuestion + vbYesNo, "Payment Processed") = vbNo Then Exit Sub
            Dim FS As New FileSystemObject
            Dim FSfolder As Folder
            Dim file As file
            'Dim i As Integer
            adoConn.Open getConnectionString
            szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn)
            adoConn.Close
            Set adoConn = Nothing
            If Right(szOutPutFileLoc, 1) = "\" Then szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)
        
            iFilecount = -1
            If szOutPutFileLoc = "" Then Exit Sub
            Set FSfolder = FS.GetFolder(szOutPutFileLoc)
            For Each file In FSfolder.Files
                    If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                           iFilecount = iFilecount + 1
                    End If
            Next file
            If iFilecount = -1 Then
                Label19(5).Caption = "File not found"
                Exit Sub
            Else
                Label19(5).Caption = "File found :" & (iFilecount + 1) & IIf((iFilecount = 0), " File", " Files")
            End If
            Dim filearray() As String
            ReDim filearray(iFilecount) As String
            iFilecount = 0
            For Each file In FSfolder.Files
                    If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                           filearray(iFilecount) = file
                           iFilecount = iFilecount + 1
                    End If
            Next file
            Call CreateNonExistsFolder(DB_PATH & "\AllStuff\BACS_Archive")
            While i < iFilecount
                iFileLoop = 1
                    While iFileLoop > 0
                        'here checking at archive location for a file before copy
                        'Now salia has defined a new location for being archived
                         szFileName = Dir$(filearray(i))
                         szOutPutFilePath = filearray(i)
                         szFileName = Left(szFileName, Len(szFileName) - 4)
                         
                         If Dir$(DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(iFileLoop) & "." & Mid(szFileEtn, 3)) = "" Then
                            'if file info equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                            'if actual file not exists in disk it shall return emptystring by DIR function
                            If (Dir$(szOutPutFilePath) <> "") Then 'if source not found check it else u shall get an error
                                 'fso.copy source, destination
                                  fso.CopyFile szOutPutFilePath, DB_PATH & "\AllStuff\BACS_Archive\" & szFileName & "_" & Format(Now, "yyyyddmm") & "_" & CStr(iFileLoop) & "." & Mid(szFileEtn, 3)
                                  fso.DeleteFile szOutPutFilePath
                                 'szOutPutFilePath value is U:\BACSFILES\Savoy Stewart\PTX\UNPROCESSED BACS FILES\436120.csv
                                  iFileLoop = 0
                            End If
                         Else
                            'if file info not equal to this E:\BOSL3\Prestige Live Code\BACS\436120_20201703_1.csv then u are here
                            'if some file exists with the name it shall return the file name and it shall come here
                            iFileLoop = iFileLoop + 1
                         End If
                    Wend
                    
                    
                i = i + 1
           Wend
            MsgBox "BACS file has been archived successfully.", vbOKOnly + vbInformation, "Information"
            Label19(5).Caption = "File not found."
            flxBACS.Clear
            flxBACS.Rows = 2
                    
   End If
End Sub

'Private Sub configflxFileLocation()
'    flxFileLocation.Clear
'    flxFileLocation.Cols = 10
'    flxFileLocation.ColWidth(0) = 0
'    flxFileLocation.ColWidth(1) = 8500
'    flxFileLocation.ColWidth(2) = 1600
'    flxFileLocation.RowHeight(0) = 0
'    flxFileLocation.ColAlignment(2) = vbLeftJustify
'    flxFileLocation.ColWidth(3) = 2200
'    flxFileLocation.ColWidth(4) = 0
'    flxFileLocation.ColWidth(5) = 0
'    flxFileLocation.ColWidth(6) = 0
'    flxFileLocation.ColWidth(7) = 0 'File name
'    flxFileLocation.ColWidth(8) = 0 'File extension
'    flxFileLocation.ColWidth(9) = 0 'Bank Ac Name
'    flxFileLocation.Rows = 1
'End Sub
'Private Sub loadflxFileLocation()
'    Dim adoConn As New ADODB.Connection
'    Dim i As Integer
'    Dim szCombinedKey As String
'    Dim szOutPutFileLoc  As String
'    Dim adoRstFileLoc  As New ADODB.Recordset
'           If txtClientList.text = "" Then
'                    MsgBox "Please select a client.", vbInformation, "Warning"
'                    FocusControl cmdClientList
'                    Exit Sub
'                End If
'           If txtBC.text = "" Then
'                MsgBox "Please select a Bank account.", vbInformation, "Warning"
'                FocusControl cmdBC
'                Exit Sub
'           End If
'           configflxFileLocation
'           adoConn.Open getConnectionString
'
'           'szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn) 'szEB parameter return value shall be assigned from the functions
'           '***************************************************************
'           Dim adoRst As New ADODB.Recordset
'           Dim szSQL As String
'           Dim szSQLFileLoc As String
'           If txtBankACNUM.text = "ALL" And txtClientList.Tag = "ALL" Then
'                        szSQLFileLoc = "SELECT Distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc,BANK_AC_NUM  " & _
'                        "FROM tlbClientBanks where FileLoc='" & txtOutputFileLocation.text & "';"
'           ElseIf txtBankACNUM.text = "ALL" And txtClientList.Tag <> "ALL" Then
'                        szSQLFileLoc = "SELECT Distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc,BANK_AC_NUM " & _
'                        "FROM tlbClientBanks where clientID='" & txtClientList.Tag & "' AND FileLoc='" & txtOutputFileLocation.text & "';"
'           ElseIf txtBankACNUM.text = "ALL" And txtClientList.Tag = "ALL" Then
'                        szSQLFileLoc = "SELECT Distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc ,BANK_AC_NUM " & _
'                        "FROM tlbClientBanks where FileLoc='" & txtOutputFileLocation.text & "';"
'           Else
'                        szSQLFileLoc = "SELECT Distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc,BANK_AC_NUM " & _
'                        "FROM tlbClientBanks " & _
'                        "WHERE BANK_AC_NUM = '" & txtBankACNUM.text & "'  AND FileLoc='" & txtOutputFileLocation.text & "' ;"
'           End If
'           adoRstFileLoc.Open szSQLFileLoc, adoConn, adOpenStatic, adLockReadOnly
'           If Not adoRstFileLoc.EOF Then
'                flxFileLocation.Rows = RecordCount(adoRstFileLoc) + 1
'           End If
'
''           If txtBC.Tag = "ALL" And txtClientList.Tag = "ALL" Then
''                        szSQL = "SELECT distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc " & _
''                        "FROM tlbClientBanks order by FileLoc"
''           ElseIf txtBC.Tag = "ALL" And txtClientList.Tag <> "ALL" Then
''                        szSQL = "SELECT distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc " & _
''                        "FROM tlbClientBanks where clientID='" & txtClientList.Tag & "'  order by FileLoc ;"
''           ElseIf txtBC.Tag <> "ALL" And txtClientList.Tag = "ALL" Then
''                        szSQL = "SELECT distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc " & _
''                        "FROM tlbClientBanks order by FileLoc;"
''           Else
''                szSQL = "SELECT distinct FileLoc, EB, Indentifier, FileExten,ProcessFileLoc " & _
''                        "FROM tlbClientBanks " & _
''                        "WHERE My_ID = " & txtBC.Tag & "  order by FileLoc;"
''           End If
''           flxFileLocation.Height = 835
''           Frame1.Top = 1650
'           i = 1
'           While Not adoRstFileLoc.EOF
'                    szFileName = adoRstFileLoc.Fields.Item("Indentifier").Value
'                    szOutPutFileLoc = adoRstFileLoc("FileLoc").Value
'                    If Right(szOutPutFileLoc, 1) = "\" Then
'                        szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)
'                    End If
'                    szOutPutFilePath = szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
'                    szCombinedKey = szOutPutFilePath & adoRstFileLoc("EB").Value & adoRstFileLoc("ProcessFileLoc").Value
'
'                         szFileEtn = adoRstFileLoc.Fields.Item("FileExten").Value
'                         flxFileLocation.TextMatrix(i, 1) = IIf(IsNull(adoRstFileLoc("FileLoc").Value), "", adoRstFileLoc("FileLoc").Value)
'                         flxFileLocation.TextMatrix(i, 2) = IIf(IsNull(adoRstFileLoc("Indentifier").Value), "", adoRstFileLoc("Indentifier").Value) & "." & Mid(szFileEtn, 3)
'                         If IsNull(adoRstFileLoc("FileLoc").Value) = True Then
'                               flxFileLocation.TextMatrix(i, 3) = "File Location Not set"
'                         ElseIf adoRstFileLoc("FileLoc").Value = "" Then
'                               flxFileLocation.TextMatrix(i, 3) = "File Location Not set"
'                         ElseIf FolderExists(adoRstFileLoc("FileLoc").Value) = False Then
'                               flxFileLocation.TextMatrix(i, 3) = "Folder does not exists"
'                         ElseIf FolderExists(adoRstFileLoc("FileLoc").Value) = True Then
'                               szFileName = adoRstFileLoc.Fields.Item("Indentifier").Value
'                               szOutPutFileLoc = adoRstFileLoc("FileLoc").Value
'                               If Right(szOutPutFileLoc, 1) = "\" Then
'                                     szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)
'                               End If
'                               szOutPutFilePath = szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
'                               If FileExists(szOutPutFilePath) = True Then
'                                    flxFileLocation.TextMatrix(i, 3) = "File exists"
'                                    flxFileLocation.TextMatrix(i, 6) = szOutPutFilePath
'                               Else
'                                    flxFileLocation.TextMatrix(i, 3) = "File Does not exists"
'                               End If
'                         End If
'                         flxFileLocation.TextMatrix(i, 4) = adoRstFileLoc.Fields.Item("EB").Value
'                         flxFileLocation.TextMatrix(i, 5) = szCombinedKey ' adoRstFileLoc.Fields.Item("MY_ID").Value
'                         flxFileLocation.TextMatrix(i, 7) = adoRstFileLoc("Indentifier").Value
'                         flxFileLocation.TextMatrix(i, 8) = IIf(IsNull(adoRstFileLoc("FileExten").Value), "", adoRstFileLoc("FileExten").Value)
'                         flxFileLocation.TextMatrix(i, 9) = IIf(IsNull(adoRstFileLoc("BANK_AC_NUM").Value), "", adoRstFileLoc("BANK_AC_NUM").Value)
'                         'flxFileLocation.TextMatrix(i, 2) = ""
''                         If i > 1 Then
''
''                           flxFileLocation.Height = flxFileLocation.Height + 240
''                           Frame1.Top = Frame1.Top + 240
''                         End If
'
'                 ' End If
'                  i = i + 1
'                  adoRstFileLoc.MoveNext
'           Wend
'           adoRstFileLoc.Close
'End Sub
Private Function ReturnLastIDBACSPaymentRun(adoConn As ADODB.Connection) As String
    Dim rsBACSPaymentRun As New ADODB.Recordset
    rsBACSPaymentRun.Open "Select RunNo from BACSPaymentRun order by RunNo desc", adoConn, adOpenKeyset, adLockReadOnly
    If rsBACSPaymentRun.EOF Then
        ReturnLastIDBACSPaymentRun = "Not found"
        rsBACSPaymentRun.Close
        Exit Function
    End If
    If Not rsBACSPaymentRun.EOF Then
        ReturnLastIDBACSPaymentRun = rsBACSPaymentRun("RunNo").Value
        rsBACSPaymentRun.Close
    End If
    
End Function
Private Function ValidBACSPaymentRun(adoConn As ADODB.Connection, szRunNumber As String) As Boolean
    Dim rsBACSPaymentRun As New ADODB.Recordset
    rsBACSPaymentRun.Open "Select RunNo from BACSPaymentRun where   RunNo =" & szRunNumber & "", adoConn, adOpenKeyset, adLockReadOnly
    If rsBACSPaymentRun.EOF Then
        ValidBACSPaymentRun = False
        rsBACSPaymentRun.Close
        Exit Function
    End If
    If Not rsBACSPaymentRun.EOF Then
        ValidBACSPaymentRun = True
        rsBACSPaymentRun.Close
    End If
    
End Function
Private Sub viewTab1()
   Call ConfigflxBACSTab1
   Dim adoConn As New ADODB.Connection
   Dim rsBACSPaymentRun As New ADODB.Recordset
   Dim szaFormatLine
   Dim iRow As Long
   Dim szOutPutLine As String
   adoConn.Open getConnectionString
   Label19(7).Caption = ReturnLastIDBACSPaymentRun(adoConn)
   If txtRunNumberSearch.text = "" Then Exit Sub
   rsBACSPaymentRun.Open "Select * from BACSPaymentRun where RunNo=" & txtRunNumberSearch.text & "", adoConn, adOpenKeyset, adLockReadOnly
   flxBACSTab1.Rows = 1
   iRow = 1
   While Not rsBACSPaymentRun.EOF
              If "1" = rsBACSPaymentRun("EB").Value Then                                     'Barclays Business Master
                    szOutPutLine = rsBACSPaymentRun("Description").Value        'Reading the header
                    szaFormatLine = Split(szOutPutLine)
                    flxBACSTab1.AddItem ""
                    flxBACSTab1.TextMatrix(iRow, 1) = "N/A"
                    flxBACSTab1.TextMatrix(iRow, 2) = "N/A"
                    flxBACSTab1.TextMatrix(iRow, 3) = szaFormatLine(1)
                    flxBACSTab1.TextMatrix(iRow, 4) = szaFormatLine(0)
                    flxBACSTab1.TextMatrix(iRow, 5) = szaFormatLine(2)
                    flxBACSTab1.TextMatrix(iRow, 6) = Format(szaFormatLine(3), "0.00")
                    flxBACSTab1.TextMatrix(iRow, 7) = szaFormatLine(4)
                    iRow = iRow + 1
              End If
              If "2" = rsBACSPaymentRun("EB").Value Then                                       'Albany BACS System
                    szOutPutLine = rsBACSPaymentRun("Description").Value
                    flxBACSTab1.AddItem ""
                    flxBACSTab1.TextMatrix(iRow, 1) = Mid(szOutPutLine, 24, 8)
                    flxBACSTab1.TextMatrix(iRow, 2) = Mid(szOutPutLine, 47, 18)
                    flxBACSTab1.TextMatrix(iRow, 3) = Mid(szOutPutLine, 83, 18)
                    flxBACSTab1.TextMatrix(iRow, 4) = Mid(szOutPutLine, 1, 6)
                    flxBACSTab1.TextMatrix(iRow, 5) = Mid(szOutPutLine, 7, 8)
                    flxBACSTab1.TextMatrix(iRow, 6) = Format(Val(Mid(szOutPutLine, 36, 11)) / 100, "0.00")
                    flxBACSTab1.TextMatrix(iRow, 7) = Mid(szOutPutLine, 65, 18)
                    iRow = iRow + 1
             End If
             rsBACSPaymentRun.MoveNext
   Wend
   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub
Private Sub viewTab0()
    'On Error GoTo Err
   Call ConfigFlxBACS
   
   Dim szEB As String
   Dim szOutPutFileLoc  As String, i As Integer
   Dim szFormatLine As String, szColHeading As String
   Dim adoRstFileLoc As New ADODB.Recordset
   Dim iRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim iFilecount As Integer
   adoConn.Open getConnectionString
   'Label19(7).Caption = ReturnLastIDBACSPaymentRun
   'szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn) 'szEB parameter return value shall be assigned from the functions
   '***************************************************************
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim szSQLFileLoc As String
   
   flxBACS.Rows = 2
   iRow = 1
   
'   adoConn.Open getConnectionString
'First get all the files in the folder
'If file count is more that one then halt and give a warning
'Read the file name and give a warning if file count is more than one
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim file As file
    'Dim i As Integer
    
    szOutPutFileLoc = BACS_OPFLocation(adoConn, szEB, szFileName, szFileEtn)
    If Right(szOutPutFileLoc, 1) = "\" Then szOutPutFileLoc = Left(szOutPutFileLoc, Len(szOutPutFileLoc) - 1)

    iFilecount = -1
    Call CreateNonExistsFolder(szOutPutFileLoc)
    If szOutPutFileLoc = "" Then
            MsgBox "Please enter output file location", vbInformation, "Warning"
            Exit Sub
    End If
    Set FSfolder = FS.GetFolder(szOutPutFileLoc)
    For Each file In FSfolder.Files
            If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                   iFilecount = iFilecount + 1
            End If
    Next file
    If iFilecount = -1 Then
        Label19(5).Caption = "File not found"
        Exit Sub
    Else
        Label19(5).Caption = "File found :" & (iFilecount + 1) & IIf((iFilecount = 0), " File", " Files")
    End If
    Dim filearray() As String
    ReDim filearray(iFilecount) As String
    iFilecount = 0
    For Each file In FSfolder.Files
            If UCase(Right(file, 3)) = UCase(Right(szFileEtn, 3)) Then
                   filearray(iFilecount) = file
                   iFilecount = iFilecount + 1
            End If
    Next file
    If iFilecount = -1 Then
        Label19(5).Caption = "File not found"
        Exit Sub
    Else
        szFileName = filearray(0)
        txtOutputFileLocation.text = szFileName
    End If
 
   'szOutPutFilePath = szOutPutFileLoc & "\" & szFileName & "." & Mid(szFileEtn, 3)
   
   szOutPutFilePath = szFileName
   If FileExists(szOutPutFilePath) = False Then
        adoConn.Close
        Label19(5).Caption = "File found :" & (iFilecount + 1) & IIf((iFilecount = 0), " File", " Files")
        Exit Sub
   End If
   iRow = 1
   i = 0
   Dim bSwitch As Boolean
   While i < iFilecount
        Debug.Print i
        If i Mod 2 = 0 Then
            bSwitch = False
        Else
            bSwitch = True
        End If
        Call DisplayBACSInformationInGrid(filearray(i), iRow, szEB, bSwitch)
        i = i + 1
   Wend
   adoConn.Close
   Set adoConn = Nothing
   
   Exit Sub
Err:
   MsgBox Err.description
   Me.MousePointer = vbArrow
End Sub
Private Sub DisplayBACSInformationInGrid(szOutPutFilePath As String, ByRef iRow As Integer, szEB As String, ByRef bSwitch As Boolean)
    Dim szOutPutLine As String
    Dim szaFormatLine() As String
    If Dir$(szOutPutFilePath) <> "" Then                                                  'BACS file found in the folder
      Open szOutPutFilePath For Input As #1
      If szEB = "1" Then                                    'Barclays Business Master
         Line Input #1, szOutPutLine         'Reading the header
         While Not EOF(1)
   '     Read from the file.
            Line Input #1, szOutPutLine         'Reading the data
            szaFormatLine = Split(szOutPutLine)
            flxBACS.AddItem ""
            If bSwitch Then UMarkRowFlxGrid flxBACS, iRow
            flxBACS.TextMatrix(iRow, 1) = "N/A"
            flxBACS.TextMatrix(iRow, 2) = "N/A"
            flxBACS.TextMatrix(iRow, 3) = szaFormatLine(1)
            flxBACS.TextMatrix(iRow, 4) = szaFormatLine(0)
            flxBACS.TextMatrix(iRow, 5) = szaFormatLine(2)
            flxBACS.TextMatrix(iRow, 6) = Format(szaFormatLine(3), "0.00")
            flxBACS.TextMatrix(iRow, 7) = szaFormatLine(4)
            iRow = iRow + 1
         Wend
      End If

      If szEB = "2" Then                                       'Albany BACS System
         'flxBACS.Rows = 1
         While Not EOF(1)
   '     Read from the file.
            Line Input #1, szOutPutLine
            flxBACS.AddItem ""
            If bSwitch Then UMarkRowFlxGrid flxBACS, iRow
            flxBACS.TextMatrix(iRow, 1) = Mid(szOutPutLine, 24, 8)
            flxBACS.TextMatrix(iRow, 2) = Mid(szOutPutLine, 47, 18)
            flxBACS.TextMatrix(iRow, 3) = Mid(szOutPutLine, 83, 18)
            flxBACS.TextMatrix(iRow, 4) = Mid(szOutPutLine, 1, 6)
            flxBACS.TextMatrix(iRow, 5) = Mid(szOutPutLine, 7, 8)
            flxBACS.TextMatrix(iRow, 6) = Format(Val(Mid(szOutPutLine, 36, 11)) / 100, "0.00")
            flxBACS.TextMatrix(iRow, 7) = Mid(szOutPutLine, 65, 18)
            iRow = iRow + 1
         Wend
      End If
      If szEB = "3" Then                                    'Barclays Business Master
         Line Input #1, szOutPutLine         'Reading the header
         While Not EOF(1)
   '     Read from the file.
            Line Input #1, szOutPutLine         'Reading the data
            szaFormatLine = Split(szOutPutLine)
            flxBACS.AddItem ""
            If bSwitch Then UMarkRowFlxGrid flxBACS, iRow
            flxBACS.TextMatrix(iRow, 1) = "N/A"
            flxBACS.TextMatrix(iRow, 2) = "N/A"
            flxBACS.TextMatrix(iRow, 3) = szaFormatLine(1)
            flxBACS.TextMatrix(iRow, 4) = szaFormatLine(0)
            flxBACS.TextMatrix(iRow, 5) = szaFormatLine(2)
            flxBACS.TextMatrix(iRow, 6) = Format(szaFormatLine(3), "0.00")
            flxBACS.TextMatrix(iRow, 7) = szaFormatLine(4)
            iRow = iRow + 1
         Wend
      End If
      Close 1
   Else
     ' lblFileName.Caption = "NO FILE FOUND"
   End If
End Sub
Private Sub cmdView_Click()
   If SSTab1.Tab = 0 Then
        Call viewTab0
   Else
        Call viewTab1
   End If

End Sub

Private Function BACS_OPFLocation(adoConn As ADODB.Connection, ByRef szEB As String, ByRef szFileName As String, ByRef szFileEtn As String) As String
   'On Error GoTo ERR_HANDLER
   'If cboEbanking.ListIndex < 0 Then Exit Function
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
  ' If txtBC.Tag = "ALL" Then
                szSQL = "SELECT FileLoc, EB, Indentifier, FileExten,ProcessFileLoc " & _
                "FROM tlbClientBanks where eb='" & szEBSel & "';"

   'Else
    '    szSQL = "SELECT FileLoc, EB, Indentifier, FileExten " & _
'                "FROM tlbClientBanks " & _
'                "WHERE My_ID = " & txtBC.Tag & ";"
'   End If
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
     If Not adoRst.EOF Then  'if there is only one single bank selected
            If IsNull(adoRst.Fields.Item("EB").Value) Then
               MsgBox "Client's e-banking details are not updated.", vbCritical + vbOKOnly, "Batch Payment"
               adoRst.Clone
               Set adoRst = Nothing
               Exit Function
            End If
            szEB = adoRst.Fields.Item("EB").Value
            szFileName = adoRst.Fields.Item("Indentifier").Value
            szFileEtn = adoRst.Fields.Item("FileExten").Value
            BACS_OPFLocation = adoRst.Fields.Item("FileLoc").Value
            BACS_ProcessFileLocation = adoRst.Fields.Item("ProcessFileLoc").Value
     Else
        'When there is multiple banks
     End If

   adoRst.Close
   Set adoRst = Nothing
   Exit Function

ERR_HANDLER:
   
   Set adoRst = Nothing
End Function
'
Private Sub Form_Activate()
'   ConfigFlxBACS
'   Call loadflxFileLocation
    'MsgBox szTLS
End Sub
Public Sub loadflxEbanking()
    Dim szLine As String
    Dim Data() As String
    Dim i As Integer
    flxEbanking.Clear
    flxEbanking.Rows = 2
    flxEbanking.Cols = 3
    flxEbanking.ColWidth(0) = 150
    flxEbanking.ColWidth(1) = 1200
    flxEbanking.ColWidth(2) = 3300
    flxEbanking.ColAlignment(1) = vbLeftJustify
    flxEbanking.ColAlignment(2) = vbLeftJustify
    
   On Error GoTo CatchErr
   Open App.Path & "\BACS\bacs.txt" For Input As #1
   ReDim Data(1, 0) As String

   i = 0
   While Not EOF(1)
      Line Input #1, szLine
      If Val(szLine) > 0 Then
           flxEbanking.TextMatrix(i, 1) = szLine
         Data(0, i) = szLine
         Line Input #1, szLine
         Data(1, i) = szLine
            flxEbanking.TextMatrix(i, 2) = szLine
         Line Input #1, szLine
         i = i + 1
         ReDim Preserve Data(1, i) As String
      End If
   Wend
   
   
CatchErr:
   Close #1
End Sub
Public Sub E_BankingService()
'loadflxEbanking
   Dim szLine As String
   Dim Data() As String
   Dim i As Integer

   On Error GoTo CatchErr

   Open App.Path & "\BACS\bacs.txt" For Input As #1

   ReDim Data(1, 0) As String

   i = 0
   While Not EOF(1)
      Line Input #1, szLine
      If Val(szLine) > 0 Then
         Data(0, i) = szLine
         Line Input #1, szLine
         Data(1, i) = szLine
         Line Input #1, szLine
         i = i + 1
         ReDim Preserve Data(1, i) As String
      End If
   Wend
   'cboEbanking.Column() = Data()
    If Len(szLine) > 0 Then
            'cboEbanking.ListIndex = 0
    End If
CatchErr:
   Close #1
End Sub
Private Sub Form_Load()
'   Dim adoConn As New ADODB.Connection
'   Dim szSQL As String
   SSTab1.Tab = 0
   Me.Top = 0
   Me.Left = 0
   Me.Width = 13920
   Me.BackColor = MODULEBACKCOLOR
   Call E_BankingService
   Call loadFirstOutputLocation
  ' Call cmdView_Click
   
'   adoConn.Open getConnectionString
'
'   PrepareList adoConn, cboClient, cboProperty
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

'Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
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
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboC.Column() = Data()
'   adoRst.Close
''*************************************** PROPERTY ******************************************
'   If cboC.text <> "" Then
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "WHERE ClientID = '" & cboC.Column(0) & "' " & _
'              "ORDER BY PropertyID;"
'   Else
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "ORDER BY PropertyID;"
'   End If
''   Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
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
'
'   cboP.Column() = Data()
'   cboP.ListIndex = 0
'
'NoRes:
'   adoRst.Close
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

Private Sub LoadBank(adoConn As ADODB.Connection)
   
End Sub

Public Sub TestingCommand()
'   OKButton_Click
'   frmBatchPayment.Testing_Method
End Sub

'Private Sub cboClient_Click()
'   If txtClientList.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   'LoadProperties adoConn, cboProperty, txtClientList.Tag
'   LoadBank adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub LoadProperties(adoConn As ADODB.Connection, cboP As Control, szClientID As String)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, j As Integer
'   Dim i As Integer, Data() As String
'   Dim TotalRow As Integer, TotalCol As Integer
'
'   On Error GoTo ErrorHandler
'
''***************************************  PROPERTY  ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & szClientID & "' " & _
'           "ORDER BY PropertyID;"
''   Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   cboP.Clear
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
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
'
'   cboP.Column() = Data()
'   cboP.ListIndex = 0
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox Err.description & "::" & Err.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub FileFormat(ByRef szFormatLine As String, ByVal szEB As String, ByRef szColHeading As String)
   Dim szLine As String

   Open App.Path & "\BACS\bacs.txt" For Input As #2

   While Not EOF(2)
      Line Input #2, szLine
      
      If szLine = szEB Then
         Line Input #2, szLine
         Line Input #2, szLine
         Line Input #2, szColHeading
         
         szFormatLine = szLine
      Else
         Line Input #2, szLine
         Line Input #2, szLine
         Line Input #2, szColHeading
      End If
   Wend

   Close #2
End Sub
