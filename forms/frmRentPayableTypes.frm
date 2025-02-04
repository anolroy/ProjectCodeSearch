VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentPayableTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payable Types"
   ClientHeight    =   12045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21525
   Icon            =   "frmRentPayableTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12045
   ScaleWidth      =   21525
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   14580
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   39
      Top             =   945
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
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   24
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
      Begin MSForms.Label Label4 
         Height          =   195
         Left            =   4230
         TabIndex        =   83
         Top             =   90
         Visible         =   0   'False
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   4185
         TabIndex        =   82
         Top             =   375
         Width           =   1845
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3254;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   23
         Top             =   375
         Width           =   2520
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4445;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   22
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
         Left            =   1665
         TabIndex        =   43
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   90
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
         TabIndex        =   41
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   40
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   280
         Index           =   15
         Left            =   45
         Top             =   50
         Width           =   5850
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Index           =   1
      Left            =   18945
      TabIndex        =   47
      Top             =   7830
      Visible         =   0   'False
      Width           =   6360
      Begin VB.OptionButton optCopyDemandTemplate 
         Caption         =   "Copy demand template only"
         Height          =   330
         Left            =   1845
         TabIndex        =   57
         Top             =   2295
         Width           =   2310
      End
      Begin VB.OptionButton optCopyDemandType 
         Caption         =   "Copy demand type"
         Height          =   195
         Left            =   1845
         TabIndex        =   56
         Top             =   1890
         Value           =   -1  'True
         Width           =   3480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   135
         TabIndex        =   53
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2175
         TabIndex        =   52
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   3555
         TabIndex        =   51
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4935
         TabIndex        =   50
         Top             =   4320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   45
         ScaleHeight     =   825
         ScaleWidth      =   6210
         TabIndex        =   48
         Top             =   135
         Width           =   6240
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
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Demand Type Copy Wizard"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   150
            TabIndex        =   49
            Top             =   225
            Width           =   6000
         End
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   4875
         Left            =   45
         Top             =   135
         Width           =   6270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   0
         X2              =   6840
         Y1              =   4095
         Y2              =   4095
      End
   End
   Begin VB.Frame fraCommand 
      Height          =   825
      Left            =   45
      TabIndex        =   73
      Top             =   9000
      Width           =   17790
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Co&py"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   2370
         TabIndex        =   11
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   495
         Left            =   405
         TabIndex        =   9
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1380
         TabIndex        =   10
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4350
         TabIndex        =   14
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000009&
         Caption         =   "&Close"
         Height          =   495
         Index           =   0
         Left            =   13860
         TabIndex        =   15
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy  From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Index           =   2
      Left            =   18945
      TabIndex        =   67
      Top             =   5175
      Visible         =   0   'False
      Width           =   6285
      Begin VB.CommandButton cmdFilterbyProperty 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5265
         TabIndex        =   70
         Top             =   315
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperty 
         Height          =   2400
         Left            =   45
         TabIndex        =   68
         Top             =   720
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   4233
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   1260
         TabIndex        =   72
         Top             =   315
         Width           =   630
      End
      Begin MSForms.TextBox txtFilterbyProperty 
         Height          =   285
         Left            =   2340
         TabIndex        =   71
         Tag             =   "ALL"
         Top             =   315
         Width           =   2925
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5159;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         Caption         =   "Properties"
         Height          =   240
         Left            =   180
         TabIndex        =   69
         Top             =   315
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Index           =   5
      Left            =   90
      TabIndex        =   63
      Top             =   10215
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CheckBox chkDemandall2 
         Caption         =   "All Demand"
         Height          =   195
         Left            =   180
         TabIndex        =   64
         Top             =   270
         Width           =   1905
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypeList2 
         Height          =   2580
         Left            =   90
         TabIndex        =   65
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4551
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
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
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Index           =   4
      Left            =   18900
      TabIndex        =   61
      Top             =   7200
      Visible         =   0   'False
      Width           =   6285
      Begin VB.CheckBox chkAllProperties 
         Caption         =   "All Properties"
         Height          =   195
         Left            =   90
         TabIndex        =   66
         Top             =   315
         Width           =   1905
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperty1 
         Height          =   2400
         Left            =   45
         TabIndex        =   62
         Top             =   720
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   4233
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
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
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Index           =   3
      Left            =   6480
      TabIndex        =   58
      Top             =   10215
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CheckBox chkAllDemand 
         Caption         =   "All Demand"
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   270
         Width           =   1905
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypeList 
         Height          =   2580
         Left            =   90
         TabIndex        =   60
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4551
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
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
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5505
      Left            =   45
      TabIndex        =   27
      Top             =   630
      Width           =   17805
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypes 
         Height          =   5010
         Left            =   45
         TabIndex        =   21
         Top             =   405
         Width           =   17640
         _ExtentX        =   31115
         _ExtentY        =   8837
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
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
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   33
         Left            =   17595
         TabIndex        =   79
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   32
         Left            =   14040
         TabIndex        =   78
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   31
         Left            =   12330
         TabIndex        =   25
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   30
         Left            =   10485
         TabIndex        =   77
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   28
         Left            =   7425
         TabIndex        =   76
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   27
         Left            =   6570
         TabIndex        =   31
         Top             =   180
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Payable Type"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   29
         Left            =   8325
         TabIndex        =   30
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   26
         Left            =   3195
         TabIndex        =   29
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   25
         Left            =   360
         TabIndex        =   28
         Top             =   135
         Width           =   450
      End
   End
   Begin VB.Frame fraDemandType 
      Enabled         =   0   'False
      Height          =   2850
      Left            =   45
      TabIndex        =   32
      Top             =   6165
      Width           =   17805
      Begin VB.CommandButton cmdBrowsFile 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2280
         Width           =   390
      End
      Begin VB.TextBox txtDemandTemplate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1950
         Width           =   3285
      End
      Begin VB.CommandButton cmdBrowsFile 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1935
         Width           =   390
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   ".."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   13875
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   675
         Width           =   345
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10350
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   675
         Width           =   3525
      End
      Begin VB.CommandButton cmdproperty 
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
         Height          =   320
         Left            =   7425
         TabIndex        =   0
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   1
         Top             =   510
         Width           =   3825
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10050
         MaxLength       =   4
         TabIndex        =   16
         Top             =   225
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payment Dates:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   33
         Top             =   1170
         Width           =   7725
         Begin VB.CommandButton cmdDemandTypePayDates 
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
            Height          =   320
            Left            =   7200
            TabIndex        =   5
            Top             =   540
            Width           =   345
         End
         Begin VB.TextBox txtDemandTypePayDates 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   540
            Width           =   3825
         End
         Begin MSForms.OptionButton optAuto 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   195
            Width           =   2535
            VariousPropertyBits=   746588179
            BackColor       =   12632256
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4471;661"
            Value           =   "0"
            Caption         =   "Use Automatic Payment Dates"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optPreset 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   495
            Width           =   2175
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3836;661"
            Value           =   "1"
            Caption         =   "Use Preset Payment Dates"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtPrefix 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11865
         MaxLength       =   4
         TabIndex        =   6
         Top             =   225
         Width           =   2025
      End
      Begin MSForms.TextBox txtEmailTemplate 
         Height          =   315
         Left            =   10620
         TabIndex        =   89
         Top             =   2280
         Width           =   3285
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "5794;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email TMP:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   9315
         TabIndex        =   88
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Print TMP:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   9315
         TabIndex        =   87
         Top             =   1935
         Width           =   885
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Use Control Account"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   9315
         TabIndex        =   81
         Top             =   1260
         Visible         =   0   'False
         Width           =   1785
      End
      Begin MSForms.CheckBox chkUseControlAccount 
         Height          =   255
         Left            =   11160
         TabIndex        =   80
         Top             =   1260
         Visible         =   0   'False
         Width           =   255
         VariousPropertyBits=   746588179
         BackColor       =   16764879
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "450;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   315
         Left            =   3600
         TabIndex        =   17
         Tag             =   "ALL"
         Top             =   180
         Width           =   3825
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6747;556"
         Value           =   "All Properties"
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
         Caption         =   "Rent PayableType:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   38
         Top             =   540
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
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
         Index           =   7
         Left            =   9330
         TabIndex        =   37
         Top             =   225
         Width           =   255
      End
      Begin MSForms.CheckBox chkGroup 
         Height          =   255
         Left            =   10035
         TabIndex        =   7
         Top             =   690
         Width           =   255
         VariousPropertyBits=   746588179
         BackColor       =   16764879
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "450;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   9315
         TabIndex        =   36
         Top             =   690
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   11265
         TabIndex        =   35
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   330
         TabIndex        =   34
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   45
      TabIndex        =   44
      Top             =   0
      Width           =   17790
      Begin VB.CommandButton cmdPropertyFilter 
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
         Left            =   12780
         TabIndex        =   75
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton cmdClientFilter 
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
         TabIndex        =   74
         Top             =   180
         Width           =   345
      End
      Begin MSForms.TextBox txtPropertyList 
         Height          =   300
         Left            =   9360
         TabIndex        =   20
         Tag             =   "ALL"
         Top             =   180
         Width           =   3420
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6032;538"
         Value           =   "ALL Properties"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   305
         Left            =   2520
         TabIndex        =   19
         Tag             =   "ALL"
         Top             =   180
         Width           =   3240
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5715;538"
         Value           =   "ALL Clients"
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
         Caption         =   "By Property"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   8325
         TabIndex        =   46
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By client"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   1080
         TabIndex        =   45
         Top             =   225
         Width           =   1110
      End
   End
   Begin VB.Label lblFrameIndex 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "1"
      Height          =   195
      Left            =   18945
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmRentPayableTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
           As Long) As Long
Const SW_SHOW = 5
Dim MAX_NUMBER_FRAME_INDEX As Integer
Dim szDemandId          As String
Dim szDemandType        As String
Dim bAddNew             As Boolean
Dim iSelRow             As Integer
Dim bSortingCol1        As Boolean
Dim bSortingCol2        As Boolean
Dim bSortingCol3        As Boolean
Dim bSortingCol4        As Boolean
Dim szExistingProperty  As String
Dim strCommandSource As String
Dim EditMode As Boolean
Dim copyfromDemandTypeID As String
Dim intTopRow As Integer
Dim szSelectedClient As String
'Private Sub cboProperty_Click()
'   'Resolved by BOSL
'   'issue 475
'   'Modified by anol 23 Oct 2014
'   Call LoadNCinCombo1
'End Sub
'Private Sub LoadNCinCombo1()
'   Dim adoconn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, TotalRow As Integer
'   Dim Data() As String, i As Integer
'   If adoconn.State = 0 Then
'      adoconn.Open getConnectionString
'   End If
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger,property where property.clientID=NominalLedger.clientID AND property.PropertyID='" & txtProperty.Tag & "' order by NominalLedger.code;"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   If Not adoRst.EOF Then
'        txtDemandTypeNCAmt.text = adoRst.Fields.Item("Code").Value
'        txt1.text = adoRst.Fields.Item("Name").Value
'   End If
''   TotalRow = adoRst.RecordCount
''   ReDim Data(2, TotalRow) As String
''
''   i = 0
''   While Not adoRst.EOF
''      Data(0, i) = adoRst.Fields.Item("Code").Value
''      Data(1, i) = adoRst.Fields.Item("Name").Value
''      i = i + 1
''      adoRst.MoveNext
''   Wend
''   cboDemandTypeNCAmt.Column() = Data()
'''   cboDemandTypeNCvat.Column() = Data()
'''   cboDemandTypeNCTotal.Column() = Data()
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   If adoconn.State = 1 Then
'      adoconn.Close
'   End If
'End Sub
Public Sub GetRecord()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
'This sub procedure is called when you are changing the rowcol of the grid
'   On Error GoTo KatchErr
   If szDemandId = "" Then Exit Sub
   adoConn.Open getConnectionString
   Dim rsPaymentDates As New ADODB.Recordset
   
'    szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, " & _
'               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
'               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
'               "CategoryCode, PaymentDates, Spare1, GroupID, DemandReportName, " & _
'               "EmailInvoiceTemplate,A.Value,StatementTemplate,Consolidated " & _
'            "FROM PayableTypes,SecondaryCode A " & _
'            "WHERE CSTR(PayableTypes.CategoryCode)=A.Code AND  PrimaryCode= 'DCTG' AND ID = " & szDemandID & ""
            
    szSQL = "SELECT ID,CL.ClientName,CL.ClientID,P.PropertyName,P.PropertyID, PayType, PayIC, PaySagePrefix," & _
                  "PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, TransactionType, " & _
                  "CategoryCode, PaymentDates,GroupID,PayNCAmt,PayNNAmt,isUseControlAccount " & _
                "FROM PayableTypes C,Property P,Client CL where CL.ClientID=C.ClientID  " & _
                "and P.ClientID=CL.ClientID AND P.PropertyID=C.PropertyID  " & _
                "AND ID = " & szDemandId & " ORDER BY ID"

           
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If adoRst.EOF Then
        MsgBox "Current record not found", vbCritical, "Failed"
        adoRst.Close
        Exit Sub
   End If
'    Exit Sub
   txtType.text = adoRst!PayType
   txtID.text = adoRst!ID
   
   szSelectedClient = adoRst!clientID
   txtProperty.Tag = adoRst!propertyID
   txtProperty.text = adoRst!PropertyName
  
'   If IsNull(adoRst!PayNCAmt) = False Then
'        txtDemandTypeNCAmt.text = adoRst!PayNCAmt
'   End If
'   If IsNull(adoRst!PayNNAmt) = False Then
'            txt1.text = adoRst!PayNNAmt
'   End If
   txtPrefix.text = IIf(IsNull(adoRst!PaySagePrefix), "", IIf(adoRst!PaySagePrefix = "NULL", "", adoRst!PaySagePrefix))
'   If Not IsNull(adoRst!CategoryCode) Then txtDemandCategoryCode.text = adoRst!CategoryCode
'   txtDemandCategoryName.text = adoRst!CategoryName
   If adoRst!PaymentDates = 255 Then
      optAuto.Value = True
   Else
      optPreset.Value = True
      rsPaymentDates.Open "Select NameofSet from PaymentDates where DateSetID=" & CInt(adoRst!PaymentDates) & "", adoConn, adOpenStatic, adLockReadOnly
      If Not rsPaymentDates.EOF Then
            txtDemandTypePayDates.Tag = CInt(adoRst!PaymentDates)
            txtDemandTypePayDates.text = rsPaymentDates!NameOfSet
      Else
            txtDemandTypePayDates.Tag = "0"
            txtDemandTypePayDates.text = "DEFAULT"
      End If
      rsPaymentDates.Close
      Set rsPaymentDates = Nothing
   End If
   If Val(IIf(IsNull(adoRst!GroupID), 0, adoRst!GroupID)) > 0 Then
      txtGroup.text = IIf(IsNull(adoRst!GroupID), 0, adoRst!GroupID)
      chkGroup.Value = True
   Else
      chkGroup.Value = False
   End If
   If IIf(IsNull(adoRst!isUseControlAccount), False, adoRst!isUseControlAccount) = True Then
        chkUseControlAccount.Value = True
   Else
        chkUseControlAccount.Value = False
   End If
    
   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

KatchErr:
   ShowMsgInTaskBar "Could not load Pay Charges due to to some error", "Y", "N"
   adoConn.Close
   Set adoConn = Nothing
End Sub





'Private Sub cboDemandTypeCategory_Change()
'   On Error GoTo ErorrHandler
'
'   txt4.text = cboDemandTypeCategory.Column(1)
'
'   Exit Sub
'ErorrHandler:
'   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
'End Sub

'Private Sub cboDemandTypeNCAmt_Change()
'   On Error GoTo ErorrHandler
'
'   txt1.text = cboDemandTypeNCAmt.Column(1)
'
'   Exit Sub
'ErorrHandler:
'   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
'End Sub

'Private Sub cboDemandTypeNCTotal_Change()
'   On Error GoTo ErorrHandler
'
'   txt3.text = cboDemandTypeNCTotal.Column(1)
'
'   Exit Sub
'ErorrHandler:
'   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
'End Sub

'Private Sub cboDemandTypeNCvat_Change()
'   On Error GoTo ErorrHandler
'
'   txt2.text = cboDemandTypeNCvat.Column(1)
'
'   Exit Sub
'ErorrHandler:
'   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
'End Sub

Private Function SecCodeValue(szPrimaryCode As String, szCode As String) As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT Value " & _
            "FROM SecondaryCode " & _
            "WHERE PrimaryCode = '" & szPrimaryCode & "' AND Code = '" & szCode & "'"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   SecCodeValue = adoRst!Value
   adoRst.Close
   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub cboDemandTypeCategory_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
'        cboBank.SetFocus
    End If
End Sub

Private Sub cboDemandTypePayDates_Click()
   optPreset.Value = 1
End Sub

'Private Sub cboGroup_GotFocus()
'   If cboDemandTypeCategory.text = "" Then
'      MsgBox "Please select the Demand Category first.", vbCritical + vbOKOnly, "Demand Type"
'      cboDemandTypeCategory.SetFocus
'   End If
'End Sub

Private Sub cboProperty_Change()
   Dim adoConn As New ADODB.Connection
   If txtProperty.Tag = "" Then
      If txtProperty.text <> "" Then txtProperty.text = ""
      Exit Sub
   End If

   If szExistingProperty <> "" And txtProperty.text <> "" And EditMode Then
      If txtProperty.Tag = "ALL" Then
         MsgBox "Do not set this demand type to all properties.", vbCritical + vbOKOnly, "Demand Type"
         cmdproperty.SetFocus
         txtProperty.Tag = ""
         txtProperty.text = ""
      Else
'         adoConn.Open getConnectionString
'         LoadClientBankDetails adoConn 'LoadClientBankDetails of that property/client anol
'         adoConn.Close
'         Set adoConn = Nothing
      End If
   End If

   If txtProperty.text <> "" And cmdSaveNew.Enabled Then
'      adoConn.Open getConnectionString
'      LoadClientBankDetails adoConn 'LoadClientBankDetails of that property/client anol
'      adoConn.Close
'      Set adoConn = Nothing
   End If
End Sub

Private Sub cboProperty_GotFocus()
   If txtProperty.Tag = "" Then Exit Sub

   If EditMode Then
      If MsgBox("Do you want to change the Property?", _
                 vbQuestion + vbYesNo, _
                "Demand Type - Allocation") = vbYes Then
         cmdproperty.SetFocus
         If txtProperty.text <> "" Then
            szExistingProperty = txtProperty.Tag
         Else
            szExistingProperty = ""
         End If
         Exit Sub
      End If
      cmdSaveNew.SetFocus
   End If
End Sub

'Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If txtProperty.text = "" Then
'      If txtProperty.text <> "" Then txtProperty.text = ""
'      KeyAscii = 0
'   End If
'End Sub
Private Sub setStatementonPropertyClick()
'            Dim adoConn As New ADODB.Connection
'            Dim adoRst  As New ADODB.Recordset
'            If txtProperty.Tag = "ALL" Then Exit Sub
'            adoConn.Open getConnectionString
'            adoRst.Open "Select StatementTemplate from PayableTypes where PropertyID='" & txtProperty.Tag & "' AND (StatementTemplate<>'' or StatementTemplate is not null) order by ID DESC", adoConn, adOpenStatic, adLockReadOnly
'            If Not adoRst.EOF Then
'                If FileExists(App.Path & "\CompanyReports\" & adoRst.Fields("StatementTemplate").Value & "") Then
'                     txtStatementTemplate.text = adoRst.Fields("StatementTemplate").Value
'                Else
'                     txtStatementTemplate.text = ""
'                End If
'            End If
'            adoRst.Close
'            adoConn.Close
            
            
End Sub
Private Function checkExistingLeaseBeforeChangeProperty() As Boolean
    'Mark 1 Leave this for not until salia give you full spec on this module
    Exit Function
'   If txtProperty.text = "" Then Exit Function
'
'   If szExistingProperty <> "" And txtProperty.text <> "" And EditMode Then
'      If txtProperty.Tag = "ALL" Then
'         MsgBox "Do not set this demand type to all properties.", vbCritical + vbOKOnly, "Demand Type"
'         cmdProperty.SetFocus
'         txtProperty.text = ""
'         txtProperty.Tag = ""
'      End If
'
'      If szExistingProperty <> txtProperty.Tag Then
'         Dim adoconn As New ADODB.Connection
'         Dim adoRst  As New ADODB.Recordset
'         Dim szSQL   As String
'
'         adoconn.Open getConnectionString
'
'         If txtDemandCategoryCode.text = "1" Then                              'Rent Charge
'            szSQL = "SELECT PropertyName " & _
'                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
'                          "FROM   PayableTypes AS T, LRentCharges AS R, " & _
'                               "LeaseDetails AS L, Units AS U, Property AS P " & _
'                          "WHERE  T.ID = R.BRDemandType AND R.LeaseID = L.LeaseID AND " & _
'                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
'                    "WHERE Q.PropertyID <> '" & txtProperty.Tag & "';"
''Debug.Print szSQL
'         End If
'         If txtDemandCategoryCode.text = "2" Then                              'Service Charge
'            szSQL = "SELECT PropertyName " & _
'                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
'                          "FROM   PayableTypes AS T, LServiceCharges AS S, " & _
'                               "LeaseDetails AS L, Units AS U, Property AS P " & _
'                          "WHERE  T.ID = S.SCDemandType AND S.LeaseID = L.LeaseID AND " & _
'                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
'                    "WHERE Q.PropertyID <> '" & txtProperty.Tag & "';"
''Debug.Print szSQL
'         End If
'         If txtDemandCategoryCode.text = "3" Then                              'Insurance Charge
'            szSQL = "SELECT PropertyName " & _
'                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
'                          "FROM   PayableTypes AS T, LInsuranceCharges AS I, " & _
'                               "LeaseDetails AS L, Units AS U, Property AS P " & _
'                          "WHERE  T.ID = I.InsuranceDemandType AND I.LeaseID = L.LeaseID AND " & _
'                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
'                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
'                    "WHERE Q.PropertyID <> '" & txtProperty.Tag & "';"
''Debug.Print szSQL
'         End If
'        If szSQL = "" Then
'            adoconn.Close
'            Set adoconn = Nothing
'            Exit Function
'        End If
'         adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'         If Not adoRst.EOF Then
'            szSQL = ""
'            While Not adoRst.EOF
'               szSQL = adoRst.Fields.Item(0).Value
'               adoRst.MoveNext
'               If Not adoRst.EOF Then szSQL = szSQL & ", "
'            Wend
'
'            MsgBox "This demand type is being used in leases by " & szSQL & "." & Chr(13) & _
'                   "Please reschedule the demand type in the lease first.", vbCritical + vbOKOnly, "Demand Types"
'            txtProperty.text = ""
'            txtProperty.Tag = ""
'            checkExistingLeaseBeforeChangeProperty = True
'         End If
'         adoRst.Close
'
'
'         Set adoRst = Nothing
'         adoconn.Close
'         Set adoconn = Nothing
'      End If
'   End If
End Function

Private Sub cboDemandTypePayDates_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
'        FocusControl cmdBrowsFile(0)
    End If
End Sub

Private Sub chkAllDemand_Click()
        Dim i As Integer
        Dim iIncDec As Integer

          For i = 1 To flxDemandTypeList.Rows - 1
          If flxDemandTypeList.TextMatrix(i, 1) <> "" Then
                iIncDec = iIncDec + SelectALLFlxGridRow(0, flxDemandTypeList, i, chkAllDemand.Value)
          End If
          Next i
End Sub

Private Sub chkAllProperties_Click()
    Dim i As Integer
    Dim iIncDec As Integer
    
    For i = 1 To flxProperty1.Rows - 1
       If flxProperty1.TextMatrix(i, 1) <> "" Then
          iIncDec = iIncDec + SelectALLFlxGridRow(0, flxProperty1, i, chkAllProperties.Value)
       End If
    Next i

End Sub
Private Function SelectALLFlxGridRow(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer, Sel As Boolean) As Integer
   Dim iRow As Integer
    
   If Not Sel Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      For iRow = 1 To conFlxGrid.Cols - 1
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iRow
      SelectALLFlxGridRow = 0
   Else
      conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
      conFlxGrid.row = iSelRow
      For iRow = 1 To conFlxGrid.Cols - 1
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(174, 179, 233)
      Next iRow
      SelectALLFlxGridRow = SelectALLFlxGridRow + 1
   End If
End Function

Private Sub chkDemandall2_Click()
     Dim i As Integer
        Dim iIncDec As Integer

          For i = 1 To flxDemandTypeList2.Rows - 1
          If flxDemandTypeList2.TextMatrix(i, 1) <> "" Then
'                iIncDec = iIncDec + SelectALLFlxGridRow(0, flxDemandTypeList2, i, chkDemandall2.Value)
          End If
          Next i
End Sub

Private Sub chkGroup_Click()
   If chkGroup.Value Then
      cmdGroup.Enabled = True
      If EditMode Or cmdSaveNew.Enabled Then FocusControl cmdGroup
   Else
      txtGroup.text = ""
      cmdGroup.Enabled = False
   End If
End Sub

Private Sub cmdAdd_Click()
   EditMode = False
   Call AddNewDemandType
   flxDemandTypes.Enabled = False
   txtType.SetFocus
   szDemandId = ""
   txtProperty.Tag = ""
   txtProperty.text = ""
   
   cmdproperty.SetFocus
End Sub

Public Sub AddNewDemandType()
   Call EmptyBoxes

   cmdDelete.Enabled = False
   cmdEdit.Enabled = False
   cmdSaveNew.Enabled = True
   cmdSaveNew.ZOrder 0
   cmdCancelNew.Enabled = True
   cmdAdd.Enabled = False

'   Call EnableBoxes
   fraDemandType.Enabled = True
   chkGroup.Enabled = True
End Sub

Private Sub cmdBack_Click()
    

     If Val(lblFrameIndex.Caption) = 2 Then
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
'       Frame1(Val(lblFrameIndex.Caption)).Top = 1800
'       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
     
    End If
   If Val(lblFrameIndex.Caption) = 3 Then
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
       Frame1(Val(lblFrameIndex.Caption)).Top = 1800
'       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
       
    End If
   If Val(lblFrameIndex.Caption) = 4 Then
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
       Frame1(Val(lblFrameIndex.Caption)).Top = Frame1(Val(lblFrameIndex.Caption) - 1).Top
'       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
     
    End If
   If Val(lblFrameIndex.Caption) = 5 Then
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
       Frame1(Val(lblFrameIndex.Caption)).Top = Frame1(Val(lblFrameIndex.Caption) - 1).Top
'       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0

    End If
   cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
   cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
   cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
End Sub

Private Sub cmdBank_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "3"
    If txtProperty.Tag = "ALL" Or txtProperty.Tag = "" Then
        LoadflxBank "ALL"
    Else
        LoadflxBank txtProperty.Tag
    End If
    fraDemandType.Enabled = False
    fraCommands.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdBrowsFile_Click(Index As Integer)
'   Dim ofn As OPENFILENAME
'   Dim lHwnd As Long
'   Const HKEY_LOCAL_MACHINE As Long = &H80000002
'   Dim szOldFile_PathName As String
'   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
'   Dim fso As Object
'
'   ofn.lStructSize = Len(ofn)
'   ofn.hwndOwner = lHwnd
'   ofn.hInstance = App.hInstance
'   ofn.lpstrFilter = "All Files (*.rpt)" + Chr$(0) + "*.rpt" + Chr$(0)
'   ofn.lpstrFile = Space$(254)
'   ofn.nMaxFile = 255
'   ofn.lpstrFileTitle = Space$(254)
'   ofn.nMaxFileTitle = 255
'   ofn.lpstrInitialDir = CurDir & "\CompanyReports"
'   ofn.lpstrTitle = "Select a Report file"
'   ofn.Flags = 0
'
'   If GetOpenFileName(ofn) = 0 Then Exit Sub
'
'   If Index = 0 Then txtDemandTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 1 Then txtEmailTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 2 Then txtStatementTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 0 Then
'        FocusControl cmdBrowsFile(1)
'   ElseIf Index = 1 Then
'        FocusControl cmdBrowsFile(2)
'   ElseIf Index = 2 Then
'        FocusControl cmdSaveNew
'   End If
End Sub
Private Sub Loadflxgroup()
    Label4.Visible = False
    Dim szSQL As String
    Dim TotalRow As Integer, TotalCol As Integer
    Dim i As Integer, j As Integer
    Dim adoRst As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim rRow As Integer
    Dim iSt As Integer, iEnd As Integer
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 4500
   flxClient.ColWidth(2) = 0
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Group Name"
   lblClientName.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   adoConn.Open getConnectionString
  szSQL = "SELECT CODE, VALUE " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'GR';"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("Code").Value = "ENDRNG" Then
         iEnd = adoRst.Fields.Item("VALUE").Value
      Else
         iSt = adoRst.Fields.Item("VALUE").Value
      End If
      adoRst.MoveNext
   Wend
'    rRow = 1
'    flxClient.TextMatrix(rRow, 0) = ""
'    flxClient.TextMatrix(rRow, 1) = "ALL"
'    flxClient.TextMatrix(rRow, 2) = "All Properties"
'    flxClient.RowHeight(rRow) = 280
'    flxClient.AddItem ""
    rRow = 1
    For i = iSt To iEnd
        flxClient.row = 1
        flxClient.RowSel = 1
        flxClient.ColSel = 1
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = i
        flxClient.TextMatrix(rRow, 2) = ""
        flxClient.TextMatrix(rRow, 3) = ""
        flxClient.TextMatrix(rRow, 3) = ""
        flxClient.RowHeight(rRow) = 280
'        adoRST.MoveNext
'        If Not adoRST.EOF Then
        flxClient.AddItem ""
        rRow = rRow + 1
    Next i
    adoRst.Close
    Set adoRst = Nothing
End Sub
Private Sub LoadflxBank(strPropertyID As String)
    Dim szSQL As String
    Dim TotalRow As Integer, TotalCol As Integer
    Dim i As Integer, j As Integer
    Dim rstRec As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim rRow As Integer
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 4500
   flxClient.ColWidth(2) = 0
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Bank_AC_Name"
   lblClientName.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   adoConn.Open getConnectionString
  If strPropertyID = "ALL" Then
            szSQL = "SELECT My_ID, Bank_AC_Name, BANK_AC_NUM, BANK_SC " & _
            "FROM tlbClientBanks;"
  Else
            szSQL = "SELECT My_ID, Bank_AC_Name, BANK_AC_NUM, BANK_SC " & _
            "FROM tlbClientBanks, Property " & _
            "WHERE tlbClientBanks.CLIENT_ID = Property.ClientID AND " & _
                  "Property.PropertyID = '" & strPropertyID & "';"
  End If
    
    rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    rRow = 1
'    flxClient.TextMatrix(rRow, 0) = ""
'    flxClient.TextMatrix(rRow, 1) = "ALL"
'    flxClient.TextMatrix(rRow, 2) = "All Properties"
'    flxClient.RowHeight(rRow) = 280
'    flxClient.AddItem ""
    rRow = 1
    While Not rstRec.EOF
        flxClient.row = 1
        flxClient.RowSel = 1
        flxClient.ColSel = 1
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("Bank_AC_Name").Value
        flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("My_ID").Value
        flxClient.TextMatrix(rRow, 3) = rstRec.Fields.Item("BANK_AC_NUM").Value
        flxClient.TextMatrix(rRow, 3) = rstRec.Fields.Item("BANK_SC").Value
        flxClient.RowHeight(rRow) = 280
        rstRec.MoveNext
        If Not rstRec.EOF Then flxClient.AddItem ""
        rRow = rRow + 1
     Wend
    rstRec.Close
    Set rstRec = Nothing
End Sub
Private Sub cmdCancel_Click()
   Call GetRecord
'   Call DisableBoxes
   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   EditMode = False
   szExistingProperty = ""
End Sub

Private Sub cmdCancelNew_Click()
    If EditMode Then
        cmdCancel_Click
        Exit Sub
   End If
   Call EmptyBoxes
'   Call DisableBoxes
   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True
   
   cmdPropertyFilter.Enabled = True
   cmdClientFilter.Enabled = True
   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False
   fraDemandType.Enabled = False
End Sub


Private Sub LoadflxProperty()
    
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   Label4.Visible = True
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 2800
   flxClient.ColWidth(3) = 1700
   flxClient.ColWidth(4) = 0
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""

   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   szSQL = "SELECT PropertyID, PropertyName,ClientID " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                    rRow = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "ALL"
                    flxClient.TextMatrix(rRow, 2) = "All Properties"
                    flxClient.RowHeight(rRow) = 280
                    flxClient.AddItem ""
                    rRow = 2
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("PropertyID").Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("PropertyName").Value
                    flxClient.TextMatrix(rRow, 3) = rstRec.Fields.Item("ClientID").Value
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

Private Sub LoadflxPropertyFilter()
    
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
    Label4.Visible = True
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 2800
   flxClient.ColWidth(3) = 1700
   flxClient.ColWidth(4) = 0
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   flxClient.ColAlignment(3) = vbLeftJustify

   
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""

   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   If txtClientList.Tag = "ALL" Then
        szSQL = "SELECT PropertyID, PropertyName,CLIENTID " & _
                "FROM Property " & _
                "ORDER BY PropertyID;"
   Else
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
   End If
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                    rRow = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "ALL"
                    flxClient.TextMatrix(rRow, 2) = "All Properties"
                     flxClient.RowHeight(rRow) = 280
                    flxClient.AddItem ""
                    rRow = 2
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("PropertyID").Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("PropertyName").Value
                     flxClient.TextMatrix(rRow, 3) = rstRec.Fields.Item("CLientid").Value
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
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Label4.Visible = False
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
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "ALL"
                    flxClient.TextMatrix(rRow, 2) = "All Clients"
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
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub LoadflxNC()
    
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

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   
   lblClientID.Caption = "N/C"
   lblClientName.Caption = "Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""

   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger,property where property.clientID=NominalLedger.clientID AND property.PropertyID='" & _
           txtProperty.Tag & "' order by NominalLedger.code;"
  

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                    rRow = 1
'                    flxClient.TextMatrix(rRow, 0) = ""
'                    flxClient.TextMatrix(rRow, 1) = "ALL"
'                    flxClient.TextMatrix(rRow, 2) = "All Properties"
'                     flxClient.RowHeight(rRow) = 280
'                    flxClient.AddItem ""
'                    rRow = 2
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("Code").Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("Name").Value
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
Private Sub LoadflxDemandcategory()
    
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

   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   
   lblClientID.Caption = "Code"
   lblClientName.Caption = "Category Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""

   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   szSQL = "SELECT Code, Value FROM SecondaryCode WHERE PrimaryCode = 'DCTG';"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                rRow = 1
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("Code").Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("Value").Value
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
Private Sub cmdClient_Click()
    txtClientList.text = ""
End Sub
'Private Sub loadflxclient()
'   Dim rRow As Integer
'   Dim szSQL As String
'
'   Dim adoConn As New ADODB.Connection
'   Dim rstRec As New ADODB.Recordset
'
'   flxClient.RowHeight(0) = 0
'   flxClient.Cols = 3
'   flxClient.ColWidth(0) = 1500
'   flxClient.ColWidth(1) = 3600
'   flxClient.ColWidth(2) = 0
'
'
'   txtSearchClientID.Width = 1530
'   txtSearchClientName.Visible = True
'   picClient.Width = 5295
'   cmdPicCLose.Left = 5010
'
'   flxClient.Clear
'   flxClient.Rows = 2
'   flxClient.ColAlignment(0) = vbLeftJustify
'   flxClient.ColAlignment(1) = vbLeftJustify
'   flxClient.ColAlignment(2) = vbLeftJustify
'
'   '~~~ Added by Anol Configuring width and position of labels and search boxes.
'   lblClientID.Caption = "Client ID"
'   lblClientName.Caption = "Client Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
'   txtSearchClientName.Left = 1620
'   txtSearchClientName.text = ""
'   txtSearchClientID.text = ""
'   txtSearchClientName.Width = 3240
'   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'
'   'lblJobName.Visible = False
'   adoConn.Open getConnectionString
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"
'
'   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'
'      If strCommandSource = "1" Or strCommandSource = "3" Then
'           flxClient.TextMatrix(1, 0) = "ALL"
'           flxClient.TextMatrix(1, 1) = "All Client"
'           flxClient.TextMatrix(1, 2) = ""
'           flxClient.AddItem ""
'           rRow = 2
'           While Not rstRec.EOF
'               flxClient.row = 1
'               flxClient.RowSel = 1
'               flxClient.ColSel = 1
'               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
'               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
'               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
'               flxClient.RowHeight(rRow) = 280
'               rstRec.MoveNext
'               If Not rstRec.EOF Then flxClient.AddItem ""
'               rRow = rRow + 1
'            Wend
'      End If
'   rstRec.Close
'   adoConn.Close
'   Set rstRec = Nothing
'   Set adoConn = Nothing
'
'End Sub

Private Sub cmdClientFilter_Click()
        fraCommands.Enabled = False
        fraDemandType.Enabled = False
        fraCommand.Enabled = False
        picClient.Left = 269.029
        picClient.Top = 155.299
        strCommandSource = "ClientFilter"
        LoadflxClient
        picClient.Visible = True
        txtSearchClientID.SetFocus
End Sub

Private Sub cmdCopy_Click()
    Dim a As Integer
   chkAllProperties.Value = 0
   chkAllDemand.Value = 0
   lblFrameIndex.Caption = "1"
''   If txtType.text = "" Then
''      MsgBox "You must enter Demand Type", vbOKOnly + vbCritical, "Demand Type"
'''      txtType.SetFocus
''      Exit Sub
''   End If
''   If txtPrefix.text = "" Then
''      MsgBox "You must enter Demand Prefix", vbOKOnly + vbCritical, "Demand Type"
'''      txtPrefix.SetFocus
''      Exit Sub
''   End If
''   If txtDemandTypeNCAmt.text = "" Then
''      MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'''      cmdDemandTypeNCAmt.SetFocus
''      Exit Sub
''   End If
'''   If cboDemandTypeNCvat.text = "" Then
'''       MsgBox "You must select a Nominal Account for VAT Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'''       cboDemandTypeNCvat.SetFocus
'''       Exit Sub
'''   End If
'''   If cboDemandTypeNCTotal.text = "" Then
'''      MsgBox "You must select a Nominal Account for Total Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'''      cboDemandTypeNCTotal.SetFocus
'''      Exit Sub
'''   End If
''   If txtDemandCategoryCode.text = "" Then
''      MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
'''      cboDemandTypeCategory.SetFocus
''      Exit Sub
''   End If
''   If cboDemandTypePayDates.text = "" And optPreset.Value Then
''      MsgBox "You must select a demand payment date.", vbOKOnly + vbCritical, "Payment Date"
'''      cboDemandTypePayDates.SetFocus
''      Exit Sub
''   End If
''   If txtBank.text = "" Then
'''      If cboBank.ListCount = 1 Then
'''         cboBank.ListIndex = 0
'''      Else
''         MsgBox "You must select a bank details.", vbOKOnly + vbCritical, "Bank Details"
'''         cboBank.SetFocus
'''         Exit Sub
'''      End If
''   End If
''   If txtDemandTemplate.text = "" Then
''      MsgBox "You must enter a demand template file name.", vbOKOnly + vbCritical, "Demand Template"
'''      cmdBrowsFile(0).SetFocus
''      Exit Sub
''   End If
''   If txtEmailTemplate.text = "" Then
''      MsgBox "You must enter a demand email template file name.", vbOKOnly + vbCritical, "Demand Email Template"
'''      cmdBrowsFile(1).SetFocus
''      Exit Sub
''   End If
''   If chkGroup.Value = 1 And txtGroup.text = "" Then
''      MsgBox "You must select a group id the demand type.", vbOKOnly + vbCritical, "Group"
'''      cboGroup.SetFocus
''      Exit Sub
''   End If

'   If txtStatementTemplate.text = "" Then
'      txtStatementTemplate.text = "InvDemandStatement.rpt"
'   Else
'      If Not FileExists(App.Path & "\CompanyReports\" & txtStatementTemplate.text) Then
'         If MsgBox("The statement template file does not exist. Do you wish to save the default template?", vbQuestion + vbYesNo, "Demand Statement") = vbNo Then
'            FocusControl txtStatementTemplate
'            Exit Sub
'         Else
'            If FileExists(App.Path & "\CompanyReports\InvDemandStatement.rpt") Then
'               txtStatementTemplate.text = "InvDemandStatement.rpt"
'            Else
'               ShowMsgInTaskBar "Please contact with PCM Support, demand statement template missing", "Y", "N"
'            End If
'         End If
'      End If
'   End If
'   cmdSaveNew.Enabled = True
'   cmdCancelNew.Enabled = True
'   cmdAdd.Enabled = False
'   cmdDelete.Enabled = False
'   cmdEdit.Enabled = False
'   cmdSave.Enabled = False
'   cmdCancel.Enabled = False
'   cmdExit(0).Enabled = True
    Frame1(1).Visible = True
    Frame1(1).Left = 2835
    Frame1(1).Top = 1080
    Frame1(1).ZOrder 0
    cmdBack.Enabled = False
    cmdNext.Enabled = True
'    cmdCopy.Enabled = False
    fraCommands.Enabled = False
    optCopyDemandType.SetFocus
End Sub

Private Sub cmdDelete_Click()
   Dim Response As Byte
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   If szDemandId = "" Then
       MsgBox "You must select a Rent Payable Type to delete!", vbOKOnly + vbCritical, "No Rent Payable Type selected"
       Exit Sub
   End If
   'fixed by anol 13 apr 2015
   'issue 0000554: Demand Types can be deleted even when being used in the system
   adoConn.Open getConnectionString
   'Mark 1
   If isPayableIDUsedinagrement(adoConn, szDemandId) = True Then
       MsgBox "You cannot delete this payable type because it is linked to an agreement. " & _
        "To delete it, please remove it from the agreement!", vbOKOnly + vbCritical, "Cannot delete Rent Payable Type"
       Exit Sub
   End If
   'End of addition
   Response = MsgBox("Are you sure you want to delete Rent Payable Type: " & szDemandType, vbYesNo + vbQuestion, "Delete")
   If Response = vbNo Then Exit Sub

   'delete record.
   
   
   szSQL = "SELECT * " & _
            "FROM PayableTypes WHERE ID = " & szDemandId & ""
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
   adoRst.Delete
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadFlxDemandTypes(adoConn, "")
   txtProperty.text = ""
   txtProperty.Tag = ""
   MsgBox "Selected Rent Payable Type has been deleted", vbInformation, "Deleted"
   adoConn.Close
   Set adoConn = Nothing

   Call EmptyBoxes
End Sub
Private Function isDemandIdUsedinTrans(adoConn As ADODB.Connection, szDemandId As String) As Boolean
    'Written by anol 13 Apr 2015
    'issue 0000554: Rent Payable Types can not be deleted when being used in the system
    'This function shall check demandtype before delete.
    'If demand type is used in transaction,this function won't let user to delete demand type.
    'Mark 1 When salia give you full spec impletement this method fully
    
    
'    Dim rsCheck As New ADODB.Recordset
'    rsCheck.Open "Select * from DemandsplitRecords where TypeofDemand=" & szDemandID & "", adoConn, adOpenStatic, adLockReadOnly
'    If Not rsCheck.EOF Then
'        isDemandIdUsedinTrans = True
'        Exit Function
'    End If
'    rsCheck.Close
'    rsCheck.Open "Select * from LRentCharges where BrDemandType=" & szDemandID & "", adoConn, adOpenStatic, adLockReadOnly
'    If Not rsCheck.EOF Then
'        isDemandIdUsedinTrans = True
'        Exit Function
'    End If
'    rsCheck.Close
'    rsCheck.Open "Select * from LInsuranceCharges where InsuranceDemandType=" & szDemandID & "", adoConn, adOpenStatic, adLockReadOnly
'    If Not rsCheck.EOF Then
'        isDemandIdUsedinTrans = True
'        Exit Function
'    End If
'    rsCheck.Close
'    rsCheck.Open "Select * from LServiceCharges where ScDemandType=" & szDemandID & "", adoConn, adOpenStatic, adLockReadOnly
'    If Not rsCheck.EOF Then
'        isDemandIdUsedinTrans = True
'        Exit Function
'    End If
'    rsCheck.Close
    
End Function
Private Function isPayableIDUsedinagrement(adoConn As ADODB.Connection, szDemandId As String) As Boolean
    'Written by anol 13 Apr 2015
    'issue 0000554: Rent Payable Types can not be deleted when being used in the system
    'This function shall check demandtype before delete.
    'If demand type is used in transaction,this function won't let user to delete demand type.
    'Mark 1 When salia give you full spec impletement this method fully
    
    
    Dim rsCheck As New ADODB.Recordset
    rsCheck.Open "Select * from tlbPayable P where P.PAYABLE_TYPE=" & szDemandId & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsCheck.EOF Then
        isPayableIDUsedinagrement = True
        Exit Function
    End If
    rsCheck.Close
End Function
Private Sub cmdDemandCategory_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "Pay charge Category"
    LoadflxDemandcategory
    fraDemandType.Enabled = False
    fraCommands.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdDemandTypeNCAmt_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "2"
    LoadflxNC
    fraDemandType.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdDemandTypePayDates_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "DemandTypePayDates"
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadPaymentDates(adoConn)
    adoConn.Close
    Set adoConn = Nothing
    fraDemandType.Enabled = False
    fraCommands.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdEdit_Click()
   If szDemandId = "" Then
      MsgBox "You must select a demand type to edit", vbOKOnly + vbCritical, "No demand type selected"
      Exit Sub
   End If

'   If Left(lblProperty.Caption, 3) = "All Properties" Then
'      MsgBox "Please select property in the previous screen.", vbCritical + vbOKOnly, "Demand Type"
'      Exit Sub
'   End If

'   lblClient.Caption = frmDCTypesPre.cboClientList.Column(1)
'   lblProperty.Caption = frmDCTypesPre.cboPropertyList.Column(1)

'   Call EnableBoxes
   fraDemandType.Enabled = True
   flxDemandTypes.Enabled = False

   cmdEdit.Enabled = False
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   cmdSaveNew.Enabled = True 'added by anol 20171004
   cmdCancelNew.Enabled = True
   cmdPropertyFilter.Enabled = False
   cmdClientFilter.Enabled = False
'   cmdSave.Enabled = True
   EditMode = True
   
   chkGroup.Enabled = True
   txtType.SetFocus
End Sub

Private Function IsBankEdit() As Boolean
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String ', szaDemandID() As String

   adoConn.Open getConnectionString

   szSQL = "SELECT DR.DemandId " & _
           "FROM DemandRecords as DR, DemandSplitRecords as DSR " & _
           "WHERE DSR.TypeOfDemand = " & szDemandId & " AND " & _
               "(DR.IsPrinted = FALSE OR " & _
               "DR.UPDATE_SAGE = FALSE) AND " & _
               "DR.DemandHistory = FALSE AND " & _
               "DR.DemandId = DSR.DemandId;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      IsBankEdit = True
   Else
      IsBankEdit = False
   End If

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Function

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
End Sub
'
'Private Sub cmdGSCancel_Click(Index As Integer)
'   If MsgBox("Are you sure to cancel the changes?", vbQuestion + vbYesNo, "Cancel Saving") = vbNo Then Exit Sub
'   ButtonMode DefaultMode, Index
'End Sub

Private Sub cmdFinish_Click()
        Dim a As Integer
        Dim i As Integer
        Dim iCount As Integer
        Dim szSQL As String
        Dim adoConn As New ADODB.Connection
        Dim adoRst As New ADODB.Recordset
        Dim rsSourceDemand As New ADODB.Recordset
        Dim j As Integer
        iCount = 0
        Dim strTemp As String
        For i = 1 To flxProperty1.Rows - 1
            If flxProperty1.TextMatrix(i, 0) = "X" Then
                      iCount = iCount + 1
            End If
        Next i
            If iCount = 0 Then
                MsgBox "Please select at least one property for copying", vbInformation, "Warning!"
                Exit Sub
            End If
            iCount = 0
        If optCopyDemandType.Value Then
                strTemp = "type"
        Else
                strTemp = "template"
        End If
     If MsgBox("Are you sure you wish to copy the demand " & strTemp & " selected to the selected properties?", vbQuestion + vbYesNo, IIf(optCopyDemandTemplate, optCopyDemandTemplate.Caption, optCopyDemandType.Caption)) = vbNo Then Exit Sub
    
       
       adoConn.Open getConnectionString
       If optCopyDemandType.Value Then 'copy demand type for each of properties
            If Len(copyfromDemandTypeID) = 0 Then
                  adoConn.Close
                  Exit Sub
            End If

            Dim tmpDemand
            tmpDemand = Split(copyfromDemandTypeID, " OR ")
            For j = 0 To UBound(tmpDemand)
                    For i = 1 To flxProperty1.Rows - 1
                      If flxProperty1.TextMatrix(i, 0) = "X" Then
                           iCount = iCount + 1
                           szSQL = "SELECT MAX(ID) FROM PayableTypes"
                           adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                           a = CInt(IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value))
                           adoRst.Close
                           
                           szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, InvCrd, " & _
                                       "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
                                       "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
                                       "CategoryCode, PaymentDates, GroupID, DemandReportName, " & _
                                       "Spare1, PropertyID, EmailInvoiceTemplate, StatementTemplate " & _
                                    "FROM PayableTypes"
                           adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                           szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, InvCrd, " & _
                                       "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
                                       "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
                                       "CategoryCode, PaymentDates, GroupID, DemandReportName, " & _
                                       "Spare1, PropertyID, EmailInvoiceTemplate, StatementTemplate " & _
                                    "FROM PayableTypes where ID=" & tmpDemand(j) & ""
                           rsSourceDemand.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                           If rsSourceDemand.EOF Then
                                MsgBox "Demand type for this ID not found " & tmpDemand(j), vbInformation
                                rsSourceDemand.Close
                                fraCommands.Enabled = True
                                adoConn.Close
                                Exit Sub
                           End If
                           adoRst.AddNew
                           adoRst!ID = a + 1
                         
                           
                           Dim rsCheck As New ADODB.Recordset
                           szSQL = "SELECT T.CAName, S.Value, T.Code AS NCode, T.Name AS NName, T.ClientID, T.CAFixed AS Fixed," _
                        & "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder FROM NominalLedger AS T," _
                        & "SecondaryCode AS S,Property  WHERE Property.ClientID=T.ClientID AND T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND Property.PropertyID = '" & flxProperty1.TextMatrix(i, 1) & "' ORDER By t.CADisOrder"
        '                   rsCheck.Close
                           rsCheck.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                           Dim x1 As String
                           Dim x2 As String
                           Dim x3 As String
                           Dim x4 As String
                          
                           
                           While Not rsCheck.EOF
                                If rsCheck("CAName").Value = "Sales Ledger Control" Then
                                    x1 = rsCheck("NCODE").Value
                                    x2 = rsCheck("NName").Value
                                End If
                                If rsCheck("CAName").Value = "Output VAT" Then
                                    x3 = rsCheck("NCODE").Value
                                    x4 = rsCheck("NName").Value
                                End If
                                rsCheck.MoveNext
                           Wend
                           rsCheck.Close
                           If x1 = "" Then
                                MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & _
                                flxProperty1.TextMatrix(i, 3) & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
                                adoConn.Close
                                Command1_Click
                                Exit Sub
                           End If
                        '  ---------------------------------------------------
                        '  'InvCrd' field is required in the table. This field is no longer in use.
                        '  Thats why I am saving a charecter. I have changed the field as not required (03/09/2009).
                           adoRst!InvCrd = rsSourceDemand!InvCrd
                           adoRst!Type = rsSourceDemand!Type
                           adoRst!prefix = rsSourceDemand!prefix
                           adoRst!NominalCodeforAmount = rsSourceDemand!NominalCodeforAmount
                           adoRst!NominalNameforAmount = rsSourceDemand!NominalNameforAmount
                           adoRst!NominalCodeForVAT = x3 ''cboDemandTypeNCvat.text
                           adoRst!NominalNameforVAT = x4 '' txt2.text
                           adoRst!NominalCodeForTotal = x1 'cboDemandTypeNCTotal.text & ""
                           adoRst!NominalNameforTotal = x2 ''txt3.text
                        '   adoRst!prefix = "NULL"
                           adoRst!CategoryCode = rsSourceDemand!CategoryCode
'                           If optAuto.Value Then
'                              adoRST!PaymentDates = CByte(255)
'                           Else
'                              adoRST!PaymentDates = CByte(cboDemandTypePayDates.Value)
'                           End If
                           adoRst!PaymentDates = rsSourceDemand!PaymentDates
'                           If chkGroup.Value = 1 Then adoRST!GroupID = txtGroup.text
                           adoRst!GroupID = rsSourceDemand!GroupID
                           adoRst!DemandReportName = rsSourceDemand!DemandReportName
                           adoRst!spare1 = RetrnBankID(flxProperty1.TextMatrix(i, 1), adoConn) 'cboBank.Value ' Here you need to bring the default bank account for that property
                           adoRst!propertyID = flxProperty1.TextMatrix(i, 1)
                           adoRst!EmailInvoiceTemplate = rsSourceDemand!EmailInvoiceTemplate
                           adoRst!StatementTemplate = rsSourceDemand!StatementTemplate
                        
                           adoRst.Update
                           rsSourceDemand.Close
                           adoRst.Close
                           Set adoRst = Nothing
                     End If
                    Next i
                 Next j
         MsgBox "Demand type has been created"
   Else 'copy Demand Template
        iCount = 0
         For i = 1 To flxDemandTypeList2.Rows - 1
            If flxDemandTypeList2.TextMatrix(i, 0) = "X" Then
                      iCount = iCount + 1
            End If
        Next i
            If iCount = 0 Then
                MsgBox "Please select at least one demand type for copying", vbInformation, "Warning!"
                fraCommands.Enabled = True
                Exit Sub
            End If
            
         For i = 1 To flxDemandTypeList.Rows - 1
              If flxDemandTypeList.TextMatrix(i, 0) = "X" Then
                     szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, InvCrd, " & _
                                       "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
                                       "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
                                       "CategoryCode, PaymentDates, GroupID, DemandReportName, " & _
                                       "Spare1, PropertyID, EmailInvoiceTemplate, StatementTemplate " & _
                                    "FROM PayableTypes where ID=" & flxDemandTypeList.TextMatrix(i, 1) & ""
                    rsSourceDemand.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                    
                  Exit For
              End If
         Next i
         If rsSourceDemand.EOF Then
            MsgBox "Please select a demand type copy from"
            adoConn.Close
            fraCommands.Enabled = True
            Exit Sub
         End If
         For i = 1 To flxDemandTypeList2.Rows - 1
              If flxDemandTypeList2.TextMatrix(i, 0) = "X" Then
                    adoConn.Execute "Update PayableTypes set DemandReportName = '" & rsSourceDemand("DemandReportName").Value & "'," & _
                    "EmailInvoiceTemplate= '" & rsSourceDemand("EmailInvoiceTemplate").Value & "',StatementTemplate='" & rsSourceDemand("StatementTemplate").Value & "' " & _
                    " where ID =" & Val(flxDemandTypeList2.TextMatrix(i, 1)) & ""
               End If
         Next i
         MsgBox "Demand type has been updated"
'         If iCount = 0 Then
'             MsgBox "Please select at least one demand type to copy template"
'         Else
'             MsgBox "Demand type has been updated"
'         End If
        
   End If
   Call LoadFlxDemandTypes(adoConn, "")
   adoConn.Close
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    Frame1(3).Visible = False
    Frame1(4).Visible = False
    Frame1(5).Visible = False
   fraCommands.Enabled = True
End Sub
Private Function RetrnBankID(propertyID As String, adoConn As ADODB.Connection) As String
    Dim rsBankID As New ADODB.Recordset
    rsBankID.Open "Select MY_ID from tlbClientBanks A,Property B where A.CLIENT_ID=B.ClientID and B.PropertyID='" & propertyID & "' order by DEFAULT_AC ASC", adoConn, adOpenKeyset, adLockOptimistic
    If Not rsBankID.EOF Then
        RetrnBankID = rsBankID("MY_ID").Value
    End If
    rsBankID.Close
End Function

Private Sub cmdGroup_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "4"
    Loadflxgroup
    fraDemandType.Enabled = False
    fraCommands.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdNext_Click()
    Dim iDemandCount As Integer
    If optCopyDemandType.Value Then
        MAX_NUMBER_FRAME_INDEX = 4
        chkAllDemand.Visible = True
         Label2.Caption = "Demand Type Copy Wizard"
    Else
        MAX_NUMBER_FRAME_INDEX = 5
        chkAllDemand.Visible = False
        Label2.Caption = "Demand Type Template Copy Wizard"
    End If
    
   If Val(lblFrameIndex.Caption) = 4 Then
       If flxProperty1.TextMatrix(flxProperty1.row, 0) = "" Or flxProperty1.TextMatrix(flxProperty1.row, 1) = "" Or IsNull(flxProperty1.TextMatrix(flxProperty1.row, 1)) Then
            MsgBox "Please select a property from the list.", vbCritical + vbOKOnly, "Select demand type"
'            flxDemandTypeList2.SetFocus
            Exit Sub
       End If
        'Collect the demand types copy from
'       For iDemandCount = 1 To flxDemandTypeList.Rows - 1
'           If flxDemandTypeList.TextMatrix(iDemandCount, 0) = "X" Then
'              copyfromDemandTypeID = copyfromDemandTypeID & flxDemandTypeList.TextMatrix(iDemandCount, 1)
'              copyfromDemandTypeID = copyfromDemandTypeID & " OR "
'           End If
'       Next iDemandCount
'      If Right(copyfromDemandTypeID, 4) = " OR " Then
'         copyfromDemandTypeID = Left(copyfromDemandTypeID, Len(copyfromDemandTypeID) - 4)
'      End If
       
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1 'keeping the ongoing screen ID
       Frame1(Val(lblFrameIndex.Caption)).Top = 2000
       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(1).Left + 80
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
       'load the PayableTypes
       LoadDemandType2
    End If
    If Val(lblFrameIndex.Caption) = 3 Then
       If flxDemandTypeList.TextMatrix(flxDemandTypeList.row, 0) = "" Or flxDemandTypeList.TextMatrix(flxDemandTypeList.row, 1) = "" Or IsNull(flxDemandTypeList.TextMatrix(flxDemandTypeList.row, 1)) Then
            MsgBox "Please select a demand type from the list.", vbCritical + vbOKOnly, "Select demand type"
            flxProperty.SetFocus
            Exit Sub
       End If
        'Collect the demand types copy from
       For iDemandCount = 1 To flxDemandTypeList.Rows - 1
           If flxDemandTypeList.TextMatrix(iDemandCount, 0) = "X" Then
              copyfromDemandTypeID = copyfromDemandTypeID & flxDemandTypeList.TextMatrix(iDemandCount, 1)
              copyfromDemandTypeID = copyfromDemandTypeID & " OR "
           End If
       Next iDemandCount
      If Right(copyfromDemandTypeID, 4) = " OR " Then
         copyfromDemandTypeID = Left(copyfromDemandTypeID, Len(copyfromDemandTypeID) - 4)
      End If
       
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1 'keeping the ongoing screen ID
       Frame1(Val(lblFrameIndex.Caption)).Top = 2000
       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(1).Left + 80
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
       'load the PayableTypes
       LoadProperties2
    End If
    If Val(lblFrameIndex.Caption) = 2 Then
      'clearing prevoiusly holded demand ID
       copyfromDemandTypeID = ""
       
       If flxProperty.TextMatrix(flxProperty.row, 0) = "" Or flxProperty.TextMatrix(flxProperty.row, 1) = "" Or IsNull(flxProperty.TextMatrix(flxProperty.row, 1)) Then
            MsgBox "Please select a property from the list.", vbCritical + vbOKOnly, "Select Property"
            flxProperty.SetFocus
            Exit Sub
       End If
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1 'keeping the ongoing screen ID
       Frame1(Val(lblFrameIndex.Caption)).Top = 2000
       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(1).Left + 80
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
       'load the PayableTypes
       LoadDemandType
    End If
    If Val(lblFrameIndex.Caption) = 1 Then
    
       lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1 'keeping the ongoing screen ID
       Frame1(Val(lblFrameIndex.Caption)).Top = 2000
       Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left + 80
       Frame1(Val(lblFrameIndex.Caption)).Visible = True
       Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
      'load the properties
      LoadProperties
    End If
    cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
    cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
    cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
End Sub
Private Sub LoadDemandType()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szHeader As String
   flxDemandTypeList.Clear
   Dim iProp As Integer
   Dim szDes As String
'   connect to database
   adoConn.Open getConnectionString
   For iProp = 1 To flxProperty.Rows - 1
      If flxProperty.TextMatrix(iProp, 0) = "X" Then
         szDes = szDes & "PayableTypes.PropertyID = '" & flxProperty.TextMatrix(iProp, 1) & "'"
         szDes = szDes & " OR "
      End If
   Next iProp
    'Fixed by anol 20170326
   If Right(szDes, 4) = " OR " Then
        szDes = Left(szDes, Len(szDes) - 4)
   End If
    
   szSQL = "SELECT ID, TYPE FROM PayableTypes WHERE (" & szDes & ");"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
   szHeader$ = "|< ID|< TYPE"
   flxDemandTypeList.FormatString = szHeader$
   flxDemandTypeList.ColWidth(0) = 200
   flxDemandTypeList.ColWidth(1) = 1500
   flxDemandTypeList.ColWidth(2) = 4000
   flxDemandTypeList.Rows = 2
   flxDemandTypeList.Cols = 3
   r = 1
   While Not adoRst.EOF
      flxDemandTypeList.TextMatrix(r, 1) = adoRst.Fields.Item("ID").Value
      flxDemandTypeList.TextMatrix(r, 2) = adoRst.Fields.Item("TYPE").Value
      flxDemandTypeList.AddItem ""
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadDemandType2()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szHeader As String
   flxDemandTypeList2.Clear
   Dim iProp As Integer
   Dim szDes As String
'   connect to database
   adoConn.Open getConnectionString
   For iProp = 1 To flxProperty.Rows - 1
      If flxProperty1.TextMatrix(iProp, 0) = "X" Then
         szDes = szDes & "PayableTypes.PropertyID = '" & flxProperty1.TextMatrix(iProp, 1) & "'"
         szDes = szDes & " OR "
      End If
   Next iProp
    'Fixed by anol 20170326
   If Right(szDes, 4) = " OR " Then
        szDes = Left(szDes, Len(szDes) - 4)
   End If
    
   szSQL = "SELECT ID, TYPE FROM PayableTypes WHERE (" & szDes & ");"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
   szHeader$ = "|< ID|< TYPE"
   flxDemandTypeList2.FormatString = szHeader$
   flxDemandTypeList2.ColWidth(0) = 200
   flxDemandTypeList2.ColWidth(1) = 1500
   flxDemandTypeList2.ColWidth(2) = 4000
   flxDemandTypeList2.Rows = 2
   flxDemandTypeList2.Cols = 3
   r = 1
   While Not adoRst.EOF
      flxDemandTypeList2.TextMatrix(r, 1) = adoRst.Fields.Item("ID").Value
      flxDemandTypeList2.TextMatrix(r, 2) = adoRst.Fields.Item("TYPE").Value
      flxDemandTypeList2.AddItem ""
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadProperties()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szHeader As String
   flxProperty.Clear

'   connect to database
   adoConn.Open getConnectionString

'   szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
'           "FROM PROPERTY " & _
'           "WHERE PROPERTYID <> '" & flxDemandTypes.TextMatrix(flxDemandTypes.row, 5) & "';"
   szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
           "FROM PROPERTY ;"
           
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
    szHeader$ = "|< PROPERTY ID|< PROPERTY NAME"
    flxProperty.FormatString = szHeader$
    flxProperty.ColWidth(0) = 200
    flxProperty.ColWidth(1) = 1500
    flxProperty.ColWidth(2) = 4000
    flxProperty.Rows = 2
    flxProperty.Cols = 3
    r = 1
   While Not adoRst.EOF
      flxProperty.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
      flxProperty.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      flxProperty.AddItem ""
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadProperties2()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szHeader As String
   flxProperty1.Clear

'   connect to database
   adoConn.Open getConnectionString

'   szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
'           "FROM PROPERTY " & _
'           "WHERE PROPERTYID <> '" & flxProperty.TextMatrix(flxProperty.row, 5) & "';"
   szSQL = "SELECT PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM PROPERTY ;"
           
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
    szHeader$ = "|< PROPERTY ID|< PROPERTY NAME"
    flxProperty1.FormatString = szHeader$
    flxProperty1.ColWidth(0) = 200
    flxProperty1.ColWidth(1) = 1500
    flxProperty1.ColWidth(2) = 4000
    flxProperty1.ColWidth(3) = 0
    flxProperty1.Rows = 2
    flxProperty1.Cols = 4
    r = 1
   While Not adoRst.EOF
      flxProperty1.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
      flxProperty1.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      flxProperty1.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value
      
      flxProperty1.AddItem ""
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    fraCommands.Enabled = True
    fraDemandType.Enabled = True
    fraCommand.Enabled = True
    fraDemandType.Enabled = True
    fraCommands.Enabled = True
End Sub

Private Sub cmdProperties_Click()
   txtPropertyList.text = ""
End Sub

Private Sub cmdproperty_Click()
    picClient.Left = 2880
    picClient.Top = 4230
    strCommandSource = "SelectedProperty"
    LoadflxProperty
    fraDemandType.Enabled = False
    fraCommands.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus

End Sub

'Private Sub LoadNominalCode()
'    flxSupplier(0).Cols = 3
'   flxSupplier(0).ColWidth(0) = 1500
'   flxSupplier(0).ColWidth(1) = 2700
'    flxSupplier(0).ColWidth(2) = 0
'   flxSupplier(0).ColAlignment = vbLeftJustify
'
'    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
'   lblSearch0(0).Width = 1400
'   lblSearch0(0).Left = 50
'   lblSearch1(0).Width = 2600
'   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
'
'   lblSearch0(0).Caption = "N/C"
'   lblSearch1(0).Caption = "Name"
'   lblSearch2(0).Visible = False
'
'   flxSupplier(0).RowHeight(0) = 0
'
'' Error Handler
'   On Error GoTo Error_Handler
'
'   Dim adoConn As ADODB.Connection
'   Dim rRow As Integer, iRec As Integer
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   Set adoConn = New ADODB.Connection
'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim iRows As Integer
'
'   flxSupplier(0).Rows = 2
'   iRows = 1
'   While Not adoRst.EOF
'      flxSupplier(0).TextMatrix(iRows, 0) = adoRst.Fields.Item("Code").Value
'      flxSupplier(0).TextMatrix(iRows, 1) = adoRst.Fields.Item("Name").Value
'      If Not adoRst.EOF Then flxSupplier(0).AddItem ""
'      iRows = iRows + 1
'      adoRst.MoveNext
'   Wend
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'
'   Exit Sub
'
'' Error Handling Code
'Error_Handler:
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'End Sub

Private Sub cmdSave_Click()
'   On Error Resume Next
   Dim i As Integer
   If txtPrefix.text = "" Then
      MsgBox "You must enter Demand Prefix", vbOKOnly + vbCritical, "Rent Payable Type"
      txtPrefix.SetFocus
      Exit Sub
   End If
'   If txtDemandTypeNCAmt.text = "" Then
'      MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'      cmdDemandTypeNCAmt.SetFocus
'      Exit Sub
'   End If
'   If txtDemandTypeNCAmt.text <> "" And txt1.text = "" Then
'      MsgBox "You must select a correct Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'      txtDemandTypeNCAmt.text = ""
'      cmdDemandTypeNCAmt.SetFocus
'      Exit Sub
'   End If

'   If txtDemandCategoryCode.text = "" Then
'      MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
'      Exit Sub
'   End If

   If txtDemandTypePayDates.text = "" And optPreset.Value Then
      MsgBox "You must select a demand payment date.", vbOKOnly + vbCritical, "Payment Date"
      FocusControl cmdDemandTypePayDates
      Exit Sub
   End If

'   If txtBank.text = "" Then
'         MsgBox "You must select a bank details.", vbOKOnly + vbCritical, "Bank Details"
'         Exit Sub
'   End If

   If txtProperty.text = "" Or txtProperty.Tag = "" Then
      MsgBox "You must select a property.", vbOKOnly + vbCritical, "Demand Type property"
      cmdproperty.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString
   szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, " & _
               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
               "CategoryCode, PaymentDates, GroupID, DemandReportName, Spare1, " & _
               "PropertyID, EmailInvoiceTemplate, StatementTemplate,Consolidated " & _
            "FROM PayableTypes WHERE ID =" & szDemandId & ""
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   adoRst!Type = txtType.text
   adoRst!prefix = txtPrefix.text
'   adoRst!NominalCodeforAmount = txtDemandTypeNCAmt.text
'   adoRst!NominalNameforAmount = txt1.text
'   adoRst!CategoryCode = txtDemandCategoryCode.text
   If optAuto.Value Then
      adoRst!PaymentDates = CByte(255)
   Else
      adoRst!PaymentDates = Val(txtDemandTypePayDates.Tag)
   End If
   adoRst!GroupID = Val(txtGroup.text)
   'adoRst!DemandReportName = txtDemandTemplate.text
   'adoRst!spare1 = txtBank.Tag
   
   'adoRst!Consolidated = chkConsolidated.Value

   adoRst!propertyID = txtProperty.Tag

   'adoRst!EmailInvoiceTemplate = txtEmailTemplate.text
   'adoRst!StatementTemplate = txtStatementTemplate.text

   adoRst.Update
   adoRst.Close
   Set adoRst = Nothing


    szSQL = "SELECT D.ID, D.Type, D.PropertyID, " & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
           "FROM (PayableTypes AS D LEFT JOIN Property AS P ON " & _
                 "D.PropertyID = P.PropertyID) LEFT JOIN Client AS C ON P.ClientID = C.ClientID where ID =" & szDemandId & " " & _
           "ORDER BY ClientName, PropertyName, D.Type, D.ID;"
           
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   With adoRst.Fields
      For i = 0 To adoRst.RecordCount - 1
         
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) = IIf(IsNull(.Item("ClientID")), "", .Item("ClientID"))
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 2) = IIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 3) = IIf(IsNull(.Item("PropertyID")), "", .Item("PropertyID"))
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 4) = IIf(IsNull(.Item("PropertyName")), "", .Item("PropertyName"))
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 5) = IIf(IsNull(.Item("Type")), "", .Item("Type"))
         flxDemandTypes.TextMatrix(flxDemandTypes.row, 6) = IIf(IsNull(.Item("ID")), "", .Item("ID"))
         adoRst.MoveNext
         
      Next i
   End With
   adoRst.Close
   
   adoConn.Close
   Set adoConn = Nothing

    MsgBox "Edited charge type details have been saved.", vbInformation, "Saved"


   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSaveNew.Enabled = False
   cmdPropertyFilter.Enabled = True
   cmdClientFilter.Enabled = True
   
   iSelRow = 0
   szExistingProperty = ""
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
   flxClient.ColWidth(0) = 0
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
            flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL Properties"
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
   
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdPropertyFilter_Click()
        fraCommands.Enabled = False
        fraDemandType.Enabled = False
        fraCommand.Enabled = False
        picClient.Left = 4269.029
        picClient.Top = 155.299
        strCommandSource = "PropertyFilter"
        Call LoadflxPropertyFilter
        picClient.Visible = True
        txtSearchClientID.SetFocus
End Sub

Private Sub cmdSaveNew_Click()
  
'   If EditMode Then
'        cmdSave_Click
'        Exit Sub
'   End If
   Dim a As Integer
   
   If txtType.text = "" Then
      MsgBox "You must enter a Rent Payable Type", vbOKOnly + vbCritical, "Rent Payable Type"
      txtType.SetFocus
      Exit Sub
   End If
   If txtPrefix.text = "" Then
      MsgBox "You must enter a Pay charge Prefix", vbOKOnly + vbCritical, "Rent Payable Type"
      txtPrefix.SetFocus
      Exit Sub
   End If
'   If txtDemandTypeNCAmt.text = "" Then
'      MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
'      cmdDemandTypeNCAmt.SetFocus
'      Exit Sub
'   End If

'   If txtDemandCategoryCode.text = "" Then
'      MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
'      Exit Sub
'   End If
   If txtDemandTypePayDates.text = "" And optPreset.Value Then
      MsgBox "You must select a demand payment date.", vbOKOnly + vbCritical, "Payment Date"
      FocusControl cmdDemandTypePayDates
      Exit Sub
   End If

   If chkGroup.Value = 1 And txtGroup.text = "" Then
      MsgBox "You must select a group id the demand type.", vbOKOnly + vbCritical, "Group"
      cmdGroup.SetFocus
      Exit Sub
   End If
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   adoConn.Open getConnectionString
   If EditMode = False Then
           szSQL = "SELECT MAX(ID) FROM PayableTypes"
           adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           a = 0
           a = CInt(IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value))
           adoRst.Close
           szDemandId = a + 1
   End If



    szSQL = "SELECT ID, PayType, PayIC, PaySagePrefix, PayNCAmt, PayNNAmt, " & _
                  "PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, TransactionType, " & _
                  "CategoryCode, PaymentDates,  GroupID,PayNCAmt,PayNNAmt,PropertyID,clientID,isUseControlAccount,CreatedBy,CreatedDate,LastModifiedBy,LastModifiedDate " & _
                  "FROM PayableTypes C where ID=" & szDemandId & ""
                  
    adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
    If EditMode = False Then
        adoRst.AddNew
        adoRst!ID = szDemandId
        adoRst!CreatedBy = User
        adoRst!CreatedDate = Now
    End If
    'Modified by Anol 10 Sep 2014
   
   Dim rsCheck As New ADODB.Recordset
   szSQL = "SELECT T.CAName, S.Value, T.Code AS NCode, T.Name AS NName, T.ClientID, T.CAFixed AS Fixed," _
            & "IIF(T.CAPosting, 'YES', 'NO') AS P, T.CAType AS Type, T.CADisOrder FROM NominalLedger AS T," _
            & "SecondaryCode AS S,Property  WHERE Property.ClientID=T.ClientID AND T.CAType = S.Code AND S.PrimaryCode = 'CAT' AND Property.PropertyID = '" & txtProperty.Tag & "' ORDER By t.CADisOrder"
   rsCheck.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   Dim x1 As String
   Dim x2 As String
   Dim x3 As String
   Dim x4 As String
   
   While Not rsCheck.EOF
        If rsCheck("CAName").Value = "Sales Ledger Control" Then
            x1 = rsCheck("NCODE").Value
            x2 = rsCheck("NName").Value
        End If
        If rsCheck("CAName").Value = "Output VAT" Then
             x3 = rsCheck("NCODE").Value
            x4 = rsCheck("NName").Value
        End If
        rsCheck.MoveNext
   Wend
   rsCheck.Close
   Set rsCheck = Nothing
'  ---------------------------------------------------
'  'InvCrd' field is required in the table. This field is no longer in use.
'  Thats why I am saving a charecter. I have changed the field as not required (03/09/2009).
   adoRst!PayIC = "X"
   adoRst!PayType = txtType.text
   adoRst!PaySagePrefix = txtPrefix.text
'   adoRst!PayNCAmt = txtDemandTypeNCAmt.text
'   adoRst!PayNNAmt = txt1.text
   adoRst!PayNCVat = x3 ''cboDemandTypeNCvat.text
   adoRst!PayNNVat = x4 '' txt2.text
   adoRst!PayNCTotal = x1 'cboDemandTypeNCTotal.text & ""
   adoRst!PayNNTotal = x2 ''txt3.text
'   adoRst!prefix = "NULL"
'   adoRst!CategoryCode = txtDemandCategoryCode.text
   If optAuto.Value Then
      adoRst!PaymentDates = CByte(255)
   Else
      adoRst!PaymentDates = Val(txtDemandTypePayDates.Tag)
   End If
   If chkGroup.Value = 1 Then adoRst!GroupID = txtGroup.text
   adoRst!propertyID = txtProperty.Tag
   adoRst!clientID = szSelectedClient
   adoRst!isUseControlAccount = chkUseControlAccount.Value
   adoRst!LastModifiedBy = User
   adoRst!LastModifiedDate = Now
   
   adoRst.Update
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadFlxDemandTypes(adoConn, "")
'   flxDemandTypes.TopRow = intTopRow
'   flxDemandTypes.row = intTopRow
   adoConn.Close
   Set adoConn = Nothing

   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False

   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True
   fraDemandType.Enabled = False
   FocusControl cmdAdd
   cmdPropertyFilter.Enabled = True
   cmdClientFilter.Enabled = True
   If EditMode = False Then
        MsgBox "Your New Rent Payable Type details have been saved.", vbInformation, "Saved"
   Else
        MsgBox "Edited charge type details have been saved.", vbInformation, "Saved"
   End If
   Exit Sub
   
End Sub

Private Sub Command1_Click()
     Frame1(1).Visible = False
     Frame1(2).Visible = False
     Frame1(3).Visible = False
     Frame1(4).Visible = False
     Frame1(5).Visible = False
'     cmdCopy.Enabled = True
     fraCommands.Enabled = True
End Sub

Private Sub Command2_Click()
     Frame1(1).Visible = False
     Frame1(2).Visible = False
     Frame1(3).Visible = False
     Frame1(4).Visible = False
     Frame1(5).Visible = False
'     cmdCopy.Enabled = True
     fraCommands.Enabled = True
End Sub

Private Sub flxClient_Click()
        Dim adoConn As New ADODB.Connection
        fraDemandType.Enabled = True
        fraCommands.Enabled = True
        fraCommands.Enabled = True
        
        fraCommand.Enabled = True
        If strCommandSource = "SelectedProperty" Then
                txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                szSelectedClient = flxClient.TextMatrix(flxClient.row, 3)
                 'Written by anol 20170601
                setStatementonPropertyClick
                'Written by anol 20160926
                
                If checkExistingLeaseBeforeChangeProperty = True Then Exit Sub ' you cannot change the property when existing lease is assocaited with demandtype
                cboProperty_Change
'                txtDemandTypeNCAmt.text = ""
'                txt1.text = ""
        '        LoadNCinCombo1
                cboProperty_GotFocus
                
                If txtType.Enabled Then txtType.SetFocus
        ElseIf strCommandSource = "2" Then      'Nominal Code for Amount
''                txtDemandTypeNCAmt.text = flxClient.TextMatrix(flxClient.row, 1)
''                txt1.text = flxClient.TextMatrix(flxClient.row, 2)
''                FocusControl cmdDemandCategory
'                'this check is nuts becuase you are adding always 1 item by deafault
'                If cboDemandTypePayDates.ListCount > 0 Then
'                    cboDemandTypePayDates.ListIndex = 0
'                End If
                 'Here I am going to set first data anyway
'        ElseIf strCommandSource = "3" Then
'                txtBank.text = flxClient.TextMatrix(flxClient.row, 1)
'                txtBank.Tag = flxClient.TextMatrix(flxClient.row, 2)
                
'                FocusControl cmdGroup
'                 FocusControl cmdBrowsFile(0)
'                  FocusControl cmdDemandTypePayDates
'        ElseIf strCommandSource = "4" Then
'                txtGroup.text = flxClient.TextMatrix(flxClient.row, 1)
'                txtGroup.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                FocusControl cmdBank
               
        ElseIf strCommandSource = "Pay charge Category" Then
''                txtDemandCategoryCode.text = flxClient.TextMatrix(flxClient.row, 1)
''                txtDemandCategoryName.text = flxClient.TextMatrix(flxClient.row, 2)
''                FocusControl cmdDemandTypePayDates
                
        ElseIf strCommandSource = "ClientFilter" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                FocusControl cmdPropertyFilter
                fraDemandType.Enabled = False
                '                If txtClientList.Tag = "ALL" Then when you select a client property selection must reset
                txtPropertyList.Tag = "ALL"
                txtPropertyList.text = "ALL Properties"
                
                adoConn.Open getConnectionString
                Call LoadFlxDemandTypes(adoConn, "")
                adoConn.Close
                Set adoConn = Nothing
                '                End If
        ElseIf strCommandSource = "PropertyFilter" Then
                txtPropertyList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtPropertyList.text = flxClient.TextMatrix(flxClient.row, 2)
                FocusControl cmdPropertyFilter
                
                adoConn.Open getConnectionString
                Call LoadFlxDemandTypes(adoConn, "")
                adoConn.Close
                Set adoConn = Nothing
                fraDemandType.Enabled = False
        ElseIf strCommandSource = "DemandTypePayDates" Then
                txtDemandTypePayDates.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtDemandTypePayDates.text = flxClient.TextMatrix(flxClient.row, 2)
                FocusControl cmdSaveNew
        End If
        picClient.Visible = False
        
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub flxDemandTypeList_Click()
   Dim i As Integer
   If flxDemandTypeList.TextMatrix(flxDemandTypeList.row, 1) = "" Then Exit Sub
   If optCopyDemandType.Value Then
         i = SelectFlxGridRow(0, flxDemandTypeList, flxDemandTypeList.row)
   Else
         Call SelectOnly1RowFlxGrid(flxDemandTypeList, flxDemandTypeList.row, 0)
   End If
   
End Sub

Private Sub flxDemandTypeList2_Click()
    Dim i As Integer
    If flxDemandTypeList2.TextMatrix(flxDemandTypeList2.row, 1) = "" Then Exit Sub
    i = SelectFlxGridRow(0, flxDemandTypeList2, flxDemandTypeList2.row)
End Sub

Private Sub flxDemandTypes_RowColChange()
   With flxDemandTypes
      szDemandId = .TextMatrix(.row, 6)
      szDemandType = .TextMatrix(.row, 8)
      iSelRow = .row
      txtProperty.text = .TextMatrix(.row, 13)
      txtProperty.Tag = .TextMatrix(.row, 14)
      szExistingProperty = txtProperty.Tag
      Me.Caption = "Rent Payable Types - " & txtProperty.text
   End With

   EmptyBoxes
   Call GetRecord
End Sub

Private Sub flxProperty_Click()
   Dim i As Integer
   If flxProperty.TextMatrix(flxProperty.row, 1) = "" Then Exit Sub
   Call SelectOnly1RowFlxGrid(flxProperty, flxProperty.row, 0)
End Sub

Private Sub flxProperty1_Click()
   Dim i As Integer
   If flxProperty1.TextMatrix(flxProperty1.row, 1) = "" Then Exit Sub
   i = SelectFlxGridRow(0, flxProperty1, flxProperty1.row)
End Sub

Private Sub Form_Load()
   Me.Width = 17970 '10590
   Me.Height = 10350 '7020
'   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   fraDemandType.Enabled = False
   Me.BackColor = MODULEBACKCOLOR
   fraCommands.BackColor = MODULEBACKCOLOR
   fraDemandType.BackColor = MODULEBACKCOLOR
   Frame3.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   Label3.BackColor = MODULEBACKCOLOR
   Frame1(1).BackColor = MODULEBACKCOLOR
   Frame1(2).BackColor = MODULEBACKCOLOR
   Frame1(3).BackColor = MODULEBACKCOLOR
   Frame1(4).BackColor = MODULEBACKCOLOR
   Frame1(5).BackColor = MODULEBACKCOLOR
   fraCommand.BackColor = MODULEBACKCOLOR
   optCopyDemandType.BackColor = MODULEBACKCOLOR
   optCopyDemandTemplate.BackColor = MODULEBACKCOLOR
   chkAllDemand.BackColor = MODULEBACKCOLOR
   chkAllProperties.BackColor = MODULEBACKCOLOR
   chkDemandall2.BackColor = MODULEBACKCOLOR
'   fraDemandType.Top = 120
   fraDemandType.Left = 40
   cmdPropertyFilter.Enabled = True
   cmdClientFilter.Enabled = True
   chkGroup.Enabled = False

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

'   LoadNCinCombo adoConn         'all nominal code and name are collecting in all combos from sage

   Call LoadFlxDemandTypes(adoConn, "")

   LoadPaymentDates adoConn

'   LoadDemandCategory adoConn  'all category

'   LoadGroup adoConn

'   LoadBankDetails adoConn      'all clints' bank details

'   LoadProperty adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

'Private Sub LoadProperty(adoConn As Adodb.Connection)
'   Dim adoRst As New Adodb.Recordset
'   Dim szSQL As String
'   Dim TotalRow As Integer, TotalCol As Integer, i As Integer, j As Integer
'
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProPostCode " & _
'           "FROM Property " & _
'           "ORDER BY PropertyID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol - 1, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'      For j = 0 To TotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboProperty.Column() = Data()
'   cboProperty.ListIndex = 0
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

'Private Sub LoadGroup(adoConn As ADODB.Connection)
''   cboGroup.Clear
''
''   Dim szSQL As String, iSt As Integer, iEnd As Integer
''   Dim adoRST As New ADODB.Recordset
''   Dim i As Integer
''
''   szSQL = "SELECT CODE, VALUE " & _
''           "FROM SecondaryCode " & _
''           "WHERE PrimaryCode = 'GR';"
''
''   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   While Not adoRST.EOF
''      If adoRST.Fields.Item("Code").Value = "ENDRNG" Then
''         iEnd = adoRST.Fields.Item("VALUE").Value
''      Else
''         iSt = adoRST.Fields.Item("VALUE").Value
''      End If
''      adoRST.MoveNext
''   Wend
''
''   For i = iSt To iEnd
''      cboGroup.AddItem i
''   Next i
''
''   adoRST.Close
''   Set adoRST = Nothing
'
'
'End Sub

Private Function Get_Bank_AC_Name(My_ID As String, adoConn As ADODB.Connection) As String


   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT My_ID, Bank_AC_Name, BANK_AC_NUM, BANK_SC " & _
           "FROM tlbClientBanks " & _
           "WHERE tlbClientBanks.My_ID = " & My_ID & ";"
                 

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly



   If Not adoRst.EOF Then
        Get_Bank_AC_Name = adoRst.Fields("Bank_AC_Name").Value
   End If
   adoRst.Close
   Set adoRst = Nothing
End Function

'Private Sub LoadBankDetails(adoConn As ADODB.Connection)
''   cboBank.Clear
'
'   Dim szSQL As String
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'   Dim adoRST As New ADODB.Recordset
'
'   szSQL = "SELECT My_ID, Bank_AC_Name, BANK_AC_NUM, BANK_SC " & _
'               "FROM tlbClientBanks;"
'
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRST.RecordCount < 1 Then
'      MsgBox "There are no client's bank details has been setup.", vbCritical + vbOKOnly, "Bank Details Missing"
'      cmdCancelNew_Click
'      Exit Sub
'   End If
'
''   TotalRow = adoRST.RecordCount
''   TotalCol = adoRST.Fields.count
''
''   Dim Data() As String
''   ReDim Data(TotalCol - 1, TotalRow - 1) As String
''
''   For i = 0 To adoRST.RecordCount - 1
''      For j = 0 To adoRST.Fields.count - 1
''         Data(j, i) = adoRST.Fields(j)
''      Next j
''      adoRST.MoveNext
''   Next i
''
''   cboBank.Column() = Data()
'
'   adoRST.Close
'   Set adoRST = Nothing
'End Sub

Private Sub LoadDemandCategory(adoConn As ADODB.Connection)
'   Dim adoRST As New ADODB.Recordset
'   Dim szSQL As String
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   cboDemandTypeCategory.Clear
'
'   szSQL = "SELECT Code, Value " & _
'           "FROM SecondaryCode " & _
'           "WHERE PrimaryCode = 'DCTG';"
'
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRST.RecordCount < 1 Then
'      adoRST.Close
'      Set adoRST = Nothing
'      Exit Sub
'   End If
'
'   TotalRow = adoRST.RecordCount
'   TotalCol = adoRST.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To adoRST.RecordCount - 1
'       For j = 0 To adoRST.Fields.count - 1
'           Data(j, i) = adoRST.Fields(j)
'       Next j
'       adoRST.MoveNext
'   Next i
'
'   cboDemandTypeCategory.Column() = Data()
'
'   adoRST.Close
'   Set adoRST = Nothing
End Sub

Private Sub LoadPaymentDates(adoConn As ADODB.Connection)
'   cboDemandTypePayDates.Clear
'Write flexgrid congfig here
   Dim rRow As Integer
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.Cols = 5
   flxClient.RowHeight(0) = 0
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "ID"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Caption = "Name of Date Set"
   lblClientName.Width = 2600
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientName.Visible = True
   
   txtSearchClientID.Width = 1530
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
'End if config flexgrid

   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim Data() As String
   Dim TotalRow, TotalCol As Integer

   ReDim Data(1, 0) As String

   Data(0, 0) = "0"
   Data(1, 0) = "DEFAULT"

   szSQL = "SELECT NameOfSet " & _
               "FROM PaymentDates " & _
               "ORDER BY DateSetID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount = 0 Then
'      cboDemandTypePayDates.Column() = Data()
        rRow = 1
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = Data(0, 0)
        flxClient.TextMatrix(rRow, 2) = Data(1, 0)
        flxClient.RowHeight(rRow) = 280
        
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   TotalRow = adoRst.RecordCount

   ReDim Data(1, TotalRow) As String
   Dim i As Integer

        adoRst.MoveFirst
        Data(0, 0) = "0"
        Data(1, 0) = "DEFAULT"
        
        rRow = 1
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = Data(0, 0)
        flxClient.TextMatrix(rRow, 2) = Data(1, 0)
        flxClient.RowHeight(rRow) = 280
        flxClient.AddItem ""
        
        rRow = 2
   For i = 1 To adoRst.RecordCount
        Data(0, i) = i
        Data(1, i) = adoRst("NameOfSet").Value
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = Data(0, i)
        flxClient.TextMatrix(rRow, 2) = Data(1, i)
        flxClient.RowHeight(rRow) = 280
        If Not adoRst.EOF Then flxClient.AddItem ""
        rRow = rRow + 1
        adoRst.MoveNext
   Next i

'   cboDemandTypePayDates.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadChargeType(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow, TotalCol As Integer
   Dim Data() As String

   szSQL = "SELECT ID, PayType, PayIC, PaySagePrefix, PayNCAmt, PayNNAmt, " & _
                  "PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, TransactionType, " & _
                  "CategoryCode, PaymentDates " & _
                "FROM PayableTypes ORDER BY ID"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount < 1 Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub

'Private Sub LoadNCinCombo(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, TotalRow As Integer
'   Dim Data() As String, i As Integer
'
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   TotalRow = adoRst.RecordCount
'   ReDim Data(2, TotalRow) As String
'
'   i = 0
'   While Not adoRst.EOF
'      Data(0, i) = adoRst.Fields.Item("Code").Value
'      Data(1, i) = adoRst.Fields.Item("Name").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cboDemandTypeNCAmt.Column() = Data()
''   cboDemandTypeNCvat.Column() = Data()
''   cboDemandTypeNCTotal.Column() = Data()
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Private Sub ConfigFlxDemandTypes()
    Dim szHeader As String
    flxDemandTypes.Clear
    flxDemandTypes.Rows = 2
    szHeader$ = "|<ClientID|<ClientName|<PropertyID|<PropertyName|<ID|<Type"
    flxDemandTypes.FormatString = szHeader$
    With flxDemandTypes
        .Cols = 17
        .RowHeight(0) = 0
        .ColWidth(0) = 200 'Label1(20).Left - .Left                 'Selection column
        .ColWidth(1) = Label1(26).Left - Label1(25).Left + 100        'Client Name
        .ColWidth(2) = 0 '          'empty
        .ColWidth(3) = Label1(27).Left - Label1(26).Left - 100         'Property Name
        .ColWidth(4) = 0 '           'empty
        .ColWidth(5) = 0 '           'empty
        .ColWidth(6) = Label1(28).Left - Label1(27).Left            'ID
        .ColWidth(7) = Label1(29).Left - Label1(28).Left            'Prefix
        .ColWidth(8) = Label1(30).Left - Label1(29).Left            'PayType
        .ColWidth(9) = Label1(31).Left - Label1(30).Left             'Charge Category Name
        .ColWidth(10) = 0                                            'Group ID
        .ColWidth(11) = 0                                            'Client Name
        .ColWidth(12) = 0                                            'Client ID
        .ColWidth(13) = 0                                            'Property Name
        .ColWidth(14) = 0                                            'Property ID
        .ColWidth(15) = Label1(32).Left - Label1(31).Left                                             'Nominal Code
        .ColAlignment(15) = vbAlignLeft
        .ColWidth(16) = Label1(33).Left - Label1(32).Left                                            'Nominal Name
    End With
End Sub
Private Function DemandtypeSQL() As String
'    Dim szSQL1 As String
        If txtClientList.Tag = "ALL" And txtPropertyList.Tag = "ALL" Then
                DemandtypeSQL = "SELECT D.*, " & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
                "FROM (PayableTypes AS D INNER JOIN Property AS P ON " & _
                      "D.PropertyID = P.PropertyID ) INNER JOIN Client AS C ON P.ClientID = C.ClientID " & _
                "ORDER BY ClientName, PropertyName, D.PayType, D.ID;"
         ElseIf txtPropertyList.Tag <> "ALL" Then
                DemandtypeSQL = "SELECT D.*," & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
                "FROM (PayableTypes AS D INNER JOIN Property AS P ON " & _
                      "D.PropertyID = P.PropertyID) INNER JOIN Client AS C ON P.ClientID = C.ClientID " & _
                      "WHERE D.PropertyID='" & txtPropertyList.Tag & "' " & _
                "ORDER BY ClientName, PropertyName, D.PayType, D.ID;"
         ElseIf txtClientList.Tag <> "ALL" Then
                    DemandtypeSQL = "SELECT D.*, " & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
                "FROM (PayableTypes AS D  INNER JOIN Property AS P ON " & _
                      "D.PropertyID = P.PropertyID) INNER JOIN Client AS C ON P.ClientID = C.ClientID " & _
                      "where C.ClientID='" & txtClientList.Tag & "' " & _
                "ORDER BY ClientName, PropertyName, D.PayType, D.ID;"
         End If

End Function
Public Sub LoadFlxDemandTypes(adoConn As ADODB.Connection, Filter As String)
   Dim Data() As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim szSQL As String, szHeader As String
   Dim adoRst As New ADODB.Recordset
   
'    szSQL = "SELECT ID,CL.ClientName,CL.ClientID,P.PropertyName,P.PropertyID, PayType, FeeIC, FeeSagePrefix, FeeNCAmt, FeeNNAmt, " & _
'                  "FeeNCVat, FeeNNVat, FeeNCTotal, FeeNNTotal, TransactionType, " & _
'                  "CategoryCode, PaymentDates, RecoverableExp, Value as CategoryName,GroupID " & _
'                "FROM PayableTypes C,SecondaryCode S,Property P,Client CL where CL.ClientID=C.ClientID  " & _
'                "and P.ClientID=CL.ClientID AND P.PropertyID=C.PropertyID AND  S.Code=cstr(C.CategoryCode) AND  S.PrimaryCode='DCTG' ORDER BY ID"

   szSQL = DemandtypeSQL
   ConfigFlxDemandTypes
   'adoRst.Close
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
   If Filter <> "" Then
        adoRst.Filter = Filter
   End If
   With adoRst.Fields
      For i = 0 To adoRst.RecordCount - 1
         flxDemandTypes.TextMatrix(i + 1, 1) = IIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(i + 1, 2) = "" 'Keep number 2 empty for future useIIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(i + 1, 3) = IIf(IsNull(.Item("PropertyName")), "", .Item("PropertyName"))
         flxDemandTypes.TextMatrix(i + 1, 4) = "" 'Keep number 4 empty for future useIIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(i + 1, 5) = "" 'Keep number 5 empty for future use
         flxDemandTypes.TextMatrix(i + 1, 6) = IIf(IsNull(.Item("ID")), "", .Item("ID"))
         flxDemandTypes.TextMatrix(i + 1, 7) = IIf(IsNull(.Item("PaySagePrefix")), "", .Item("PaySagePrefix"))
         flxDemandTypes.TextMatrix(i + 1, 8) = IIf(IsNull(.Item("PayType")), "", .Item("PayType"))
'         flxDemandTypes.TextMatrix(i + 1, 9) = IIf(IsNull(.Item("CategoryName")), "", .Item("CategoryName"))
         flxDemandTypes.TextMatrix(i + 1, 10) = IIf(IsNull(.Item("GroupID")), "", .Item("GroupID"))
         
         flxDemandTypes.TextMatrix(i + 1, 11) = IIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(i + 1, 12) = IIf(IsNull(.Item("ClientID")), "", .Item("ClientID"))
         flxDemandTypes.TextMatrix(i + 1, 13) = IIf(IsNull(.Item("PropertyName")), "", .Item("PropertyName"))
         flxDemandTypes.TextMatrix(i + 1, 14) = IIf(IsNull(.Item("PropertyID")), "", .Item("PropertyID"))
         flxDemandTypes.TextMatrix(i + 1, 15) = IIf(IsNull(.Item("PayNCAmt")), "", .Item("PayNCAmt"))
         flxDemandTypes.TextMatrix(i + 1, 16) = IIf(IsNull(.Item("PayNNAmt")), "", .Item("PayNNAmt"))
         If Val(IIf(IsNull(.Item("ID")), "", .Item("ID"))) = Val(txtID.text) Then
                intTopRow = i + 1
         End If
         adoRst.MoveNext
         If Not adoRst.EOF Then flxDemandTypes.AddItem ""
      Next i
   End With
   
   flxDemandTypes.row = 0
   iSelRow = 0
   'Now build tree view
   'this will remove the duplicate values in client and properties
                Dim a As Integer
                Dim b As Integer
                Dim PropertyArray() As String
                Dim ClientArray() As String
                ReDim PropertyArray(flxDemandTypes.Rows - 1, 0)
                ReDim ClientArray(flxDemandTypes.Rows - 1, 0)
                'saving all property ID,client ID in an array
                For a = 1 To flxDemandTypes.Rows - 2
                       PropertyArray(a, 0) = flxDemandTypes.TextMatrix(a, 14)
                       ClientArray(a, 0) = flxDemandTypes.TextMatrix(a, 12)
                Next a
        
            'tree building only for property Name
                For a = 1 To flxDemandTypes.Rows - 1
                        For b = a + 1 To flxDemandTypes.Rows - 1
                                If flxDemandTypes.TextMatrix(a, 14) = flxDemandTypes.TextMatrix(b, 14) Then
                                     flxDemandTypes.TextMatrix(b, 14) = ""
                                End If
                        Next b
                Next a
            'tree building only for Client name
                For a = 1 To flxDemandTypes.Rows - 2
                    For b = a + 1 To flxDemandTypes.Rows - 1
                        If flxDemandTypes.TextMatrix(a, 1) = flxDemandTypes.TextMatrix(b, 1) Then
                             ' duplicate value is found  in client
                             flxDemandTypes.TextMatrix(b, 12) = ""
                        End If
                    Next b
                Next a

NoRes:
   adoRst.Close
End Sub

Public Sub EmptyBoxes()
'   txtDemandTypeNCAmt.text = ""
'   cboDemandTypeNCvat.text = ""
'   cboDemandTypeNCTotal.text = ""
'   txtDemandCategoryCode.text = ""
'   txtDemandCategoryName.text = ""
'   cboBank.text = ""
   txtGroup.text = ""
   txtType.text = ""
   txtID.text = ""
   txtPrefix.text = ""
   txtDemandTemplate.text = ""
   txtEmailTemplate.text = ""

   chkGroup.Value = False
'   txt1.text = ""
'   txt2.text = ""
'   txt3.text = ""
'   txt4.text = ""
End Sub

Public Sub EnableBoxes()
   cmdAdd.Enabled = False

'   cmdDemandTypeNCAmt.Enabled = True
'   cboDemandTypeNCvat.Enabled = True
'   cboDemandTypeNCTotal.Enabled = True
'   cmdDemandCategory.Enabled = True
  ' cmdBank.Enabled = True

   'cmdBrowsFile(0).Enabled = True
   'cmdBrowsFile(1).Enabled = True
   chkGroup.Enabled = True
End Sub

Public Sub DisableBoxes()
   cmdAdd.Enabled = True

'   cmdDemandTypeNCAmt.Enabled = False
'   cboDemandTypeNCvat.Enabled = False
'   cboDemandTypeNCTotal.Enabled = False
'   cmdDemandCategory.Enabled = False
   'cmdBank.Enabled = False

'   cmdBrowsFile(0).Enabled = False
 '  cmdBrowsFile(1).Enabled = False
   chkGroup.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub fraCommands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub fraDemandType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Label1_Click(Index As Integer)
   Dim i As Integer
   For i = 1 To flxDemandTypes.Rows - 1
        flxDemandTypes.TextMatrix(i, 2) = flxDemandTypes.TextMatrix(i, 7)
        flxDemandTypes.TextMatrix(i, 4) = flxDemandTypes.TextMatrix(i, 8)
   Next i
   If Index = 23 Then                                       ' Sort ID
      SortingGrid flxDemandTypes, 6, bSortingCol1, "Integer"
      If flxDemandTypes.Rows > 1 Then
            flxDemandTypes.TopRow = 1
      End If
      bSortingCol1 = IIf(bSortingCol1, False, True)
      Label1(23).FontBold = True
      Label1(23).ForeColor = RGB(0, 0, 255)
      Label1(20).FontBold = False
      Label1(21).FontBold = False
      Label1(22).FontBold = False
   End If
   If Index = 22 Then                                       ' Sort Type
      SortingGrid flxDemandTypes, 5, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      Label1(23).FontBold = False
      Label1(20).FontBold = False
      Label1(21).FontBold = False
      Label1(22).FontBold = True
      Label1(22).BackColor = RGB(0, 0, 255)
   End If
   If Index = 21 Then                                       ' Property Name
      SortingGrid flxDemandTypes, 4, bSortingCol3
      bSortingCol3 = IIf(bSortingCol3, False, True)
      Label1(23).FontBold = False
      Label1(20).FontBold = False
      Label1(21).FontBold = True
      Label1(21).BackColor = RGB(0, 0, 255)
      Label1(22).FontBold = False
   End If
   If Index = 20 Then                                       ' Client name
      SortingGrid flxDemandTypes, 2, bSortingCol4
      bSortingCol4 = IIf(bSortingCol4, False, True)
      Label1(23).FontBold = False
      Label1(20).FontBold = True
      Label1(20).BackColor = RGB(0, 0, 255)
      Label1(21).FontBold = False
      Label1(22).FontBold = False
   End If
   flxDemandTypes.row = 0
End Sub

Private Sub optAuto_Click()
    txtDemandTypePayDates.text = ""
    txtDemandTypePayDates.Tag = ""
   cmdDemandTypePayDates.Enabled = Not optAuto.Value
End Sub

Private Sub optCopyDemandTemplate_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl cmdNext
    End If
End Sub

Private Sub optCopyDemandType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdNext
    End If
End Sub

Private Sub optPreset_Click()
   cmdDemandTypePayDates.Enabled = optPreset.Value
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






Private Sub TextBox1_Change()
           'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 3)), UCase(TextBox1.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtClientList_Change()
'This one is taking a long time
''   'search in the demand type grid by client
''    Dim i As Integer
''    txtClientList.text = UCase(txtClientList.text)
''   If Len(txtClientList.text) > 0 Then
''        txtPropertyList.text = ""
''   End If
''
''   For i = flxDemandTypes.Rows - 1 To 1 Step -1
''        flxDemandTypes.RowHeight(i) = 240
''        If InStr(1, UCase(flxDemandTypes.TextMatrix(i, 2)), UCase(txtClientList.text), vbTextCompare) = 0 Then
''              flxDemandTypes.RowHeight(i) = 0
''        End If
''        Debug.Print UCase(flxDemandTypes.TextMatrix(i, 2))
''        If flxDemandTypes.RowHeight(i) = 240 Then
''              flxDemandTypes.row = i
''        End If
''   Next i
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
'    If Len(txtClientList.text) > 0 Then
'        Call LoadFlxDemandTypes(adoConn, " ClientID Like '%" + UCase(txtClientList.text) + "*'")
'    Else
        Call LoadFlxDemandTypes(adoConn, "")
'    End If
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtPropertyList.SetFocus
    End If
End Sub

Private Sub txtFilterbyProperty_Change()
    'search in the demand type grid by client
    Dim i As Integer
    txtFilterbyProperty.text = UCase(txtFilterbyProperty.text)
   

   For i = flxProperty.Rows - 1 To 1 Step -1
        flxProperty.RowHeight(i) = 240
        If InStr(1, UCase(flxProperty.TextMatrix(i, 2)), UCase(txtFilterbyProperty.text), vbTextCompare) = 0 Then
              flxProperty.RowHeight(i) = 0
        End If
        Debug.Print UCase(flxProperty.TextMatrix(i, 2))
        If flxProperty.RowHeight(i) = 240 Then
              flxProperty.row = i
        End If
   Next i
End Sub

Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        cmdDemandTypeNCAmt.SetFocus
    End If
End Sub

Private Sub txtPropertyList_Change()
''search in the demand type grid by Property
'      Dim i As Integer
'    txtPropertyList.text = UCase(txtPropertyList.text)
'   If Len(txtPropertyList.text) > 0 Then
'        txtClientList.text = ""
'   End If
'
'   For i = flxDemandTypes.Rows - 1 To 1 Step -1
'        flxDemandTypes.RowHeight(i) = 240
'        If InStr(1, UCase(flxDemandTypes.TextMatrix(i, 4)), UCase(txtPropertyList.text), vbTextCompare) = 0 Then
'              flxDemandTypes.RowHeight(i) = 0
'        End If
'        If flxDemandTypes.RowHeight(i) = 240 Then
'              flxDemandTypes.row = i
'        End If
'   Next i
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
'    Call LoadFlxDemandTypes(adoConn, "D.PropertyID Like'%" + UCase(txtPropertyList.text) + "*'")
    Call LoadFlxDemandTypes(adoConn, "")
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub txtPropertyList_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        If KeyCode = 13 Then
                flxDemandTypes.SetFocus
        End If
    
End Sub

Private Sub txtPropertyList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxDemandTypes.SetFocus
    End If
    
End Sub

Private Sub txtSearchClientID_Change()
        'Updated by anol 22 Dec 2015
   Dim i As Integer
'    txtSearchClientID.text = UCase(txtSearchClientID.text)
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

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        If Len(txtSearchClientID) > 0 Then
            flxClient.SetFocus
        Else
            txtSearchClientName.SetFocus
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

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxClient.SetFocus
    End If
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPrefix.SetFocus
    End If
End Sub
