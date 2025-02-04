VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNominalLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nominal Accounts"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17235
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNominalLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   17235
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Index           =   2
      Left            =   3600
      TabIndex        =   68
      Top             =   2790
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CommandButton cmdShowAllclient 
         Caption         =   "All"
         Height          =   300
         Left            =   5265
         TabIndex        =   69
         Top             =   315
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientCopy 
         Height          =   1905
         Left            =   45
         TabIndex        =   70
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3360
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
      Begin VB.Label Label4 
         Caption         =   "Copy  From Client"
         Height          =   240
         Left            =   180
         TabIndex        =   73
         Top             =   315
         Width           =   1950
      End
      Begin MSForms.TextBox txtFilterbyProperty 
         Height          =   285
         Left            =   3285
         TabIndex        =   72
         Tag             =   "ALL"
         Top             =   315
         Width           =   1935
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3413;503"
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
         Index           =   0
         Left            =   2430
         TabIndex        =   71
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Index           =   1
      Left            =   3510
      TabIndex        =   57
      Top             =   1395
      Visible         =   0   'False
      Width           =   6360
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   1
         Left            =   45
         ScaleHeight     =   555
         ScaleWidth      =   6210
         TabIndex        =   64
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
            TabIndex        =   65
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Create Chart of Accounts Wizard"
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
            TabIndex        =   66
            Top             =   135
            Width           =   6000
         End
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   4935
         TabIndex        =   63
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   3555
         TabIndex        =   62
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2175
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelWizard 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   135
         TabIndex        =   60
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optCopyDefltCOA 
         Caption         =   "Copy  from default Chart of Accounts"
         Height          =   195
         Left            =   360
         TabIndex        =   59
         Top             =   810
         Value           =   -1  'True
         Width           =   3480
      End
      Begin VB.OptionButton optCopyDemandTemplate 
         Caption         =   "Copy  from existing Client's Chart of Accounts"
         Height          =   330
         Left            =   360
         TabIndex        =   58
         Top             =   1080
         Width           =   5145
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   4785
         Left            =   45
         Top             =   180
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
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   13815
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   39
      Top             =   7875
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
         TabIndex        =   18
         Top             =   60
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   90
         TabIndex        =   17
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   43
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   42
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   41
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
         Left            =   1620
         TabIndex        =   40
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
         TabIndex        =   15
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
         TabIndex        =   16
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
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5850
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   90
      TabIndex        =   44
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton cmdConvertNetAmountToDRCR 
         Caption         =   "Convert"
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
         Left            =   12105
         TabIndex        =   11
         Top             =   225
         Width           =   1200
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
         Height          =   345
         Left            =   8445
         TabIndex        =   1
         Top             =   230
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
         Left            =   4245
         TabIndex        =   0
         Top             =   225
         Width           =   300
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fund:"
         Height          =   195
         Index           =   1
         Left            =   8820
         TabIndex        =   50
         Top             =   285
         Width           =   390
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Property:"
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   49
         Top             =   285
         Width           =   645
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Client:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   48
         Top             =   285
         Width           =   465
      End
      Begin MSForms.CommandButton cmdFundLookUp 
         Height          =   345
         Left            =   11730
         TabIndex        =   2
         Top             =   225
         Width           =   300
         Caption         =   ".."
         Size            =   "529;609"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   645
         TabIndex        =   47
         Top             =   225
         Width           =   3780
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6667;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   5445
         TabIndex        =   46
         Tag             =   "ALL"
         Top             =   225
         Width           =   3105
         VariousPropertyBits=   746604571
         Size            =   "5477;556"
         Value           =   "ALL Properties"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtFundName 
         Height          =   315
         Left            =   9345
         TabIndex        =   45
         Tag             =   "ALL"
         Top             =   225
         Width           =   2430
         VariousPropertyBits=   746604571
         Size            =   "4286;556"
         Value           =   "All Funds"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   19
      Top             =   660
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   14208
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Chart of Accounts"
      TabPicture(0)   =   "frmNominalLedger.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGridCaption(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(66)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbPeriodFrom"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(12)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbPeriodTo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDebitTotal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblCurrentPnL"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblCreditTotal"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Shape1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtBudgetYears"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdBudgetYears"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "flxNominalCode"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkYtD"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdFilter"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdPrint"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdDelete"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdAddNew"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdEdit"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdClose"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdClear"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "picBankCode"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdCopy"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdSearch"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "fraSearch"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "Default Chart of Accounts"
      TabPicture(1)   =   "frmNominalLedger.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEdit2"
      Tab(1).Control(1)=   "cmdDeleteDefaultCOA"
      Tab(1).Control(2)=   "cmdAddnewDefaultCOA"
      Tab(1).Control(3)=   "cmdClose1"
      Tab(1).Control(4)=   "flxChartOfACCDefault"
      Tab(1).Control(5)=   "Shape1(5)"
      Tab(1).Control(6)=   "Label1(16)"
      Tab(1).Control(7)=   "Label1(15)"
      Tab(1).Control(8)=   "Label1(11)"
      Tab(1).Control(9)=   "Label1(10)"
      Tab(1).Control(10)=   "Label1(9)"
      Tab(1).Control(11)=   "lblGridCaption(1)"
      Tab(1).ControlCount=   12
      Begin VB.CommandButton cmdEdit2 
         Caption         =   "Edit"
         Height          =   375
         Left            =   -73560
         TabIndex        =   95
         Top             =   7380
         Width           =   1080
      End
      Begin VB.CommandButton cmdDeleteDefaultCOA 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   -71685
         TabIndex        =   94
         Top             =   7365
         Width           =   1035
      End
      Begin VB.CommandButton cmdAddnewDefaultCOA 
         Caption         =   "Add &New"
         Height          =   375
         Left            =   -74835
         TabIndex        =   93
         Top             =   7365
         Width           =   1080
      End
      Begin VB.CommandButton cmdClose1 
         Cancel          =   -1  'True
         Caption         =   "Cl&ose"
         Height          =   375
         Left            =   -63180
         TabIndex        =   92
         Top             =   7365
         Width           =   1125
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Automatic Demand Generate:"
         ForeColor       =   &H00FF00FF&
         Height          =   2220
         Left            =   540
         TabIndex        =   75
         Top             =   2925
         Visible         =   0   'False
         Width           =   3715
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E5E5E5&
            Height          =   2100
            Index           =   0
            Left            =   45
            ScaleHeight     =   2040
            ScaleWidth      =   3555
            TabIndex        =   76
            Top             =   50
            Width           =   3615
            Begin VB.CommandButton cmdSearchOK 
               Caption         =   "&OK"
               Height          =   375
               Left            =   120
               TabIndex        =   81
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CommandButton cmdSearchCancel 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   2055
               TabIndex        =   80
               Top             =   1635
               Width           =   1200
            End
            Begin VB.TextBox txtSearchNo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1575
               MaxLength       =   10
               TabIndex        =   79
               Top             =   450
               Width           =   1830
            End
            Begin VB.TextBox txtSearchRef 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1575
               MaxLength       =   20
               TabIndex        =   78
               Top             =   790
               Width           =   1830
            End
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
               TabIndex        =   77
               Top             =   0
               Width           =   255
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H00FFC0C0&
               FillStyle       =   0  'Solid
               Height          =   55
               Index           =   4
               Left            =   0
               Top             =   240
               Width           =   3855
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Search Options"
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
               Index           =   1
               Left            =   300
               TabIndex        =   84
               Top             =   0
               Width           =   1200
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               Height          =   1155
               Index           =   2
               Left            =   75
               Top             =   360
               Width           =   3450
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00FFC0C0&
               BorderWidth     =   3
               Height          =   1155
               Index           =   7
               Left            =   75
               Top             =   360
               Width           =   3450
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Nominal code"
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
               Index           =   2
               Left            =   180
               TabIndex        =   83
               Top             =   450
               Width           =   1110
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Nominal Name"
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
               Index           =   3
               Left            =   180
               TabIndex        =   82
               Top             =   810
               Width           =   1185
            End
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4455
         TabIndex        =   74
         Top             =   7290
         Width           =   1125
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Co&py"
         Height          =   375
         Left            =   1350
         TabIndex        =   67
         Top             =   7290
         Width           =   930
      End
      Begin VB.PictureBox picBankCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   9720
         ScaleHeight     =   3180
         ScaleWidth      =   2640
         TabIndex        =   51
         Top             =   1305
         Visible         =   0   'False
         Width           =   2670
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
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   90
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
            Height          =   2715
            Left            =   45
            TabIndex        =   56
            Top             =   450
            Width           =   2565
            _ExtentX        =   4524
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
         Begin MSForms.Label Label8 
            Height          =   195
            Left            =   1725
            TabIndex        =   54
            Top             =   150
            Width           =   1185
            VariousPropertyBits=   8388627
            Caption         =   "Bank Name"
            Size            =   "2090;344"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label9 
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   135
            Width           =   1230
            VariousPropertyBits=   8388627
            Caption         =   "Bank Code"
            Size            =   "2170;344"
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
            Width           =   2205
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12020
         TabIndex        =   6
         Top             =   495
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cl&ose"
         Height          =   375
         Left            =   11985
         TabIndex        =   10
         Top             =   7335
         Width           =   1125
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
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   6750
         TabIndex        =   14
         Top             =   7290
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add &New"
         Height          =   375
         Left            =   225
         TabIndex        =   7
         Top             =   7290
         Width           =   1080
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3390
         TabIndex        =   9
         Top             =   7290
         Width           =   1035
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2325
         TabIndex        =   8
         Top             =   7290
         Width           =   1035
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Display"
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
         Left            =   10800
         TabIndex        =   5
         Top             =   495
         Width           =   1200
      End
      Begin VB.CheckBox chkYtD 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Y&TD"
         Height          =   255
         Left            =   9960
         TabIndex        =   13
         Top             =   525
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNominalCode 
         Height          =   5880
         Left            =   120
         TabIndex        =   20
         Top             =   1155
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   10372
         _Version        =   393216
         Cols            =   6
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton cmdBudgetYears 
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
         Left            =   3825
         TabIndex        =   3
         Top             =   495
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChartOfACCDefault 
         Height          =   6420
         Left            =   -74955
         TabIndex        =   91
         Top             =   810
         Width           =   13170
         _ExtentX        =   23230
         _ExtentY        =   11324
         _Version        =   393216
         Cols            =   6
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   615
         Index           =   5
         Left            =   -74955
         Top             =   7290
         Width           =   13125
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   16
         Left            =   -70275
         TabIndex        =   89
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name"
         Height          =   195
         Index           =   15
         Left            =   -73155
         TabIndex        =   88
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
         Height          =   195
         Index           =   11
         Left            =   -74955
         TabIndex        =   87
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "  Posting      Account"
         Height          =   195
         Index           =   10
         Left            =   -63525
         TabIndex        =   86
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Type"
         Height          =   195
         Index           =   9
         Left            =   -68475
         TabIndex        =   85
         Top             =   600
         Width           =   810
      End
      Begin MSForms.TextBox txtBudgetYears 
         Height          =   285
         Left            =   1485
         TabIndex        =   55
         Top             =   495
         Width           =   2340
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4128;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Surplus(Deficit) :"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5805
         TabIndex        =   38
         Top             =   7335
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   840
         Index           =   3
         Left            =   120
         Top             =   7125
         Width           =   13125
      End
      Begin VB.Label lblCreditTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9765
         TabIndex        =   37
         Top             =   7335
         Width           =   1455
      End
      Begin VB.Label lblCurrentPnL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6660
         TabIndex        =   36
         Top             =   7350
         Width           =   1860
      End
      Begin VB.Label lblDebitTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8055
         TabIndex        =   35
         Top             =   7335
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Type"
         Height          =   195
         Index           =   4
         Left            =   6600
         TabIndex        =   34
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "  Posting Account"
         Height          =   195
         Index           =   7
         Left            =   11700
         TabIndex        =   29
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Period To:"
         Height          =   195
         Index           =   14
         Left            =   7200
         TabIndex        =   28
         Top             =   540
         Width           =   705
      End
      Begin MSForms.ComboBox cmbPeriodTo 
         Height          =   285
         Left            =   7920
         TabIndex        =   4
         Top             =   495
         Width           =   1920
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Period From:"
         Height          =   195
         Index           =   12
         Left            =   4320
         TabIndex        =   27
         Top             =   540
         Width           =   885
      End
      Begin MSForms.ComboBox cmbPeriodFrom 
         Height          =   285
         Left            =   5220
         TabIndex        =   12
         Top             =   480
         Width           =   1920
         VariousPropertyBits=   1753237529
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year:"
         Height          =   195
         Index           =   66
         Left            =   240
         TabIndex        =   26
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   24
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Credit"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   10290
         TabIndex        =   23
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Debit"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   8610
         TabIndex        =   22
         Top             =   945
         Width           =   435
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   21
         Top             =   945
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00F7F7F7&
         FillStyle       =   0  'Solid
         Height          =   480
         Index           =   0
         Left            =   120
         Top             =   420
         Width           =   13125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   480
         Index           =   1
         Left            =   120
         Top             =   420
         Width           =   13125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   840
         Index           =   2
         Left            =   120
         Top             =   7125
         Width           =   13125
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FCE0E0&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   33
         Top             =   945
         Width           =   13125
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FCE0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -74955
         TabIndex        =   90
         Top             =   540
         Width           =   13170
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code"
      Height          =   195
      Index           =   77
      Left            =   4920
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Type ID"
      Height          =   195
      Index           =   8
      Left            =   6120
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "DrCr"
      Height          =   195
      Index           =   13
      Left            =   6960
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "frmNominalLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Asif. Issue 0000476

Option Explicit

Public CALLER_FORM   As String
Public MiLoading     As Boolean
Public dtStartPnL    As Date
Public dtStartBS     As Date
Public dtEnd         As Date

Private iNewEdit     As Byte
Private TotalNC      As Integer
Private RE           As Double

Dim reportingDate As String
Dim sessionID As String
Dim sTextBox  As String
Public Form_Activated As Boolean
Dim sMode As String
Dim iFinancialWarning As Integer
Public isFinancialYearCreated As Boolean


Private Sub chkYtD_Click()

   If chkYtD.Value = 1 Then
      cmbPeriodFrom.ListIndex = -1
      cmbPeriodFrom.Enabled = False
      
      If cmbPeriodTo.ListCount > 0 Then
        cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If
   Else
      cmbPeriodFrom.Enabled = True
      If cmbPeriodFrom.ListCount > 0 Then
        cmbPeriodFrom.ListIndex = 0
      End If
      
      If cmbPeriodTo.ListCount > 0 Then
        cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If
   End If
End Sub

'Private Sub cmbClient_Change()
''   If MiLoading Then Exit Sub
'   If txtClientlist.text = "" Then Exit Sub
'
'  ' On Error GoTo ERR_HANDER
'
'   Dim K          As Integer
'   Dim iRow       As Integer
'   Dim adoConn    As New ADODB.Connection
'   Dim szSQL      As String
'   Dim adoRstSrc  As New ADODB.Recordset
'   Dim adoRstDst  As New ADODB.Recordset
'
'   adoConn.Open getConnectionString
'
'   LoadCmbFinancialYear adoConn
'
'   LoadCmbProperties adoConn, cmbProperty
'
''   GenerateNominalAccounts adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   Exit Sub
'
'ERR_HANDER:
'   ShowMsgInTaskBar "Nominal Ledger could not be loaded for the selected client", "Y", "N"
'   MsgBox ERR.description
''   adoConn.RollbackTrans
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub
Private Function NominalLedgerCreationFromDedault(adoConn As ADODB.Connection) As Boolean

Dim szSQL      As String

Dim adoRstCheck  As New ADODB.Recordset
Dim adoRstSrc  As New ADODB.Recordset
Dim adoRstDst  As New ADODB.Recordset

On Error GoTo ERR_HANDER
'Resolved by BOSL
'issue 532 Cannot copy default chart of account for new client
'Modified by anol 04 Deb 2015

adoConn.BeginTrans
szSQL = "SELECT * FROM NOMINALLEDGER WHERE CLIENTID = '" & txtClientList.Tag & "'"
adoRstCheck.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly

If Not adoRstCheck.EOF Then
    adoConn.RollbackTrans
    MsgBox "Chart of accounts already exists for this client", vbInformation, "Warning!!"
    Exit Function
End If

' If MsgBox(txtClientList.text & " does not have a set of nominal codes." & Chr(13) & _
'               "Do you wish to copy default chart of accounts?", vbQuestion + vbYesNo, "Nominal Ledger Not Setup") = vbYes Then
         
'         iRow = flxNominalCode.Rows

        ' adoConn.BeginTrans

         szSQL = "SELECT Code, Name, Type, DrCr " & _
                 "FROM   NominalLedger " & _
                 "WHERE  ClientID = 'NONE';"
         adoRstSrc.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If adoRstSrc.EOF Then
             MsgBox "Default chart of account not found!!", vbOKOnly + vbCritical, "Warming!!"
             adoRstSrc.Close
             adoConn.RollbackTrans
             Exit Function
        End If
         szSQL = "SELECT * " & _
                 "FROM   NominalLedger;"
         adoRstDst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
         'adoConn.CursorLocation = adUseClient
         'adoConn.BeginTrans
         While Not adoRstSrc.EOF
            adoRstDst.AddNew
            
            adoRstDst.Fields.Item("Code").Value = adoRstSrc.Fields.Item("Code").Value
            adoRstDst.Fields.Item("Name").Value = adoRstSrc.Fields.Item("Name").Value
            adoRstDst.Fields.Item("Type").Value = adoRstSrc.Fields.Item("Type").Value
            adoRstDst.Fields.Item("DrCr").Value = adoRstSrc.Fields.Item("DrCr").Value
            adoRstDst.Fields.Item("Posting").Value = True
            adoRstDst.Fields.Item("ClientID").Value = txtClientList.Tag
            adoRstSrc.MoveNext
            adoRstDst.Update
            NominalLedgerCreationFromDedault = True
         Wend

         adoRstDst.Close
         Set adoRstDst = Nothing
         adoRstSrc.Close
         Set adoRstSrc = Nothing

         adoConn.CommitTrans
         'Below line is added by anol 04 Feb 2015
         Exit Function
'      End If

ERR_HANDER:
   ShowMsgInTaskBar "There was a problem to create a set of nominal code", "Y", "N"
   'MsgBox ERR.description
   adoConn.RollbackTrans

End Function
Private Function NominalLedgerCreationFromClient(adoConn As ADODB.Connection, ClientID As String) As Boolean

Dim szSQL      As String

Dim adoRstCheck  As New ADODB.Recordset
Dim adoRstSrc  As New ADODB.Recordset
Dim adoRstDst  As New ADODB.Recordset

On Error GoTo ERR_HANDER
'Resolved by BOSL
'issue 532 Cannot copy default chart of account for new client
'Modified by anol 04 Deb 2015

adoConn.BeginTrans
szSQL = "SELECT * FROM NOMINALLEDGER WHERE CLIENTID = '" & txtClientList.Tag & "'"
adoRstCheck.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly

If Not adoRstCheck.EOF Then
    adoConn.RollbackTrans
    MsgBox "Chart of accounts already exists for this client", vbInformation, "Warning!!"
    Exit Function
End If

' If MsgBox(txtClientList.text & " does not have a set of nominal codes." & Chr(13) & _
'               "Do you wish to copy default chart of accounts?", vbQuestion + vbYesNo, "Nominal Ledger Not Setup") = vbYes Then
         
'         iRow = flxNominalCode.Rows

        ' adoConn.BeginTrans

         szSQL = "SELECT Code, Name, Type, DrCr " & _
                 "FROM   NominalLedger " & _
                 "WHERE  ClientID = '" & ClientID & "';"
         adoRstSrc.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         If adoRstSrc.EOF Then
             MsgBox "Chart of account not found on source client '" & ClientID & "' !!", vbOKOnly + vbCritical, "Warning!!"
             adoRstSrc.Close
             adoConn.RollbackTrans
             Exit Function
        End If

         szSQL = "SELECT * " & _
                 "FROM   NominalLedger;"
         adoRstDst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
         'adoConn.CursorLocation = adUseClient
         'adoConn.BeginTrans
         While Not adoRstSrc.EOF
            adoRstDst.AddNew
            adoRstDst.Fields.Item("CreatedBy").Value = User
            adoRstDst.Fields.Item("CreatedDate").Value = Now
            adoRstDst.Fields.Item("Code").Value = adoRstSrc.Fields.Item("Code").Value
            adoRstDst.Fields.Item("Name").Value = adoRstSrc.Fields.Item("Name").Value
            adoRstDst.Fields.Item("Type").Value = adoRstSrc.Fields.Item("Type").Value
            adoRstDst.Fields.Item("DrCr").Value = adoRstSrc.Fields.Item("DrCr").Value
            adoRstDst.Fields.Item("CAType").Value = ""
            adoRstDst.Fields.Item("Posting").Value = True
            adoRstDst.Fields.Item("ClientID").Value = txtClientList.Tag
            adoRstSrc.MoveNext
            adoRstDst.Update
            NominalLedgerCreationFromClient = True
         Wend

         adoRstDst.Close
         Set adoRstDst = Nothing
         adoRstSrc.Close
         Set adoRstSrc = Nothing

         adoConn.CommitTrans
         'Below line is added by anol 04 Feb 2015
         Exit Function
'      End If

ERR_HANDER:
   ShowMsgInTaskBar "There was a problem to create a set of nominal code", "Y", "N"
   'MsgBox ERR.description
   adoConn.RollbackTrans

End Function

Public Sub NominalLedgerSetupForNewClient()

Dim szSQL      As String
Dim adoConn As New ADODB.Connection
Dim adoRstCheck  As New ADODB.Recordset
Dim adoRstSrc  As New ADODB.Recordset
Dim adoRstDst  As New ADODB.Recordset

On Error GoTo ERR_HANDER
'Resolved by BOSL
'issue 532 Cannot copy default chart of account for new client
'Modified by anol 04 Deb 2015
adoConn.Open getConnectionString
'adoConn.BeginTrans
szSQL = "SELECT * FROM NOMINALLEDGER WHERE CLIENTID = '" & txtClientList.Tag & "'"
adoRstCheck.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly

If Not adoRstCheck.EOF Then
'    adoConn.RollbackTrans
    Exit Sub
End If
adoConn.Close
Set adoConn = Nothing
    MsgBox "'" & txtClientList.text & "' does not have a Chart of Accounts set up. " & Chr(13) & _
            "Please create a new Chart of Accounts for this client", vbInformation, "Wanrning!"
            
   cmdCopy_Click
' If MsgBox(txtClientList.text & " does not have a set of nominal codes." & Chr(13) & _
'               "Do you wish to copy default chart of accounts?", vbQuestion + vbYesNo, "Nominal Ledger Not Setup") = vbYes Then
         
'         iRow = flxNominalCode.Rows

        ' adoConn.BeginTrans

''         szSQL = "SELECT Code, Name, Type, DrCr " & _
''                 "FROM   NominalLedger " & _
''                 "WHERE  ClientID = 'NONE';"
''         adoRstSrc.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''         szSQL = "SELECT * " & _
''                 "FROM   NominalLedger;"
''         adoRstDst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
''         'adoConn.CursorLocation = adUseClient
''         'adoConn.BeginTrans
''         While Not adoRstSrc.EOF
''            adoRstDst.AddNew
''            adoRstDst.Fields.Item("Code").Value = adoRstSrc.Fields.Item("Code").Value
''            adoRstDst.Fields.Item("Name").Value = adoRstSrc.Fields.Item("Name").Value
''            adoRstDst.Fields.Item("Type").Value = adoRstSrc.Fields.Item("Type").Value
''            adoRstDst.Fields.Item("DrCr").Value = adoRstSrc.Fields.Item("DrCr").Value
''            adoRstDst.Fields.Item("Posting").Value = True
''            adoRstDst.Fields.Item("ClientID").Value = txtClientList.Tag
''            adoRstSrc.MoveNext
''            adoRstDst.Update
''         Wend
''
''         adoRstDst.Close
''         Set adoRstDst = Nothing
''         adoRstSrc.Close
''         Set adoRstSrc = Nothing

         
         'Below line is added by anol 04 Feb 2015
        
'      End If
'      adoConn.CommitTrans
      Exit Sub
ERR_HANDER:
   ShowMsgInTaskBar "There was a problem to create a set of nominal code", "Y", "N"
   'MsgBox ERR.description
'   adoConn.RollbackTrans

End Sub
'
Public Function LoadFirstFinancialYear() As Boolean
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer              'Open Flag index
   Dim adoRST     As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
'Desc added by anol 02 Nov 2015
   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status,setasCurrent " & _
           "FROM   FinancialYear " & _
           "WHERE  ClientID = '" & txtClientList.Tag & "' " & _
           "ORDER BY FY_StDate Desc;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount - 1
   TotalCol = adoRST.Fields.Count - 1
   ReDim Data(TotalCol, TotalRow) As String
   LoadFirstFinancialYear = True
   K = -1
   
   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
         If K = -1 And j = 4 Then
            'Code changes by anol 2020-10-19
            If adoRST.Fields("setasCurrent").Value Then 'If adoRst.Fields("Status").Value Then
               K = i
               dtStartPnL = CDate(adoRST.Fields("FY_StDate").Value)
               dtStartBS = CDate("01 January 2000")
               txtBudgetYears.text = adoRST.Fields("FinancialYear").Value
               txtBudgetYears.Tag = adoRST.Fields("FYrID").Value
               cmdBudgetYears.Tag = CDate(adoRST.Fields("FY_StDate").Value)
            End If
         End If
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   adoRST.MoveFirst
   If txtBudgetYears.text = "" Then
        If Not adoRST.EOF Then
               dtStartPnL = CDate(adoRST.Fields("FY_StDate").Value)
               dtStartBS = CDate("01 January 2000")
               txtBudgetYears.text = adoRST.Fields("FinancialYear").Value
               txtBudgetYears.Tag = adoRST.Fields("FYrID").Value
               cmdBudgetYears.Tag = CDate(adoRST.Fields("FY_StDate").Value)
        End If
        
   End If
'   cmbFinancialYear.Column() = Data()
'   cmbFinancialYear.ListIndex = k

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
   Exit Function

NoRes:
    Set adoRST = Nothing
    adoConn.Close
    Set adoConn = Nothing
    FocusControl cmdAddNew
    'MsgBox "No Financial years exist for this Client. Please create at least one Financial Year", vbInformation, "Warning"
    frmFinancialYearCreate.txtStDate.Locked = False
    frmFinancialYearCreate.lblClientName.Caption = txtClientList.text
    frmFinancialYearCreate.lblClientName.Tag = txtClientList.Tag
    frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Add New"
    frmFinancialYearCreate.FinancialYearID = UniqueID()
    frmFinancialYearCreate.CallingFrom = "NominalLedger"
    frmFinancialYearCreate.Show
    frmFinancialYearCreate.ZOrder 0
   
'    frmFinancialYearCreate.ZOrder 1
    txtBudgetYears.text = ""
    txtBudgetYears.Tag = ""
    cmdBudgetYears.Tag = ""
    iFinancialWarning = iFinancialWarning + 1
    If iFinancialWarning < 2 Then
        MsgBox "No Financial years exist for the Client : '" & txtClientList.text & "'. Please create at least one Financial Year", vbInformation, "Warning"
    End If
   If isFinancialYearCreated = True Then
        isFinancialYearCreated = False
        Call LoadFirstFinancialYear
       
        
        isFinancialYearCreated = False
        
    End If
   Exit Function
End Function

'Private Sub UpdateNominalBalance(adoConn As ADODB.Connection)
'   Dim iRow       As Integer
'   Dim cBal       As Currency
'
'   For iRow = 1 To flxNominalCode.Rows - 1
'      If flxNominalCode.RowHeight(iRow) = 240 And flxNominalCode.TextMatrix(iRow, 0) <> "" And flxNominalCode.TextMatrix(iRow, 0) <> cmbClient.Column(2) Then
''If UCase(SystemUser) = "SAMRAT" And flxNominalCode.TextMatrix(iRow, 0) = "1100" Then
''MsgBox ""
''End If
'         flxNominalCode.TextMatrix(iRow, 5) = ""
'         flxNominalCode.TextMatrix(iRow, 6) = ""
'         If flxNominalCode.TextMatrix(iRow, 2) = 1 Then        'Balance Sheet
'            cBal = CalcuateBalanceDr_BS(flxNominalCode.TextMatrix(iRow, 0), flxNominalCode.TextMatrix(iRow, 7), adoConn)
'            If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 5) = Format(cBal, "0.00")
'            cBal = CalcuateBalanceCr_BS(flxNominalCode.TextMatrix(iRow, 0), flxNominalCode.TextMatrix(iRow, 7), adoConn)
'            If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 6) = Format(cBal, "0.00")
'         End If
'         If flxNominalCode.TextMatrix(iRow, 2) = 2 Then        'Profit and Loss
'            cBal = CalcuateBalanceDr_PnL(flxNominalCode.TextMatrix(iRow, 0), flxNominalCode.TextMatrix(iRow, 7), adoConn)
'            If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 5) = Format(cBal, "0.00")
'            cBal = CalcuateBalanceCr_PnL(flxNominalCode.TextMatrix(iRow, 0), flxNominalCode.TextMatrix(iRow, 7), adoConn)
'            If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 6) = Format(cBal, "0.00")
'         End If
'
'         If flxNominalCode.TextMatrix(iRow, 5) <> "" And flxNominalCode.TextMatrix(iRow, 6) <> "" Then
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) > Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 5) = Format(Val(flxNominalCode.TextMatrix(iRow, 5)) - Val(flxNominalCode.TextMatrix(iRow, 6)), "0.00")
'               flxNominalCode.TextMatrix(iRow, 6) = ""
'            End If
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) < Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 6) = Format(Val(flxNominalCode.TextMatrix(iRow, 6)) - Val(flxNominalCode.TextMatrix(iRow, 5)), "0.00")
'               flxNominalCode.TextMatrix(iRow, 5) = ""
'            End If
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) = Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 5) = ""
'               flxNominalCode.TextMatrix(iRow, 6) = ""
'            End If
'         End If
'      End If
''  Retained Earnings
'      If flxNominalCode.RowHeight(iRow) = 240 And flxNominalCode.TextMatrix(iRow, 0) <> "" And flxNominalCode.TextMatrix(iRow, 0) = cmbClient.Column(2) Then
'         flxNominalCode.TextMatrix(iRow, 5) = ""
'         flxNominalCode.TextMatrix(iRow, 6) = ""
'         cBal = CalculateRetainedEarnings_Dr(flxNominalCode.TextMatrix(iRow, 7), adoConn)
'         If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 5) = Format(cBal, "0.00")
'         cBal = CalculateRetainedEarnings_Cr(flxNominalCode.TextMatrix(iRow, 7), adoConn)
'         If cBal <> 0 Then flxNominalCode.TextMatrix(iRow, 6) = Format(cBal, "0.00")
'
'         If flxNominalCode.TextMatrix(iRow, 5) <> "" And flxNominalCode.TextMatrix(iRow, 6) <> "" Then
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) > Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 5) = Format(Val(flxNominalCode.TextMatrix(iRow, 5)) - Val(flxNominalCode.TextMatrix(iRow, 6)), "0.00")
'               flxNominalCode.TextMatrix(iRow, 6) = ""
'               RE = flxNominalCode.TextMatrix(iRow, 5)
'            End If
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) < Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 6) = Format(Val(flxNominalCode.TextMatrix(iRow, 6)) - Val(flxNominalCode.TextMatrix(iRow, 5)), "0.00")
'               flxNominalCode.TextMatrix(iRow, 5) = ""
'               RE = (-1) * flxNominalCode.TextMatrix(iRow, 6)
'            End If
'            If Val(flxNominalCode.TextMatrix(iRow, 5)) = Val(flxNominalCode.TextMatrix(iRow, 6)) Then
'               flxNominalCode.TextMatrix(iRow, 5) = ""
'               flxNominalCode.TextMatrix(iRow, 6) = ""
'               RE = 0
'            End If
'         End If
'      End If
'   Next iRow
'End Sub
'
'Private Function CalculateRetainedEarnings_Cr(szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalculateRetainedEarnings_Cr = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N, NominalLedger AS L " & _
'            "WHERE N.NOMINAL_CODE = L.Code AND N.ClientID = '" & szClient & "' AND L.ClientID = '" & szClient & "' AND L.Type = 2 AND " & _
'               "N.POSTED_DATE <= #" & Format(DateAdd("d", -1, cmbFinancialYear.Column(3)), "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 7 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 16 OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 12) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalculateRetainedEarnings_Cr = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalculateRetainedEarnings_Cr = 0
'      Else
'         CalculateRetainedEarnings_Cr = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'
'Private Function CalculateRetainedEarnings_Dr(szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalculateRetainedEarnings_Dr = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N, NominalLedger AS L " & _
'            "WHERE N.NOMINAL_CODE = L.Code AND N.ClientID = '" & szClient & "' AND L.ClientID = '" & szClient & "' AND L.Type = 2 AND " & _
'               "N.POSTED_DATE <= #" & Format(DateAdd("d", -1, cmbFinancialYear.Column(3)), "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 2 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 15 OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 11) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalculateRetainedEarnings_Dr = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalculateRetainedEarnings_Dr = 0
'      Else
'         CalculateRetainedEarnings_Dr = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'
'Private Function CalcuateBalanceDr_BS(szCode As String, szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalcuateBalanceDr_BS = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE = '" & szCode & "' AND " & _
'               "N.POSTED_DATE >= #" & Format(dtStartBS, "dd mmmm yyyy") & "# AND N.POSTED_DATE <= #" & Format(dtEnd, "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 2 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 15 OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 11) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalcuateBalanceDr_BS = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceDr_BS = 0
'      Else
'         CalcuateBalanceDr_BS = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'
'   adoConn.Execute "UPDATE NominalLedger " & _
'                   "SET    Debit = " & CalcuateBalanceDr_BS & " " & _
'                   "WHERE  Code = '" & szCode & "' AND " & _
'                          "ClientID = '" & szClient & "';"
'End Function
'
'Private Function CalcuateBalanceCr_BS(szCode As String, szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalcuateBalanceCr_BS = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE = '" & szCode & "' AND " & _
'               "N.POSTED_DATE >= #" & Format(dtStartBS, "dd mmmm yyyy") & "# AND N.POSTED_DATE <= #" & Format(dtEnd, "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 7 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 16 OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 12) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalcuateBalanceCr_BS = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceCr_BS = 0
'      Else
'         CalcuateBalanceCr_BS = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'
'   adoConn.Execute "UPDATE NominalLedger " & _
'                   "SET    Credit = " & CalcuateBalanceCr_BS & " " & _
'                   "WHERE  Code = '" & szCode & "' AND " & _
'                          "ClientID = '" & szClient & "';"
'End Function
'
'Private Function CalcuateBalanceDr_PnL(szCode As String, szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalcuateBalanceDr_PnL = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE = '" & szCode & "' AND " & _
'               "N.POSTED_DATE >= #" & Format(dtStartPnL, "dd mmmm yyyy") & "# AND N.POSTED_DATE <= #" & Format(dtEnd, "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 2 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 15 OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 11) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalcuateBalanceDr_PnL = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceDr_PnL = 0
'      Else
'         CalcuateBalanceDr_PnL = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'
'   adoConn.Execute "UPDATE NominalLedger " & _
'                   "SET    Debit = " & CalcuateBalanceDr_PnL & " " & _
'                   "WHERE  Code = '" & szCode & "' AND " & _
'                          "ClientID = '" & szClient & "';"
'End Function
'
'Private Function CalcuateBalanceCr_PnL(szCode As String, szClient As String, adoConn As ADODB.Connection) As Currency
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim adoRst     As New ADODB.Recordset
'
'   CalcuateBalanceCr_PnL = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE = '" & szCode & "' AND " & _
'               "N.POSTED_DATE >= #" & Format(dtStartPnL, "dd mmmm yyyy") & "# AND N.POSTED_DATE <= #" & Format(dtEnd, "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 7 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 16 OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 12) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalcuateBalanceCr_PnL = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceCr_PnL = 0
'      Else
'         CalcuateBalanceCr_PnL = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'
'   adoConn.Execute "UPDATE NominalLedger " & _
'                   "SET    Credit = " & CalcuateBalanceCr_PnL & " " & _
'                   "WHERE  Code = '" & szCode & "' AND " & _
'                          "ClientID = '" & szClient & "';"
'End Function
'
'Public Sub RefreshGrid(adoConn As ADODB.Connection)
'   Dim iRow As Integer
'
'   LoadFlxNominalCode adoConn
'
'   For iRow = 1 To TotalNC
'      If flxNominalCode.TextMatrix(iRow, 7) = txtClientlist.tag Then
'         flxNominalCode.RowHeight(iRow) = 240
'      End If
'   Next
'End Sub

Public Sub LoadPeriods(adoConn As ADODB.Connection)
  
   Dim adoRST     As New ADODB.Recordset
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim szSQL      As String
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer                    'Open flag index

   If txtBudgetYears.text <> "" Then
      

      szSQL = "SELECT PeriodID, Period_Descp, P_StDate, P_EndDate, Status " & _
              "FROM   Periods " & _
              "WHERE  FYrID = '" & txtBudgetYears.Tag & "' " & _
              "ORDER BY P_StDate;"


      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      If adoRST.EOF Then GoTo NoRes

      TotalRow = adoRST.RecordCount - 1
      TotalCol = adoRST.Fields.Count - 1
      ReDim Data(TotalCol, TotalRow) As String

      K = -1
      For i = 0 To TotalRow
         For j = 0 To TotalCol
            Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
            If K = -1 And j = 4 Then
               If adoRST.Fields("Status").Value Then
                  K = i
                  dtEnd = CDate(adoRST.Fields("P_EndDate").Value)
               End If
            End If
         Next j
         adoRST.MoveNext
         If adoRST.EOF Then Exit For
      Next i

      cmbPeriodFrom.Column() = Data()
      cmbPeriodTo.Column() = Data()

      cmbPeriodFrom.ListIndex = 0
      If (cmbPeriodTo.ListCount > 0) Then
         cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If

      chkYtD_Click
   Else
       ' MsgBox "PLease set the budget to load periods", vbInformation, "Warning"
      
   End If
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Periods are not found. Please contact with system support", "Y", "N"
   
End Sub

Private Sub cmbPeriodFrom_Change()
'If cmbPeriodFrom.ListCount > 0 And cmbPeriodTo.ListCount > 0 Then
'    If cmbPeriodFrom.Column(2) > cmbPeriodTo.Column(2) Then
'        cmbPeriodTo.ListIndex = cmbPeriodFrom.ListIndex
'    End If
'End If
End Sub

Private Sub cmbPeriodTo_Change()
'If cmbPeriodFrom.ListCount > 0 And cmbPeriodTo.ListCount > 0 Then
'    If cmbPeriodFrom.Column(2) > cmbPeriodTo.Column(2) Then
'        cmbPeriodTo.ListIndex = cmbPeriodFrom.ListIndex
'    End If
'End If
End Sub

Private Sub SetCurrentFY(adoConn As ADODB.Connection)
On Error GoTo ErrorHandler


Dim adoRST As New ADODB.Recordset
Dim szSQL As String



If Not IsNull(txtPropertyName.Tag) And txtPropertyName.Tag <> "ALL" Then
   'cmbFinancialYear.Value = GetCurrentFYFromProperty(adoConn, txtPropertyName.Tag)
   'adoConn.Open getConnectionString
   'szSQL = "SELECT P.CBY FROM Property AS P WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
   szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID,F.FY_StDate " & _
           "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
           "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"

   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
   While Not adoRST.EOF
        If IsNull(adoRST.Fields.Item("FYrID").Value) = False Then
            txtBudgetYears.Tag = adoRST.Fields.Item("FYrID").Value
            txtBudgetYears.text = adoRST.Fields.Item("CBY").Value
            cmdBudgetYears.Tag = adoRST.Fields.Item("FY_StDate").Value
       End If
       adoRST.MoveNext
   Wend

   adoRST.Close
   Set adoRST = Nothing
   
End If
   Exit Sub
ErrorHandler:

   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Could not update the NLPosting Table"
   Set adoRST = Nothing
   


End Sub

Private Sub cmdAddNew_Click()
   If IsNull(txtClientList.Tag) Then
      ShowMsgInTaskBar "Please select a client to add new Nominal Code", "Y", "N"
      cmdClientList.SetFocus
       
      Exit Sub
   End If

   frmNLAmendment.AddNew = True
   Load frmNLAmendment
   frmNLAmendment.lblClient = txtClientList.text
'   frmNLAmendment.StartUpPosition = StartUpPositionConstants.vbStartUpScreen
   frmNLAmendment.Show
   
'   Me.Enabled = False
   
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    
    Call GenerateNominalAccounts(adoConn, "")
    
    adoConn.Close
    Set adoConn = Nothing
End Sub


Private Sub cmdAddnewDefaultCOA_Click()
   If IsNull(txtClientList.Tag) Then
        ShowMsgInTaskBar "Please select a client to add new Nominal Code", "Y", "N"
        cmdClientList.SetFocus
        Exit Sub
   End If

   frmNLAmendment.AddNew = True
   Load frmNLAmendment
   frmNLAmendment.lblClient = "NONE"
'   frmNLAmendment.StartUpPosition = StartUpPositionConstants.vbStartUpScreen
   frmNLAmendment.Show
   
'   Me.Enabled = False
   
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    
    Call GenerateNominalAccounts(adoConn, "")
    
        'MsgBox "Selected Nominal Code has been deleted", vbInformation, "Nominal code Deleted"
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdBankClose_Click()
    picBankCode.Visible = False
    Frame2.Enabled = True
End Sub

Private Sub cmdBudgetYears_Click()
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    sTextBox = "3"
    picBankCode.Left = txtBudgetYears.Left
    picBankCode.Top = txtBudgetYears.Top + 5
    Call LoadGridFY
    
    
    
End Sub

Private Sub cmdCancelWizard_Click()
    Frame1(1).Visible = False
    Frame1(2).Visible = False
End Sub

Private Sub cmdClose_Click()
   If iNewEdit <> 0 Then
      If MsgBox("Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Nominal Ledger") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String
   Dim K As Integer

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   
   configGridFY
   adoConn.Open getConnectionString
           

   szSQL = "SELECT F.FYrID, F.FY_StDate,FinancialYear,F.FY_Description,setascurrent " & _
           "FROM FinancialYear AS F " & _
           "WHERE F.ClientID = '" & txtClientList.Tag & "'  " & _
           "ORDER BY FY_EndDate DESC;"


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
        MsgBox "No Financial years exist for this Client. Please create at least one Financial Year", vbInformation, "Warning!"
        frmFinancialYearCreate.lblClientName.Caption = txtClientList.text
        frmFinancialYearCreate.lblClientName.Tag = txtClientList.Tag
        frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Add New"
        frmFinancialYearCreate.FinancialYearID = UniqueID()
        frmFinancialYearCreate.Show
   Else
        rRow = 1
        gridBankCode.Rows = rstRec.RecordCount + 1
        While Not rstRec.EOF
           gridBankCode.TextMatrix(rRow, 0) = ""
           gridBankCode.TextMatrix(rRow, 1) = Trim(rstRec.Fields.Item("FY_StDate").Value)
           gridBankCode.TextMatrix(rRow, 2) = "  " & Replace(Trim(rstRec.Fields.Item("FY_Description").Value), "-", "  -  ")
           gridBankCode.TextMatrix(rRow, 3) = Trim(rstRec.Fields.Item("FYrID").Value)
           gridBankCode.TextMatrix(rRow, 4) = Trim(rstRec.Fields.Item("setascurrent").Value)
           If gridBankCode.TextMatrix(rRow, 4) = True Then
                 gridBankCode.row = rRow
                 For K = 1 To 4
                        gridBankCode.col = K
                        gridBankCode.CellFontBold = True
                 Next
           Else
                 gridBankCode.row = rRow
                 For K = 1 To 4
                        gridBankCode.col = K
                        gridBankCode.CellFontBold = False
                 Next
           End If
           gridBankCode.RowHeight(rRow) = 240
           rstRec.MoveNext
           rRow = rRow + 1
        Wend
        gridBankCode.RowSel = 1
        picBankCode.Visible = True
        gridBankCode.SetFocus
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub configGridFY()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 5
   gridBankCode.TextMatrix(0, 0) = "Nominal Code"
   gridBankCode.TextMatrix(0, 1) = "Name"
   gridBankCode.ColWidth(0) = 100
   gridBankCode.ColWidth(1) = 0
   gridBankCode.ColAlignment(1) = vbLeftJustify
   gridBankCode.ColAlignment(2) = vbLeftJustify
   gridBankCode.ColWidth(2) = 2700
   gridBankCode.ColWidth(3) = 0
   gridBankCode.ColWidth(4) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "Financial Year"
   Label8.Caption = ""
   
End Sub

Private Sub cmdClose1_Click()
    Unload Me
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdConvertNetAmountToDRCR_Click()
Dim adoConn As New ADODB.Connection
adoConn.Open getConnectionString

If MsgBox("This will convert the figure of all the nominal journal lines to (+ve/-ve) format to represent DR/CR." & vbNewLine & vbNewLine & "The process may take upto several minutes. Do you really want to continue", vbYesNo, "Conversion to new Nominal Ledger Format") = vbYes Then
    If ConvertNLPostingAmountToDRCR(adoConn) = True Then
        MsgBox "All the nominal journal lines are successfully converted to DR/CR (+ve/-ve) format.", vbOKOnly, "Conversion to new Nominal Ledger Format"
    End If
End If

adoConn.Close
Set adoConn = Nothing

End Sub

Private Sub cmdCopy_Click()
    Frame1(1).Top = 1575
    Frame1(1).Left = 3510
    Frame1(2).Left = 3570
    Frame1(2).Top = 3015
    Frame1(1).Visible = True
    Frame1(2).Visible = False
    optCopyDefltCOA.Value = True
End Sub

Private Sub cmdDelete_Click()
   If IsNull(txtClientList.Tag) Then Exit Sub
   If flxNominalCode.TextMatrix(flxNominalCode.row, 0) = "" Then
        MsgBox "Please select a nominal code to delete", vbInformation, "Please select a nominal code"
        Exit Sub
   End If
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset
   Dim lnRecordcount As Long
   Dim lnRecordcount2 As Long
   
   adoConn.Open getConnectionString
    szSQL = "SELECT Code AS NC, 'NominalJournal' AS TN " & _
           "FROM NominalLedger " & _
           "WHERE CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' and CATYPE in ('I','O','R','S','P') AND clientID='" & txtClientList.Tag & "'"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount2 = adoRST.RecordCount
    adoRST.Close
    
    szSQL = "SELECT NOMINAL_CODE AS NC, 'NLPOSTING' AS TN " & _
           "FROM NLPOSTING " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' AND clientID='" & txtClientList.Tag & "' "
  ' szSQL = szSQL + "UNION "
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = adoRST.RecordCount
    adoRST.Close
   szSQL = "SELECT NominalCodeforAmount AS NC, 'DemandSplitRecords' AS TN " & _
           "FROM DemandSplitRecords " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   'szSQL = szSQL + "UNION "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCodeforAmount AS NC, 'DemandTypes' AS TN " & _
           "FROM DemandTypes D, Property P " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' AND P.PropertyID=D.PropertyID  and P.ClientID='" & txtClientList.Tag & "'"
   'szSQL = szSQL + "UNION "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If adoRST.RecordCount > 0 Then
        MsgBox "There are records in Demand Types using this nominal code.", vbInformation, "Warning"
   End If
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT GlobalBankCode AS NC, 'GlobalData' AS TN " & _
           "FROM GlobalData " & _
           "WHERE GlobalBankCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   'szSQL = szSQL + "UNION "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NC, 'GlobalSCDtls' AS TN " & _
           "FROM GlobalSCDtls " & _
           "WHERE NC = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT PayNCAmt AS NC, 'PayableTypes' AS TN " & _
           "FROM PayableTypes " & _
           "WHERE PayNCAmt = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'PayTransactions' AS TN " & _
           "FROM PayTransactions " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'RptTransactions' AS TN " & _
           "FROM RptTransactions " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'tblPoA' AS TN " & _
           "FROM tblPoA " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NOMINAL_CODE AS NC, 'tblPurInvSRec' AS TN " & _
           "FROM tblPurInvSRec " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'TenantDeposit' AS TN " & _
           "FROM TenantDeposit " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NOMINAL_CODE AS NC, 'tlbBankPayment' AS TN " & _
           "FROM tlbBankPayment " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCodeforAmount AS NC, 'tlbChildDemandRecord' AS TN " & _
           "FROM tlbChildDemandRecord " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'tlbClientBanks' AS TN " & _
           "FROM tlbClientBanks " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NOMINAL_CODE AS NC, 'tlbCreditNote' AS TN " & _
           "FROM tlbCreditNote " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'tlbPayment' AS TN " & _
           "FROM tlbPayment " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NominalCode AS NC, 'tlbReceipt' AS TN " & _
           "FROM tlbReceipt " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NOMINAL_CODE AS NC, 'tlbRecharged' AS TN " & _
           "FROM tlbRecharged " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close

   szSQL = "SELECT NOMINAL_CODE AS NC, 'tlbRechargePre' AS TN " & _
           "FROM tlbRechargePre " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "';"
'Debug.Print szSQL
  adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    lnRecordcount = lnRecordcount + adoRST.RecordCount
    adoRST.Close
    If lnRecordcount2 > 0 Then
      MsgBox "This nominal code has been set as a control account and cannot be deleted.", vbCritical + vbOKOnly, "Nominal Code : " & flxNominalCode.TextMatrix(flxNominalCode.row, 0)
      Exit Sub
   End If

   If lnRecordcount > 0 Then
      MsgBox "This nominal code has transactions posted against it and cannot be deleted.", vbCritical + vbOKOnly, "Nominal Code : " & flxNominalCode.TextMatrix(flxNominalCode.row, 0)
   Else
      If MsgBox("Do you wish to delete the nominal code?", vbQuestion + vbYesNo, "Nominal Code") = vbYes Then
         szSQL = "DELETE * " & _
                 "FROM NominalLedger " & _
                 "WHERE ClientID = '" & txtClientList.Tag & "' AND " & _
                       "Code = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "';"
         adoConn.Execute szSQL

         ShowMsgInTaskBar "Nominal Code " & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & " has been removed."

         flxNominalCode.RemoveItem flxNominalCode.row
         TotalNC = TotalNC - 1
      End If
   End If
   
'   adoRst.Close
'   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdDeleteDefaultCOA_Click()
        If IsNull(txtClientList.Tag) Then Exit Sub
        If flxChartOfACCDefault.TextMatrix(flxChartOfACCDefault.row, 0) = "" Then
             MsgBox "Please select a nominal code to delete", vbInformation, "Please select a nominal code"
             Exit Sub
        End If
        Dim adoConn As New ADODB.Connection
        Dim szSQL As String
        
        adoConn.Open getConnectionString
        szSQL = "Delete from  NominalLedger " & _
                "WHERE CODE = '" & flxChartOfACCDefault.TextMatrix(flxChartOfACCDefault.row, 0) & "' and clientID='NONE'"
        adoConn.Execute szSQL
        adoConn.Close
        Call LoadDefaultChartofAccounts
        MsgBox "Selected Nominal Code has been deleted", vbInformation, "Nominal code Deleted"
End Sub

Private Sub cmdEdit2_Click()
'       If IsNull(txtClientList.Tag) Then
'        ShowMsgInTaskBar "Please select a client to add new Nominal Code", "Y", "N"
'        cmdClientList.SetFocus
'        Exit Sub
'   End If
'
'   frmNLAmendment.AddNew = False
''   Load frmNLAmendment
''   frmNLAmendment.lblClient = "NONE"
''   frmNLAmendment.StartUpPosition = StartUpPositionConstants.vbStartUpScreen
''   frmNLAmendment.Show
'    LoadForm frmNLAmendment
'
''   Me.Enabled = False
'
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'
'    Call GenerateNominalAccounts(adoConn, "")
'
'        'MsgBox "Selected Nominal Code has been deleted", vbInformation, "Nominal code Deleted"
'    adoConn.Close
'    Set adoConn = Nothing
    If flxNominalCode.row = 0 Then
        ShowMsgInTaskBar "Please select a Nominal code from the grid.", , "N"
        Exit Sub
    End If
    
    If flxNominalCode.TextMatrix(flxNominalCode.row, 0) = "" Or _
            flxNominalCode.RowHeight(flxNominalCode.row) = 0 Then Exit Sub
    Load frmNLAmendment
    
    With flxChartOfACCDefault
          frmNLAmendment.lblClient.Caption = "NONE"
          frmNLAmendment.lblClient.Tag = "NONE"
          frmNLAmendment.txtCode.text = .TextMatrix(.row, 0)
          frmNLAmendment.txtName.text = .TextMatrix(.row, 1)
          frmNLAmendment.cboType.Value = .TextMatrix(.row, 2)
          frmNLAmendment.cmbPosting.ListIndex = IIf(.TextMatrix(.row, 8) = "YES", 0, 1)
          If frmNLAmendment.cmbPosting.ListIndex = 1 Then
                frmNLAmendment.cmbPosting.Locked = True
          End If
          frmNLAmendment.cboDrCr.Value = .TextMatrix(.row, 7)
          frmNLAmendment.chkYtD.Value = chkYtD.Value
        'adoRst.Fields.Item("DrCr").Value = cboDrCr.Value
          If .TextMatrix(.row, 4) = "" Then 'this is different from normal
             frmNLAmendment.cboSubType.ListIndex = -1
          Else
             frmNLAmendment.cboSubType.text = .TextMatrix(.row, 4)
          End If
    End With
    frmNLAmendment.AddNew = False
    '   frmNLAmendment.Show
    LoadForm frmNLAmendment
    frmNLAmendment.txtName.SetFocus
    frmNLAmendment.cmdSave.Enabled = True
    Me.Enabled = False
End Sub

Public Sub cmdFilter_Click()
'Below line added by anol 08 Apr 2015
'If you double click on display button it says user locked out. it raise double call of this procedure.So I am disabling this button and enabling while procedure call,
'so that user will not have the scope for double click
    cmdFilter.Enabled = False
    ConfigFlxNominalCode
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    
    Call ExportData2NominalLedger(adoConn)
    Call GenerateNominalAccounts(adoConn, "")
    
    adoConn.Close
    Set adoConn = Nothing
    If iFinancialWarning = 0 Then
        cmdFilter.Enabled = True
        FocusControl cmdFilter
    End If
'   If IsNull(txtClientlist.tag) Then
'      ShowMsgInTaskBar "Please select a client", "Y", "N"
'      cmdClientList.SetFocus
'      Exit Sub
'   End If
'   If IsNull(cmbFinancialYear.Value) Then
'      ShowMsgInTaskBar "Please select the financial year", "Y", "N"
'      cmbFinancialYear.SetFocus
'      Exit Sub
'   End If
'   If IsNull(cmbPeriodFrom.Value) Then
'      ShowMsgInTaskBar "Please select the period from date", "Y", "N"
'      cmbPeriodFrom.SetFocus
'      Exit Sub
'   End If
'   If chkYtD.Value = 0 And IsNull(cmbPeriodTo.Value) Then
'      ShowMsgInTaskBar "Please select the period to date", "Y", "N"
'      cmbPeriodTo.SetFocus
'      Exit Sub
'   End If
'   If chkYtD.Value = 0 Then
'      If CDate(cmbPeriodFrom.Column(2)) > CDate(cmbPeriodTo.Column(2)) Then
'         ShowMsgInTaskBar "From date cannot be after To date", "Y", "N"
'         cmbPeriodFrom.SetFocus
'         Exit Sub
'      End If
'   End If
'
'   If chkYtD.Value = 1 Then
'      dtStartPnL = cmbFinancialYear.Column(3)            'Beginning of the Financial year
'      dtStartBS = CDate("01 January 2000")               'Beginning of the System
'      dtEnd = cmbPeriodFrom.Column(3)
'   Else
'      dtStartPnL = cmbPeriodFrom.Column(2)
'      dtStartBS = cmbPeriodFrom.Column(2)
'      dtEnd = cmbPeriodTo.Column(3)
'   End If
'
'   Dim adoConn    As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   Call ExportData2NominalLedger(adoConn)
'   UpdateNominalBalance adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing

End Sub



Private Sub cmdFinish_Click()
    Dim adoConn As New ADODB.Connection
    Dim iRow    As Integer
    Dim K       As Integer
    Dim rCount As Integer
    Dim selRow As Integer
    
    adoConn.Open getConnectionString
    If optCopyDefltCOA.Value = True Then
        If NominalLedgerCreationFromDedault(adoConn) = True Then
            MsgBox "Default Chart of Accounts successfully copied.", vbInformation, "Copy successful"
        End If
     End If
     If optCopyDemandTemplate.Value = True Then
           K = 0
           For iRow = 1 To flxClientCopy.Rows - 1
              If flxClientCopy.TextMatrix(iRow, 0) = "X" Then
                 K = K + 1
                 flxClientCopy.row = iRow
              End If
           Next iRow
           If K = 0 Then
              MsgBox "Please select a Client", vbInformation, "Information"
              FocusControl flxClientCopy
              Exit Sub
           End If
           If K > 1 Then
              MsgBox "Please select only one Client", vbInformation, "Information"
              FocusControl flxClientCopy
              Exit Sub
           End If
           For rCount = 1 To flxClientCopy.Rows - 1
                 If flxClientCopy.TextMatrix(rCount, 0) = "X" Then
                     selRow = rCount
                 End If
           Next
           If NominalLedgerCreationFromClient(adoConn, flxClientCopy.TextMatrix(selRow, 1)) = True Then
                MsgBox "Copy Chart of Accounts from client '" & flxClientCopy.TextMatrix(selRow, 1) & "' successful.", vbInformation, "Copy successful"
                Frame1(1).Visible = False
                Frame1(2).Visible = False
                cmdFilter.Enabled = True
                
            End If
     End If
     Call RefreshNominalList(adoConn, "")
'     iFinancialWarning = 0
     adoConn.Close
     Set adoConn = Nothing
     Frame1(1).Visible = False
     Frame1(2).Visible = False
     cmdFilter.Enabled = True
End Sub

Private Sub cmdPrint_Click()

    Dim ClientID As String
    If IsNull(txtClientList.Tag) Then
       ShowMsgInTaskBar "Please select a client", "Y", "N"
       cmdClientList.SetFocus
       Exit Sub
    End If
    If txtBudgetYears.text = "" Then
       ShowMsgInTaskBar "Please select the financial year", "Y", "N"
       cmdBudgetYears.SetFocus
       Exit Sub
    End If
    If chkYtD.Value = 0 And IsNull(cmbPeriodFrom.Value) Then
       ShowMsgInTaskBar "Please select the period from date", "Y", "N"
       cmbPeriodFrom.SetFocus
       Exit Sub
    End If
    If IsNull(cmbPeriodTo.Value) Then
       ShowMsgInTaskBar "Please select the period to date", "Y", "N"
       cmbPeriodTo.SetFocus
       Exit Sub
    End If
    
    If chkYtD.Value = 0 Then
       If CDate(cmbPeriodFrom.Column(2)) > CDate(cmbPeriodTo.Column(2)) Then
          ShowMsgInTaskBar "From date cannot be after To date", "Y", "N"
          cmbPeriodFrom.SetFocus
          Exit Sub
       End If
    End If
    

   Dim periodFrom As String
   Dim periodTo As String

    If chkYtD.Value = 1 Then
       
       dtStartPnL = cmdBudgetYears.Tag 'cmbFinancialYear.Column(3)            'Beginning of the Financial year(FY_StDate)
       dtStartBS = CDate("01 January 2000")               'Beginning of the System
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
       
       
    Else
       dtStartPnL = cmbPeriodFrom.Column(2)
       dtStartBS = cmbPeriodFrom.Column(2)
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
    End If


Dim adoConn As New ADODB.Connection
Dim adoRST As New ADODB.Recordset
cmdPrint.Enabled = False
adoConn.Open getConnectionString

On Error GoTo CreateReportTrialBalance

   adoRST.Open "SELECT * FROM ReportTrialBalance;", adoConn, adOpenStatic, adLockReadOnly
   adoRST.Close

   GoTo LoadTBeport

CreateReportTrialBalance:
   adoConn.Execute _
      "CREATE TABLE ReportTrialBalance " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "ClientID      TEXT(10), " & _
            "SubHeader     TEXT(200), " & _
            "NominalCode   TEXT(15) NOT NULL, " & _
            "Name          TEXT(200), " & _
            "NominalType   TEXT(50), " & _
            "Balance       CURRENCY, " & _
            "Debit         CURRENCY, " & _
            "Credit        CURRENCY, " & _
            "PRIMARY KEY (ReportingDate, SessionID, NominalCode)" & _
         ");"
         
LoadTBeport:
    
    Dim szSQL As String

On Error GoTo ErrorHandler:
    
'    szSQL = GetTrialBalanceQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
'
'    adoConn.Execute _
'      "DELETE FROM ReportTrialBalance WHERE SessionID = '" & sessionID & "';"
'
'    adoConn.Execute _
'      "INSERT INTO ReportTrialBalance " & _
'      "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
'        szSQL
     
   
'    adoRst.Open "SELECT * FROM ReportTrialBalance;", adoConn, adOpenStatic, adLockReadOnly
'    adoRst.Close
'    adoConn.Close
         adoConn.Execute "DELETE FROM ReportTrialBalance WHERE SessionID = '" & sessionID & "';"
        'added by anol 20161025
        adoConn.Execute "DELETE FROM ReportTrialBalance WHERE ReportingDate < #" & reportingDate & "# ;"
    
    
    If chkYtD.Value = 1 Then
        'Geting the balance sheet Items
        szSQL = GetTrialBalanceQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
        
       
        
        adoConn.Execute "INSERT INTO ReportTrialBalance " & _
        "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
        szSQL
        'added by anol 20170201, issue 296 trial balance is not correct
        'Getting RETAINED EARNINGS BEFORE THE FINANCIAL YEAR
        szSQL = GetTrialBalanceQuery3(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
        adoConn.Execute "INSERT INTO ReportTrialBalance " & _
        "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
        szSQL
        'added by anol 20161025
        'Geting the Profit and loss  Items
        szSQL = GetTrialBalanceQuery2(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
        adoConn.Execute "INSERT INTO ReportTrialBalance " & _
        "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
        szSQL
        
        
    Else
         szSQL = GetTrialBalanceQuery4(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
        adoConn.Execute "INSERT INTO ReportTrialBalance " & _
        "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
        szSQL
    End If
'
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TrialBalanceNew.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
'
   Report.ParameterFields(1).AddCurrentValue sessionID
   Report.ParameterFields(2).AddCurrentValue txtClientList.text
   Report.ParameterFields(3).AddCurrentValue txtPropertyName.text
   Report.ParameterFields(4).AddCurrentValue txtFundName.text
   
   Report.ParameterFields(5).AddCurrentValue Format(dtStartPnL, "dd/mm/yyyy")
   Report.ParameterFields(6).AddCurrentValue Format(dtEnd, "dd/mm/yyyy")
   
   'Report.ParameterFields(4).AddCurrentValue cboClientID.Column(1)
   
'   Report.ParameterFields(6).AddCurrentValue txtFundName.text

   Load frmReport
   frmReport.LoadReportViewer Report
   cmdPrint.Enabled = True
   FocusControl cmdPrint
    Exit Sub
    
ErrorHandler:

    MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Could not load Trial Balance"
    Set adoRST = Nothing
    
End Sub

Private Sub cmdRefresh_Click()
'   Me.MousePointer = vbHourglass
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   Call ExportData2NominalLedger(adoConn)
'   UpdateNominalBalance adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'   Me.MousePointer = vbArrow
End Sub

Private Sub Command1_Click()
    'GenerateNominalAccounts
End Sub

Private Sub cmdSearch_Click()
    fraSearch.Left = 2430
'    fraSearch.Visible = True
    If cmdSearch.Caption = "Clear Sea&rch" Then
         'sMode = "NO"
         txtSearchNo.text = ""
         txtSearchRef.text = ""
         cmdSearch.Caption = "Sea&rch"
         fraSearch.Visible = False
    Else
        If fraSearch.Visible = False Then
            fraSearch.Visible = True
            FocusControl txtSearchNo
        Else
            fraSearch.Visible = False
        End If
    End If
End Sub

Private Sub cmdSearchCancel_Click()
'        ConfigFlxNominalCode
'        Dim adoConn As New ADODB.Connection
'        adoConn.Open getConnectionString
'        Call GenerateNominalAccounts(adoConn, "")
'        adoConn.Close
'        Set adoConn = Nothing

    If cmdSearch.Caption = "Clear Sea&rch" Then
         txtSearchNo.text = ""
         txtSearchRef.text = ""
         cmdSearch.Caption = "Sea&rch"
         fraSearch.Visible = False
   End If
    
        fraSearch.Visible = False
End Sub

Private Sub cmdSearchOK_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdShowAllclient_Click()
   Dim i As Integer
   For i = flxClientCopy.Rows - 1 To 1 Step -1
        flxClientCopy.RowHeight(i) = 240
        flxClientCopy.row = i
   Next i
End Sub

Private Sub Command2_Click()
    Frame1(1).Visible = False
    Frame1(2).Visible = False
End Sub

Private Sub flxClientCopy_Click()
     SelectOnly1RowFlxGrid flxClientCopy, flxClientCopy.row
End Sub

Private Sub flxNominalCode_DblClick()
    Dim adoConn As New ADODB.Connection
    Dim rsNominalLedger As New ADODB.Recordset
    If flxNominalCode.row = 0 Then
        ShowMsgInTaskBar "Please select a Nominal code from the grid.", , "N"
        Exit Sub
    End If
    
    If flxNominalCode.TextMatrix(flxNominalCode.row, 0) = "" Or _
            flxNominalCode.RowHeight(flxNominalCode.row) = 0 Then Exit Sub
    Load frmNLAmendment
    
    With flxNominalCode
          frmNLAmendment.lblClient.Caption = txtClientList.text
          frmNLAmendment.lblClient.Tag = txtClientList.Tag
          frmNLAmendment.txtCode.text = .TextMatrix(.row, 0)
          frmNLAmendment.txtName.text = .TextMatrix(.row, 1)
          frmNLAmendment.cboType.Value = .TextMatrix(.row, 2)
          frmNLAmendment.cmbPosting.ListIndex = IIf(.TextMatrix(.row, 8) = "YES", 0, 1)
'          If frmNLAmendment.cmbPosting.ListIndex = 1 Then
'                frmNLAmendment.cmbPosting.Locked = True
'          End If
'          adoConn.Open getConnectionString
'          rsNominalLedger.Open "Select * from NominalLedger where ClientID='" & txtClientList.Tag & "' AND Code='" & .TextMatrix(.row, 0) & "'", adoConn, adOpenStatic, adLockReadOnly
'          If Not rsNominalLedger.EOF Then
''                If rsNominalLedger("POSTING").Value=0 AND  Then
''                End If
'          End If
'          rsNominalLedger.Close
'          adoConn.Close
          frmNLAmendment.cboDrCr.Value = .TextMatrix(.row, 9)
          frmNLAmendment.chkYtD.Value = chkYtD.Value
        
          If .TextMatrix(.row, 10) = "" Then
             frmNLAmendment.cboSubType.ListIndex = -1
          Else
             frmNLAmendment.cboSubType.Value = .TextMatrix(.row, 10)
          End If
    End With
    frmNLAmendment.AddNew = False
    '   frmNLAmendment.Show
    LoadForm frmNLAmendment
    frmNLAmendment.txtName.SetFocus
    
    Me.Enabled = False
   

End Sub

Private Sub flxNominalCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxNominalCode.ToolTipText = flxNominalCode.TextMatrix(flxNominalCode.MouseRow, flxNominalCode.MouseCol)
End Sub

'Resolved by BOSL
'Issue No: 0000476
'Modified By: Asif. 10 Oct 2014

Private Sub Form_Activate()
   MiLoading = False
  
   Dim adoConn As New ADODB.Connection
  
  
        
        If Form_Activated = False Then
            adoConn.Open getConnectionString
            LoadFirstFinancialYear
            LoadPeriods adoConn
            If txtBudgetYears.text <> "" Then
                cmdFilter_Click
            End If
            adoConn.Close
        End If
        
   
   
    Form_Activated = True
End Sub

Public Sub LoadDefaultChartofAccounts()
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim i As Integer
    Dim j As Integer
    Dim adoRST As New ADODB.Recordset
    adoConn.Open getConnectionString
    szSQL = "SELECT NOMINALLEDGER.CODE, NOMINALLEDGER.NAME, NOMINALLEDGER.TYPE AS NLTYPECODE, " & _
    "(SELECT NLTYPE.TYPEVALUE FROM NLTYPE WHERE NLTYPE.NLTYPECODE = NOMINALLEDGER.TYPE) AS TYPEVALUE, " & _
    "(SELECT NLSUBTYPES.STNAME FROM NLSUBTYPES WHERE NLSUBTYPES.STCODE = NOMINALLEDGER.SUBTYPE) AS STNAME, " & _
    "'Default CLIENT' as ClientID, " & _
    "IIF(NOMINALLEDGER.Posting, 'YES', 'NO') AS POSTING, NOMINALLEDGER.DRCR, NOMINALLEDGER.SubType " & _
    "From " & _
    "NOMINALLEDGER where clientID='NONE' "
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    Debug.Print szSQL
    flxChartOfACCDefault.Rows = adoRST.RecordCount + 1
    For i = 0 To adoRST.RecordCount - 1
           For j = 0 To adoRST.Fields.Count - 1
                 flxChartOfACCDefault.TextMatrix(i + 1, j) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
          Next j
      adoRST.MoveNext
    Next i
    adoRST.Close
    adoConn.Close
    Set adoConn = Nothing

End Sub
Private Sub Form_Load()
   Me.Height = 9270
   Me.Width = 13650
   SSTab1.Tab = 0
'   'cascade form function created by anol 2019 -06-17
'    Dim iLeft As Integer
'    Dim iTop As Integer
'    Call BuildFormlist(Me.Name, iTop, iLeft)
'    frmBPPreForm.Top = iTop
'    frmBPPreForm.Left = iLeft
'    'Cascade Form End
' frmMMain.Arrange vbCascade
'   Me.Top = 0
'   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   SSTab1.BackColor = Me.BackColor
   Frame2.BackColor = MODULEBACKCOLOR
   MiLoading = True
   flxNominalCode.Sort = 1
   Call ConfigflxChartOfACCDefault
   Call LoadDefaultChartofAccounts
   
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset
   ConfigFlxNominalCode
   
   ControlHanlding DefaultMode

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  


    If Not adoRST.EOF Then
                iFinancialWarning = 0
                txtClientList.Tag = adoRST.Fields("CLIENTID").Value
                txtClientList.text = adoRST.Fields("CLIENTNAME").Value
                        adoRST.Close
'                        szSQL = "SELECT PropertyID, PropertyName, " & _
'                             "ProAddressLine1, ProPostCode " & _
'                            "FROM Property " & _
'                           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                           "ORDER BY PropertyID;"
'                         adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                    If Not adoRst.EOF Then
'                        txtPropertyName.text = IIf(IsNull(adoRst.Fields("PropertyName").Value), "", adoRst.Fields("PropertyName").Value)
'                        txtPropertyName.Tag = IIf(IsNull(adoRst.Fields("PropertyID").Value), "", adoRst.Fields("PropertyID").Value)
'                    Else
'                        txtPropertyName.text = ""
'                        txtPropertyName.Tag = ""
'                    End If
                    
   End If
   
   
'   If LoadCmbClient(adoConn, cmbClient) Then
''      If Not LoadFlxNominalCode(adoConn) Then
''         adoConn.Close
''         Set adoConn = Nothing
''
''         MiLoading = False
''         Exit Sub
''      End If
'   End If
   
   'LoadCmbFunds adoConn, cmbFund
   
   'Added by BOSL. Issue: 0000476.
   'Added by Asif. 20 Dec 2014
   sessionID = GetTimeStamp
'   MsgBox sessionID
   reportingDate = Format(DateValue(Now), "dd mmmm yyyy")

   adoConn.Close
   Set adoConn = Nothing

   chkYtD.Value = 1
   'If cmbClient.ListCount > 0 And cmbClient.ListIndex < 0 Then cmbClient.ListIndex = 0
   'If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
        Call WheelHook(Me.hWnd)
   'End If
End Sub
Private Function GenerateNominalAccounts(adoConn As ADODB.Connection, Filter As String)
   
   Dim szSQL As String
   Dim ClientID As String
   Dim tempstr As String
   
    If IsNull(txtClientList.Tag) Then
       MsgBox "Please select a client", vbInformation, "Warning"
       cmdClientList.SetFocus
       Exit Function
    End If
    If iFinancialWarning > 0 Then Exit Function
    If txtBudgetYears.text = "" Then
       MsgBox "Please select the financial year", vbInformation, "Warning"
       cmdBudgetYears_Click
       'cmdBudgetYears.SetFocus
       Exit Function
    End If
   
    If chkYtD.Value = 0 And IsNull(cmbPeriodFrom.Value) Then
       MsgBox "Please select the period from date", vbInformation, "Warning"
       cmbPeriodFrom.SetFocus
       Exit Function
    End If
    If IsNull(cmbPeriodTo.Value) Then
       MsgBox "Please select the period to date", vbInformation, "Warning"
       cmbPeriodTo.SetFocus
       Exit Function
    End If
    
    If chkYtD.Value = 0 Then
       If CDate(cmbPeriodFrom.Column(2)) > CDate(cmbPeriodTo.Column(2)) Then
          MsgBox "From date cannot be after To date", vbInformation, "Warning"
          cmbPeriodFrom.SetFocus
          Exit Function
       End If
    End If
    
   
   Dim periodFrom As String
   Dim periodTo As String
   
    If chkYtD.Value = 1 Then
       
       dtStartPnL = cmdBudgetYears.Tag            'Beginning of the Financial year
       dtStartBS = CDate("01 January 2000")               'Beginning of the System
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
       
       
    Else
       dtStartPnL = cmbPeriodFrom.Column(2)
       dtStartBS = cmbPeriodFrom.Column(2)
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
    End If
   
    ' Check if Nominal Ledger is setup for the selected client:
    NominalLedgerSetupForNewClient


'   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
'                   x    x        0           x         x         x              x          0        x      0     0
   
   If chkYtD.Value = 1 Then
        szSQL = GetNominalBalancesQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, True)
   Else
        szSQL = GetNominalBalancesQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, False)
   End If

'   Debug.Print szSQL
   
   Dim adoRST As New ADODB.Recordset

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
   If Filter = "1" And Len(Trim(txtSearchNo.text)) > 0 Then
       tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
       adoRST.Filter = "Code Like '%" & tempstr & "%' "
   End If
   If Filter = "2" And Len(Trim(txtSearchRef.text)) Then
       tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
       adoRST.Filter = "Name Like '%" & tempstr & "%' "
   End If
   
   If adoRST.RecordCount = 0 Then
        flxNominalCode.Rows = 2
   Else
        flxNominalCode.Rows = adoRST.RecordCount + 1
   End If
'   flxNominalCode.RowHeight(1) = RowHeight
   
   If adoRST.EOF Then
       adoRST.Close
       Set adoRST = Nothing
       Exit Function
   End If

   Dim i As Integer, j As Integer
   Dim debitTotal As Double, creditTotal As Double
   Dim currentPnL As Double
   
   debitTotal = 0
   creditTotal = 0
   
   currentPnL = 0
      
   Dim retainedEarningsControl As String
   
   retainedEarningsControl = GetNominalCodeForControlAccount(adoConn, "Retained Earnings", txtClientList.Tag)
   
   For i = 0 To adoRST.RecordCount - 1
   
      For j = 0 To adoRST.Fields.Count - 1
         flxNominalCode.TextMatrix(i + 1, j) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
         
         If UCase(adoRST.Fields(j).Name) = "DEBIT" Then
            debitTotal = debitTotal + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            
            If adoRST.Fields(0) <> retainedEarningsControl And adoRST.Fields(2) = "2" Then
                currentPnL = currentPnL + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            End If
         
         ElseIf UCase(adoRST.Fields(j).Name) = "CREDIT" Then
            
            creditTotal = creditTotal + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            
             If adoRST.Fields(0) <> retainedEarningsControl And adoRST.Fields(2) = "2" Then
                currentPnL = currentPnL - IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            End If
         End If
         
         
         
      Next j
      adoRST.MoveNext
'      If Not adoRst.EOF Then flxNominalCode.AddItem ""
   Next i
   If Filter = "" Then
        lblDebitTotal = Format(debitTotal, "#,##0.00")
        lblCreditTotal = Format(creditTotal, "#,##0.00")
        
        'Multiplied by -1 to show the debit figure (expenditure) as negative
        lblCurrentPnL = Format(currentPnL * -1, "#,##0.00")
        
        If currentPnL * -1 > 0 Then
             lblCurrentPnL.ForeColor = vbBlue
        Else
            lblCurrentPnL.ForeColor = vbRed
        End If
   Else
        lblDebitTotal = Format(0, "#,##0.00")
        lblCreditTotal = Format(0, "#,##0.00")
   End If
   adoRST.Close
   Set adoRST = Nothing
   'Formating added by anol 20160522
    For i = 0 To flxNominalCode.Rows - 1
         flxNominalCode.TextMatrix(i, 5) = Format(flxNominalCode.TextMatrix(i, 5), "#,##0.00")
         flxNominalCode.TextMatrix(i, 6) = Format(flxNominalCode.TextMatrix(i, 6), "#,##0.00")
     Next i
   Exit Function

NewNominalCode:
   Set adoRST = Nothing
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Generating Nominal Balances"
   
End Function
Private Function RefreshNominalList(adoConn As ADODB.Connection, Filter As String)
      Dim szSQL As String
   Dim ClientID As String
  
   Dim tempstr As String
    If txtBudgetYears.text = "" Then
       MsgBox "Please select the financial year", vbInformation, "Warning"
       cmdBudgetYears_Click
       'cmdBudgetYears.SetFocus
       Exit Function
    End If
   
    If chkYtD.Value = 0 And IsNull(cmbPeriodFrom.Value) Then
       MsgBox "Please select the period from date", vbInformation, "Warning"
       cmbPeriodFrom.SetFocus
       Exit Function
    End If
    If IsNull(cmbPeriodTo.Value) Then
       MsgBox "Please select the period to date", vbInformation, "Warning"
       cmbPeriodTo.SetFocus
       Exit Function
    End If
    
    If chkYtD.Value = 0 Then
       If CDate(cmbPeriodFrom.Column(2)) > CDate(cmbPeriodTo.Column(2)) Then
          MsgBox "From date cannot be after To date", vbInformation, "Warning"
          cmbPeriodFrom.SetFocus
          Exit Function
       End If
    End If
    
   
   Dim periodFrom As String
   Dim periodTo As String
   
    If chkYtD.Value = 1 Then
       
       dtStartPnL = cmdBudgetYears.Tag            'Beginning of the Financial year
       dtStartBS = CDate("01 January 2000")               'Beginning of the System
       If IsNull(cmbPeriodTo.Column(3)) Then
       End If
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
       
       
    Else
       dtStartPnL = cmbPeriodFrom.Column(2)
       dtStartBS = cmbPeriodFrom.Column(2)
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
    End If
   
    ' Check if Nominal Ledger is setup for the selected client:
    NominalLedgerSetupForNewClient


'   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
'                   x    x        0           x         x         x              x          0        x      0     0
   
   If chkYtD.Value = 1 Then
        szSQL = GetNominalBalancesQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, True)
   Else
        szSQL = GetNominalBalancesQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, False)
   End If

'   Debug.Print szSQL
   
   Dim adoRST As New ADODB.Recordset

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
   If Filter = "1" And Len(Trim(txtSearchNo.text)) > 0 Then
       tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
       adoRST.Filter = "Code Like '%" & tempstr & "%' "
   End If
   If Filter = "2" And Len(Trim(txtSearchRef.text)) Then
       tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
       adoRST.Filter = "Name Like '%" & tempstr & "%' "
   End If
   
   If adoRST.RecordCount = 0 Then
        flxNominalCode.Rows = 2
   Else
        flxNominalCode.Rows = adoRST.RecordCount + 1
   End If
'   flxNominalCode.RowHeight(1) = RowHeight
   
   If adoRST.EOF Then
       adoRST.Close
       Set adoRST = Nothing
       Exit Function
   End If

   Dim i As Integer, j As Integer
   Dim debitTotal As Double, creditTotal As Double
   Dim currentPnL As Double
   
   debitTotal = 0
   creditTotal = 0
   
   currentPnL = 0
      
   Dim retainedEarningsControl As String
   
   retainedEarningsControl = GetNominalCodeForControlAccount(adoConn, "Retained Earnings", txtClientList.Tag)
   
   For i = 0 To adoRST.RecordCount - 1
   
      For j = 0 To adoRST.Fields.Count - 1
         flxNominalCode.TextMatrix(i + 1, j) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
         
         If UCase(adoRST.Fields(j).Name) = "DEBIT" Then
            debitTotal = debitTotal + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            
            If adoRST.Fields(0) <> retainedEarningsControl And adoRST.Fields(2) = "2" Then
                currentPnL = currentPnL + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            End If
         
         ElseIf UCase(adoRST.Fields(j).Name) = "CREDIT" Then
            
            creditTotal = creditTotal + IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            
             If adoRST.Fields(0) <> retainedEarningsControl And adoRST.Fields(2) = "2" Then
                currentPnL = currentPnL - IIf(IsNull(adoRST.Fields(j)), 0, adoRST.Fields(j))
            End If
         End If
         
         
         
      Next j
      adoRST.MoveNext
'      If Not adoRst.EOF Then flxNominalCode.AddItem ""
   Next i
   If Filter = "" Then
        lblDebitTotal = Format(debitTotal, "#,##0.00")
        lblCreditTotal = Format(creditTotal, "#,##0.00")
        
        'Multiplied by -1 to show the debit figure (expenditure) as negative
        lblCurrentPnL = Format(currentPnL * -1, "#,##0.00")
        
        If currentPnL * -1 > 0 Then
             lblCurrentPnL.ForeColor = vbBlue
        Else
            lblCurrentPnL.ForeColor = vbRed
        End If
   Else
        lblDebitTotal = Format(0, "#,##0.00")
        lblCreditTotal = Format(0, "#,##0.00")
   End If
   adoRST.Close
   Set adoRST = Nothing
   'Formating added by anol 20160522
    For i = 0 To flxNominalCode.Rows - 1
         flxNominalCode.TextMatrix(i, 5) = Format(flxNominalCode.TextMatrix(i, 5), "#,##0.00")
         flxNominalCode.TextMatrix(i, 6) = Format(flxNominalCode.TextMatrix(i, 6), "#,##0.00")
     Next i
   Exit Function

NewNominalCode:
   Set adoRST = Nothing
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Generating Nominal Balances"
   
End Function
'Private Function LoadFlxNominalCode(adoConn As ADODB.Connection) As Boolean
'   Dim szSQL As String
'   Dim adoRst As New ADODB.Recordset
'
'   LoadFlxNominalCode = True
'   On Error GoTo NewNominalCode
'
''  Check: has the client id been setup with NC?
'   adoRst.Open "SELECT ClientID FROM NominalLedger;", adoConn, adOpenStatic, adLockReadOnly
'   adoRst.Close
'
''  ClientID column has been setup
'   szSQL = "SELECT N.Code, N.Name, T.NLTypeCode, T.TypeValue, S.STName, '', '', " & _
'               "N.ClientID, IIF(N.Posting, 'YES', 'NO'), DrCr, SubType " & _
'           "FROM (NominalLedger AS N LEFT JOIN NLSubTypes AS S ON N.SubType = S.STCode) " & _
'                 "INNER JOIN NLType AS T ON N.Type = T.NLTypeCode " & _
'           "WHERE N.ClientID <> 'NONE' " & _
'           "ORDER BY N.Code;"
''Debug.Print szSQL
''   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
''                   x    x        0           x         x         x              x          0        x      0     0
'   TotalNC = populateGridDefinedHeader(adoConn, szSQL, flxNominalCode, 0)
'
'   Set adoRst = Nothing
'   Exit Function
'
'NewNominalCode:
'   Set adoRst = Nothing
'   LoadFlxNominalCode = False
'
'End Function

Private Sub ConfigFlxNominalCode()
   Dim szHeader As String, iCol As Integer

   flxNominalCode.Clear
   flxNominalCode.Cols = 10
   flxNominalCode.Rows = 2
   flxNominalCode.RowHeight(0) = 0
   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
'                   x    x        0           x       x          x              x          0         x      0     0
   flxNominalCode.FormatString = szHeader$

   flxNominalCode.ColWidth(0) = Label1(2).Left - Label1(1).Left      'Code
   flxNominalCode.ColWidth(1) = Label1(3).Left - Label1(2).Left      'Nominal Name
   flxNominalCode.ColWidth(2) = 0                                    'NLTypeCode
   flxNominalCode.ColWidth(3) = Label1(4).Left - Label1(3).Left      'TypeValue
   flxNominalCode.ColWidth(4) = Label1(5).Left - Label1(4).Left      'Sub Type Name
   flxNominalCode.ColWidth(5) = Label1(6).Left - Label1(5).Left      'DebitBalance
   Label1(5).Width = flxNominalCode.ColWidth(5)
   Label1(5).Alignment = vbRightJustify
   flxNominalCode.ColWidth(6) = Label1(7).Left - Label1(6).Left      'CreditBalance
   Label1(6).Width = flxNominalCode.ColWidth(6)
   Label1(6).Alignment = vbRightJustify
   flxNominalCode.ColWidth(7) = 0                                    'ClientID
   flxNominalCode.ColWidth(8) = 1200 'flxNominalCode.Width + flxNominalCode.Left - Label1(6).Left - 300     'Posting
   flxNominalCode.ColWidth(9) = 0                                    'DrCr
   flxNominalCode.ColWidth(10) = 0                                   'SubType
   
   lblCurrentPnL.Width = Label1(5).Width
   lblDebitTotal.Width = Label1(5).Width
   lblCreditTotal.Width = Label1(6).Width
   
   lblCurrentPnL.Left = Label1(5).Left - Label1(5).Width
   lblDebitTotal.Left = Label1(5).Left
   lblCreditTotal.Left = Label1(6).Left
End Sub
Private Sub ConfigflxChartOfACCDefault()
   Dim szHeader As String, iCol As Integer

   flxChartOfACCDefault.Clear
   flxChartOfACCDefault.Cols = 10
   flxChartOfACCDefault.Rows = 2
   flxChartOfACCDefault.RowHeight(0) = 0
   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
'                   x    x        0           x       x          x              x          0         x      0     0
   flxChartOfACCDefault.FormatString = szHeader$

   flxChartOfACCDefault.ColWidth(0) = Label1(2).Left - Label1(1).Left      'Code
   flxChartOfACCDefault.ColWidth(1) = Label1(3).Left - Label1(2).Left      'Nominal Name
   flxChartOfACCDefault.ColWidth(2) = 0                                    'NLTypeCode
   flxChartOfACCDefault.ColWidth(3) = Label1(4).Left - Label1(3).Left      'TypeValue
   flxChartOfACCDefault.ColWidth(4) = Label1(5).Left - Label1(4).Left      'Sub Type Name
   flxChartOfACCDefault.ColWidth(5) = Label1(6).Left - Label1(5).Left      'DebitBalance
   Label1(5).Width = flxChartOfACCDefault.ColWidth(5)
   Label1(5).Alignment = vbRightJustify
   flxChartOfACCDefault.ColWidth(6) = Label1(7).Left - Label1(6).Left      'CreditBalance
   Label1(6).Width = flxChartOfACCDefault.ColWidth(6)
   Label1(6).Alignment = vbRightJustify
   flxChartOfACCDefault.ColWidth(7) = 0                                    'ClientID
   flxChartOfACCDefault.ColWidth(8) = 1200 'flxChartOfACCDefault.Width + flxChartOfACCDefault.Left - Label1(6).Left - 300     'Posting
   flxChartOfACCDefault.ColWidth(9) = 0                                    'DrCr
   flxChartOfACCDefault.ColWidth(10) = 0                                   'SubType
   
'   lblCurrentPnL.Width = Label1(5).Width
'   lblDebitTotal.Width = Label1(5).Width
'   lblCreditTotal.Width = Label1(6).Width
'
'   lblCurrentPnL.Left = Label1(5).Left - Label1(5).Width
'   lblDebitTotal.Left = Label1(5).Left
'   lblCreditTotal.Left = Label1(6).Left
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   UnLoadForm Me
   ClearReportData "ReportTrialBalance", sessionID
   
   If CALLER_FORM = "frmClientNew4" Then
'      frmClientNew4.LoadNCinCombo
'      frmClientNew4.Show
      CALLER_FORM = ""
   Else
      'frmMMain.fraCmdButton.Enabled = True
   End If
End Sub

Public Sub ControlHanlding(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         cmdAddNew.Enabled = True
         cmdEdit.Enabled = True
         flxNominalCode.Enabled = True

         iNewEdit = 0
   
      Case ComponentMode.NewEntryMode
         cmdAddNew.Enabled = False
         cmdEdit.Enabled = False
         flxNominalCode.Enabled = False

         iNewEdit = 1

      Case ComponentMode.EditMode
         cmdAddNew.Enabled = False
         cmdEdit.Enabled = False
         flxNominalCode.Enabled = False

         iNewEdit = 2

      Case ComponentMode.SavedMode
         cmdAddNew.Enabled = True
         cmdEdit.Enabled = True
         flxNominalCode.Enabled = True
         iNewEdit = 0
   End Select
End Sub

Private Function LoadCmbClient(adoConn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbClient = False
   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT C.CLIENTID, C.CLIENTNAME, S.Code AS NCode " & _
           "FROM CLIENT AS C LEFT JOIN NominalLedger AS S ON C.ClientID = S.ClientID " & _
           "Where S.CAType = 'R' " & _
           "ORDER BY CLIENTNAME;"

   szSQL = "SELECT C.CLIENTID, C.CLIENTNAME, Q.Code AS NCode " & _
           "FROM CLIENT AS C LEFT JOIN (" & _
               "SELECT ClientID, Code " & _
               "FROM NominalLedger AS N " & _
               "WHERE N.CAType = 'R'" & _
              ") AS Q ON C.ClientID = Q.ClientID " & _
           "ORDER BY C.CLIENTID;"


   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount - 1
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   cboC.Column() = Data()

   LoadCmbClient = True
   Exit Function

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Function


Private Function LoadCmbProperties(adoConn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbProperties = False
   On Error GoTo ErrorHandler

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property WHERE CLIENTID = '" & txtClientList.Tag & "' " & _
           "ORDER BY PropertyID;"


   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

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
   
   cboC.Column() = Data()
   cboC.ListIndex = 0
   
   LoadCmbProperties = True
   Exit Function

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   ShowMsgInTaskBar "The Nominal Ledger cannot be loaded, as no property has been setup for this client", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Function LoadCmbFunds(adoConn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbFunds = False
   On Error GoTo ErrorHandler

   szSQL = "SELECT FundID, FundCode " & _
           "FROM Fund ORDER BY FundCode;"


   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Funds"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   
   cboC.Column() = Data()
   cboC.ListIndex = 0
   
   LoadCmbFunds = True
   Exit Function

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Function



Private Sub gridBankCode_Click()
     Dim adoConn  As New ADODB.Connection
     adoConn.Open getConnectionString
     If sTextBox = "3" Then
            flxNominalCode.Clear
            flxNominalCode.Rows = 2
            lblDebitTotal.Caption = "0.00"
            lblCreditTotal.Caption = "0.00"
            lblCurrentPnL.Caption = "0.00"
            txtBudgetYears.Tag = gridBankCode.TextMatrix(gridBankCode.row, 3) 'FYID
            txtBudgetYears.text = gridBankCode.TextMatrix(gridBankCode.row, 2) 'FY description
            cmdBudgetYears.Tag = gridBankCode.TextMatrix(gridBankCode.row, 1)  'FY start date
            LoadPeriods adoConn
            cmdFilter.Enabled = True
            FocusControl cmdFilter
            picBankCode.Visible = False
            
      End If
      adoConn.Close
      Set adoConn = Nothing
End Sub

Private Sub gridBankCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gridBankCode_Click
    End If
End Sub

Private Sub optCopyDefltCOA_Click()
     Frame1(2).Visible = False
End Sub

Private Sub optCopyDemandTemplate_Click()
     Frame1(2).Visible = True
     LoadflxClientCopyForSelection
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SSTab1.MousePointer = vbArrow
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
Private Sub cmdClientList_Click()
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
    SSTab1.Enabled = False
    Frame2.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
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
Private Sub LoadflxClientCopyForSelection()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClientCopy.RowHeight(0) = 0
   flxClientCopy.Cols = 3
   flxClientCopy.ColWidth(0) = 300
   flxClientCopy.ColWidth(1) = 1500
   flxClientCopy.ColWidth(2) = 3500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClientCopy.Clear
   flxClientCopy.Rows = 2
   flxClientCopy.ColAlignment(0) = vbLeftJustify
   flxClientCopy.ColAlignment(1) = vbLeftJustify
   flxClientCopy.ColAlignment(2) = vbLeftJustify

   
   
   adoConn.Open getConnectionString
   szSQL = "SELECT DIStinct N.clientID,U.ClientNAme from nominalLedger N, (SELECT CLIENTID," & _
            "CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID) as U where N.ClientID=U.clientID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
       flxClientCopy.Rows = rstRec.RecordCount + 1
           rRow = 1
           While Not rstRec.EOF
               flxClientCopy.row = 1
               flxClientCopy.RowSel = 1
               flxClientCopy.ColSel = 1
               flxClientCopy.TextMatrix(rRow, 0) = ""
               flxClientCopy.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClientCopy.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClientCopy.RowHeight(rRow) = 280
               rstRec.MoveNext
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    SSTab1.Enabled = True
    Frame2.Enabled = True
    cmdClientList.SetFocus
End Sub

Private Sub cmdproperty_Click()
        Frame1(1).Visible = False
        Frame1(2).Visible = False
        picClient.Left = 4747.029
        picClient.Top = 155.299
        sTextBox = "2"
        LoadPropertyList
        SSTab1.Enabled = False
        Frame2.Enabled = False
        picClient.Visible = True
        txtSearchClientID.SetFocus
End Sub

Private Sub flxClient_Click()
        SSTab1.Enabled = True
        Frame2.Enabled = True
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
                flxNominalCode.Clear
                flxNominalCode.Rows = 2
                lblDebitTotal.Caption = "0.00"
                lblCreditTotal.Caption = "0.00"
                lblCurrentPnL.Caption = "0.00"
                txtBudgetYears.text = ""
                txtBudgetYears.Tag = ""
                iFinancialWarning = 0
                If LoadFirstFinancialYear = True Then
                    FocusControl cmdProperty
                End If
                
                Call LoadFirstFinancialYear
                LoadPeriods adoConn
                Call cmdFilter_Click
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                flxNominalCode.Clear
                flxNominalCode.Rows = 2
                lblDebitTotal.Caption = "0.00"
                lblCreditTotal.Caption = "0.00"
                lblCurrentPnL.Caption = "0.00"
                SetCurrentFY adoConn
                cmdFundLookUp.SetFocus
        End If
        If sTextBox = "3" Then
                txtFundName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFundName.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                flxNominalCode.Clear
                flxNominalCode.Rows = 2
                cmdBudgetYears.SetFocus
        End If
        picClient.Visible = False
        adoConn.Close
        
End Sub

Private Sub flxClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And flxClient.row = 1 Then
        txtSearchClientID.SetFocus
     End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          SSTab1.Enabled = True
          Frame2.Enabled = True
          flxClient_Click
    End If
    If KeyAscii = 27 Then
            picClient.Visible = False
            SSTab1.Enabled = True
            Frame2.Enabled = True
            If sTextBox = "1" Then
                 cmdClientList.SetFocus
            ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
            ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
            End If
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub

Private Sub txtFilterbyProperty_Change()
    Dim i As Integer

 
   For i = flxClientCopy.Rows - 1 To 1 Step -1
        flxClientCopy.RowHeight(i) = 240
        If InStr(1, UCase(flxClientCopy.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClientCopy.RowHeight(i) = 0
              flxClientCopy.TextMatrix(i, 1) = ""
        End If
        If flxClientCopy.RowHeight(i) = 240 Then
              flxClientCopy.row = i
        End If
   Next i
End Sub

Private Sub txtFundName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdFundLookUp.SetFocus
    End If
End Sub

Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdProperty.SetFocus
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
           txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
         picClient.Visible = False
          SSTab1.Enabled = True
            Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
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
Private Sub cmdFundLookUp_Click()
   Frame1(1).Visible = False
   Frame1(2).Visible = False
   If txtClientList.text = "" Then
      ShowMsgInTaskBar "Please select a client to continue.", , "N"
      Exit Sub
   End If

'   fmePropertyLookup.Top = txtFundNo.Top + txtFundNo.Height + 5
'   fmePropertyLookup.Left = txtFundNo.Left - (fmePropertyLookup.Width - txtFundNo.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   gridPropertyLookup.Visible = True
'   txtSearchProperty.text = ""
'   txtSearchProperty.Enabled = True
'   txtSearchProperty.SetFocus
'
'   LOOKUPCommand = "Fund"
'
'   PopulatePropertyLookup IIf(txtClientList.Tag = "ALL", "", " WHERE CLIENTID = '" & txtClientList.Tag & "'")
    picClient.Left = 5500.029
    picClient.Top = 225.299
     sTextBox = "3"
     picClient.Visible = True
    
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFunds(adoConn)
    adoConn.Close
    SSTab1.Enabled = False
      Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadFunds(conConnection As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim SQLStr1 As String
   SQLStr1 = "SELECT FundID, FundCode, FundName FROM Fund;"
   adoRST.Open SQLStr1, conConnection, adOpenKeyset, adLockReadOnly

   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 2700
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   
'   flxClient.ColWidth(0) = 200
'   flxClient.ColWidth(1) = 0
'   flxClient.ColWidth(2) = 3000
'   picClient.Width = 3500
'   cmdPicCLose.Left = 3200
'   txtSearchClientID.Left = 45

   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   txtSearchClientID.Width = 1500
   txtSearchClientName.Visible = True
'   picClient.Width = 5295
'   flxClient.Width = 5175
   
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   txtSearchClientName.Left = 1580
   'txtSearchClientName.Width = 3600
'   picClient.Height = 4095
'   flxClient.Height = 3345
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 250
'   lblClientName.Width = 3600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   If adoRST.RecordCount > 0 Then
        ReDim szaFundCode(adoRST.RecordCount, 2) As String
   End If
   
   Dim rRow As Integer
   If adoRST.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
       If sTextBox = "3" Then
            rRow = 1
             flxClient.TextMatrix(rRow, 0) = "ALL"
               flxClient.TextMatrix(rRow, 1) = "ALL"
               flxClient.TextMatrix(rRow, 2) = "ALL Funds"
               szaFundCode(rRow - 1, 0) = "ALL"
               szaFundCode(rRow - 1, 1) = "ALL"
               szaFundCode(rRow - 1, 2) = "ALL Funds"
               flxClient.RowHeight(rRow) = 280
               flxClient.AddItem ""
            rRow = 2
            While Not adoRST.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
'               flxClient.TextMatrix(rRow, 0) = adoRst.Fields.Item("FundCode").Value
'               flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("FundName").Value
'               flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("FundID").Value
               
               flxClient.TextMatrix(rRow, 0) = "  " & adoRST.Fields.Item("FundID").Value
               flxClient.TextMatrix(rRow, 1) = adoRST.Fields.Item("FundCode").Value
               flxClient.TextMatrix(rRow, 2) = adoRST.Fields.Item("FundName").Value
               
               szaFundCode(rRow - 1, 0) = adoRST.Fields.Item("FundCode").Value
               szaFundCode(rRow - 1, 1) = adoRST.Fields.Item("FundName").Value
               szaFundCode(rRow - 1, 2) = adoRST.Fields.Item("FundID").Value
                flxClient.RowHeight(rRow) = 280
               adoRST.MoveNext
               If Not adoRST.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
       End If
   End If
End Sub


'Private Sub cmdPropertyLookup_Click()
'   If txtClientList.text = "" Then
'      ShowMsgInTaskBar "Please select a client to continue.", , "N"
'      Exit Sub
'   End If
'
'   fmePropertyLookup.Top = txtPropertyID.Top + txtPropertyID.Height + 5
'   fmePropertyLookup.Left = txtPropertyID.Left - (fmePropertyLookup.Width - txtPropertyID.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   flxNominalCode.Visible = True
'   txtSearchNo.text = ""
'   txtSearchNo.Enabled = True
'   txtSearchNo.SetFocus
'
'   LOOKUPCommand = "Property"
'
'   PopulatePropertyLookup IIf(txtClientList.text = "ALL", "", " WHERE CLIENTID = '" & txtClientList.text & "'")
'End Sub


Private Sub txtSearchNo_Change()
    If sMode = "NO" Then
        ConfigFlxNominalCode
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        
        'Call ExportData2NominalLedger(adoConn)
        
        If Len(Trim(txtSearchNo.text)) > 0 Then
            Call GenerateNominalAccounts(adoConn, "1")
        Else
            Call GenerateNominalAccounts(adoConn, "")
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

Private Sub txtSearchNo_GotFocus()
    sMode = "NO"
End Sub

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSearchRef.SetFocus
    End If
End Sub

Private Sub txtSearchRef_Change()
    If sMode = "Name" Then
        ConfigFlxNominalCode
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        
        'Call ExportData2NominalLedger(adoConn)
        If Len(Trim(txtSearchRef.text)) > 0 Then
            txtSearchNo.text = ""
            Call GenerateNominalAccounts(adoConn, "2")
        Else
            Call GenerateNominalAccounts(adoConn, "")
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

Private Sub txtSearchRef_GotFocus()
     sMode = "Name"
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
End Sub
