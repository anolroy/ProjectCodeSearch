VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmViewBankBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Bank Balance"
   ClientHeight    =   12510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19515
   Icon            =   "frmViewBankBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12510
   ScaleWidth      =   19515
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5580
      Left            =   11400
      ScaleHeight     =   5550
      ScaleWidth      =   7470
      TabIndex        =   8
      Top             =   3720
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
         TabIndex        =   7
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4830
         Left            =   45
         TabIndex        =   6
         Top             =   675
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   8520
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
         Left            =   1350
         TabIndex        =   5
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   135
         TabIndex        =   4
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1485
         TabIndex        =   14
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
         TabIndex        =   13
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Left            =   4590
         TabIndex        =   10
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
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   4590
         TabIndex        =   9
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
   Begin TabDlg.SSTab SSTabViewBankBalance 
      Height          =   11760
      Left            =   45
      TabIndex        =   15
      Top             =   90
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   20743
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "View Bank Balance"
      TabPicture(0)   =   "frmViewBankBalance.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdPrint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Reconciled Balances"
      TabPicture(1)   =   "frmViewBankBalance.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdPrintReconciled"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -59925
         TabIndex        =   57
         Top             =   11070
         Width           =   1890
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -57810
         TabIndex        =   51
         Top             =   11070
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   10680
         Left            =   45
         TabIndex        =   24
         Top             =   315
         Width           =   19365
         Begin VB.PictureBox fmeLoading 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   7200
            ScaleHeight     =   450
            ScaleWidth      =   3195
            TabIndex        =   58
            Top             =   5160
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
               Left            =   600
               TabIndex        =   59
               Top             =   120
               Width           =   4590
            End
         End
         Begin VB.Frame fraDateRange 
            Caption         =   "Date Range"
            Height          =   1215
            Left            =   6930
            TabIndex        =   34
            Top             =   945
            Visible         =   0   'False
            Width           =   14220
            Begin VB.TextBox txtSCYRREnDt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   36
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtSCYRRStDt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   35
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "End Date"
               Height          =   255
               Index           =   2
               Left            =   840
               TabIndex        =   39
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Start Date"
               Height          =   255
               Index           =   3
               Left            =   840
               TabIndex        =   38
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.PictureBox picBankCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3210
            Left            =   810
            ScaleHeight     =   3180
            ScaleWidth      =   2640
            TabIndex        =   52
            Top             =   4005
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
               TabIndex        =   53
               Top             =   90
               Width           =   255
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
               Height          =   2715
               Left            =   45
               TabIndex        =   54
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
               TabIndex        =   56
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
               TabIndex        =   55
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
         Begin VB.CommandButton cmdOK 
            Caption         =   "&Display"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   17325
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Generate Payment later"
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Frame Frame4 
            Height          =   735
            Left            =   135
            TabIndex        =   43
            Top             =   180
            Width           =   18600
            Begin VB.CommandButton cmdClientList2 
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
               Left            =   4335
               TabIndex        =   44
               Top             =   270
               Width           =   300
            End
            Begin MSForms.TextBox txtClientList2 
               Height          =   285
               Left            =   915
               TabIndex        =   46
               Top             =   270
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
            Begin VB.Label lblClient 
               BackStyle       =   0  'Transparent
               Caption         =   "Client:"
               ForeColor       =   &H80000007&
               Height          =   255
               Index           =   1
               Left            =   135
               TabIndex        =   45
               Top             =   300
               Width           =   555
            End
         End
         Begin VB.Frame fraDateOption 
            Caption         =   "Date options"
            Height          =   510
            Left            =   135
            TabIndex        =   40
            Top             =   930
            Width           =   14280
            Begin VB.OptionButton Option2 
               Caption         =   "By Financial Year"
               Height          =   195
               Left            =   1530
               TabIndex        =   42
               Top             =   180
               Value           =   -1  'True
               Width           =   1545
            End
            Begin VB.OptionButton Option1 
               Caption         =   "By Date"
               Height          =   195
               Left            =   3195
               TabIndex        =   41
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.Frame fraFinancial 
            Caption         =   "Financial Period:"
            Height          =   1215
            Left            =   135
            TabIndex        =   25
            Top             =   1470
            Width           =   14265
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
               Left            =   3675
               TabIndex        =   27
               Top             =   315
               Width           =   300
            End
            Begin VB.CheckBox chkYtD 
               Caption         =   "Y&TD"
               Height          =   255
               Left            =   10095
               TabIndex        =   26
               Top             =   315
               Width           =   735
            End
            Begin MSForms.ComboBox cmbPeriodFrom 
               Height          =   285
               Left            =   5130
               TabIndex        =   33
               Top             =   315
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
               Caption         =   "Period From:"
               Height          =   195
               Index           =   12
               Left            =   4095
               TabIndex        =   32
               Top             =   330
               Width           =   885
            End
            Begin MSForms.ComboBox cmbPeriodTo 
               Height          =   285
               Left            =   7920
               TabIndex        =   31
               Top             =   285
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
               Caption         =   "Period To:"
               Height          =   195
               Index           =   14
               Left            =   7110
               TabIndex        =   30
               Top             =   330
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Financial Year:"
               Height          =   195
               Index           =   66
               Left            =   90
               TabIndex        =   29
               Top             =   360
               Width           =   1005
            End
            Begin MSForms.TextBox txtBudgetYears 
               Height          =   285
               Left            =   1335
               TabIndex        =   28
               Top             =   315
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
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxReconciled 
            Height          =   7770
            Left            =   240
            TabIndex        =   47
            Top             =   2760
            Width           =   19215
            _ExtentX        =   33893
            _ExtentY        =   13705
            _Version        =   393216
            Cols            =   5
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
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   16830
         TabIndex        =   23
         Top             =   11070
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintReconciled 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14535
         TabIndex        =   22
         Top             =   11070
         Width           =   1890
      End
      Begin VB.Frame Frame2 
         Height          =   8880
         Left            =   -75000
         TabIndex        =   19
         Top             =   1950
         Width           =   18645
         Begin VB.CommandButton cmdClose2 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   13680
            TabIndex        =   20
            Top             =   9990
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAllBank 
            Height          =   8580
            Left            =   90
            TabIndex        =   21
            Top             =   225
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   15134
            _Version        =   393216
            Cols            =   5
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
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1590
         Left            =   -74955
         TabIndex        =   16
         Top             =   375
         Width           =   18645
         Begin VB.CommandButton cmdDisplay1 
            Caption         =   "&Display"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   10620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Generate Payment later"
            Top             =   270
            Width           =   1200
         End
         Begin VB.Frame Frame5 
            Caption         =   "Date Range"
            Height          =   1215
            Left            =   4860
            TabIndex        =   48
            Top             =   180
            Width           =   4860
            Begin VB.TextBox txtStart1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   1
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtEnd1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   2
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Start Date"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   50
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "End Date"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   49
               Top             =   720
               Width           =   975
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
            Left            =   4335
            TabIndex        =   0
            Top             =   270
            Width           =   300
         End
         Begin MSForms.TextBox txtClientList 
            Height          =   285
            Left            =   915
            TabIndex        =   18
            Top             =   270
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
         Begin VB.Label lblClient 
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   17
            Top             =   300
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmViewBankBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTextBox As String
Dim szAllBankBalance As String
Dim SelectedConBankID As String
Dim dtEnd As Date
' adoConn.Execute "Delete From ReportViewBankBalance"
   'If txtClientList2.text = "ALL" Then
'        szSQL = "Select BankCode as BNC,BankName as BNN,conBankID, '1' as D, '' as clientID from ConsolidatedBankList" & _
'                " UNION Select CB.NominalCode AS BNC, N.Name AS BNN,0 as conBankID, '2' as D,CB.CLIENT_ID  from tlbClientBanks CB,NominalLedger  N WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code " & _
'                "" & _
'              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
             
''              szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) " & _
''                    " as Amount,X.ClientID,BankCode as BNC,N.Name AS BNN  From ( SELECT SUM(R.Amount)" & _
''                    " AS AMT, Type AS T,ClientID,BankCode  FROM tlbReceipt AS R, " & _
''                    " tlbClientBanks AS B WHERE  B.NominalCode = R.BankCode AND B.CLIENT_ID =" & _
''                    " R.ClientID AND R.Amount > 0 AND right(R.ReconNow,4)='Full' GROUP BY Type,ClientID,BankCode  " & _
''                    " UNION  SELECT SUM(P.Amount) AS AMT, Type AS T,ClientID,BankCode  FROM" & _
''                    " tlbPayment AS P, tlbClientBanks AS B WHERE B.NominalCode = P.BankCode AND" & _
''                    " B.CLIENT_ID = P.ClientID AND P.Amount > 0 AND right(P.ReconNow,4)='Full' GROUP BY TYPE,ClientID,BankCode " & _
''                    "  UNION  SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS" & _
''                    " AS T,ClientID,BANK_AC as BankCode  FROM tlbBankPayment AS BP, tlbClientBanks AS" & _
''                    " CB WHERE  CB.NominalCode = BP.BANK_AC AND CB.CLIENT_ID = BP.ClientID AND" & _
''                    " (BP.NET_AMOUNT + BP.VAT) > 0 AND right(BP.ReconNow,4)='Full' GROUP BY TRANS,ClientID,BANK_AC  ORDER BY" & _
''                    " T) X,NominalLedger N where N.ClientID = X.CLIENTID AND N.Code = X.BankCode" & _
''                    " group by X.ClientID,BankCode,Name order by  X.ClientID,BankCode,Name "
''                adoConn.Execute "Insert into ReportViewBankBalance(amount,ClientID,BankCode,BankName) " & _
''                                szSQL
''               adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''              rRow = 1
''                While Not adoRst.EOF
''                    flxReconciled.row = 1
''                    flxReconciled.TextMatrix(rRow, 0) = ""
''                    flxReconciled.TextMatrix(rRow, 1) = IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value)
''                    flxReconciled.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value)
''                    flxReconciled.TextMatrix(rRow, 3) = adoRst.Fields.Item("BNN").Value
'''                    If adoRst.Fields.Item("D").Value = "2" Then
''                            flxReconciled.TextMatrix(rRow, 4) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
'''                            Debug.Print flxReconciled.TextMatrix(rRow, 3)
'''                    Else
'''                            flxReconciled.TextMatrix(rRow, 3) = BankAccBalanceConsolidated(adoconn, adoRst.Fields.Item("conBankID").Value)
'''                            Debug.Print flxReconciled.TextMatrix(rRow, 3)
'''                    End If
''                    flxReconciled.RowHeight(rRow) = 280
''                    adoRst.MoveNext
''                    If Not adoRst.EOF Then flxReconciled.AddItem ""
''                    rRow = rRow + 1
''                 Wend
''                 adoRst.Close
'             szSQL = "Select StatementDate,BankCode as BNC,BankName as BNN,conBankID, '1' as D, '' as clientID from ConsolidatedBankList"
'             adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'             flxReconciled.AddItem ""
'             rRow = 1
'             Dim dblPayment As Currency
'             Dim dblReceipt As Currency
'                While Not adoRst.EOF
'                    flxReconciled.row = 1
'                    flxReconciled.TextMatrix(rRow, 0) = ""
'                    flxReconciled.TextMatrix(rRow, 1) = "Consolidated"
'                    flxReconciled.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value)
'                    flxReconciled.TextMatrix(rRow, 3) = adoRst.Fields.Item("BNN").Value
'                    flxReconciled.TextMatrix(rRow, 6) = Format(adoRst.Fields.Item("StatementDate").Value, "dd MMM yyyy")
'                    ''this is similar to cashbook tab
'                    dblReceipt = 0
'                    dblPayment = 0
'                    flxReconciled.TextMatrix(rRow, 7) = LastReconciledBankBalanceConsolidated(adoConn, adoRst.Fields.Item("conBankID").Value, adoRst.Fields.Item("StatementDate").Value, dblPayment, dblReceipt)
'                    flxReconciled.TextMatrix(rRow, 4) = dblReceipt
'                    flxReconciled.TextMatrix(rRow, 5) = dblPayment
'                    flxReconciled.TextMatrix(rRow, 8) = BankAccBalanceConsolidatedReconciledDated(adoConn, adoRst.Fields.Item("conBankID").Value, adoRst.Fields.Item("StatementDate").Value)
'                    flxReconciled.TextMatrix(rRow, 10) = BankAccBalanceConsolidated(adoConn, adoRst.Fields.Item("conBankID").Value)
'                    flxReconciled.TextMatrix(rRow, 9) = flxReconciled.TextMatrix(rRow, 10) - flxReconciled.TextMatrix(rRow, 8) 'Unreconciled Balance
'                    Debug.Print flxReconciled.TextMatrix(rRow, 3)
'                    flxReconciled.RowHeight(rRow) = 280
'                    adoRst.MoveNext
'                    If Not adoRst.EOF Then flxReconciled.AddItem ""
'                    rRow = rRow + 1
'                 Wend
'                 adoRst.Close

'        szSQL = "Select A.StatementDate,A.BankCode as BNC,BankName as BNN,conBankID, '1' as D, '' as clientID from " & _
'                "ConsolidatedBankList A,tlbBankReconClosingBal B where B.ClientID=A.BankCode "

Private Sub cmdBankClose_Click()
     picBankCode.Visible = False
End Sub

Private Sub cmdBudgetYears_Click()
    sTextBox = "3"
    picBankCode.Left = cmdBudgetYears.Left - 500
    picBankCode.Top = cmdBudgetYears.Top + 1500
    If txtClientList2.text = "All" Then Exit Sub
    If txtClientList2.text = "Consolidated" Then Exit Sub
    Call LoadGridFY
End Sub
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String
   Dim K As Integer

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim firstClientID As String
   Dim adocheck As New ADODB.Recordset
   configGridFY
   adoConn.Open getConnectionString
           
'Grab 1 client ID which has maximum number of Financial Year Created
'SELECT  ClientID,  Count(ClientID) from financialYear group by  ClientID  order by  Count(ClientID) desc
'   adocheck.Open "SELECT  ClientID from financialYear  order by  FY_EndDate  desc", adoconn, adOpenStatic, adLockReadOnly
'   If Not adocheck.EOF Then
'        firstClientID = adocheck("ClientID").Value
'   End If
'   adocheck.Close
   szSQL = "SELECT F.FYrID, F.FY_StDate,FinancialYear,F.FY_Description,setascurrent " & _
           "FROM FinancialYear AS F " & _
           "WHERE F.ClientID = '" & txtClientList2.Tag & "'  " & _
           "ORDER BY FY_EndDate DESC;"


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
        MsgBox "No Financial years exist for this Client. Please create at least one Financial Year", vbInformation, "Warning!"
'        frmFinancialYearCreate.lblClientName.Caption = txtClientList.text
'        frmFinancialYearCreate.lblClientName.Tag = txtClientList.Tag
'        frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Add New"
'        frmFinancialYearCreate.financialYearID = UniqueID()
'        frmFinancialYearCreate.Show
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
Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    Label2.Caption = ""
    TextBox1.Visible = False
    LoadflxClient
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Function isConsolidateExists(adoConn As ADODB.Connection) As Boolean
'This function shall check if there is any value inputed as consolidated
'Written by anol 04 May 2015

   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT   consolidated " & _
           "FROM     tlbclientbanks " & _
           "where consolidated=true;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        isConsolidateExists = True
        Exit Function
   End If
End Function
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
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   
   txtSearchClientID.Left = 45
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                     rRow = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "ALL"
                    flxClient.TextMatrix(rRow, 2) = "ALL"
                    flxClient.RowHeight(rRow) = 280
                    flxClient.AddItem ""
                    
                    
           If isConsolidateExists(adoConn) Then
                    rRow = 2
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = "Consolidated"
                    flxClient.TextMatrix(rRow, 2) = "Consolidated"
                    flxClient.RowHeight(rRow) = 280
                    flxClient.AddItem ""
                    rRow = 3
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
    
    
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Sub ConfigConsolidateBank()
    Dim szHeader As String
    flxAllBank.Clear
    flxAllBank.Rows = 2
    'szHeader$ = "|<Bank ID|<Bank Code|<Bank Name|<Bank AC Number|<Sort Code|<Statement Date|<Closing Bal|<SOB"
    szHeader$ = "|<Client ID|<Bank Code|<Bank Name |<Bank Balance"
    flxAllBank.FormatString = szHeader$
   flxAllBank.Cols = 9
   flxAllBank.ColWidth(0) = 200
   flxAllBank.ColWidth(1) = 3000
   flxAllBank.ColWidth(2) = 3000
   flxAllBank.ColWidth(3) = 2000
   flxAllBank.ColWidth(4) = 1500
   flxAllBank.ColAlignment(4) = vbLeftJustify
   flxAllBank.ColWidth(5) = 1500
   flxAllBank.ColWidth(6) = 1500
   flxAllBank.ColWidth(7) = 1500
   flxAllBank.ColWidth(8) = 1500
   flxAllBank.ColAlignment = vbLeftJustify
End Sub

Private Sub ConfigureFlxBank()
    flxAllBank.Clear
    flxAllBank.Rows = 2
    Dim szHeader As String
    szHeader$ = "|<Client ID |<Bank Code|<Bank Name|<Current Cashbook Balance|<Retentions Balance|<Available Bank Balance"

    flxAllBank.FormatString = szHeader$
    
    
   'flxAllBank.RowHeight(0) = 0
   flxAllBank.Cols = 7
   flxAllBank.ColWidth(0) = 200
   flxAllBank.ColWidth(1) = 1500
   flxAllBank.ColWidth(2) = 4500
   flxAllBank.ColWidth(3) = 4500
   flxAllBank.ColWidth(4) = 2500
   flxAllBank.ColAlignment(4) = vbLeftJustify
   flxAllBank.ColWidth(5) = 2500
   flxAllBank.ColWidth(6) = 2500
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxAllBank.ColAlignment = vbLeftJustify
End Sub
Private Sub ConfigureFlxReconciled()
        flxReconciled.Clear
        flxReconciled.Rows = 2
        Dim szHeader As String
        szHeader$ = "|<Client ID |<Bank Code|<Bank Name|<Bank Balance"
        flxReconciled.FormatString = szHeader$
        flxReconciled.Cols = 5
        flxReconciled.ColWidth(0) = 200
        flxReconciled.ColWidth(1) = 1500
        flxReconciled.ColWidth(2) = 4500
        flxReconciled.ColWidth(3) = 4500
        flxReconciled.ColWidth(4) = 4500
        flxReconciled.ColAlignment(4) = vbLeftJustify
        txtSearchClientID.Width = 1530
        txtSearchClientName.Visible = True
        flxReconciled.ColAlignment = vbLeftJustify
End Sub
Private Function BankAndBalance(adoConn As ADODB.Connection) As String
   'On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim rsConsolidated As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim rRow As Integer
   If txtClientList.text = "" Then Exit Function
   If txtClientList.text = "Consolidated" Then
        ConfigConsolidateBank
   Else
        ConfigureFlxBank
   End If
   adoConn.Execute "Delete From ReportViewBankBalance"
   If txtClientList.text = "ALL" Then
            'add only consolidated part
              szSQL = "Select BankCode as BNC,BankName as BNN,conBankID, '1' as D, '' as clientID from ConsolidatedBankList"
               rsConsolidated.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
              flxAllBank.AddItem ""
              rRow = 1
                While Not rsConsolidated.EOF
                    flxAllBank.row = 1
                    flxAllBank.TextMatrix(rRow, 0) = ""
                    flxAllBank.TextMatrix(rRow, 1) = "Consolidated"
                    flxAllBank.TextMatrix(rRow, 2) = IIf(IsNull(rsConsolidated.Fields.Item("BNC").Value), "", rsConsolidated.Fields.Item("BNC").Value)
                    flxAllBank.TextMatrix(rRow, 3) = rsConsolidated.Fields.Item("BNN").Value
                    flxAllBank.TextMatrix(rRow, 4) = BankAccBalanceConsolidated(adoConn, rsConsolidated.Fields.Item("conBankID").Value)
                    Debug.Print flxAllBank.TextMatrix(rRow, 3)
                    flxAllBank.RowHeight(rRow) = 280
                    adoConn.Execute "Insert into ReportViewBankBalance(CLIENTID,amount,BankCode,BankName) " & _
                                        "VALUES ('" & flxAllBank.TextMatrix(rRow, 1) & "'," & flxAllBank.TextMatrix(rRow, 4) & ", '" & flxAllBank.TextMatrix(rRow, 2) & "','" & flxAllBank.TextMatrix(rRow, 3) & "')"
                                         '"'" & flxAllBank.TextMatrix(rRow, 3)  & "'");"
                   flxAllBank.TextMatrix(rRow, 5) = 0 ' ReturnRetention(adoconn, flxAllBank.TextMatrix(rRow, 2), flxAllBank.TextMatrix(rRow, 2))
                    flxAllBank.AddItem ""
                   rRow = rRow + 1
                   'flxAllBank.TextMatrix(rRow, 6) = 0Val(flxAllBank.TextMatrix(rRow, 5)) - Val(flxAllBank.TextMatrix(rRow, 4))'SALIA
                        'Nested part
                                      'Show breakdown of consolidated bank account
                                  szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) " & _
                                      " as Amount,X.ClientID,BankCode as BNC,N.Name AS BNN  From ( SELECT SUM(R.Amount)" & _
                                      " AS AMT, Type AS T,ClientID,BankCode  FROM tlbReceipt AS R, " & _
                                      " tlbClientBanks AS B WHERE  B.NominalCode = R.BankCode AND B.CLIENT_ID =" & _
                                      " R.ClientID AND R.Amount > 0 AND R.RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND R.RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND ConsolidatedBankID=" & rsConsolidated.Fields.Item("conBankID").Value & " AND B.consolidated=true GROUP BY Type,ClientID,BankCode  " & _
                                      " UNION  SELECT SUM(P.Amount) AS AMT, Type AS T,ClientID,BankCode  FROM" & _
                                      " tlbPayment AS P, tlbClientBanks AS B WHERE B.NominalCode = P.BankCode AND" & _
                                      " B.CLIENT_ID = P.ClientID AND P.Amount > 0 AND P.PDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND P.PDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND ConsolidatedBankID=" & rsConsolidated.Fields.Item("conBankID").Value & " AND B.consolidated=true GROUP BY TYPE,ClientID,BankCode " & _
                                      "  UNION  SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS" & _
                                      " AS T,ClientID,BANK_AC as BankCode  FROM tlbBankPayment AS BP, tlbClientBanks AS" & _
                                      " CB WHERE  BP.TRAN_DATE >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND  BP.TRAN_DATE <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND  CB.NominalCode = BP.BANK_AC AND CB.CLIENT_ID = BP.ClientID AND" & _
                                      " (BP.NET_AMOUNT + BP.VAT) > 0 AND ConsolidatedBankID=" & rsConsolidated.Fields.Item("conBankID").Value & " AND CB.consolidated=true GROUP BY TRANS,ClientID,BANK_AC  ORDER BY" & _
                                      " T) X,NominalLedger N where N.ClientID = X.CLIENTID AND N.Code = X.BankCode" & _
                                      " group by X.ClientID,BankCode,Name order by  X.ClientID,BankCode,Name "
                                      
                                      
                                adoConn.Execute "Insert into ReportViewBankBalance(amount,ClientID,BankCode,BankName) " & _
                                                  szSQL
                                 adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                 
                                 While Not adoRst.EOF
                                      flxAllBank.row = 1
                                      flxAllBank.TextMatrix(rRow, 0) = ""
                                      flxAllBank.TextMatrix(rRow, 1) = "    > " & IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value)
                                      flxAllBank.TextMatrix(rRow, 2) = "    " & IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value)
                                      flxAllBank.TextMatrix(rRow, 3) = adoRst.Fields.Item("BNN").Value
                                      flxAllBank.TextMatrix(rRow, 4) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
                                      flxAllBank.TextMatrix(rRow, 5) = "0.00"
                                      flxAllBank.RowHeight(rRow) = 280
                                      adoRst.MoveNext
                                      If Not adoRst.EOF Then flxAllBank.AddItem ""
                                      rRow = rRow + 1
                                   Wend
                                   adoRst.Close

                    rsConsolidated.MoveNext
                    If Not rsConsolidated.EOF Then flxAllBank.AddItem ""
                    rRow = rRow + 1
                 Wend
                 rsConsolidated.Close
                 
                
               'Now here add on retention values
               Dim i As Integer
              
               szSQL = "Select  Sum(amount) as AMTT,BankCode,ClientID  FROM RetentionDetails R where R.isDeleted=false " & _
                            "AND RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# group BY BankCode,ClientID"
               adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
               While Not adoRst.EOF
                     For i = 0 To flxAllBank.Rows - 1
                         If Trim(Replace(flxAllBank.TextMatrix(i, 1), "   > ", "")) = adoRst("ClientID").Value And Trim(Replace(flxAllBank.TextMatrix(i, 2), "   > ", "")) = adoRst("BankCode").Value Then
                            flxAllBank.TextMatrix(i, 5) = Format(adoRst("AMTT").Value, "0.00")
                            Exit For
                        End If
                    Next
                    adoRst.MoveNext
                Wend
                adoRst.Close
              'Show NON consolidated bank account
            'Exit Function
             szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) " & _
                    " as Amount,X.ClientID,BankCode as BNC,N.Name AS BNN  From ( SELECT SUM(R.Amount)" & _
                    " AS AMT, Type AS T,ClientID,BankCode  FROM tlbReceipt AS R, " & _
                    " tlbClientBanks AS B WHERE  B.NominalCode = R.BankCode AND B.CLIENT_ID =" & _
                    " R.ClientID AND R.Amount > 0 AND R.RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND R.RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND B.consolidated=false GROUP BY Type,ClientID,BankCode  " & _
                    " UNION  SELECT SUM(P.Amount) AS AMT, Type AS T,ClientID,BankCode  FROM" & _
                    " tlbPayment AS P, tlbClientBanks AS B WHERE B.NominalCode = P.BankCode AND" & _
                    " B.CLIENT_ID = P.ClientID AND P.Amount > 0 AND P.PDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND P.PDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND B.consolidated=false GROUP BY TYPE,ClientID,BankCode " & _
                    "  UNION  SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS" & _
                    " AS T,ClientID,BANK_AC as BankCode  FROM tlbBankPayment AS BP, tlbClientBanks AS" & _
                    " CB WHERE  BP.TRAN_DATE >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND  BP.TRAN_DATE <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND  CB.NominalCode = BP.BANK_AC AND CB.CLIENT_ID = BP.ClientID AND" & _
                    " (BP.NET_AMOUNT + BP.VAT) > 0 AND CB.consolidated=false GROUP BY TRANS,ClientID,BANK_AC  ORDER BY" & _
                    " T) X,NominalLedger N where N.ClientID = X.CLIENTID AND N.Code = X.BankCode" & _
                    " group by X.ClientID,BankCode,Name order by  X.ClientID,BankCode,Name "
                    
                adoConn.Execute "Insert into ReportViewBankBalance(amount,ClientID,BankCode,BankName) " & _
                                szSQL
               adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
              flxAllBank.AddItem ""
               flxAllBank.AddItem ""
                While Not adoRst.EOF
                    flxAllBank.row = 1
                    flxAllBank.TextMatrix(rRow, 0) = ""
                    flxAllBank.TextMatrix(rRow, 1) = IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value)
                    flxAllBank.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value)
                    flxAllBank.TextMatrix(rRow, 3) = adoRst.Fields.Item("BNN").Value
                    flxAllBank.TextMatrix(rRow, 4) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
                    flxAllBank.TextMatrix(rRow, 5) = "0.00"
                    flxAllBank.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxAllBank.AddItem ""
                    rRow = rRow + 1
                 Wend
                 adoRst.Close
                 For i = 1 To flxAllBank.Rows - 1
                          flxAllBank.TextMatrix(i, 6) = Format(Val(flxAllBank.TextMatrix(i, 4)) - Val(flxAllBank.TextMatrix(i, 5)), "0.00")
                 Next
   
        Exit Function
   End If
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
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
            If txtClientList.text <> "Consolidated" Then
                    MsgBox "Please setup your Client Bank Accounts." & Chr(13) & _
                           "Please alsoko check the nominal chart of account for the client."
             End If
             
   Else
        If txtClientList.text = "Consolidated" Then ' this if part is for consolidated options
              rRow = 1
                While Not adoRst.EOF
                    flxAllBank.row = 1
                    flxAllBank.TextMatrix(rRow, 0) = ""
                    flxAllBank.TextMatrix(rRow, 1) = "Consolidated"
                    flxAllBank.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("BankCode").Value), "", adoRst.Fields.Item("BankCode").Value)
                    flxAllBank.TextMatrix(rRow, 3) = adoRst.Fields.Item("BankName").Value
                    flxAllBank.TextMatrix(rRow, 4) = BankAccBalanceConsolidated(adoConn, adoRst.Fields.Item("conBankID").Value)
                     adoConn.Execute "Insert into ReportViewBankBalance(CLIENTID,amount,BankCode,BankName) " & _
                                        "VALUES ('" & flxAllBank.TextMatrix(rRow, 1) & "'," & flxAllBank.TextMatrix(rRow, 4) & ", '" & flxAllBank.TextMatrix(rRow, 2) & "','" & flxAllBank.TextMatrix(rRow, 3) & "')"
                                        
'                    flxAllBank.TextMatrix(rRow, 4) = adoRst.Fields.Item("conBankID").Value
'                    flxAllBank.TextMatrix(rRow, 5) = IIf(IsNull(adoRst.Fields.Item("BankCode").Value), "", adoRst.Fields.Item("BankCode").Value)
'                    flxAllBank.TextMatrix(rRow, 6) = IIf(IsNull(adoRst.Fields.Item("StatementDate").Value), "", adoRst.Fields.Item("StatementDate").Value)
'                    flxAllBank.TextMatrix(rRow, 7) = IIf(IsNull(adoRst.Fields.Item("ClosingBal").Value), "0", adoRst.Fields.Item("ClosingBal").Value)
'                    flxAllBank.TextMatrix(rRow, 8) = IIf(IsNull(adoRst.Fields.Item("SOB").Value), "0", adoRst.Fields.Item("SOB").Value)
                    flxAllBank.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxAllBank.AddItem ""
                    rRow = rRow + 1
                 Wend
        Else
                rRow = 1
                'For single client show BAL
                While Not adoRst.EOF
                    flxAllBank.row = 1
                    flxAllBank.TextMatrix(rRow, 0) = ""
                    flxAllBank.TextMatrix(rRow, 2) = adoRst.Fields.Item("BNC").Value
                    flxAllBank.TextMatrix(rRow, 3) = adoRst.Fields.Item("BNN").Value
                    'flxAllBank.TextMatrix(rRow, 4) = Format(BankAccBalance(adoconn, IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value), txtClientList.Tag), "0.00")
                    flxAllBank.TextMatrix(rRow, 4) = Format(BankBalDated(adoConn, IIf(IsNull(adoRst.Fields.Item("BNC").Value), "", adoRst.Fields.Item("BNC").Value), txtClientList.Tag), "0.00")
                    flxAllBank.TextMatrix(rRow, 5) = "0.00"
                    flxAllBank.TextMatrix(rRow, 1) = adoRst.Fields.Item("CLIENT_ID").Value
                     adoConn.Execute "Insert into ReportViewBankBalance(CLIENTID,amount,BankCode,BankName) " & _
                                        "VALUES ('" & flxAllBank.TextMatrix(rRow, 1) & "', " & flxAllBank.TextMatrix(rRow, 4) & ", '" & flxAllBank.TextMatrix(rRow, 2) & "','" & flxAllBank.TextMatrix(rRow, 3) & "')"

                    flxAllBank.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxAllBank.AddItem ""
                    rRow = rRow + 1
                 Wend
                 
                  'Now here add on retention values
                 'Dim i As Integer
                adoRst.Close
                 szSQL = "Select  Sum(amount) as AMTT,BankCode,ClientID  FROM RetentionDetails R where R.isDeleted=false " & _
                              "AND RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# group BY BankCode,ClientID"
                 adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                 While Not adoRst.EOF
                       For i = 0 To flxAllBank.Rows - 1
                           If Trim(Replace(flxAllBank.TextMatrix(i, 1), "   => ", "")) = adoRst("ClientID").Value And Trim(Replace(flxAllBank.TextMatrix(i, 2), "   => ", "")) = adoRst("BankCode").Value Then
                              flxAllBank.TextMatrix(i, 5) = Format(adoRst("AMTT").Value, "0.00")
                              Exit For
                          End If
                      Next
                      adoRst.MoveNext
                  Wend
                  adoRst.Close
                  
                  For i = 1 To flxAllBank.Rows - 1
                          flxAllBank.TextMatrix(i, 6) = Format(Val(flxAllBank.TextMatrix(i, 4)) - Val(flxAllBank.TextMatrix(i, 5)), "0.00")
                  Next
                  
                  
         End If
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function
Private Function BankBalDated(adoConn As ADODB.Connection, strBankCode As String, szClientID As String) As Double
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    'szSQL = "Select sum(amount) as DAmt from RetentionDetails where  isDeleted=false and  BankCode='" & strBankCode & "' and ClientID='" & szClientID & "' "
    szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) " & _
                    " as Amount,X.ClientID,BankCode as BNC,N.Name AS BNN  From ( SELECT SUM(R.Amount)" & _
                    " AS AMT, Type AS T,ClientID,BankCode  FROM tlbReceipt AS R, " & _
                    " tlbClientBanks AS B WHERE  B.NominalCode = R.BankCode AND B.CLIENT_ID =" & _
                    " R.ClientID AND R.Amount > 0 AND R.RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND R.RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "#  GROUP BY Type,ClientID,BankCode  " & _
                    " UNION  SELECT SUM(P.Amount) AS AMT, Type AS T,ClientID,BankCode  FROM" & _
                    " tlbPayment AS P, tlbClientBanks AS B WHERE B.NominalCode = P.BankCode AND" & _
                    " B.CLIENT_ID = P.ClientID AND P.Amount > 0 AND P.PDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND P.PDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "#  GROUP BY TYPE,ClientID,BankCode " & _
                    "  UNION  SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS" & _
                    " AS T,ClientID,BANK_AC as BankCode  FROM tlbBankPayment AS BP, tlbClientBanks AS" & _
                    " CB WHERE  BP.TRAN_DATE >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND  BP.TRAN_DATE <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "# AND  CB.NominalCode = BP.BANK_AC AND CB.CLIENT_ID = BP.ClientID AND" & _
                    " (BP.NET_AMOUNT + BP.VAT) > 0  GROUP BY TRANS,ClientID,BANK_AC  ORDER BY" & _
                    " T) X,NominalLedger N where N.ClientID = X.CLIENTID AND N.Code = X.BankCode AND N.ClientID ='" & szClientID & "' AND X.BankCode='" & strBankCode & "' " & _
                    " group by X.ClientID,BankCode,Name order by  X.ClientID,BankCode,Name "
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        BankBalDated = IIf(IsNull(adoRst.Fields.Item("amount").Value), 0, adoRst.Fields.Item("amount").Value)
    End If
    adoRst.Close
  

End Function


Public Function BankAccBalanceReconciled(adoConn As ADODB.Connection, szBank As String, szClientID As String) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode = '" & szBank & "' AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "P.ClientID = '" & szClientID & "' AND B.NominalCode = R.BankCode AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0  AND right(R.ReconNow,3)='Full'" & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = '" & szBank & "' AND " & _
                 "P.ClientID = '" & szClientID & "' AND " & _
                 "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID AND " & _
                 "P.Amount > 0 AND right(P.ReconNow,3)='Full'" & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = '" & szBank & "' AND " & _
                  "BP.ClientID = '" & szClientID & "' AND " & _
                  "CB.NominalCode = BP.BANK_AC AND CB.CLIENT_ID = BP.ClientID AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 AND right(BP.ReconNow,3)='Full' " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalanceReconciled = BankAccBalanceReconciled + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function LastStatementBalanceConsolidated(adoConn As ADODB.Connection, ByVal SelectedConBankID As String, ByVal DTdate As Date) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND " & _
                 "B.CLIENT_ID = P.ClientID AND right(R.ReconNow,4)='Full' AND " & _
                 "R.Amount > 0  AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "#" & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND right(P.ReconNow,4)='Full' AND " & _
                 "P.Amount > 0  AND PDate<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 AND right(BP.ReconNow,4)='Full' AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL AND right(BP.ReconNow,4)='Full' AND right(P.ReconNow,4)='Full' AND right(R.ReconNow,4)='Full'
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         LastStatementBalanceConsolidated = LastStatementBalanceConsolidated + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function BankAccBalanceConsolidatedReconciled(adoConn As ADODB.Connection, ByVal SelectedConBankID As String) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0  AND right(R.ReconNow,4)='Full'" & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND " & _
                 "P.Amount > 0 AND right(P.ReconNow,4)='Full' " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 AND right(BP.ReconNow,4)='Full' " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalanceConsolidatedReconciled = BankAccBalanceConsolidatedReconciled + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function BankAccBalanceConsolidated(adoConn As ADODB.Connection, ByVal SelectedConBankID As String) As Currency
'This function is a part of Display 1st TAB
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND R.RDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND R.RDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "#   AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0 " & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND P.PDate >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND P.PDate <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "#  AND " & _
                 "P.Amount > 0 " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND BP.TRAN_DATE >= #" & Format(txtStart1.text, "dd/MM/yyyy") & "# AND  BP.TRAN_DATE <= #" & Format(txtEnd1.text, "dd/MM/yyyy") & "#  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
Private Function BankAccBalanceConsolidated2ndTAB(adoConn As ADODB.Connection, ByVal SelectedConBankID As String) As Currency
'This function is a part of Display2ndTAB
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
'AND R.RDate >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "#
'AND P.PDate >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "#
' AND BP.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "#
   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND R.RDate <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#   AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0 " & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode  AND P.PDate <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#  AND " & _
                 "P.Amount > 0 " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC AND  BP.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalanceConsolidated2ndTAB = BankAccBalanceConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
Private Function BankAccBalUnreconciledConsolidated2ndTAB(adoConn As ADODB.Connection, ByVal SelectedConBankID As String) As Currency
'This function is a part of Display2ndTAB
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND R.RDate >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "# AND R.RDate <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#   AND " & _
                 "B.CLIENT_ID = P.ClientID AND  (ReconNow is null OR ReconNow='') AND  " & _
                 "R.Amount > 0 " & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND P.PDate >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "# AND P.PDate <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#  AND " & _
                 "P.Amount > 0 AND (ReconNow is null OR ReconNow='')   " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND BP.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "dd/MM/yyyy") & "# AND  BP.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "dd/MM/yyyy") & "#  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 AND (ReconNow is null OR ReconNow='')  " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalUnreconciledConsolidated2ndTAB = BankAccBalUnreconciledConsolidated2ndTAB + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
Private Sub cmdClientList2_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "2"
    LoadflxClient
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose2_Click()
    Unload Me
End Sub

Private Sub cmdDisplay1_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    szAllBankBalance = BankAndBalance(adoConn)
    adoConn.Close
End Sub

Private Sub cmdOK_Click()
        Dim adoConn As New ADODB.Connection
        If txtClientList2.text = "" Then Exit Sub
'        If txtClientList2.text = "Consolidated" Then
         
'        Else
'            ConfigureFlxReconciled
'        End If
        picClient.Left = 5355.029
        picClient.Top = 140
        picClient.Visible = False
        adoConn.Open getConnectionString
        fmeLoading.Visible = True
        szAllBankBalance = Display2ndTAB(adoConn)
        fmeLoading.Refresh
        fmeLoading.Visible = False
        adoConn.Close
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    FocusControl cmdClientList
End Sub

Private Sub cmdPrint_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim i As Integer

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ViewBankBalance.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
    
  
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
'   Report.ParameterFields(1).AddCurrentValue strClientID
'   Report.ParameterFields(2).AddCurrentValue strPropertyID


   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdPrintReconciled_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim i As Integer

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ViewBankBalanceReconciled.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
    
  
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
'   Report.ParameterFields(1).AddCurrentValue strClientID
'   Report.ParameterFields(2).AddCurrentValue strPropertyID


   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub flxClient_Click()
           Dim adoConn As New ADODB.Connection
           ' On Error GoTo ErrorHandler
            adoConn.Open getConnectionString
            Dim rsCheck As New ADODB.Recordset
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    
                    picClient.Visible = False
                    picClient.Left = 5355.029
                    picClient.Top = 140
                    szAllBankBalance = BankAndBalance(adoConn)
                    FocusControl txtStart1
            ElseIf sTextBox = "2" Then 'bank accounts
                   picClient.Visible = False
                   txtClientList2.Tag = flxClient.TextMatrix(flxClient.row, 1)
                   txtClientList2.text = flxClient.TextMatrix(flxClient.row, 2)
            End If
            adoConn.Close
            Set adoConn = Nothing
            
End Sub
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
Private Sub Form_Load()
'     Me.BackColor = MODULEBACKCOLOR
'     Frame1.BackColor = MODULEBACKCOLOR
'     Frame2.BackColor = MODULEBACKCOLOR
'     fraFinancial.BackColor = MODULEBACKCOLOR
'     chkYtD.BackColor = MODULEBACKCOLOR
    fraFinancial.Top = 1470
    fraFinancial.Left = 140
    fraDateRange.Top = 1470
    fraDateRange.Left = 140
    Me.Width = 19605
    Me.Height = 12345
    txtClientList2.text = "All"
    txtClientList2.Tag = "All"
    txtClientList.text = "ALL"
    txtClientList.Tag = "ALL"
    txtStart1.text = Format("01/01/2000", "dd/mm/yyyy")
    txtEnd1.text = Format(Date, "dd/mm/yyyy")
    txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
    SSTabViewBankBalance.Tab = 0
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    szAllBankBalance = BankAndBalance(adoConn)
    Call WheelHook(Me.hWnd)
End Sub
Public Sub LoadPeriods(adoConn As ADODB.Connection)
  
   Dim adoRst     As New ADODB.Recordset
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


      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      If adoRst.EOF Then GoTo NoRes

      TotalRow = adoRst.RecordCount - 1
      TotalCol = adoRst.Fields.Count - 1
      ReDim Data(TotalCol, TotalRow) As String

      K = -1
      For i = 0 To TotalRow
         For j = 0 To TotalCol
            Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
            If K = -1 And j = 4 Then
               If adoRst.Fields("Status").Value Then
                  K = i
                  dtEnd = CDate(adoRst.Fields("P_EndDate").Value)
               End If
            End If
         Next j
         adoRst.MoveNext
         If adoRst.EOF Then Exit For
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

Private Sub gridBankCode_Click()
    Dim adoConn  As New ADODB.Connection
     adoConn.Open getConnectionString
     If sTextBox = "3" Then
            txtBudgetYears.Tag = gridBankCode.TextMatrix(gridBankCode.row, 3) 'FYID
            txtBudgetYears.text = gridBankCode.TextMatrix(gridBankCode.row, 2) 'FY description
            cmdBudgetYears.Tag = gridBankCode.TextMatrix(gridBankCode.row, 1)  'FY start date
            LoadPeriods adoConn
            picBankCode.Visible = False
            
      End If
      adoConn.Close
      Set adoConn = Nothing
End Sub

Private Sub ConfigConsolidateReconciled()
    Dim szHeader As String
    flxReconciled.Clear
    flxReconciled.Rows = 2
    'szHeader$ = "|<Bank ID|<Bank Code|<Bank Name|<Bank AC Number|<Sort Code|<Statement Date|<Closing Bal|<SOB"
   ' szHeader$ = "|<Client ID|<Bank Code|<Bank Name|<Bank Balance"
   szHeader$ = "|<Client ID|<Bank Account Name|<Bank Account No|<Receipts|< Payments|< Last Reconciled Date |<Last Reconciled Bank Balance" & _
            "|<Last Reconciled Statement Balance|<Current Unreconciled Cashbook balance|<Current Cashbook Balance "
    flxReconciled.FormatString = szHeader$
   flxReconciled.RowHeight(0) = 520
   flxReconciled.Cols = 11
   flxReconciled.ColWidth(0) = 200
   flxReconciled.ColWidth(1) = 1800
   flxReconciled.ColWidth(2) = 3300
   flxReconciled.ColWidth(3) = 1600
   flxReconciled.ColWidth(4) = 1500
   flxReconciled.ColAlignment(4) = vbLeftJustify
   flxReconciled.ColAlignment(5) = vbLeftJustify
   flxReconciled.ColAlignment(6) = vbLeftJustify
   flxReconciled.ColAlignment(7) = vbLeftJustify
   flxReconciled.ColAlignment(8) = vbLeftJustify
   flxReconciled.ColWidth(5) = 1500
   flxReconciled.ColWidth(6) = 1500
   flxReconciled.ColWidth(7) = 1500
   flxReconciled.ColWidth(8) = 1500
   flxReconciled.ColAlignment = vbLeftJustify
End Sub

Private Function LastReconciledBankBalance(adoConn As ADODB.Connection, ByVal ClientID As String, ByVal BankCode As String, ByVal DTdate As Date, ByRef payment As Currency, ByRef receipt As Currency) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

           
    szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & BankCode & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & ClientID & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# group by Type " & _
           "UNION "
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & BankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & ClientID & "' AND " & _
                       "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# group by TRANS " & _
                "UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                       "P.BankCode = '" & BankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & ClientID & "' AND " & _
                       "PDATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                "group by Type;"
'Debug.Print szSQL AND right(BP.ReconNow,4)='Full' AND right(P.ReconNow,4)='Full' AND right(R.ReconNow,4)='Full'
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Exit Function
   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         LastReconciledBankBalance = LastReconciledBankBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         LastReconciledBankBalance = LastReconciledBankBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         LastReconciledBankBalance = LastReconciledBankBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         LastReconciledBankBalance = LastReconciledBankBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         LastReconciledBankBalance = LastReconciledBankBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         LastReconciledBankBalance = LastReconciledBankBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         LastReconciledBankBalance = LastReconciledBankBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         LastReconciledBankBalance = LastReconciledBankBalance + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
   
        szSQL = "Select sum(amt) as amtt from(SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & BankCode & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & ClientID & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# group by Type " & _
           "UNION "
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & BankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & ClientID & "' AND " & _
                       "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# group by TRANS " & _
                "UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 24) AND " & _
                       "P.BankCode = '" & BankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & ClientID & "' AND " & _
                       "PDATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                "group by Type);"

        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not adoRst.EOF Then
                 receipt = IIf(IsNull(adoRst.Fields.Item("amtt").Value), 0, adoRst.Fields.Item("amtt").Value)
        End If
        adoRst.Close
        
        szSQL = "Select sum(amt) as amtt from(SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 23 ) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & BankCode & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & ClientID & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# group by Type " & _
           "UNION "
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11) AND " & _
                       "BP.BANK_AC = '" & BankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & ClientID & "' AND " & _
                       "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# group by TRANS " & _
                "UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 24) AND " & _
                       "P.BankCode = '" & BankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & ClientID & "' AND " & _
                       "PDATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                "group by Type);"

        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not adoRst.EOF Then
                 payment = IIf(IsNull(adoRst.Fields.Item("amtt").Value), 0, adoRst.Fields.Item("amtt").Value)
        End If
        adoRst.Close



End Function

Private Function LastReconciledBankBalanceConsolidated(adoConn As ADODB.Connection, ByVal SelectedConBankID As String, ByVal DTdate As Date, ByRef payment As Currency, ByRef receipt As Currency) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND " & _
                 "B.CLIENT_ID = P.ClientID  AND " & _
                 "R.Amount > 0  AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "#" & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND " & _
                 "P.Amount > 0  AND PDate<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"


                
                
                
'Debug.Print szSQL AND right(BP.ReconNow,4)='Full' AND right(P.ReconNow,4)='Full' AND right(R.ReconNow,4)='Full'
       adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
       receipt = 0
       payment = 0
       While Not adoRst.EOF
          If adoRst.Fields.Item("T").Value = "3" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated + adoRst.Fields.Item("AMT").Value
             receipt = receipt + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "4" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated + adoRst.Fields.Item("AMT").Value
             receipt = receipt + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "8" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated - adoRst.Fields.Item("AMT").Value
             payment = payment + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "9" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated - adoRst.Fields.Item("AMT").Value
             payment = payment + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "BP" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated - adoRst.Fields.Item("AMT").Value
             payment = payment + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "BR" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated + adoRst.Fields.Item("AMT").Value
             receipt = receipt + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "23" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated - adoRst.Fields.Item("AMT").Value
             payment = payment + adoRst.Fields.Item("AMT").Value
          If adoRst.Fields.Item("T").Value = "24" Then _
             LastReconciledBankBalanceConsolidated = LastReconciledBankBalanceConsolidated + adoRst.Fields.Item("AMT").Value
             receipt = receipt + adoRst.Fields.Item("AMT").Value
    
          adoRst.MoveNext
       Wend

            adoRst.Close
            szSQL = "Select sum(amt) as amtt from( SELECT   Sum(Amount) as amt " & _
               "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
               "WHERE (R.Type = 3 OR R.Type = 4) AND " & _
                      "TT.TYPE_ID = R.Type AND " & _
                      "R.BankCode = B.NominalCode AND " & _
                      "U.UnitNumber = R.UnitID AND " & _
                      "U.PropertyID = P.PropertyID AND " & _
                      "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                      "B.NominalCode = R.BankCode AND " & _
                      "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
               "UNION "
            szSQL = szSQL & _
                    "SELECT   Sum(NET_AMOUNT+VAT) as amt " & _
                    "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (BP.TransactionType = 12) AND " & _
                           "BP.BANK_AC =B.NominalCode AND BP.TransactionType = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                    "UNION "
            szSQL = szSQL & _
                    "SELECT  Sum(Amount) as amt " & _
                    "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (P.Type = 24) AND " & _
                           "P.BankCode = B.NominalCode AND P.Type = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID AND PDate<=#" & Format(DTdate, "dd MMM yyyy") & "#  ) " & _
                    ""
           adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not adoRst.EOF Then
                 receipt = adoRst.Fields.Item("amtt").Value
            End If
            adoRst.Close
            szSQL = " select sum(amt) as amtt from( SELECT   Sum(Amount) as amt " & _
               "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
               "WHERE (R.Type = 23) AND " & _
                      "TT.TYPE_ID = R.Type AND " & _
                      "R.BankCode = B.NominalCode AND " & _
                      "U.UnitNumber = R.UnitID AND " & _
                      "U.PropertyID = P.PropertyID AND " & _
                      "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                      "B.NominalCode = R.BankCode AND " & _
                      "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
               "UNION "
            szSQL = szSQL & _
                    "SELECT   Sum(NET_AMOUNT+VAT) as amt " & _
                    "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (BP.TransactionType = 11) AND " & _
                           "BP.BANK_AC =B.NominalCode AND BP.TransactionType = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                    "UNION "
            szSQL = szSQL & _
                    "SELECT  Sum(Amount) as amt " & _
                    "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                    "WHERE (P.Type = 8 OR P.Type = 9) AND " & _
                           "P.BankCode = B.NominalCode AND P.Type = TT.TYPE_ID AND " & _
                           "B.ConsolidatedBankID = " & SelectedConBankID & " AND " & _
                           "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID AND PDate<=#" & Format(DTdate, "dd MMM yyyy") & "#  ) " & _
                    ""
           adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not adoRst.EOF Then
                 payment = adoRst.Fields.Item("amtt").Value
            End If


   adoRst.Close
   
   Set adoRst = Nothing
End Function
Private Function LastStatementBalance(adoConn As ADODB.Connection, ByVal ClientID As String, ByVal BankCode As String, ByVal DTdate As Date) As Currency
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

'   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
'           "FROM tlbReceipt AS R, " & _
'                "Units AS U, Property AS P, tlbClientBanks AS B " & _
'           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
'                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
'                 "B.NominalCode = R.BankCode AND " & _
'                 "B.CLIENT_ID = P.ClientID AND right(R.ReconNow,4)='Full' AND " & _
'                 "R.Amount > 0  AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "#" & _
'           "GROUP BY Type " & _
'           "UNION "
'   szSQL = szSQL & _
'           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
'           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
'           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
'                 "B.NominalCode = P.BankCode AND right(P.ReconNow,4)='Full' AND " & _
'                 "P.Amount > 0  AND PDate<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
'           "GROUP BY TYPE " & _
'           "UNION "
'   szSQL = szSQL & _
'           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
'           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
'           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
'                  "CB.NominalCode = BP.BANK_AC  AND " & _
'               "(BP.NET_AMOUNT + BP.VAT) > 0 AND right(BP.ReconNow,4)='Full' AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
'           "GROUP BY TRANS " & _
'           "ORDER BY T;"
           
           
   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & BankCode & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & ClientID & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID AND RDate<=#" & Format(DTdate, "dd MMM yyyy") & "# group by Type " & _
           "UNION "
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & BankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & ClientID & "' AND right(BP.ReconNow,4)='Full' AND " & _
                       "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID AND TRAN_DATE<=#" & Format(DTdate, "dd MMM yyyy") & "# group by TRANS " & _
                "UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                       "P.BankCode = '" & BankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & ClientID & "' AND right(P.ReconNow,4)='Full' AND " & _
                       "PDATE<=#" & Format(DTdate, "dd MMM yyyy") & "# " & _
                "group by Type;"
'Debug.Print szSQL AND right(BP.ReconNow,4)='Full' AND right(P.ReconNow,4)='Full' AND right(R.ReconNow,4)='Full'
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         LastStatementBalance = LastStatementBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         LastStatementBalance = LastStatementBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         LastStatementBalance = LastStatementBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         LastStatementBalance = LastStatementBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         LastStatementBalance = LastStatementBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         LastStatementBalance = LastStatementBalance + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         LastStatementBalance = LastStatementBalance - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         LastStatementBalance = LastStatementBalance + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
Private Function Display2ndTAB(adoConn As ADODB.Connection) As String
   'On Error GoTo Error_Handler

        Dim iRec As Integer
        Dim adoRst As New ADODB.Recordset
        Dim szSQL As String, szaData() As String
        Dim rRow As Integer
        Dim rsReconciliationDate As New ADODB.Recordset
        Dim rsInsertViewBankBalanceReconciled As New ADODB.Recordset
        Dim dtCutoffDate As String
        
        
        Call ConfigConsolidateReconciled                    'Configuration flexgrid
          
        If Option2.Value = True Then
                If cmbPeriodTo.Value = Null Or cmbPeriodTo.Value = "" Then
                    MsgBox "Please enter Period", vbInformation, "warning"
                    FocusControl cmdBudgetYears
                    Exit Function
                End If
               dtCutoffDate = Format(cmbPeriodTo.text, "dd MMM yyyy")
        Else
                If txtSCYRREnDt.text = "" Then
                    MsgBox "Please enter end date", vbInformation, "warning"
                    FocusControl txtSCYRREnDt
                    Exit Function
                End If
               dtCutoffDate = Format(txtSCYRREnDt.text, "dd MMM yyyy")
        End If
      
        adoConn.Execute "Delete From ReportViewBankBalanceReconciled"
        rsInsertViewBankBalanceReconciled.Open "Select * from ReportViewBankBalanceReconciled", adoConn, adOpenDynamic, adLockOptimistic
If txtClientList2.text = "ALL" Or txtClientList2.text = "All" Or txtClientList2.text = "Consolidated" Then
        
        szSQL = "Select Distinct A.ClientID,A.BankCode,B.conBankID from " & _
                "tlbBankReconClosingBal A INNER JOIN ConsolidatedBankList B ON A.ClientID=B.BankCode "
             adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
             flxReconciled.AddItem ""
             rRow = 1
             Dim dblPayment As Currency
             Dim dblReceipt As Currency
             'Load all columns for all consolidated values
             
                While Not adoRst.EOF
                    rsInsertViewBankBalanceReconciled.AddNew
                    flxReconciled.row = 1
                    flxReconciled.TextMatrix(rRow, 0) = ""
                   
                    flxReconciled.TextMatrix(rRow, 1) = "Consolidated"
                    rsInsertViewBankBalanceReconciled!ClientID = "Consolidated"
                    rsInsertViewBankBalanceReconciled!ClientName = "Consolidated"
                    flxReconciled.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value) ' Bank Code
                    'rsInsertViewBankBalanceReconciled!BankCode = flxReconciled.TextMatrix(rRow, 2)
                    flxReconciled.TextMatrix(rRow, 3) = adoRst.Fields.Item("BankCode").Value
                    rsInsertViewBankBalanceReconciled!BankAccountName = flxReconciled.TextMatrix(rRow, 2)
                    'rsInsertViewBankBalanceReconciled!BankAccountNumber = flxReconciled.TextMatrix(rRow, 2)
                    rsReconciliationDate.Open "Select Max(StatementDate) as MaxDate from tlbBankReconClosingBal where ClientID='" & flxReconciled.TextMatrix(rRow, 2) & "'" & _
                                " and StatementDate<=#" & Format(dtCutoffDate, "dd MMM yyyy") & "#", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsReconciliationDate.EOF Then
                             flxReconciled.TextMatrix(rRow, 6) = Format(rsReconciliationDate.Fields.Item("MaxDate").Value, "dd MMM yyyy")
                    End If
                    If flxReconciled.TextMatrix(rRow, 6) = "" Then
                            flxReconciled.TextMatrix(rRow, 6) = Format(Date, "dd MMM yyyy")
                    End If
                    rsInsertViewBankBalanceReconciled!LastReconciledDate = flxReconciled.TextMatrix(rRow, 6)
                    rsReconciliationDate.Close
                    ''this is similar to cashbook tab
                    dblReceipt = 0
                    dblPayment = 0
                    flxReconciled.TextMatrix(rRow, 7) = Format(LastReconciledBankBalanceConsolidated(adoConn, adoRst.Fields.Item("conBankID").Value, flxReconciled.TextMatrix(rRow, 6), dblPayment, dblReceipt), "0.00")
                    rsInsertViewBankBalanceReconciled!LastReconciledBankBalance = flxReconciled.TextMatrix(rRow, 7)
                    rsInsertViewBankBalanceReconciled!receipt = dblReceipt
                    rsInsertViewBankBalanceReconciled!payment = dblPayment
                    flxReconciled.TextMatrix(rRow, 4) = Format(dblReceipt, "0.00")
                    flxReconciled.TextMatrix(rRow, 5) = Format(dblPayment, "0.00")
                    'Last Reconciled statement Balance LastStatementBalanceConsolidated where reconnow is marked
                    flxReconciled.TextMatrix(rRow, 8) = Format(LastStatementBalanceConsolidated(adoConn, adoRst.Fields.Item("conBankID").Value, flxReconciled.TextMatrix(rRow, 6)), "0.00")
                    rsInsertViewBankBalanceReconciled!LastReconciledStatementBalance = flxReconciled.TextMatrix(rRow, 8)
                    'Bank Balance with date Range 'this need to be newly modify part of Display2ndTAB
                    'flxReconciled.TextMatrix(rRow, 10) = Format(BankAccBalanceConsolidated2ndTAB(adoconn, adoRst.Fields.Item("conBankID").Value), "0.00") 'Current Bank Balance
                   flxReconciled.TextMatrix(rRow, 10) = Format(BankAccBalanceConsolidated2ndTAB(adoConn, adoRst.Fields.Item("conBankID").Value), "0.00") 'Current Bank Balance
                   
                    'flxReconciled.TextMatrix(rRow, 9) = Format(flxReconciled.TextMatrix(rRow, 10) - flxReconciled.TextMatrix(rRow, 8), "0.00") 'Unreconciled Balance
                    flxReconciled.TextMatrix(rRow, 9) = Format(BankAccBalUnreconciledConsolidated2ndTAB(adoConn, adoRst.Fields.Item("conBankID").Value), "0.00") 'Unreconciled Balance
                    'BankAccBalUnreconciledConsolidated2ndTAB
                    rsInsertViewBankBalanceReconciled!UnreconciledCashBookBalance = flxReconciled.TextMatrix(rRow, 9)
                    rsInsertViewBankBalanceReconciled!CashbookcurrentBalance = Format(flxReconciled.TextMatrix(rRow, 10), "0.00")
'                    Debug.Print flxReconciled.TextMatrix(rRow, 3)
                    flxReconciled.RowHeight(rRow) = 280
                    rsInsertViewBankBalanceReconciled.Update
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxReconciled.AddItem ""
                    rRow = rRow + 1
                 Wend
                 adoRst.Close
                 
                 szSQL = "Select Distinct A.ClientID,A.BankCode,C.Bank_AC_Name from " & _
                "(tlbBankReconClosingBal A INNER JOIN Client B ON A.ClientID=B.ClientID) inner join tlbClientBanks C on  C.client_ID=B.clientID and C.NominalCode=A.BankCode "
             adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
             flxReconciled.AddItem ""
             'rRow = 1
'             Dim dblPayment As Currency
'             Dim dblReceipt As Currency
                While Not adoRst.EOF
                    flxReconciled.row = 1
                    flxReconciled.TextMatrix(rRow, 0) = ""
                    flxReconciled.TextMatrix(rRow, 1) = IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value)
                    rsInsertViewBankBalanceReconciled.AddNew
                    rsInsertViewBankBalanceReconciled!ClientID = flxReconciled.TextMatrix(rRow, 1)
                    rsInsertViewBankBalanceReconciled!ClientName = flxReconciled.TextMatrix(rRow, 1)
                    flxReconciled.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("Bank_AC_Name").Value), "", adoRst.Fields.Item("Bank_AC_Name").Value)
                    rsInsertViewBankBalanceReconciled!BankAccountName = flxReconciled.TextMatrix(rRow, 2)
                    flxReconciled.TextMatrix(rRow, 3) = adoRst.Fields.Item("BankCode").Value
                     rsInsertViewBankBalanceReconciled!BankAccountNumber = flxReconciled.TextMatrix(rRow, 3)
                    rsReconciliationDate.Open "Select Max(StatementDate) as MaxDate from tlbBankReconClosingBal where ClientID='" & flxReconciled.TextMatrix(rRow, 1) & "'" & _
                                " AND BankCode='" & flxReconciled.TextMatrix(rRow, 3) & "' and StatementDate<=#" & Format(dtCutoffDate, "dd MMM yyyy") & "# ", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsReconciliationDate.EOF Then
                             flxReconciled.TextMatrix(rRow, 6) = Format(rsReconciliationDate.Fields.Item("MaxDate").Value, "dd MMM yyyy")
                    End If
                    If flxReconciled.TextMatrix(rRow, 6) = "" Then
                            flxReconciled.TextMatrix(rRow, 6) = Format(Date, "dd MMM yyyy")
                    End If
                    rsReconciliationDate.Close
                    rsInsertViewBankBalanceReconciled!LastReconciledDate = flxReconciled.TextMatrix(rRow, 6)
'                    flxReconciled.TextMatrix(rRow, 6) = Format(adoRst.Fields.Item("StatementDate").Value, "dd MMM yyyy")
'                    ''this is similar to cashbook tab
                    dblReceipt = 0
                    dblPayment = 0
                    flxReconciled.TextMatrix(rRow, 7) = LastReconciledBankBalance(adoConn, flxReconciled.TextMatrix(rRow, 1), CStr(flxReconciled.TextMatrix(rRow, 3)), flxReconciled.TextMatrix(rRow, 6), dblPayment, dblReceipt)
                    rsInsertViewBankBalanceReconciled!LastReconciledBankBalance = flxReconciled.TextMatrix(rRow, 7)
                    flxReconciled.TextMatrix(rRow, 4) = Format(dblReceipt, "0.00")
                    flxReconciled.TextMatrix(rRow, 5) = Format(dblPayment, "0.00")
                    rsInsertViewBankBalanceReconciled!receipt = dblReceipt
                    rsInsertViewBankBalanceReconciled!payment = dblPayment
                    flxReconciled.TextMatrix(rRow, 8) = Format(LastStatementBalance(adoConn, flxReconciled.TextMatrix(rRow, 1), flxReconciled.TextMatrix(rRow, 3), flxReconciled.TextMatrix(rRow, 6)), "0.00")
                     rsInsertViewBankBalanceReconciled!LastReconciledStatementBalance = flxReconciled.TextMatrix(rRow, 8)
                    flxReconciled.TextMatrix(rRow, 10) = Format(BankAccBalance(adoConn, flxReconciled.TextMatrix(rRow, 3), flxReconciled.TextMatrix(rRow, 1)), "0.00")
                    flxReconciled.TextMatrix(rRow, 9) = Format(flxReconciled.TextMatrix(rRow, 10) - flxReconciled.TextMatrix(rRow, 8), "0.00") 'Unreconciled Balance
                    rsInsertViewBankBalanceReconciled!UnreconciledCashBookBalance = flxReconciled.TextMatrix(rRow, 9)
                    rsInsertViewBankBalanceReconciled!CashbookcurrentBalance = flxReconciled.TextMatrix(rRow, 10)
                    rsInsertViewBankBalanceReconciled.Update
'                    Debug.Print flxReconciled.TextMatrix(rRow, 3)
                    flxReconciled.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxReconciled.AddItem ""
                    rRow = rRow + 1
                 Wend
                 adoRst.Close
             Else
                rRow = 1
                    szSQL = "Select Distinct A.ClientID,A.BankCode,C.Bank_AC_Name from " & _
                "(tlbBankReconClosingBal A INNER JOIN Client B ON A.ClientID=B.ClientID) inner join tlbClientBanks C on  C.client_ID=B.clientID and C.NominalCode=A.BankCode " & _
                " where A.clientID='" & txtClientList2.Tag & "' "
             adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
             flxReconciled.AddItem ""
             'rRow = 1
'             Dim dblPayment As Currency
'             Dim dblReceipt As Currency
                While Not adoRst.EOF
                    flxReconciled.row = 1
                    flxReconciled.TextMatrix(rRow, 0) = ""
                    flxReconciled.TextMatrix(rRow, 1) = IIf(IsNull(adoRst.Fields.Item("ClientID").Value), "", adoRst.Fields.Item("ClientID").Value)
                    rsInsertViewBankBalanceReconciled.AddNew
                    rsInsertViewBankBalanceReconciled!ClientID = flxReconciled.TextMatrix(rRow, 1)
                    rsInsertViewBankBalanceReconciled!ClientName = flxReconciled.TextMatrix(rRow, 1)
                    flxReconciled.TextMatrix(rRow, 2) = IIf(IsNull(adoRst.Fields.Item("Bank_AC_Name").Value), "", adoRst.Fields.Item("Bank_AC_Name").Value)
                    rsInsertViewBankBalanceReconciled!BankAccountName = flxReconciled.TextMatrix(rRow, 2)
                    flxReconciled.TextMatrix(rRow, 3) = adoRst.Fields.Item("BankCode").Value
                     rsInsertViewBankBalanceReconciled!BankAccountNumber = flxReconciled.TextMatrix(rRow, 3)
                    rsReconciliationDate.Open "Select Max(StatementDate) as MaxDate from tlbBankReconClosingBal where ClientID='" & flxReconciled.TextMatrix(rRow, 1) & "'" & _
                                " AND BankCode='" & flxReconciled.TextMatrix(rRow, 3) & "' and StatementDate<=#" & Format(dtCutoffDate, "dd MMM yyyy") & "# ", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsReconciliationDate.EOF Then
                             flxReconciled.TextMatrix(rRow, 6) = Format(rsReconciliationDate.Fields.Item("MaxDate").Value, "dd MMM yyyy")
                    End If
                    If flxReconciled.TextMatrix(rRow, 6) = "" Then
                            flxReconciled.TextMatrix(rRow, 6) = Format(Date, "dd MMM yyyy")
                    End If
                    rsReconciliationDate.Close
                    rsInsertViewBankBalanceReconciled!LastReconciledDate = flxReconciled.TextMatrix(rRow, 6)
'                    flxReconciled.TextMatrix(rRow, 6) = Format(adoRst.Fields.Item("StatementDate").Value, "dd MMM yyyy")
'                    ''this is similar to cashbook tab
                    dblReceipt = 0
                    dblPayment = 0
                    flxReconciled.TextMatrix(rRow, 7) = LastReconciledBankBalance(adoConn, flxReconciled.TextMatrix(rRow, 1), CStr(flxReconciled.TextMatrix(rRow, 3)), flxReconciled.TextMatrix(rRow, 6), dblPayment, dblReceipt)
                    rsInsertViewBankBalanceReconciled!LastReconciledBankBalance = flxReconciled.TextMatrix(rRow, 7)
                    flxReconciled.TextMatrix(rRow, 4) = Format(dblReceipt, "0.00")
                    flxReconciled.TextMatrix(rRow, 5) = Format(dblPayment, "0.00")
                    rsInsertViewBankBalanceReconciled!receipt = dblReceipt
                    rsInsertViewBankBalanceReconciled!payment = dblPayment
                    flxReconciled.TextMatrix(rRow, 8) = Format(LastStatementBalance(adoConn, flxReconciled.TextMatrix(rRow, 1), flxReconciled.TextMatrix(rRow, 3), flxReconciled.TextMatrix(rRow, 6)), "0.00")
                     rsInsertViewBankBalanceReconciled!LastReconciledStatementBalance = flxReconciled.TextMatrix(rRow, 8)
                    flxReconciled.TextMatrix(rRow, 10) = Format(BankAccBalance(adoConn, flxReconciled.TextMatrix(rRow, 3), flxReconciled.TextMatrix(rRow, 1)), "0.00")
                    flxReconciled.TextMatrix(rRow, 9) = Format(flxReconciled.TextMatrix(rRow, 10) - flxReconciled.TextMatrix(rRow, 8), "0.00") 'Unreconciled Balance
                    rsInsertViewBankBalanceReconciled!UnreconciledCashBookBalance = flxReconciled.TextMatrix(rRow, 9)
                    rsInsertViewBankBalanceReconciled!CashbookcurrentBalance = flxReconciled.TextMatrix(rRow, 10)
                    rsInsertViewBankBalanceReconciled.Update
'                    Debug.Print flxReconciled.TextMatrix(rRow, 3)
                    flxReconciled.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxReconciled.AddItem ""
                    rRow = rRow + 1
                 Wend
                 adoRst.Close
                 
                 
             End If
             rsInsertViewBankBalanceReconciled.Close
             
                 
   
End Function
Private Sub Option1_Click()
    fraFinancial.Visible = False
    fraDateRange.Visible = True
    fraFinancial.Left = 100
    fraDateRange.Left = 100
    FocusControl txtSCYRRStDt
End Sub
Private Sub Option2_Click()
    fraFinancial.Visible = True
    fraDateRange.Visible = False
    fraFinancial.Left = 100
    fraDateRange.Left = 100
    FocusControl cmdBudgetYears
End Sub

Private Sub SSTabViewBankBalance_Click(PreviousTab As Integer)
'    If SSTabViewBankBalance.Tab = 0 Then
'        fraFinancial.Visible = False
'     Else
'        fraFinancial.Visible = True
'   End If
End Sub

Private Sub txtSCYRREnDt_Change()
    TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
   If Len(txtSCYRREnDt.text) < 10 Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdOK
    End If
    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
    If txtSCYRREnDt.text <> "" Then TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
   If Len(txtSCYRRStDt.text) < 10 Then txtSCYRRStDt.text = Format("01/01/2000", "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtSCYRREnDt
    End If
    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
    If txtSCYRRStDt.text <> "" Then TextBoxFormatDate txtSCYRRStDt
    ' If txtSCYRRStDt.text <> "" Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
End Sub

'*********************
Private Sub txtEnd1_Change()
    TextBoxChangeDate txtEnd1
End Sub

Private Sub txtEnd1_GotFocus()
   If Len(txtEnd1.text) < 10 Then txtEnd1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtEnd1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdDisplay1
    End If
    TextBoxKeyPrsDate txtEnd1, KeyAscii
End Sub

Private Sub txtEnd1_LostFocus()
    If txtEnd1.text <> "" Then TextBoxFormatDate txtEnd1
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
Private Sub txtStart1_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtStart1_GotFocus()
   If Len(txtStart1.text) < 10 Then txtStart1.text = Format("01/01/2000", "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtStart1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtEnd1
    End If
    TextBoxKeyPrsDate txtStart1, KeyAscii
End Sub

Private Sub txtStart1_LostFocus()
    If txtStart1.text <> "" Then TextBoxFormatDate txtStart1
     If txtStart1.text <> "" Then txtEnd1.text = Format(Date, "dd/mm/yyyy")
End Sub

