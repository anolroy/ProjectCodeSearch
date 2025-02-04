VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNJ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nominal Journal"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16290
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNJ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   2220
      Left            =   3555
      TabIndex        =   72
      Top             =   5490
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         Height          =   2100
         Index           =   0
         Left            =   40
         ScaleHeight     =   2040
         ScaleWidth      =   3555
         TabIndex        =   73
         Top             =   50
         Width           =   3615
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   135
            TabIndex        =   81
            Top             =   1605
            Width           =   1200
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2055
            TabIndex        =   83
            Top             =   1635
            Width           =   1200
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   75
            Top             =   450
            Width           =   2685
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   76
            Top             =   790
            Width           =   2685
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   80
            TabIndex        =   77
            Top             =   1125
            Width           =   1290
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2025
            MaxLength       =   80
            TabIndex        =   79
            Top             =   1125
            Width           =   1380
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
            TabIndex        =   74
            Top             =   0
            Width           =   255
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0FFFF&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   30
            Left            =   0
            Top             =   260
            Width           =   3855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Search Options"
            Height          =   195
            Index           =   9
            Left            =   765
            TabIndex        =   84
            Top             =   45
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   82
            Top             =   495
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   80
            Top             =   810
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   12
            Left            =   135
            TabIndex        =   78
            Top             =   1125
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5220
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   58
      Top             =   8280
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
         TabIndex        =   59
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   60
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
         TabIndex        =   66
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   65
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
         Left            =   1650
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   61
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
         Width           =   4950
      End
   End
   Begin TabDlg.SSTab tabNJ 
      Height          =   7815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      MousePointer    =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nominal Journal"
      TabPicture(0)   =   "frmNJ.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGridCaption(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblGridCaption(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblGridCaption(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblGridCaption(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblGridCaption(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkSelAll"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblGridCaption(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "flxNJ"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Nominal Journal History"
      TabPicture(1)   =   "frmNJ.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSearchHistory"
      Tab(1).Control(1)=   "cmdRevJournal"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdClose(1)"
      Tab(1).Control(4)=   "flxNJ_Hist"
      Tab(1).Control(5)=   "flxNJ_Split"
      Tab(1).Control(6)=   "chkSelallHistory"
      Tab(1).Control(7)=   "lblGridCaption(22)"
      Tab(1).Control(8)=   "lblGridCaption(18)"
      Tab(1).Control(9)=   "lblGridCaption(17)"
      Tab(1).Control(10)=   "lblGridCaption(16)"
      Tab(1).Control(11)=   "lblGridCaption(21)"
      Tab(1).Control(12)=   "lblGridCaption(19)"
      Tab(1).Control(13)=   "lblGridCaption(20)"
      Tab(1).Control(14)=   "lblGridCaption(11)"
      Tab(1).Control(15)=   "lblGridCaption(12)"
      Tab(1).Control(16)=   "lblGridCaption(13)"
      Tab(1).Control(17)=   "lblGridCaption(14)"
      Tab(1).Control(18)=   "lblGridCaption(15)"
      Tab(1).Control(19)=   "lblGridCaption(9)"
      Tab(1).Control(20)=   "lblGridCaption(7)"
      Tab(1).ControlCount=   21
      Begin VB.CommandButton cmdSearchHistory 
         Caption         =   "Sea&rch"
         Height          =   435
         Left            =   -73155
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   7155
         Width           =   1080
      End
      Begin VB.CommandButton cmdRevJournal 
         Caption         =   "Reverse History"
         Height          =   435
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   7155
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         Height          =   7335
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1335
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Sea&rch"
            Height          =   375
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   5130
            Width           =   1080
         End
         Begin VB.CommandButton cmdPrintNJ 
            Caption         =   "Print Journal"
            Height          =   435
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2856
            Width           =   1080
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   435
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1608
            Width           =   1080
         End
         Begin VB.CommandButton cmdPrintNJList 
            Caption         =   "Print List"
            Height          =   435
            Left            =   1125
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4095
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "C&lose"
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   6600
            Width           =   1095
         End
         Begin VB.CommandButton cmdPostHist 
            Caption         =   "&Post to Hist."
            Height          =   435
            Left            =   120
            TabIndex        =   10
            Top             =   4095
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "&Add New"
            Height          =   435
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filter"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   12375
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
            Left            =   5895
            TabIndex        =   70
            Top             =   720
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
            Left            =   5895
            TabIndex        =   68
            Top             =   315
            Width           =   300
         End
         Begin VB.CommandButton cmdClearHist 
            Caption         =   "Cle&ar"
            Height          =   315
            Left            =   10560
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtDateFromHist 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   21
            Top             =   320
            Width           =   1455
         End
         Begin VB.TextBox txtDateToHist 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilterHist 
            Caption         =   "&Ok"
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
            Left            =   10575
            TabIndex        =   23
            Top             =   320
            Width           =   1455
         End
         Begin MSForms.TextBox txtPropertyList 
            Height          =   285
            Left            =   1080
            TabIndex        =   69
            Top             =   720
            Width           =   4815
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "8493;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtClientList 
            Height          =   285
            Left            =   1080
            TabIndex        =   67
            Top             =   315
            Width           =   4815
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "8493;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propert&y:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Client:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   39
            Top             =   320
            Width           =   465
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date From:"
            Height          =   195
            Index           =   3
            Left            =   7560
            TabIndex        =   38
            Top             =   320
            Width           =   765
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date To:"
            Height          =   195
            Index           =   2
            Left            =   7560
            TabIndex        =   37
            Top             =   720
            Width           =   585
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   435
         Index           =   1
         Left            =   -64200
         TabIndex        =   25
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filter"
         Height          =   975
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   10935
         Begin VB.CheckBox chkProperty 
            Caption         =   "Excl."
            Height          =   195
            Left            =   6120
            TabIndex        =   71
            Top             =   270
            Width           =   780
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
            Left            =   5805
            TabIndex        =   2
            Top             =   225
            Width           =   300
         End
         Begin VB.CommandButton cmdClient 
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
            Left            =   2835
            TabIndex        =   1
            Top             =   225
            Width           =   300
         End
         Begin VB.TextBox txtDateTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9840
            TabIndex        =   4
            Top             =   240
            Width           =   1000
         End
         Begin VB.TextBox txtPropID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "&Ok"
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
            Left            =   8160
            TabIndex        =   5
            Top             =   600
            Width           =   1245
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clea&r"
            Height          =   315
            Left            =   9595
            TabIndex        =   6
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtDateFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8160
            TabIndex        =   3
            Top             =   240
            Width           =   1000
         End
         Begin VB.TextBox txtClientID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSForms.TextBox txtProperty 
            Height          =   285
            Left            =   3960
            TabIndex        =   57
            Top             =   225
            Width           =   1845
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3254;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtClient 
            Height          =   285
            Left            =   630
            TabIndex        =   56
            Top             =   225
            Width           =   2295
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "4048;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date To:"
            Height          =   195
            Index           =   1
            Left            =   9240
            TabIndex        =   42
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date From:"
            Height          =   195
            Index           =   0
            Left            =   7320
            TabIndex        =   17
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Client:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propert&y:"
            Height          =   195
            Index           =   4
            Left            =   3195
            TabIndex        =   15
            Top             =   240
            Width           =   645
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNJ 
         Height          =   6000
         Left            =   1560
         TabIndex        =   18
         Top             =   1680
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10583
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNJ_Hist 
         Height          =   3120
         Left            =   -74865
         TabIndex        =   20
         Top             =   1935
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   5503
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNJ_Split 
         Height          =   1680
         Left            =   -74865
         TabIndex        =   46
         Top             =   5445
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   2963
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
      Begin MSForms.CheckBox chkSelallHistory 
         Height          =   255
         Left            =   -74865
         TabIndex        =   87
         Top             =   1620
         Width           =   375
         VariousPropertyBits=   746588179
         BackColor       =   15781855
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "661;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   195
         Index           =   22
         Left            =   -63975
         TabIndex        =   53
         Top             =   5140
         Width           =   1065
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   18
         Left            =   -70335
         TabIndex        =   52
         Top             =   5140
         Width           =   360
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   17
         Left            =   -73695
         TabIndex        =   51
         Top             =   5140
         Width           =   840
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal A/C"
         Height          =   255
         Index           =   16
         Left            =   -74880
         TabIndex        =   50
         Top             =   5140
         Width           =   1095
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Amount"
         Height          =   195
         Index           =   21
         Left            =   -64935
         TabIndex        =   49
         Top             =   5140
         Width           =   885
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   19
         Left            =   -67695
         TabIndex        =   48
         Top             =   5140
         Width           =   1305
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   20
         Left            =   -66375
         TabIndex        =   47
         Top             =   5140
         Width           =   1290
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   11
         Left            =   -74520
         TabIndex        =   45
         Top             =   1650
         Width           =   210
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1840
         TabIndex        =   44
         Top             =   1410
         Width           =   495
      End
      Begin MSForms.CheckBox chkSelAll 
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1370
         Width           =   375
         VariousPropertyBits=   746588179
         BackColor       =   15781855
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "661;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   255
         Index           =   12
         Left            =   -73785
         TabIndex        =   34
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   13
         Left            =   -71520
         TabIndex        =   33
         Top             =   1650
         Width           =   615
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   14
         Left            =   -69240
         TabIndex        =   32
         Top             =   1650
         Width           =   345
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   195
         Index           =   15
         Left            =   -68160
         TabIndex        =   31
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   195
         Index           =   5
         Left            =   8280
         TabIndex        =   29
         Top             =   1410
         Width           =   330
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   4
         Left            =   7200
         TabIndex        =   28
         Top             =   1410
         Width           =   345
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   27
         Top             =   1410
         Width           =   615
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   255
         Index           =   2
         Left            =   2895
         TabIndex        =   26
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H80000000&
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
         Index           =   0
         Left            =   1560
         TabIndex        =   30
         Top             =   1365
         Width           =   10935
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H80000000&
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
         Index           =   9
         Left            =   -74880
         TabIndex        =   35
         Top             =   1600
         Width           =   12375
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H80000000&
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
         Index           =   7
         Left            =   -74880
         TabIndex        =   54
         Top             =   5100
         Width           =   12375
      End
   End
End
Attribute VB_Name = "frmNJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTextBox As String
Public UserSessionID As String
Public iSelRow As Integer
Dim colTransactionIDOtherPIGrid As String
Private Sub chkProperty_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    If chkProperty.Value = 0 Then
        txtProperty.text = "ALL"
        cmdProperty.Enabled = True
    Else
        txtProperty.text = ""
        cmdProperty.Enabled = False
    End If
    LoadFlxNJ1 adoconn, txtClient.text, txtProperty.text
    adoconn.Close
    Set adoconn = Nothing
End Sub

Private Sub chkSelallHistory_Click()
    Dim iRow As Integer
   
   If Not chkSelallHistory.Value Then
      For iRow = 1 To flxNJ_Hist.Rows - 1
         flxNJ_Hist.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxNJ_Hist.Rows - 1
         flxNJ_Hist.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub cmdSearchCancel_Click()
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        txtSearchRef.text = ""
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        LoadFlxNJ adoconn, ""
        adoconn.Close
        Set adoconn = Nothing
        cmdSearch.Caption = "Sea&rch"
        fraSearch.Visible = False
End Sub

Private Sub cmdSearchHistory_Click()
    fraSearch.Left = 3555
    fraSearch.Top = 5490
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'
    
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    If cmdSearch.Caption = "Clear Sea&rch" Then
         txtSearchNo.text = ""
         txtSearchRef.text = ""
         'fmeLoading.Visible = False
         cmdSearch.Caption = "Sea&rch"
         fraSearch.Visible = False

    Else
        If fraSearch.Visible = False Then
            fraSearch.Visible = True
            txtSearchNo.SetFocus
        Else
            fraSearch.Visible = False
        End If
    End If
'    adoconn.Close
End Sub

Private Sub flxNJ_RowColChange()
    Call InstantLockingCheck
End Sub

Private Sub txtSearchFromD_Change()
    TextBoxChangeDate txtSearchFromD
    txtSearchNo.text = ""
    txtSearchRef.text = ""
End Sub

Private Sub txtSearchFromD_GotFocus()
'    If Len(txtSearchFromD.text) < 10 Then txtSearchFromD.text = Format(Date, "dd/mm/yyyy")
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
Private Sub txtSearchToD_Change()
     TextBoxChangeDate txtSearchToD
     txtSearchNo.text = ""
     txtSearchRef.text = ""
End Sub

Private Sub txtSearchToD_GotFocus()
'    If Len(txtSearchToD.text) < 10 Then txtSearchToD.text = Format(Date, "dd/mm/yyyy")
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
Private Sub chkSelAll_Click()
   Dim iRow As Integer
   
   If Not chkSelAll.Value Then
      For iRow = 1 To flxNJ.Rows - 1
         flxNJ.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxNJ.Rows - 1
         flxNJ.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub chkSelAll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   Dim iRow As Integer
'
'   If Not chkSelAll.Value Then
'      For iRow = 1 To flxNJ.Rows - 1
'         flxNJ.TextMatrix(iRow, 0) = ""
'      Next iRow
'   Else
'      For iRow = 1 To flxNJ.Rows - 1
'         flxNJ.TextMatrix(iRow, 0) = "X"
'      Next iRow
'   End If

'   For iRow = 1 To flxNJ.Rows - 1
'      If flxNJ.RowHeight(iRow) > 0 Then
'         If Not chkSelAll.Value Then
'            'SelectFlxGridRow 0, flxNJ, iRow
'         Else
'            flxNJ.TextMatrix(iRow, 0) = "X"
'           ' SelectFlxGridRow 0, flxNJ, iRow
'         End If
'      Else
'         flxNJ.TextMatrix(iRow, 0) = "X"
'        ' SelectFlxGridRow 0, flxNJ, iRow
'      End If
'   Next iRow
End Sub

'Private Sub cmbClient_Click()
'   Dim adoConn    As New ADODB.Connection
'
'   txtClientID.text = cmbClient.Value
''   connect to database
'   adoConn.Open getConnectionString
'   LoadCmbProperty adoConn, cmbProperty, txtClientID.text
'   txtPropID.text = ""
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub cmbClientHist_Click()
'   Dim adoConn    As New ADODB.Connection
'
'   txtClientIDHist.text = txtClientList.Tag
'   adoConn.Open getConnectionString
'   LoadCmbProperty adoConn, cmbPropHist, txtClientIDHist.text
'   txtPropIDHist.text = ""
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub cmbProperty_Click()
'   txtPropID.text = cmbProperty.Value
'End Sub

'Private Sub cmbPropHist_Click()
'   txtPropIDHist.text = cmbPropHist.Value
'End Sub

Public Sub cmdAddNew_Click()
   Dim adoconn    As New ADODB.Connection

   adoconn.Open getConnectionString

   If Not AreCA_Setup(adoconn) Then
      ShowMsgInTaskBar "Please setup control accounts for the client(s)", "Y", "N"
      adoconn.Close
      Set adoconn = Nothing
      Exit Sub
   End If

   adoconn.Close
   Set adoconn = Nothing
   frmNJ_Entry.bEditMode = False
   Load frmNJ_Entry
   frmNJ_Entry.lHeaderID = 0
   frmNJ_Entry.Left = 100
   frmNJ_Entry.Top = 200
   frmNJ_Entry.bFunction = True
'   frmNJ_Entry.Show
    LoadForm frmNJ_Entry
   Me.Enabled = False
End Sub

Private Sub cmdClear_Click()
'   cmbClient.ListIndex = -1
'   txtClientID.text = ""
'   cmbProperty.ListIndex = -1
   txtPropID.text = ""
   txtDateFrom.text = ""
   txtDateTo.text = ""

   cmdFilter_Click
End Sub

Private Sub cmdClearHist_Click()
    txtClientList.text = "ALL Client"
    txtClientList.Tag = "ALL"
    txtPropertyList.text = "ALL Properties"
    txtPropertyList.Tag = "ALL"
    
    txtDateFromHist.text = ""
    txtDateToHist.text = ""
    
    cmdFilterHist_Click
End Sub

Private Sub cmdClient_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 270
    picClient.Visible = True
    LoadflxClient
    tabNJ.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClientList_Click()
    sTextBox = "3"
    picClient.Left = 915
    picClient.Top = 570
    picClient.Visible = True
    LoadflxClient
    tabNJ.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClose_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub
Private Function InstantLockingCheck() As Boolean 'unlocking for all row
    Dim adoPay As New ADODB.Recordset
    Dim rsLockDialog As New ADODB.Recordset
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String, iRow As Integer
    Dim selRow As Integer
    Dim selcol As Integer
    Dim i As Integer
    Dim j As Integer
    Dim strSQL As String
    Dim szRecordID As String
    Dim tempstr As String
    szRecordID = flxNJ.TextMatrix(flxNJ.row, 1)
    tempstr = Replace(UCase(szRecordID), "'", "''")
    tempstr = Replace(UCase(tempstr), "N", "")
    tempstr = Replace(UCase(tempstr), "L", "")
    tempstr = Replace(UCase(tempstr), "J", "")
    szRecordID = tempstr
    
    selRow = flxNJ.row
    selcol = flxNJ.col
   If IsNumeric(szRecordID) = False Then Exit Function
   
   adoconn.Open getConnectionString
   
'   ' I am doing some test here
'   ' on loading time full table vs selected row
'
'    szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MYid= '" & flxNJ.TextMatrix(flxNJ.row, 10) & "'"
'
'   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   adoPay.Close
'   szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment"
'
'   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   adoPay.Close
'
'   Exit Sub
   'first part is instant lock
   szSQL = " SELECT BP.RecordID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM NJ_Header BP WHERE RecordID= " & szRecordID & ""
   
   adoPay.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

  'locking status show for current row
   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then   'szSQL <> UserSessionID shall be always true bcoz PI is generating only one session thrgh out this module.
                flxNJ.col = 0
                'flxNJ.row = flxNJ.row
                flxNJ.CellBackColor = vbRed
                InstantLockingCheck = False
                colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoPay("RecordID").Value), "", adoPay("RecordID").Value) & ","
            Else 'lock for this user
                        flxNJ.col = 0
                        i = flxNJ.row
                        flxNJ.CellBackColor = vbWhite
    '                    adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
    '                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxNJ.TextMatrix(iPIEdit, 0) & "'"
                        'Need to clear the locking flag
                        flxNJ.TextMatrix(i, 8) = ""
                        flxNJ.TextMatrix(i, 9) = ""
                        flxNJ.TextMatrix(i, 10) = ""
                        flxNJ.TextMatrix(i, 11) = ""
            End If
           
   End If
   'second part instant unlock
       If Len(colTransactionIDOtherPIGrid) > 0 Then
            szSQL = "SELECT RecordID,UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
                 "FROM NJ_Header where (isnull(UserSessionID) OR UserSessionID='') " & _
                 " AND RecordID in (" & colTransactionIDOtherPIGrid & ") order by 1,2 Desc;"
            rsLockDialog.Open szSQL, adoconn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
             While Not rsLockDialog.EOF
                      flxNJ.col = 0
                      For j = 1 To flxNJ.Rows - 1
                          If flxNJ.TextMatrix(j, 1) = "NLJ" & rsLockDialog("RecordID").Value And i <> j Then 'no need to update row of first part check
                                flxNJ.row = j
                                flxNJ.CellBackColor = vbWhite
                          End If
                       Next j
                    rsLockDialog.MoveNext
              Wend
        End If
        'second part ends here
        
        flxNJ.col = selcol
        flxNJ.row = selRow
        
        
        
        
   adoPay.Close
   adoconn.Close
   flxNJ.row = selRow
   flxNJ.col = selcol
   Set adoPay = Nothing
   Set adoconn = Nothing
End Function

Private Function IsPossible2Edit() As Boolean
    Dim adoPay As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String, iRow As Integer
    Dim szRecordID As String
    Dim tempstr As String
    szRecordID = flxNJ.TextMatrix(flxNJ.row, 1)
    tempstr = Replace(UCase(szRecordID), "'", "''")
    tempstr = Replace(UCase(tempstr), "N", "")
    tempstr = Replace(UCase(tempstr), "L", "")
    tempstr = Replace(UCase(tempstr), "J", "")
    szRecordID = tempstr
                

   adoconn.Open getConnectionString
   szSQL = " SELECT BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM NJ_Header BP WHERE RecordID= " & szRecordID & ""
   adoPay.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then 'szSQL <> UserSessionID shall be always true bcoz PI is generating only one seeion through out this module.
                flxNJ.col = 0
                flxNJ.CellBackColor = vbRed
                MsgBox "The selected invoice is currently locked by '" & IIf(IsNull(adoPay("WindowsUserName").Value), "", adoPay("WindowsUserName").Value) & _
                "' on '" & IIf(IsNull(adoPay("MachineName").Value), "", adoPay("MachineName").Value) & "' in the '" & IIf(IsNull(adoPay("Module").Value), "", adoPay("Module").Value) & "'" & vbCrLf & "" & _
                        "screen for the Client '" & IIf(IsNull(adoPay("ClientID").Value), "", adoPay("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
                IsPossible2Edit = False
            Else 'lock row for this user in database
                flxNJ.col = 0
                IsPossible2Edit = True
                flxNJ.CellBackColor = vbWhite
            End If
   End If
   adoPay.Close
   Set adoPay = Nothing
   
   If IsPossible2Edit = True Then 'the reason I need to lock it here because it has passed all the tests here to open PI and now I can lock
             adoconn.Execute "Update NJ_Header Set  DateTimeStamp='" & Now & "',Module='Nominal Journal',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where RecordID = " & szRecordID & ""
   End If
   adoconn.Close
   Set adoconn = Nothing
End Function
Private Sub cmdEdit_Click()
   Dim iRow    As Integer
   Dim k       As Integer
  
   k = 0
   For iRow = 1 To flxNJ.Rows - 1
      If flxNJ.TextMatrix(iRow, 0) = "X" Then
         k = k + 1
         iSelRow = iRow
         flxNJ.row = iRow
      End If
   Next iRow
   If k = 0 Then
      MsgBox "Please select a journal to edit", vbInformation, "Information"
      Exit Sub
   End If
   If k > 1 Then
      MsgBox "Please select only one journal to edit", vbInformation, "Information"
      For iRow = 1 To flxNJ.Rows - 1
            flxNJ.TextMatrix(iRow, 0) = ""
      Next iRow
      Exit Sub
   End If
   
   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   If Not IsPossible2Edit Then
        Exit Sub
   End If
  
   Load frmNJ_Entry
   frmNJ_Entry.bEditMode = True
   frmNJ_Entry.cmdSave.Enabled = False
   frmNJ_Entry.bEdit = "1"
   frmNJ_Entry.bFunction = False
   frmNJ_Entry.lHeaderID = Mid(flxNJ.TextMatrix(flxNJ.row, 1), 4)
   frmNJ_Entry.lblNJ_Id.Caption = "Journal No: " & flxNJ.TextMatrix(flxNJ.row, 1)
   frmNJ_Entry.Caption = "Nominal Journal: " & flxNJ.TextMatrix(flxNJ.row, 1)
   frmNJ_Entry.lblNJ_Id.Visible = True
   frmNJ_Entry.Left = 100
   frmNJ_Entry.Top = 100
   
   frmNJ_Entry.Show
   Me.Enabled = False
End Sub

Private Sub cmdFilter_Click()
   Dim iRow As Integer

   For iRow = 1 To flxNJ.Rows - 1
      flxNJ.RowHeight(iRow) = 240
   Next iRow

   For iRow = 1 To flxNJ.Rows - 1
      If flxNJ.TextMatrix(iRow, 1) = "" Then Exit For
      If txtClientID.text <> "" Then
         If flxNJ.TextMatrix(iRow, 6) <> txtClientID.text Then flxNJ.RowHeight(iRow) = 0
      End If
      If txtPropID.text <> "" Then
         If flxNJ.TextMatrix(iRow, 7) <> txtPropID.text And flxNJ.TextMatrix(iRow, 7) <> "" Then flxNJ.RowHeight(iRow) = 0
      End If
      If txtDateFrom.text <> "" Then
         If CDate(flxNJ.TextMatrix(iRow, 4)) < CDate(txtDateFrom.text) Then flxNJ.RowHeight(iRow) = 0
      End If
      If txtDateTo.text <> "" Then
         If CDate(flxNJ.TextMatrix(iRow, 4)) > CDate(txtDateTo.text) Then flxNJ.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub

Private Sub cmdFilterHist_Click()
   Dim iRow As Integer

   For iRow = 1 To flxNJ_Hist.Rows - 1
      flxNJ_Hist.RowHeight(iRow) = 240
   Next iRow

   For iRow = 1 To flxNJ_Hist.Rows - 1
      If flxNJ_Hist.TextMatrix(iRow, 1) = "" Then Exit For
      If txtClientList.Tag <> "" And txtClientList.Tag <> "ALL" Then
         If flxNJ_Hist.TextMatrix(iRow, 6) <> txtClientList.Tag Then flxNJ_Hist.RowHeight(iRow) = 0
      End If
      If txtPropertyList.Tag <> "" And txtPropertyList.Tag <> "ALL" Then
         If flxNJ_Hist.TextMatrix(iRow, 7) <> txtPropertyList.Tag And flxNJ_Hist.TextMatrix(iRow, 7) <> "" Then flxNJ_Hist.RowHeight(iRow) = 0
      End If
      If txtDateFromHist.text <> "" Then
         If CDate(flxNJ_Hist.TextMatrix(iRow, 4)) < CDate(txtDateFromHist.text) Then flxNJ_Hist.RowHeight(iRow) = 0
      End If
      If txtDateToHist.text <> "" Then
         If CDate(flxNJ_Hist.TextMatrix(iRow, 4)) > CDate(txtDateToHist.text) Then flxNJ_Hist.RowHeight(iRow) = 0
      End If
   Next iRow

   For iRow = 1 To flxNJ_Split.Rows - 1
      flxNJ_Split.RowHeight(iRow) = 0
   Next iRow
End Sub

Private Sub cmdPostHist_Click()
   Dim szID As String
   Dim iRow As Integer

   For iRow = 1 To flxNJ.Rows - 1
      If flxNJ.TextMatrix(iRow, 0) = "X" And flxNJ.RowHeight(iRow) > 0 And flxNJ.TextMatrix(iRow, 1) <> "" Then
         szID = Mid(flxNJ.TextMatrix(iRow, 1), 4) & ", " & szID
      End If
   Next iRow

   If szID = "" Then Exit Sub

   If MsgBox("Do you wish to post selected transactions to history?", vbQuestion + vbYesNo, "Nominal Journal History") = vbNo Then Exit Sub

   szID = Left(szID, Len(szID) - 2)

   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

   adoconn.Execute "UPDATE NJ_Header SET History = TRUE WHERE RecordID IN (" & szID & ");"

   Call LoadFlxNJ(adoconn, "")       'Reload the grid
   Call LoadFlxNJ_Hist(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing

   chkSelAll.Value = 0
   ShowMsgInTaskBar (iRow - 1) & " Transactions have been posted to history successfully", "Y", "P"
End Sub

Private Sub cmdPrintNJ_Click()
'   Dim iRow    As Integer
'   Dim K       As Integer
'
'   K = 0
'   For iRow = 1 To flxNJ.Rows - 1
'      If flxNJ.TextMatrix(iRow, 0) = "X" Then
'         K = K + 1
'         flxNJ.row = iRow
'      End If
'   Next iRow
'   If K = 0 Then
'      MsgBox "Please select a journal to edit", vbInformation, "Information"
'      Exit Sub
'   End If
'   If K > 1 Then
'      MsgBox "Please select only one journal to print", vbInformation, "Information"
'      Exit Sub
'   End If
''   MsgBox Val(Mid(flxNJ.TextMatrix(flxNJ.row, 1), 4))
''    Exit Sub
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NominalJournal.rpt")
'
'   Report.EnableParameterPrompting = False
'   Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue Val(Mid(flxNJ.TextMatrix(flxNJ.row, 1), 4))
'
'   Load frmReport
'   frmReport.LoadReportViewer Report

    cmdPrintNJList_Click
End Sub

Private Sub cmdPrintNJList_Click()

'   If Not chkSelAll.Value Then chkSelAll_MouseDown 1, 0, 0, 0
'
   Dim iRow As Integer
   Dim szID As String
   Dim k    As Integer

   For iRow = flxNJ.Rows - 1 To 1 Step -1
      If flxNJ.TextMatrix(iRow, 0) = "X" And flxNJ.RowHeight(iRow) > 0 Then
         k = k + 1
         szID = Mid(flxNJ.TextMatrix(iRow, 1), 4) & ", " & szID
      End If
   Next iRow

   

   If szID = "" Then
      MsgBox "No Journal has been selected", vbInformation, "Warning"
      Exit Sub
   End If
   
   szID = Left(szID, Len(szID) - 2)
    'Below part has been comment out by anol 17 Sep 2015
   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString
    'Debug.Print "UPDATE NJ_Header SET History = TRUE WHERE RecordID IN (" & szID & ");"
   adoconn.Execute "UPDATE NJ_Header SET PrintThis = FALSE;"
   adoconn.Execute "UPDATE NJ_Header SET PrintThis = TRUE WHERE RecordID IN (" & szID & ");"
    
  

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NominalJournalList.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report

'
'    Load frmNLAnalysis
'    frmNLAnalysis.Caption = "NL Listing Report"
'    frmNLAnalysis.LOOKUPparam = "NLListing"
'    frmNLAnalysis.Show
 'if I only code for color fix this will take huge time rather than loading the whole list
   ConfigFlxNJ
   Call LoadFlxNJ(adoconn, "")
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdproperty_Click()
    picClient.Left = 7100
    picClient.Top = 270
    picClient.Visible = True
    sTextBox = "2"
    LoadPropertyList
    tabNJ.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   
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
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345
   
   'lblJobName.Visible = False
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

     
      If sTextBox = "1" Or sTextBox = "3" Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 280
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub

Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345
   'lblJobName.Visible = False
   
      adoconn.Open getConnectionString
           
     If sTextBox = "2" Then
             If txtClient.text = "ALL" Then
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "ORDER BY PropertyID;"
             Else
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "WHERE ClientID = '" & txtClient.text & "' " & _
                 "ORDER BY PropertyID;"

             End If
     ElseIf sTextBox = "4" Then
          If txtClientList.Tag = "ALL" Then
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "ORDER BY PropertyID;"
             Else
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "WHERE ClientID = '" & txtClientList.text & "' " & _
                 "ORDER BY PropertyID;"

             End If

    End If

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
  
   If sTextBox = "2" Or sTextBox = "4" Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Property"
           flxClient.AddItem ""
           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

Private Sub cmdPropertyList_Click()
    picClient.Left = 915
    picClient.Top = 770
    picClient.Visible = True
    sTextBox = "4"
    LoadPropertyList
    tabNJ.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdRevJournal_Click()
    Dim iRow As Integer
   Dim szID As String
   Dim k    As Integer

   For iRow = flxNJ_Hist.Rows - 1 To 1 Step -1
      If flxNJ_Hist.TextMatrix(iRow, 0) = "X" And flxNJ_Hist.RowHeight(iRow) > 0 Then
         k = k + 1
         szID = Mid(flxNJ_Hist.TextMatrix(iRow, 1), 4) & ", " & szID
      End If
   Next iRow
   If szID <> "" Then
        szID = Left(szID, Len(szID) - 2)
   End If

   If szID = "" Then
      ShowMsgInTaskBar "No Journal has been selected", "Y", "N"
      Exit Sub
   End If

   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString
'Debug.Print "UPDATE NJ_Header SET History = TRUE WHERE RecordID IN (" & szID & ");"
   'adoConn.Execute "UPDATE NJ_Header SET History = FALSE;"
   adoconn.Execute "UPDATE NJ_Header SET History = False WHERE RecordID IN (" & szID & ");"
   Call LoadFlxNJ(adoconn, "")
   Call LoadFlxNJ_Hist(adoconn, "")
   flxNJ_Split.Clear
   flxNJ_Split.Rows = 1
   ShowMsgInTaskBar k & " Records has been transferred", "Y", "N"
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub cmdSearch_Click()
'        Dim adoconn As New ADODB.Connection
'        adoconn.Open getConnectionString
        fraSearch.Left = 1665
        fraSearch.Top = 5625
        
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        If cmdSearch.Caption = "Clear Sea&rch" Then
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             'fmeLoading.Visible = False
             cmdSearch.Caption = "Sea&rch"
             fraSearch.Visible = False
'             Call LoadFlxACHistory(adoconn, "")
        Else
            If fraSearch.Visible = False Then
                fraSearch.Visible = True
                txtSearchNo.SetFocus
            Else
                fraSearch.Visible = False
            End If
        End If
'        adoconn.Close
End Sub

Private Sub cmdSearchOK_Click()
        fraSearch.Visible = False
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        If Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
            If tabNJ.Tab = 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
                LoadFlxNJ adoconn, "3"
            Else
                cmdSearchHistory.Caption = "Clear Sea&rch"
                LoadFlxNJ_Hist adoconn, "3"
            End If
        End If
        adoconn.Close
        Set adoconn = Nothing
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Dim adoconn As New ADODB.Connection
         tabNJ.Enabled = True
         ConfigFlxNJ
        
          adoconn.Open getConnectionString
         If sTextBox = "1" Then
             txtClient.text = flxClient.TextMatrix(flxClient.row, 0)
             LoadFlxNJ1 adoconn, txtClient.text, txtProperty.text
             cmdClient.SetFocus
         ElseIf sTextBox = "2" Then
             txtProperty.text = flxClient.TextMatrix(flxClient.row, 0)
             LoadFlxNJ1 adoconn, txtClient.text, txtProperty.text
             txtPropID.text = flxClient.TextMatrix(flxClient.row, 0)
             cmdProperty.SetFocus
         End If
         adoconn.Close
         Set adoconn = Nothing
         picClient.Visible = False
    End If
End Sub

Private Sub flxNJ_Click()
'   SelectFlxGridRow 7, flxNJ, flxNJ.row
    If flxNJ.TextMatrix(flxNJ.row, 0) = "X" Then
        flxNJ.TextMatrix(flxNJ.row, 0) = ""
        flxNJ.CellBackColor = vbWhite
        Exit Sub
    Else
        flxNJ.TextMatrix(flxNJ.row, 0) = "X"
    End If
    'SelectFlxGridRow 0, flxNJ, flxNJ.row
'    If flxNJ.TextMatrix(flxNJ.row, 1) <> "" Then
'        SelectOnly1RowFlxGrid flxNJ, flxNJ.row
'    End If
   If chkSelAll.Value Then chkSelAll.Value = 0
End Sub

Private Sub flxNJ_DblClick()
   flxNJ.TextMatrix(flxNJ.row, 0) = "X"
   cmdEdit_Click
End Sub

Public Sub ReloadFlxNJ()
   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString

   Call LoadFlxNJ(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub flxNJ_Hist_Click()
'    If flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = "X" Then
'        flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = ""
'        flxNJ_Hist.CellBackColor = vbWhite
'        Exit Sub
'    End If
 Dim iRow As Integer
   
'   If Not chkSelAll.Value Then
'      For iRow = 1 To flxNJ.Rows - 1
'         flxNJ.TextMatrix(iRow, 0) = ""
'      Next iRow
'   Else
'      For iRow = 1 To flxNJ.Rows - 1
'         flxNJ.TextMatrix(iRow, 0) = "X"
'      Next iRow
'   End If
    If flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = "X" Then
       flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = ""
    Else
       flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = "X"
    End If
'    Debug.Print time
    ConfigFlxNJ_Split
'    Debug.Print time
    'SelectFlxGridRow 0, flxNJ_Hist, flxNJ_Hist.row
       If flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 0) = "X" Then
             Call loadNJHIstDeatail(Replace(flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 1), "NLJ", ""))
       Else
            flxNJ_Split.Clear
            flxNJ_Split.Rows = 2
       End If
'      Debug.Print time
End Sub

Private Sub flxNJ_Hist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   flxNJ_Hist.MousePointer = flexArrow
End Sub

Private Sub flxNJ_Hist_RowColChange()
'   Dim iRow As Integer
'
'   For iRow = 1 To flxNJ_Split.Rows - 1
'      If flxNJ_Split.TextMatrix(iRow, 0) = flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 1) Then
'         flxNJ_Split.RowHeight(iRow) = 240
'      Else
'         flxNJ_Split.RowHeight(iRow) = 0
'      End If
'   Next iRow
End Sub

Private Sub flxNJ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   flxNJ.MousePointer = flexArrow
End Sub

Private Sub ConfigFlxNJ_Split()
   Dim szHeader As String, iCol As Integer

   flxNJ_Split.Clear
   flxNJ_Split.Cols = 8
   flxNJ_Split.Rows = 2
   flxNJ_Split.RowHeight(0) = 0

   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund" & _
               "|>Dr|>Cr|>VATAmt|>TotalAmt"

   flxNJ_Split.FormatString = szHeader$
   flxNJ_Split.ColWidth(0) = 0
   For iCol = 1 To flxNJ_Split.Cols - 2
      flxNJ_Split.ColWidth(iCol) = lblGridCaption(iCol + 1 + 15).Left - lblGridCaption(iCol + 15).Left
      lblGridCaption(iCol + 15).Width = flxNJ_Split.ColWidth(iCol)
   Next iCol

   flxNJ_Split.ColWidth(iCol) = flxNJ_Split.Width + flxNJ_Split.Left - lblGridCaption(22).Left - 340  'TotalAmount
   lblGridCaption(iCol + 15).Width = flxNJ_Split.ColWidth(iCol)
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 13035
   Me.Height = 8520
   tabNJ.Tab = 0
   UserSessionID = GetTimeStamp
   Me.BackColor = MODULEBACKCOLOR
   tabNJ.BackColor = MODULEBACKCOLOR

   ConfigFlxNJ
   ConfigFlxNJ_Hist
   ConfigFlxNJ_Split

    Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    Dim szSQL As String
    adoconn.Open getConnectionString
 ' loading the first client

'   szSQL = "SELECT  CLIENTID " & _
'           "FROM  Client " & _
'           "ORDER BY CLIENTID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   If adoRst.RecordCount = 0 Then
'            MsgBox "You must create a Client before entering Nominal.", vbCritical + vbOKOnly, "Nominal Journal"
'            adoRst.Close
'            Set adoRst = Nothing
'            adoConn.Close
'            Set adoConn = Nothing
'            Exit Sub
'   Else
            txtClient.text = "ALL"
            'adoRst.Close
            txtProperty.text = "ALL"
            txtClientList.text = "All Client"
             txtClientList.Tag = "ALL"
             txtPropertyList.text = "ALL Properties"
              txtPropertyList.Tag = "ALL"
                    
'   End If
   'End of adtion
   'LoadCmbClient adoConn, cmbClient
'   LoadCmbClient adoConn, cmbClientHist
   Call LoadFlxNJ(adoconn, "")
   Call LoadFlxNJ_Hist(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigFlxNJ_Hist()
   Dim szHeader As String

   flxNJ_Hist.Clear
   flxNJ_Hist.Cols = 8
   flxNJ_Hist.Rows = 2
   flxNJ_Hist.RowHeight(0) = 0

   szHeader$ = "<HeaderID|<Client|<Property|<Date|<Title" & _
               "|CleintID|PropertyID"

   flxNJ_Hist.FormatString = szHeader$
   flxNJ_Hist.ColWidth(0) = 250
   flxNJ_Hist.ColWidth(1) = lblGridCaption(12).Left - lblGridCaption(11).Left
   flxNJ_Hist.ColWidth(2) = lblGridCaption(13).Left - lblGridCaption(12).Left
   flxNJ_Hist.ColWidth(3) = lblGridCaption(14).Left - lblGridCaption(13).Left
   flxNJ_Hist.ColWidth(4) = lblGridCaption(15).Left - lblGridCaption(14).Left
   flxNJ_Hist.ColWidth(5) = flxNJ_Hist.Width + flxNJ_Hist.Left - lblGridCaption(15).Left - 340
   flxNJ_Hist.ColWidth(6) = 0
   flxNJ_Hist.ColWidth(7) = 0
End Sub

Private Sub ConfigFlxNJ()
   Dim szHeader As String

   flxNJ.Clear
   flxNJ.Cols = 12
   flxNJ.Rows = 2
   flxNJ.RowHeight(0) = 0

   szHeader$ = "X|<HeaderID|<Client|<Property|<Date|<Title" & _
               "|CleintID|PropertyID"

   flxNJ.FormatString = szHeader$
   flxNJ.ColWidth(0) = lblGridCaption(1).Left - flxNJ.Left                          'X --> SelectionFlag
   flxNJ.ColWidth(1) = lblGridCaption(2).Left - lblGridCaption(1).Left              'HeaderID
   flxNJ.ColWidth(2) = lblGridCaption(3).Left - lblGridCaption(2).Left              'Client
   flxNJ.ColWidth(3) = lblGridCaption(4).Left - lblGridCaption(3).Left              'Property
   flxNJ.ColWidth(4) = lblGridCaption(5).Left - lblGridCaption(4).Left              'Date
   flxNJ.ColWidth(5) = flxNJ.Width + flxNJ.Left - lblGridCaption(5).Left - 340      'Title
   flxNJ.ColWidth(6) = 0                                                            'CleintID
   flxNJ.ColWidth(7) = 0                                                            'PropertyID
   
   flxNJ.ColWidth(8) = 0
   flxNJ.ColWidth(9) = 0
   flxNJ.ColWidth(10) = 0
   flxNJ.ColWidth(11) = 0
End Sub

Private Sub LoadFlxNJ_Hist(adoconn As ADODB.Connection, Filter As String)
   Dim adoRst  As New ADODB.Recordset
   Dim iRow    As Integer
   Dim jRow    As Integer
   Dim szSQL   As String
   Dim lID     As Long
   Dim tempstr As String
   flxNJ_Hist.Rows = 1
   flxNJ_Hist.AddItem ""
   iRow = 1
   jRow = 1

'   szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
'           "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
'                 "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
'           "WHERE History;"
'   szSQL = "SELECT DISTINCT S2.*, C.ClientName, P.PropertyName, S3.NC, S3.SpLineDes, S3.NetAmt, " & _
'                  "S3.VATAmt, S3.TotalAmt, N.TRANSACTION_TYPE, F.FundName " & _
'           "FROM ((((NJ_Header AS S2 INNER JOIN Client AS C ON S2.ClientID = C.ClientID) " & _
'                 "LEFT JOIN Property AS P ON S2.PropertyID = P.PropertyID) INNER JOIN NJ_Split AS S3 " & _
'                 "ON S2.RecordID = S3.ParentID) INNER JOIN NLPosting AS N ON N.PARENT_RECORD = S3.RecordID) " & _
'                 "INNER JOIN Fund AS F ON F.FundID = S3.FundID " & _
'           "WHERE History;"
 szSQL = "SELECT DISTINCT S2.*, C.ClientName, P.PropertyName " & _
           "FROM ((NJ_Header AS S2 INNER JOIN Client AS C ON S2.ClientID = C.ClientID) " & _
                 "LEFT JOIN Property AS P ON S2.PropertyID = P.PropertyID) WHERE History order by RecordID DESC"
'Debug.Print szSQL
  ' adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
If Filter = "1" Then
        If txtSearchNo.text <> "" Then
                tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
                tempstr = Replace(UCase(txtSearchNo.text), "N", "")
                tempstr = Replace(UCase(tempstr), "L", "")
                tempstr = Replace(UCase(tempstr), "J", "")
                  szSQL = "SELECT S.*,'NLJ'& S.RecordID as InvID, C.ClientName, P.PropertyName " & _
                "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                      "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                "WHERE NOT History and RecordID Like '%" & tempstr & "%' order by RecordID DESC;"
                If tempstr = "" Then
                        szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
                              "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                                    "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                              "WHERE History order by RecordID DESC;"
                End If
        End If
    End If
    If Filter = "3" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
                    "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                    "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                    "WHERE  History AND S.NJDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
                    "S.NJDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#  order by RecordID DESC;"
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
    
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoRst.Filter = "NJTitle Like '%" & tempstr & "%'"
        End If
    End If
    
    If adoRst.RecordCount = 0 Then
        flxNJ_Hist.Rows = 2
    Else
        flxNJ_Hist.Rows = adoRst.RecordCount + 1
    End If
   While Not adoRst.EOF
      flxNJ_Hist.TextMatrix(iRow, 0) = ""
      flxNJ_Hist.TextMatrix(iRow, 1) = "NLJ" & adoRst.Fields.Item("RecordID").Value
      flxNJ_Hist.TextMatrix(iRow, 2) = adoRst.Fields.Item("ClientName").Value
      flxNJ_Hist.TextMatrix(iRow, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyName").Value), "", adoRst.Fields.Item("PropertyName").Value)
      flxNJ_Hist.TextMatrix(iRow, 4) = Format(adoRst.Fields.Item("NJDate").Value, "dd/mm/yyyy")
      flxNJ_Hist.TextMatrix(iRow, 5) = adoRst.Fields.Item("NJTitle").Value
      flxNJ_Hist.TextMatrix(iRow, 6) = adoRst.Fields.Item("ClientID").Value
      flxNJ_Hist.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)
      iRow = iRow + 1
      

      adoRst.MoveNext
      'If Not adoRst.EOF Then flxNJ_Hist.AddItem ""
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub loadNJHIstDeatail(id As String)
    '      szHeader$ = "HeaderID|<NominalCode|<Description|<Fund" & _
'                  "|>Dr|>Cr|>VATAmt|>TotalAmt"
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    Dim adoconn As New ADODB.Connection
    Dim jRow As Integer
    adoconn.Open getConnectionString
     Debug.Print 1
     Debug.Print time
    szSQL = "SELECT * from  NLPosting where trans_ID='210'"
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    adoRst.Close
     Debug.Print 2
     Debug.Print time
    szSQL = "SELECT DISTINCT S2.*, C.ClientName, P.PropertyName, S3.NC, S3.SpLineDes, S3.NetAmt, " & _
                  "S3.VATAmt, S3.TotalAmt, N.TRANSACTION_TYPE, F.FundName " & _
           "FROM ((((NJ_Header AS S2 INNER JOIN Client AS C ON S2.ClientID = C.ClientID) " & _
                 "LEFT JOIN Property AS P ON S2.PropertyID = P.PropertyID) INNER JOIN NJ_Split AS S3 " & _
                 "ON S2.RecordID = S3.ParentID) INNER JOIN NLPosting AS N ON N.PARENT_RECORD = S3.RecordID) " & _
                 "INNER JOIN Fund AS F ON F.FundID = S3.FundID " & _
           "WHERE History and S2.recordID=" & id & " ;"
           'AND NJ_Header.RecordID='" & Mid(flxNJ_Hist.TextMatrix(flxNJ_Hist.row, 1), 4) & "'
            'Debug.Print time
            adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
             Debug.Print time
      jRow = 1
      Do While Not adoRst.EOF
       
            flxNJ_Split.TextMatrix(jRow, 0) = "NLJ" & adoRst.Fields.Item("RecordID").Value
            flxNJ_Split.TextMatrix(jRow, 1) = adoRst.Fields.Item("NC").Value
            flxNJ_Split.TextMatrix(jRow, 2) = adoRst.Fields.Item("SpLineDes").Value
            flxNJ_Split.TextMatrix(jRow, 3) = adoRst.Fields.Item("FundName").Value
            If Val(adoRst.Fields.Item("TRANSACTION_TYPE").Value) = 15 Then
               flxNJ_Split.TextMatrix(jRow, 4) = Format(adoRst.Fields.Item("NetAmt").Value, "0.00")
            Else
               flxNJ_Split.TextMatrix(jRow, 5) = Format(adoRst.Fields.Item("NetAmt").Value, "0.00")
            End If
            flxNJ_Split.TextMatrix(jRow, 6) = Format(adoRst.Fields.Item("VATAmt").Value, "0.00")
            flxNJ_Split.TextMatrix(jRow, 7) = Format(adoRst.Fields.Item("TotalAmt").Value, "0.00")
            adoRst.MoveNext
'            flxNJ_Split.RowHeight(jRow) = 0
            jRow = jRow + 1
            If Not adoRst.EOF Then flxNJ_Split.AddItem ""
         
      Loop
      adoRst.Close
      adoconn.Close
End Sub
Public Sub LoadFlxNJ(adoconn As ADODB.Connection, Filter As String)
   Dim adoRst As New ADODB.Recordset
   Dim iRow As Integer
   Dim szSQL As String
   Dim tempstr As String

   flxNJ.Rows = 1
   flxNJ.AddItem ""
   iRow = 1
   
   
    szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
           "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                 "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
           "WHERE NOT History order by RecordID DESC;"

  
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
                tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
                tempstr = Replace(UCase(txtSearchNo.text), "N", "")
                tempstr = Replace(UCase(tempstr), "L", "")
                tempstr = Replace(UCase(tempstr), "J", "")
                  szSQL = "SELECT S.*,'NLJ'& S.RecordID as InvID, C.ClientName, P.PropertyName " & _
                "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                      "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                "WHERE NOT History and RecordID Like '%" & tempstr & "%' order by RecordID DESC;"
                If tempstr = "" Then
                        szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
                              "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                                    "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                              "WHERE NOT History order by RecordID DESC;"
                End If
        End If
    End If
    If Filter = "3" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
                    "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
                    "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
                    "WHERE NOT History AND S.NJDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
                    "S.NJDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#  order by RecordID DESC;"
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
    
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoRst.Filter = "NJTitle Like '%" & tempstr & "%'"
        End If
    End If
    colTransactionIDOtherPIGrid = ""
    If adoRst.RecordCount = 0 Then
        flxNJ.Rows = 2
    Else
        flxNJ.Rows = adoRst.RecordCount + 1
    End If
   While Not adoRst.EOF
      flxNJ.TextMatrix(iRow, 1) = "NLJ" & CStr(adoRst.Fields.Item("RecordID").Value)
     
        flxNJ.TextMatrix(iRow, 2) = adoRst.Fields.Item("ClientName").Value
        flxNJ.TextMatrix(iRow, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyName").Value), "", adoRst.Fields.Item("PropertyName").Value)
        flxNJ.TextMatrix(iRow, 4) = Format(adoRst.Fields.Item("NJDate").Value, "dd/mm/yyyy")
        flxNJ.TextMatrix(iRow, 5) = adoRst.Fields.Item("NJTitle").Value
        flxNJ.TextMatrix(iRow, 6) = adoRst.Fields.Item("ClientID").Value
        flxNJ.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)
        flxNJ.TextMatrix(iRow, 8) = IIf(IsNull(adoRst.Fields.Item("UserSessionID").Value), "", adoRst.Fields.Item("UserSessionID").Value)
        If flxNJ.TextMatrix(iRow, 8) <> "" Then
            flxNJ.col = 0
            flxNJ.row = iRow
            flxNJ.CellBackColor = vbRed
            colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & CStr(adoRst.Fields.Item("RecordID").Value) & ","
        End If
        flxNJ.TextMatrix(iRow, 9) = IIf(IsNull(adoRst.Fields.Item("WindowsUserName").Value), "", adoRst.Fields.Item("WindowsUserName").Value)
        flxNJ.TextMatrix(iRow, 10) = IIf(IsNull(adoRst.Fields.Item("MachineName").Value), "", adoRst.Fields.Item("MachineName").Value)
        flxNJ.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("Module").Value), "", adoRst.Fields.Item("Module").Value)
        iRow = iRow + 1
        adoRst.MoveNext
      'If Not adoRst.EOF Then flxNJ.AddItem ""BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module'colTransactionIDOtherPIGrid
   Wend
    If Len(colTransactionIDOtherPIGrid) > 0 Then
          colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
    End If
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadCmbProperty(adoconn As ADODB.Connection, cboP As Control, szClientID)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & szClientID & "' " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboP.Clear
   cboP.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadCmbClient(adoconn As ADODB.Connection, cboC As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame1.MousePointer = vbArrow
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)
   Frame2.MousePointer = vbArrow
End Sub



Private Sub Label50_Click(Index As Integer)
   If Index = 4 Then cmdProperty.SetFocus
   If Index = 5 Then cmdClient.SetFocus
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub tabNJ_Click(PreviousTab As Integer)
'   If tabNJ.Tab = 1 Then cmbClientHist.SetFocus
End Sub

Private Sub tabNJ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = vbArrow
   tabNJ.MousePointer = vbArrow
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   If txtDateFrom.text = "dd/mm/yyyy" Then
      txtDateFrom.text = ""
      Exit Sub
   End If
   If Len(txtDateFrom.text) < 10 Then txtDateFrom.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   If txtDateFrom.text = "" Then Exit Sub

   If TextBoxFormatDate(txtDateFrom) Then
      If txtDateTo.text <> "" Then
         If CDate(txtDateFrom.text) > CDate(txtDateTo.text) Then
            txtDateFrom.text = ""
            ShowMsgInTaskBar "From date should not be after the To date.", "Y", "N"
            txtDateFrom.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtDateFromHist_Change()
   TextBoxChangeDate txtDateFromHist
End Sub

Private Sub txtDateFromHist_GotFocus()
   If txtDateFromHist.text = "dd/mm/yyyy" Then
      txtDateFromHist.text = ""
      Exit Sub
   End If
   If Len(txtDateFromHist.text) < 10 Then txtDateFromHist.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDateFromHist
End Sub

Private Sub txtDateFromHist_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFromHist, KeyAscii
End Sub

Private Sub txtDateFromHist_LostFocus()
   If txtDateFromHist.text = "" Then Exit Sub

   If TextBoxFormatDate(txtDateFromHist) Then
      If txtDateTo.text <> "" Then
         If CDate(txtDateFromHist.text) > CDate(txtDateTo.text) Then
            txtDateFromHist.text = ""
            ShowMsgInTaskBar "From date should not be after the To date.", "Y", "N"
            txtDateFromHist.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   If txtDateTo.text = "dd/mm/yyyy" Then
      txtDateTo.text = ""
      Exit Sub
   End If
   If Len(txtDateTo.text) < 10 Then txtDateTo.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   If txtDateTo.text = "" Then Exit Sub

   If TextBoxFormatDate(txtDateTo) Then
      If txtDateFrom.text <> "" Then
         If CDate(txtDateFrom.text) > CDate(txtDateTo.text) Then
            txtDateTo.text = ""
            ShowMsgInTaskBar "To date should not be before the From date.", "Y", "N"
            txtDateTo.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtDateToHist_Change()
   TextBoxChangeDate txtDateToHist
End Sub

Private Sub txtDateToHist_GotFocus()
   If txtDateToHist.text = "dd/mm/yyyy" Then
      txtDateToHist.text = ""
      Exit Sub
   End If
   If Len(txtDateToHist.text) < 10 Then txtDateToHist.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDateToHist
End Sub

Private Sub txtDateToHist_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateToHist, KeyAscii
End Sub

Private Sub txtDateToHist_LostFocus()
   If txtDateToHist.text = "" Then Exit Sub

   If TextBoxFormatDate(txtDateToHist) Then
      If txtDateFrom.text <> "" Then
         If CDate(txtDateFrom.text) > CDate(txtDateToHist.text) Then
            txtDateToHist.text = ""
            ShowMsgInTaskBar "To date should not be before the From date.", "Y", "N"
            txtDateToHist.SetFocus
         End If
      End If
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

Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
      flxClient.RowHeight(i) = 240
     ' If sTextBox = "1" Then
            If InStr(1, UCase(flxClient.TextMatrix(i, 0)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
                flxClient.RowHeight(i) = 0
            End If
       'End If
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
          '  If sTextBox = "11" Or sTextBox = "14" Then
                
                'flxClient.SetFocus
'                flxClient.row = 1
'                flxClient.RowSel = 1
'                flxClient.ColSel = 1
'                flxClient.CellBackColor = RGB(174, 179, 233)
                
'                Dim iRow As Integer
'                flxClient.row = 1
'                For iRow = 1 To flxClient.Cols - 1
'                   flxClient.col = iRow
'                   flxClient.CellBackColor = RGB(174, 179, 233)
'                Next iRow
               ' SelectOnly1RowFlxGrid flxClient, 1 'flxClient.row
           ' Else
                txtSearchClientName.SetFocus
           
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'         txtSearchClientName.SetFocus
'    End If
    If KeyAscii = 27 Then
          flxClient.Clear
          flxClient.Cols = 2
          flxClient.Rows = 2
          picClient.Visible = False
          tabNJ.Enabled = True
          If sTextBox = "1" Then
                cmdClient.SetFocus
          ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
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
            If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
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

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'         flxClient.SetFocus
'    End If
End Sub
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    tabNJ.Enabled = True
    cmdProperty.SetFocus
End Sub
Private Sub flxClient_Click()
    Dim adoconn As New ADODB.Connection
    tabNJ.Enabled = True
    ConfigFlxNJ
   
     adoconn.Open getConnectionString
    If sTextBox = "1" Then
        txtClient.text = flxClient.TextMatrix(flxClient.row, 0)
        LoadFlxNJ1 adoconn, txtClient.text, txtProperty.text
        txtProperty.Tag = "ALL"
        txtProperty.text = "ALL Properties"
        cmdClient.SetFocus
    ElseIf sTextBox = "2" Then
        txtProperty.text = flxClient.TextMatrix(flxClient.row, 0)
        LoadFlxNJ1 adoconn, txtClient.text, txtProperty.text
        txtPropID.text = flxClient.TextMatrix(flxClient.row, 0)
        cmdProperty.SetFocus
    ElseIf sTextBox = "3" Then
        txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 0)
        txtClientList.text = flxClient.TextMatrix(flxClient.row, 1)
        txtPropertyList.Tag = "ALL"
        txtPropertyList.text = "ALL Properties"
        flxNJ_Split.Clear
    ElseIf sTextBox = "4" Then
        txtPropertyList.Tag = flxClient.TextMatrix(flxClient.row, 0)
        txtPropertyList.text = flxClient.TextMatrix(flxClient.row, 1)
        flxNJ_Split.Clear
    End If
    adoconn.Close
    Set adoconn = Nothing
    picClient.Visible = False
End Sub
Private Sub SortTheGrid(flxGrid As MSHFlexGrid, cmbClientCombo As Control, cmbPropCombo As Control, cmbSuppCombo As Control)
   Dim sFlag As Single, iRow As Integer
   Dim szSQL As String

   sFlag = 0
   For iRow = 1 To flxGrid.Rows - 1
      If cmbClientCombo.text = "ALL" Then
         sFlag = 100
      Else
         If flxGrid.TextMatrix(iRow, 14) = cmbClientCombo.text Then sFlag = 100
      End If

      If Len(cmbPropCombo.text) > 0 Then
         If cmbPropCombo.text = "ALL" Then
            sFlag = sFlag + 10
         Else
            If flxGrid.TextMatrix(iRow, 11) = cmbPropCombo.text Then sFlag = sFlag + 10
         End If
      Else
         sFlag = sFlag + 10
      End If

     ' If Len(txtSupplierSearc.text) > 0 Then
         If cmbSuppCombo.text = "ALL" Then
            sFlag = sFlag + 1
         Else
            If flxGrid.TextMatrix(iRow, 5) = cmbSuppCombo.text Then sFlag = sFlag + 1
         End If
'      Else
'         sFlag = sFlag + 1
'      End If

      If sFlag = 111 Then
         flxGrid.RowHeight(iRow) = 240
      Else
         flxGrid.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub
Private Sub LoadFlxNJ1(adoconn As ADODB.Connection, clientID As String, propertyID As String)
   Dim adoRst As New ADODB.Recordset
   Dim iRow As Integer
   Dim szSQL As String

   flxNJ.Rows = 1
   flxNJ.AddItem ""
   iRow = 1

         szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
           "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID ) " & _
                 "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
           "WHERE NOT History;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If clientID = "ALL" And propertyID <> "ALL" Then
            adoRst.Filter = " PropertyID ='" & propertyID & "' "
   ElseIf clientID <> "ALL" And propertyID = "ALL" Then
            adoRst.Filter = " ClientID='" & clientID & "'"
   ElseIf clientID <> "ALL" And propertyID <> "ALL" Then
            adoRst.Filter = " ClientID='" & clientID & "' AND propertyID ='" & propertyID & "' "
   End If
            
   While Not adoRst.EOF
      flxNJ.TextMatrix(iRow, 1) = "NLJ" & CStr(adoRst.Fields.Item("RecordID").Value)
      flxNJ.TextMatrix(iRow, 2) = adoRst.Fields.Item("ClientName").Value
      flxNJ.TextMatrix(iRow, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyName").Value), "", adoRst.Fields.Item("PropertyName").Value)
      flxNJ.TextMatrix(iRow, 4) = Format(adoRst.Fields.Item("NJDate").Value, "dd/mm/yyyy")
      flxNJ.TextMatrix(iRow, 5) = adoRst.Fields.Item("NJTitle").Value
      flxNJ.TextMatrix(iRow, 6) = adoRst.Fields.Item("ClientID").Value
      flxNJ.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)
      iRow = iRow + 1
      adoRst.MoveNext
      If Not adoRst.EOF Then flxNJ.AddItem ""
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub txtSearchNo_Change()
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        txtSearchRef.text = ""
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        If tabNJ.Tab = 0 Then
            If Len(txtSearchNo.text) > 0 Then
                LoadFlxNJ adoconn, "1"
            Else
                LoadFlxNJ adoconn, ""        'False - uploading history, which are already posted and printed and exported to sage
            End If
           
            If Len(txtSearchNo.text) > 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
            Else
                cmdSearch.Caption = "Sea&rch"
            End If
        Else
            If Len(txtSearchNo.text) > 0 Then
                LoadFlxNJ_Hist adoconn, "1"
            Else
                LoadFlxNJ_Hist adoconn, ""        'False - uploading history, which are already posted and printed and exported to sage
            End If
           
            If Len(txtSearchNo.text) > 0 Then
                cmdSearchHistory.Caption = "Clear Sea&rch"
            Else
                cmdSearchHistory.Caption = "Sea&rch"
            End If
        
        End If
         adoconn.Close
            Set adoconn = Nothing
End Sub

Private Sub txtSearchRef_Change()
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    txtSearchNo.text = ""
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    If tabNJ.Tab = 0 Then
        If Len(txtSearchRef.text) > 0 Then
            LoadFlxNJ adoconn, "2"
        Else
            LoadFlxNJ adoconn, ""      'False - uploading history, which are already posted and printed and exported to sage
        End If
        
        
        If Len(txtSearchRef.text) > 0 Then
             cmdSearch.Caption = "Clear Sea&rch"
        Else
             cmdSearch.Caption = "Sea&rch"
        End If
     Else
            If Len(txtSearchRef.text) > 0 Then
                LoadFlxNJ_Hist adoconn, "2"
            Else
                LoadFlxNJ_Hist adoconn, ""        'False - uploading history, which are already posted and printed and exported to sage
            End If
           
            If Len(txtSearchRef.text) > 0 Then
                cmdSearchHistory.Caption = "Clear Sea&rch"
            Else
                cmdSearchHistory.Caption = "Sea&rch"
            End If
     End If
     adoconn.Close
     Set adoconn = Nothing
    
End Sub
