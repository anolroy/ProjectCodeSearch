VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BB5807FE-DBD2-11D3-87C1-4C980CC10374}#1.0#0"; "MyHover.ocx"
Begin VB.Form frmMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance"
   ClientHeight    =   12495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15825
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12495
   ScaleWidth      =   15825
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   2160
      ScaleHeight     =   4425
      ScaleWidth      =   5580
      TabIndex        =   66
      Top             =   2970
      Visible         =   0   'False
      Width           =   5610
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3750
         Left            =   45
         TabIndex        =   68
         Top             =   675
         Width           =   5490
         _ExtentX        =   9684
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   74
         Top             =   375
         Width           =   3915
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6906;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   69
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5220
      End
   End
   Begin VB.Frame fraJS_PO 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1440
      TabIndex        =   59
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton cmdAsPO 
         Caption         =   "Job P/Order"
         Height          =   300
         Left            =   60
         TabIndex        =   62
         Top             =   420
         Width           =   1335
      End
      Begin VB.CommandButton cmdAsJS 
         Caption         =   "Job Sheet"
         Height          =   300
         Left            =   60
         TabIndex        =   61
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuoteReq 
         Caption         =   "Job Quote Req"
         Height          =   300
         Left            =   60
         TabIndex        =   60
         Top             =   780
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Height          =   1140
         Left            =   0
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter"
      Height          =   1455
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   14655
      Begin VB.CommandButton cmdSupplier 
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
         Left            =   10035
         TabIndex        =   4
         Top             =   675
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
         Left            =   4455
         TabIndex        =   1
         Top             =   570
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
         Left            =   4455
         TabIndex        =   0
         Top             =   225
         Width           =   300
      End
      Begin VB.CommandButton cmdUnitLookup 
         Caption         =   """"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1035
         Width           =   315
      End
      Begin VB.TextBox txtUnitNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
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
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   49
         Top             =   1035
         Width           =   3600
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11610
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11610
         MaxLength       =   10
         TabIndex        =   7
         Top             =   660
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Ok"
         Height          =   355
         Left            =   13290
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   660
         Width           =   1080
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   355
         Left            =   13290
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1035
         Width           =   1080
      End
      Begin MSForms.TextBox txtSupplier 
         Height          =   285
         Left            =   6120
         TabIndex        =   75
         Top             =   675
         Width           =   3915
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6906;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   855
         TabIndex        =   65
         Top             =   225
         Width           =   3600
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6350;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   855
         TabIndex        =   64
         Top             =   570
         Width           =   3600
         VariousPropertyBits=   746604575
         Size            =   "6350;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   58
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   660
         Width           =   645
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   1095
         Width           =   675
      End
      Begin MSForms.ComboBox cboType 
         Height          =   315
         Left            =   6135
         TabIndex        =   5
         Top             =   1080
         Width           =   4215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7435;556"
         TextColumn      =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   5040
         TabIndex        =   55
         Top             =   1080
         Width           =   495
         VariousPropertyBits=   8388627
         Caption         =   "Type:"
         Size            =   "873;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboStatus 
         Height          =   315
         Left            =   11610
         TabIndex        =   6
         Top             =   240
         Width           =   2760
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4868;556"
         TextColumn      =   1
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "-1;0"
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   10515
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5040
         TabIndex        =   53
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Task Owner:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5040
         TabIndex        =   52
         Top             =   240
         Width           =   915
      End
      Begin MSForms.ComboBox cboReportedBy 
         Height          =   315
         Left            =   6135
         TabIndex        =   3
         Top             =   240
         Width           =   4215
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7435;556"
         TextColumn      =   2
         ColumnCount     =   8
         ListRows        =   20
         MatchEntry      =   1
         ShowDropButtonWhen=   2
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
         Caption         =   "From Date:"
         Height          =   195
         Index           =   19
         Left            =   10515
         TabIndex        =   51
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         Height          =   195
         Index           =   8
         Left            =   10515
         TabIndex        =   50
         Top             =   1080
         Width           =   585
      End
   End
   Begin VB.Frame fraEditDemand 
      Height          =   7140
      Left            =   120
      TabIndex        =   46
      Top             =   1440
      Width           =   1215
      Begin VB.CommandButton cmdAddDiary 
         Caption         =   "Add Diary Entry"
         Height          =   555
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2060
         Width           =   1080
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1215
         Begin VB.OptionButton optDiary 
            Caption         =   "Diary Only"
            Height          =   255
            Left            =   40
            TabIndex        =   13
            Top             =   765
            Width           =   1095
         End
         Begin VB.OptionButton optJobs 
            Caption         =   "Jobs Only"
            Height          =   255
            Left            =   40
            TabIndex        =   12
            Top             =   462
            Width           =   1095
         End
         Begin VB.OptionButton optAll 
            Caption         =   "View All"
            Height          =   255
            Left            =   40
            TabIndex        =   11
            Top             =   160
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdPO 
         Caption         =   "Create Purchase Order"
         Height          =   735
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5380
         Width           =   1080
      End
      Begin VB.CommandButton cmdEditMHistory 
         Caption         =   "&Edit"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2980
         Width           =   1080
      End
      Begin VB.CommandButton cmdPrintJobSheet 
         Caption         =   "Print"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3780
         Width           =   1080
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6480
         Width           =   1080
      End
      Begin VB.CommandButton cmdEmailJS_PO 
         Caption         =   "Email"
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4580
         Width           =   1080
      End
      Begin MyHoverButton.Button cmdAddJob 
         Height          =   375
         Left            =   60
         TabIndex        =   14
         Top             =   1320
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMaintenance.frx":0442
         HoverPicture    =   "frmMaintenance.frx":045E
         DisabledPicture =   "frmMaintenance.frx":047A
         DownPicture     =   "frmMaintenance.frx":0496
         MouseIcon       =   "frmMaintenance.frx":04B2
         Caption         =   "Add Job"
         HoverCaption    =   "Add Job"
         DownCaption     =   "Add Job"
      End
   End
   Begin VB.PictureBox fmeUnitLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   2385
      Left            =   8520
      ScaleHeight     =   2355
      ScaleWidth      =   7905
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   7935
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtSearchUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   50
         TabIndex        =   37
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox txtSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   980
         TabIndex        =   36
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txtSearchAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2430
         TabIndex        =   35
         Top             =   240
         Width           =   2100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUnitLookup 
         Height          =   1755
         Left            =   45
         TabIndex        =   39
         Top             =   560
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3096
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   13553358
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6300
         TabIndex        =   45
         Top             =   30
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   5460
         TabIndex        =   44
         Top             =   30
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PostCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4620
         TabIndex        =   43
         Top             =   30
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2430
         TabIndex        =   42
         Top             =   30
         Width           =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   920
         TabIndex        =   41
         Top             =   30
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   40
         Top             =   30
         Width           =   285
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   50
         Top             =   30
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Property Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7125
      Left            =   1395
      TabIndex        =   21
      Top             =   1440
      Width           =   13350
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMaintenance 
         Height          =   6285
         Left            =   120
         TabIndex        =   22
         Top             =   690
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   11086
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
         Caption         =   "Budget / Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   12195
         TabIndex        =   63
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Next Reminder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   9
         Left            =   9600
         TabIndex        =   33
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   32
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reported"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   2145
         TabIndex        =   31
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Alarm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   9600
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   8400
         TabIndex        =   27
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   26
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Task Owner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   25
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Reported by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   7200
         TabIndex        =   24
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   10920
         TabIndex        =   23
         Top             =   255
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
Private Type SendJobsByEmail
        szSuppID      As String
        szSuppEmail   As String
        szClient      As String
        colAtt        As Collection
        lURN          As Long
        szSuppName    As String
End Type
Private uSupplier(0)        As SendJobsByEmail
Dim lblAsJS_PO             As String
Private iLes               As Integer
Dim sTextBox As String
Dim bEmailResult  As Boolean

Private Sub LoadCboType(adoConn As ADODB.Connection)
   Dim sSQLQuery As String

   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MTYP'"
   populateCombo adoConn, sSQLQuery, cboType
End Sub

'Private Sub cboClientList_Change()
'   If cboClientList.ListIndex < 0 Then Exit Sub
'
'   Dim adoConn    As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   LoadProperty adoConn, cboPropertyList
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

'Private Sub LoadProperty(adoConn As ADODB.Connection, cboP As Control)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY PropertyID;"
''   Debug.Print szSQL
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
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow - 1
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboP.Clear
'   cboP.Column() = Data()
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

'Private Sub cboPropertyList_Click()
'   If txtClientList.Tag < 0 Then
'      cboClientList.SetFocus
'      Exit Sub
'   End If
'End Sub

Private Sub cmdAddDiary_Click()
   Load frmMaintananceDairy
   With frmMaintananceDairy
      .CallingForm = "M"       'Calling from Maintenance form
      .isEdit = False
      .RecordType = "D"
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With
   Me.Enabled = False
End Sub

Private Sub cmdAsJS_Click()
   Dim szFileName    As String

   fraEditDemand.Enabled = True
   fraJS_PO.Visible = False
   
   If lblAsJS_PO = "Print" Then
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "JOB SHEET"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      Else
         Report.ParameterFields(1).AddCurrentValue flxMaintenance.TextMatrix(flxMaintenance.row, 3)

         Report.ParameterFields(2).AddCurrentValue "Diary Entry"
         Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   If lblAsJS_PO = "Email" Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData
      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "JOB SHEET"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      Else
         Report.ParameterFields(1).AddCurrentValue flxMaintenance.TextMatrix(flxMaintenance.row, 3)

         Report.ParameterFields(2).AddCurrentValue "Diary Entry"
         Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If
      
      Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf"
      Report.ExportOptions.DestinationType = crEDTDiskFile
      Report.ExportOptions.FormatType = crEFTPortableDocFormat
      Report.ExportOptions.PDFExportAllPages = True
      Report.Export False

      Set Report = Nothing
      
      SaveAttachment DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf", flxMaintenance.TextMatrix(flxMaintenance.row, 16), flxMaintenance.TextMatrix(flxMaintenance.row, 28)
      
      EmailDelay 20
      Set Report = Nothing
      
'  Sending Email with demand invoice as attachments
      bEmailResult = SendJobByE_Mail("A Job has been assigned to you", _
                                        "Please find the job details in the attachment")
      ShowMsgInTaskBar "The email has been sent", "Y", "P"
   End If
End Sub

Private Function SendJobByE_Mail(szSub As String, szBody As String) As Boolean
   Dim i As Integer

   i = 0
      SendJobByE_Mail = SendEmail(szFromEmail, Trim(uSupplier(i).szSuppEmail), _
                                     szSub, _
                                     szBody, , , _
                                     uSupplier(i).colAtt, uSupplier(i).szSuppID, "SI")
End Function

Private Sub SaveAttachment(szFile As String, szSupplier As String, szSupplierEmail As String)
   Dim i As Integer

   i = 0
   Set uSupplier(i).colAtt = New Collection
   uSupplier(i).colAtt.Add szFile
   uSupplier(i).szSuppEmail = szSupplierEmail
End Sub

Private Sub cmdAsJS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraEditDemand.Enabled = True
      fraJS_PO.Visible = False
'      Frame1(0).Enabled = True
   End If
End Sub

Private Sub cmdAsPO_Click()
   fraEditDemand.Enabled = True
   fraJS_PO.Visible = False

   If lblAsJS_PO = "Print" Then
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "PURCHASE ORDER"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   If lblAsJS_PO = "Email" Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData
      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "PURCHASE ORDER"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      Else
         Report.ParameterFields(1).AddCurrentValue flxMaintenance.TextMatrix(flxMaintenance.row, 3)

         Report.ParameterFields(2).AddCurrentValue "Diary Entry"
         Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If
      
      Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf"
      Report.ExportOptions.DestinationType = crEDTDiskFile
      Report.ExportOptions.FormatType = crEFTPortableDocFormat
      Report.ExportOptions.PDFExportAllPages = True
      Report.Export False

      Set Report = Nothing
      
      SaveAttachment DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf", flxMaintenance.TextMatrix(flxMaintenance.row, 16), flxMaintenance.TextMatrix(flxMaintenance.row, 28)
      
      EmailDelay 20
      Set Report = Nothing
      
'  Sending Email with demand invoice as attachments
      bEmailResult = SendJobByE_Mail("A purchase order has been sent to you", _
                                        "Please find the purchase order details in the attachment")
      ShowMsgInTaskBar "The purchase order has been sent by email", "Y", "P"
   End If
End Sub

Private Sub cmdAsPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraEditDemand.Enabled = True
      fraJS_PO.Visible = False
   End If
End Sub

Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 585
    picClient.Visible = True
    LoadflxClient
    cmdProperty.Enabled = False
    cmdClientList.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdEditMHistory_Click()
   If flxMaintenance.TextMatrix(1, 0) = "" Then Exit Sub
   'Resolved by BOSL
   'Issue 474
   'Modified by anol 26 Oct 2014
   If flxMaintenance.row = 0 Then
      ShowMsgInTaskBar "Please select any job/diary from list", "Y", "N"
      Exit Sub
   End If
   'End of modification
   If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
      frmMaintenanceJob.isEdit = True
      frmMaintenanceJob.CallingForm = "M"                         'Maintenance
      frmMaintenanceJob.UpdateRow = flxMaintenance.row
      Load frmMaintenanceJob
      frmMaintenanceJob.ZOrder 0
      frmMaintenanceJob.Show
   Else
      frmMaintananceDairy.isEdit = True
      frmMaintananceDairy.CallingForm = "M"                         'Maintenance
      frmMaintananceDairy.UpdateRow = flxMaintenance.row
      Load frmMaintananceDairy
      frmMaintananceDairy.ZOrder 0
      frmMaintananceDairy.Show
   End If
   Me.Enabled = False
End Sub

Private Sub cmdEditMHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdEmailJS_PO_Click()
   If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
      fraJS_PO.Top = fraEditDemand.Top + cmdEmailJS_PO.Top
      fraJS_PO.Left = fraEditDemand.Left + cmdEmailJS_PO.Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
'      Frame1(0).Enabled = False
      lblAsJS_PO = "Email"
      fraEditDemand.Enabled = False
   End If
End Sub

Private Sub cmdEmailJS_PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdGridUnitLookup_Click()
   fmeUnitLookup.Visible = False
End Sub

Private Sub cmdAddJob_Click()
'   Load frmMaintenanceJob
   With frmMaintenanceJob
      .CallingForm = "M"          'Calling from property form
      .RecordType = "J"
      .lblJobName.Caption = "Job Name"
      .Label1.Caption = "Job No."
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With

   Me.Enabled = False
End Sub

Private Sub cmdAddJob_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = False
    cmdClientList.Enabled = True
    cmdProperty.Enabled = True
End Sub

Private Sub cmdProperty_Click()
    sTextBox = "2"
    picClient.Left = 915
    picClient.Top = 685
    picClient.Visible = True
    LoadPropertyList
    cmdProperty.Enabled = False
    cmdClientList.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
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


Private Sub LoadSupplierList()
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
   lblClientID.Caption = "Supplier ID"
   lblClientName.Caption = "Supplier Name"
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
   szSQL = "SELECT SupplierID, SupplierName  " & _
           "FROM Supplier where Type='SUPPLIER' " & _
           "ORDER BY SupplierID;"
'        szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
          
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



Private Sub cmdSupplier_Click()
    sTextBox = "3"
    picClient.Left = 6020
    picClient.Top = 685
    picClient.Visible = True
    LoadSupplierList
    cmdProperty.Enabled = False
    cmdClientList.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
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
          Frame2.Enabled = True
          
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
Private Sub flxClient_Click()
    Frame1.Enabled = True
    Frame2.Enabled = True
        
            cmdClientList.Enabled = True
            cmdProperty.Enabled = True
            
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        If sTextBox = "1" Then
               
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = ""
                txtPropertyName.Tag = ""
               
                Dim adoRst As New ADODB.Recordset
                Dim szSQL As String

                szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
                adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields(1).Value
                        txtPropertyName.Tag = adoRst.Fields(0).Value
                        'LoadSingleFY adoConn
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
                End If
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
               ' LoadSingleFY adoConn
               
        End If
        If sTextBox = "3" Then
                txtSupplier.text = flxClient.TextMatrix(flxClient.row, 2)
                txtSupplier.Tag = Trim(flxClient.TextMatrix(flxClient.row, 1))

        End If
'        If sTextBox = "4" Then
'                txtRCFund.text = flxClient.TextMatrix(flxClient.row, 2)
'                txtRCFund.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
'                If txtBudget.Enabled Then txtBudget.SetFocus
'        End If
        adoConn.Close
       
        picClient.Visible = False
End Sub
Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
Private Sub gridUnitLookup_DblClick()
      'Resolved By BOSL
      'issue 474
      'Modified by anol 28 Sep 2014
      If gridUnitLookup.TextMatrix(gridUnitLookup.row, 0) = "" Then Exit Sub
      txtUnitNo.text = gridUnitLookup.TextMatrix(gridUnitLookup.row, 0)
      fmeUnitLookup.Visible = False
      txtUnitNo.Locked = False
End Sub
Private Sub cmdOK_Click()
   'Resolved By BOSL
      'issue 474 Note 1
      'Modified by anol 28 Sep 2014
   ConfigFlxMaintenance
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim szSQL As String
   Dim strWHR As String
   strWHR = ""
   If txtPropertyName.Tag <> "" Then
      strWHR = " AND H.PropertyID='" & txtPropertyName.Tag & "'"
   End If
   If txtClientList.Tag <> "" Then
      strWHR = strWHR & " AND P.ClientID='" & txtClientList.Tag & "'"
   End If
   If Not optAll.Value Then
      If optJobs.Value = True Then
         strWHR = strWHR & " And H.RecordType = 'J'"
      End If
      If optDiary.Value = True Then
         strWHR = strWHR & " And H.RecordType = 'D'"
      End If
   End If
   If txtDateFrom.text <> "" Then
      strWHR = strWHR & " And H.ReportedDate >= #" & CDate(txtDateFrom.text) & "# "
   End If
   If txtDateTo.text <> "" Then
      strWHR = strWHR & " And H.ReportedDate <= #" & CDate(txtDateTo.text) & "# "
   End If
   If cboStatus.text = "PENDING" Then
      strWHR = strWHR & " AND H.Urgent='N'"
   ElseIf cboStatus.text = "URGENT" Then
      strWHR = strWHR & " AND H.Urgent<>'N' AND len(H.Urgent)>0"
   ElseIf cboStatus.text = "COMPLETED" Then
      strWHR = strWHR & " AND H.Urgent is null"
   End If
   If cboReportedBy.ListIndex <> -1 And cboReportedBy.text <> "" Then
      strWHR = strWHR & " AND H.ReportedBy='" & cboReportedBy.text & "'"
   End If
   'Modified by anol 26 Nov 2014
   If cboType.ListIndex <> -1 And cboType.text <> "" Then
      strWHR = strWHR & " AND UCASE(H.MaintenanceType)='" & UCase(cboType.Value) & "'"
   End If
   If txtSupplier.text <> "" Then
      strWHR = strWHR & " AND H.AssignedTo='" & txtSupplier.Tag & "'"
   End If
   'H.ActualCost
   szSQL = ""
   szSQL = "SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner,H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "IIF(H.RecordType = 'J',H.BudgetCost,H.Location), H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost,  H.AssignedIL, " & _
                "H.ReportedIS, H.RemindTime, H.Urgent, H.MaintenanceType, " & _
                "H.ReportedFrom, H.FundID, H.OverrideBudget, H.FYrID, " & _
                "H.BudgetPassed, P.PropertyID, P.ClientID, " & _
                "IIf(AssignedIL='S',U.SupplierOfficeEmail,S1.Description) AS EmailAdd,P.PropertyName,U.SupplierID,U.SupplierName,(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName,( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear  " & _
           "FROM (((PropertyMaintHistory AS H INNER JOIN SecondaryCode AS S ON " & _
                "H.MaintenanceType = S.Code) INNER JOIN " & _
                "Property AS P ON H.PropertyID = P.PropertyID )    LEFT JOIN " & _
                "(select Code, Description from SecondaryCode where PrimaryCode = 'MNTJOB') AS S1 ON " & _
                "H.AssignedTo = S1.Code) LEFT JOIN " & _
                "Supplier AS U ON H.AssignedTo = U.SupplierID " & _
           "WHERE S.PrimaryCode = 'MTYP'" & strWHR & " AND " & _
               "(H.ReportedFrom = 'P' OR H.ReportedFrom = 'M')"

   szSQL = szSQL & " UNION "
'szSQL = ""
   szSQL = szSQL & _
           "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy," & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "IIF(H.RecordType = 'J',H.BudgetCost,H.Location), H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost,  H.AssignedIL, " & _
                "H.ReportedIS, H.RemindTime, H.Urgent, H.MaintenanceType, " & _
                "H.ReportedFrom, H.FundID, H.OverrideBudget, H.FYrID, " & _
                "H.BudgetPassed, P.PropertyID, P.ClientID,  '', P.PropertyName, '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName,( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear   " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S, Units AS U, Property AS P " & _
           "WHERE S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' AND " & _
               "H.ReportedFrom = 'U' AND " & _
               "U.UnitNumber = H.PropertyID AND H.PropertyID = P.PropertyID"

   szSQL = szSQL & " UNION "
'szSQL = ""'added ReportedBy Field in the query  by anol 26 Nov 2014 issue 474
   szSQL = szSQL & _
           "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy," & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "IIF(H.RecordType = 'J',H.BudgetCost,H.Location), H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID,  '', P.PropertyName, '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S, Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' AND " & _
               "H.ReportedFrom = 'L' AND " & _
               "L.Status AND " & _
               "U.UnitNumber = L.UnitNumber AND " & _
               "L.SageAccountNumber = H.PropertyID AND H.PropertyID = P.PropertyID " & _
           "ORDER BY H.ReportedDate DESC;"
'Debug.Print szSQL

   populateGridDefinedHeader adoConn, szSQL, flxMaintenance

   flxMaintenance.row = 0
   flxMaintenance.col = 0
   adoConn.Close
   Set adoConn = Nothing
End Sub


Private Sub cmdPO_Click()
   If IsLoadedAndVisible("frmPO_Amend") Then
      ShowMsgInTaskBar "Purchase Order form is already open", "Y", "N"
      Exit Sub
   End If
   
   If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
      If flxMaintenance.TextMatrix(flxMaintenance.row, 16) <> "S" Then
         ShowMsgInTaskBar "A purchase order cannot be created, as  this job is assigned internally", "Y", "N"
         Exit Sub
      End If
      
      If MsgBox("Do you wish to create a purchase order from this job?", vbQuestion + vbYesNo, "Purchase Order") = vbNo Then Exit Sub

      Load frmPO_Amend

      With frmPO_Amend
         .Caption = "Create Purchase Order"
         .txtClientID.text = txtClientList.text
         'Fixed by anol 1 Jan 2015
         'error invalid use of null while create button was pressed for PO
         .txtAc(0).text = flxMaintenance.TextMatrix(flxMaintenance.row, 30) 'txtSupplier.Tag
         .txtSupplierName.text = flxMaintenance.TextMatrix(flxMaintenance.row, 31) 'cboSupplier.text
         'End of modification
         .txtDate.text = flxMaintenance.TextMatrix(flxMaintenance.row, 2)
         .txtDueDate.text = flxMaintenance.TextMatrix(flxMaintenance.row, 2)
         .txtProperty.text = txtPropertyName.text
         .txtNet_(0).text = flxMaintenance.TextMatrix(flxMaintenance.row, 10)
         .txtTotal.text = flxMaintenance.TextMatrix(flxMaintenance.row, 10)
         .txtJobNo.text = flxMaintenance.TextMatrix(flxMaintenance.row, 3)
         .szPropertyID = flxMaintenance.TextMatrix(flxMaintenance.row, 26)
         .txtProperty.text = flxMaintenance.TextMatrix(flxMaintenance.row, 29)
         .szClientID = flxMaintenance.TextMatrix(flxMaintenance.row, 27)
         .bEditMode = False
         .szCallerForm = "M"

         .Show
         .txtInv(0).SetFocus
      End With
      Me.Enabled = False

Exit Sub
      Dim adoConn    As New ADODB.Connection
      Dim adoRST_h   As New ADODB.Recordset     'Header
      Dim adoRST_c   As New ADODB.Recordset     'Child
      Dim adoSrcPIHd As New ADODB.Recordset     'Source
      Dim szHeaderID As String
      Dim lSlNumber  As Long

      adoConn.Open getConnectionString

      adoSrcPIHd.Open "SELECT * FROM PropertyMaintHistory WHERE ID = '" & _
                       Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), _
                       Len(txtPropertyName.Tag) + 2) & "';", adoConn, adOpenStatic, adLockReadOnly

      adoRST_h.Open "SELECT * FROM tblPurInv;", adoConn, adOpenDynamic, adLockOptimistic
      adoRST_c.Open "SELECT * FROM tblPurInvSRec;", adoConn, adOpenDynamic, adLockOptimistic

'     Creating the header
      With adoRST_h
         .AddNew
         szHeaderID = UniqueID()
         .Fields.Item("MY_ID").Value = szHeaderID

         lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
         .Fields.Item("SlNumber").Value = lSlNumber
         .Fields.Item("SUPP_AC").Value = adoSrcPIHd.Fields.Item("AssignedTo").Value
         .Fields.Item("TRAN_DATE").Value = Format(Now, "DD/MMMM/YYYY")
         .Fields.Item("TransactionType").Value = 25
         .Fields.Item("INV_NO").Value = flxMaintenance.TextMatrix(flxMaintenance.row, 3)
         .Fields.Item("TOTAL_AMOUNT").Value = adoSrcPIHd.Fields.Item("BudgetCost").Value
         .Fields.Item("TTP").Value = "JOBS"
         .Fields.Item("History").Value = False
         .Fields.Item("TrfPayment").Value = False
         .Fields.Item("PropertyID").Value = txtPropertyName.Tag
         .Fields.Item("DueDate").Value = Format(Now, "DD/MMMM/YYYY")

         .Update
         .Close
      End With

      With adoSrcPIHd
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ParentID").Value = szHeaderID
         .Fields.Item("TRAN_ID").Value = 1
         .Fields.Item("TRANS").Value = txtPropertyName.Tag
         
'         .Fields.Item("UNIT_ID").Value = adoSrcPIHd.Fields.Item("UNIT_ID").Value
         .Fields.Item("NOMINAL_CODE").Value = adoSrcPIHd.Fields.Item("NOMINAL_CODE").Value
         .Fields.Item("DEPT_ID").Value = adoSrcPIHd.Fields.Item("DEPT_ID").Value
         
         .Fields.Item("JOB_ID").Value = flxMaintenance.TextMatrix(flxMaintenance.row, 3)
'         .Fields.Item("COST_CODE").Value = adoSrcPIHd.Fields.Item("COST_CODE").Value
         .Fields.Item("description").Value = adoSrcPIHd.Fields.Item("description").Value
         .Fields.Item("NET_AMOUNT").Value = CCur(adoSrcPIHd.Fields.Item("NET_AMOUNT").Value)
         .Fields.Item("TAX_CODE").Value = adoSrcPIHd.Fields.Item("TAX_CODE").Value
         .Fields.Item("VAT").Value = CCur(adoSrcPIHd.Fields.Item("VAT").Value)
         .Fields.Item("ScheduleID").Value = adoSrcPIHd.Fields.Item("ScheduleID").Value
         .Fields.Item("TOTAL_AMOUNT").Value = CCur(adoSrcPIHd.Fields.Item("TOTAL_AMOUNT").Value)

         .Update
         .Close
      End With

      adoConn.Close
      Set adoConn = Nothing
      
      
      
      ShowMsgInTaskBar "The purchase order has been created", "Y", "P"
   End If
End Sub

Private Sub cmdPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdPrintJobSheet_Click()
   If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
      fraJS_PO.Top = fraEditDemand.Top + cmdPrintJobSheet.Top
      fraJS_PO.Left = fraEditDemand.Left + cmdPrintJobSheet.Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
'      Frame1(0).Enabled = False
      lblAsJS_PO = "Print"
      fraEditDemand.Enabled = False
   End If
End Sub

Private Sub cmdPrintJobSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdQuoteReq_Click()
   fraEditDemand.Enabled = True
   fraJS_PO.Visible = False

   If lblAsJS_PO = "Print" Then
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "JOB QUOTE REQUEST"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   If lblAsJS_PO = "Email" Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData
      If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
         Report.ParameterFields(1).AddCurrentValue Mid(flxMaintenance.TextMatrix(flxMaintenance.row, 3), 6)

         Report.ParameterFields(2).AddCurrentValue "Job Name"
         Report.ParameterFields(3).AddCurrentValue "JOB QUOTE REQUEST"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      Else
         Report.ParameterFields(1).AddCurrentValue flxMaintenance.TextMatrix(flxMaintenance.row, 3)

         Report.ParameterFields(2).AddCurrentValue "Diary Entry"
         Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
         Report.ParameterFields(4).AddCurrentValue txtClientList.text
      End If
      
      Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf"
      Report.ExportOptions.DestinationType = crEDTDiskFile
      Report.ExportOptions.FormatType = crEFTPortableDocFormat
      Report.ExportOptions.PDFExportAllPages = True
      Report.Export False

      Set Report = Nothing
      
      SaveAttachment DB_PATH & "\AllStuff\Temp\" & flxMaintenance.TextMatrix(flxMaintenance.row, 3) & ".pdf", flxMaintenance.TextMatrix(flxMaintenance.row, 16), flxMaintenance.TextMatrix(flxMaintenance.row, 28)
      
      EmailDelay 20
      Set Report = Nothing
      
'  Sending Email with demand invoice as attachments
      bEmailResult = SendJobByE_Mail("JOB QUOTE REQUEST", _
                                        "Please find a job in the attachment and provide a quote request")
      ShowMsgInTaskBar "The quote request has been sent by email", "Y", "P"
   End If
End Sub

Private Sub cmdQuoteReq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraEditDemand.Enabled = True
      fraJS_PO.Visible = False
   End If
End Sub

Private Sub cmdReset_Click()
   Dim iRow As Integer

   For iRow = 1 To flxMaintenance.Rows - 1
      flxMaintenance.RowHeight(iRow) = 240            'Reset the row height
   Next iRow
   flxMaintenance.row = 0

   txtClientList.text = ""
   txtClientList.Tag = ""
   txtPropertyName.text = ""
   txtPropertyName.Tag = ""
   txtUnitNo.text = ""
   cboReportedBy.ListIndex = -1
   txtSupplier.Tag = ""
   txtSupplier.text = ""
   cboType.ListIndex = -1
   cboStatus.ListIndex = -1
   txtDateFrom.text = ""
   txtDateTo.text = ""
   optAll.Value = True
End Sub

Private Sub cmdUnitLookup_Click()
   If txtPropertyName.Tag = "" Then
      cmdProperty.SetFocus
      Exit Sub
   End If

   fmeUnitLookup.Left = txtUnitNo.Left
   fmeUnitLookup.Top = txtUnitNo.Top

   fmeUnitLookup.Visible = True
   fmeUnitLookup.ZOrder 0
   gridUnitLookup.Visible = True

   txtSearchUnit.SetFocus
   txtSearchUnit.text = ""
   txtSearchAddress.Enabled = True
   txtSearchName.Enabled = True

   LoadGridUnitLookup "WHERE (((UNITS.PROPERTYID) = '" & txtPropertyName.Tag & "'));"
End Sub

Private Function LoadGridUnitLookup(ByVal strFilter_ As String)
  'cmdClientID.Default = True
   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
   conUnit_.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE, UNITTYPE, " & _
           "Occupied " & _
           "FROM UNITS " & strFilter_

'Debug.Print sSQLQuery_
   rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   gridUnitLookup.Clear
   gridUnitLookup.Rows = 2
   gridUnitLookup.Cols = 6
   ConfigGridUnitLookup
   While Not rstUnit_.EOF
      gridUnitLookup.TextMatrix(iRow, 0) = rstUnit_!UnitNumber
      gridUnitLookup.TextMatrix(iRow, 1) = rstUnit_!UnitName
      gridUnitLookup.TextMatrix(iRow, 2) = IIf(IsNull(rstUnit_!Address), "", rstUnit_!Address)
      gridUnitLookup.TextMatrix(iRow, 3) = IIf(IsNull(rstUnit_!UnitPostCode), "", rstUnit_!UnitPostCode)
      gridUnitLookup.TextMatrix(iRow, 4) = IIf(IsNull(rstUnit_!UNITTYPE), "", rstUnit_!UNITTYPE)
      gridUnitLookup.TextMatrix(iRow, 5) = IIf(rstUnit_!OCCUPIED = "N", "Vacant", "Occupied")
      rstUnit_.MoveNext
      If Not rstUnit_.EOF Then gridUnitLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstUnit_.Close
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing
End Function

Private Sub ConfigGridUnitLookup()
   fmeUnitLookup.Visible = True
   gridUnitLookup.Visible = True
   gridUnitLookup.RowHeight(0) = 0

   gridUnitLookup.ColWidth(0) = 900
   gridUnitLookup.TextMatrix(0, 0) = "Unit Number"
   gridUnitLookup.ColAlignment(0) = vbLeftJustify

   gridUnitLookup.ColWidth(1) = 1500
   gridUnitLookup.TextMatrix(0, 1) = "Name"
   gridUnitLookup.ColAlignment(1) = vbLeftJustify

   gridUnitLookup.ColWidth(2) = 2200
   gridUnitLookup.TextMatrix(0, 2) = "Address"
   gridUnitLookup.ColAlignment(2) = vbLeftJustify

   gridUnitLookup.ColWidth(3) = 800
   gridUnitLookup.TextMatrix(0, 3) = "PostCode"
   gridUnitLookup.ColAlignment(3) = vbLeftJustify

   gridUnitLookup.ColWidth(4) = 800
   gridUnitLookup.TextMatrix(0, 4) = "Unit Type"
   gridUnitLookup.ColAlignment(4) = vbLeftJustify

   gridUnitLookup.ColWidth(5) = 800
   gridUnitLookup.TextMatrix(0, 5) = "Status"
   gridUnitLookup.ColAlignment(5) = vbLeftJustify
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub flxMaintenance_DblClick()
   cmdEditMHistory_Click
End Sub

Private Sub flxMaintenance_RowColChange()
'   If flxMaintenance.TextMatrix(flxMaintenance.row, 16) = "S" Then
'      txtSupplier.Tag = flxMaintenance.TextMatrix(flxMaintenance.row, 7)
'
'   Else
'      cboSupplier.ListIndex = -1
'   End If
   txtClientList.Tag = flxMaintenance.TextMatrix(flxMaintenance.row, 27)
   txtPropertyName.Tag = flxMaintenance.TextMatrix(flxMaintenance.row, 26)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraJS_PO.Visible Then
      fraEditDemand.Enabled = True
      fraJS_PO.Visible = False
   End If
End Sub

Private Sub Form_Load()
   Me.Width = 14880
   Me.Height = 9180
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   
   ConfigFlxMaintenance

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadCombos adoConn
   LoadFlxMaintenance adoConn

   adoConn.Close
   Set adoConn = Nothing

   LoadPreviousSelection
End Sub

Private Sub LoadCombos(adoConn As ADODB.Connection)
   Dim szSQL As String
'Modied below line by anol 25 Nov 2014 issue 474
   szSQL = "SELECT SupplierID, SupplierName  " & _
           "FROM Supplier where Type='SUPPLIER' " & _
           "ORDER BY SupplierName;"
           
   'populateCombo adoConn, szSQL, cboSupplier

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

  ' populateCombo adoConn, szSQL, cboClientList

'*************************************** TYPE COMBO ******************************************
   szSQL = "SELECT CODE, VALUE " & _
           "FROM SECONDARYCODE " & _
           "WHERE PRIMARYCODE = 'MTYP'"

   populateCombo adoConn, szSQL, cboType

'*************************************** STATUS COMBO ******************************************
   szSQL = "SELECT   Code, Value " & _
           "FROM     SecondaryCode " & _
           "WHERE    PrimaryCode = 'MTYS';"
'Debug.Print szSQL
   populateCombo adoConn, szSQL, cboStatus

'*************************************** REPORTED BY COMBO ******************************************
'   szSQL = "SELECT S.Code, S.Value AS V " & _
'           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
'           "WHERE S.Code = H.ReportedBy AND " & _
'                 "H.ReportedIS = 'I' AND S.PrimaryCode = 'MNTJOB' " & _
'           "GROUP BY S.Code, S.Value"
'
'   szSQL = szSQL & " UNION "
'
'   szSQL = szSQL & _
'           "SELECT T.SageAccountNumber AS Code, T.Name AS V " & _
'           "FROM PropertyMaintHistory AS H, Tenants AS T " & _
'           "WHERE T.SageAccountNumber = H.ReportedBy AND " & _
'                 "H.ReportedIS = 'L' " & _
'           "GROUP BY T.SageAccountNumber, T.Name"
'Resolved by BOSL
'modified by anol 10 Dec 2014
'issue 474
   szSQL = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   populateCombo adoConn, szSQL, cboReportedBy
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim szChoice As String

   If optAll.Value Then szChoice = "A"
   If optJobs.Value Then szChoice = "J"
   If optDiary.Value Then szChoice = "D"

   SaveSetting "PropertyManagement", "ChoosedOption", "MNT-c" & CStr(SCID), szChoice
End Sub

Private Function LoadPreviousSelection() As Boolean
   Dim szChoice As String

   szChoice = GetSetting("PropertyManagement", "ChoosedOption", "MNT-c" & CStr(SCID))

   If szChoice = "A" Then optAll.Value = True
   If szChoice = "J" Then optJobs.Value = True
   If szChoice = "D" Then optDiary.Value = True
End Function

Private Sub fraEditDemand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Public Sub RefreshMaintenanceGrid(adoConn As ADODB.Connection)
   LoadFlxMaintenance adoConn
End Sub

Public Sub LoadFlxMaintenance(ByVal conMHistory_ As ADODB.Connection)
   Dim szSQL As String

   szSQL = "SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy," & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost,  H.AssignedIL, " & _
                "H.ReportedIS, H.RemindTime, H.Urgent, H.MaintenanceType, " & _
                "H.ReportedFrom, H.FundID, H.OverrideBudget, H.FYrID, " & _
                "H.BudgetPassed, P.PropertyID, P.ClientID, " & _
                "IIf(AssignedIL='S',U.SupplierOfficeEmail,S1.Description) AS EmailAdd, " & _
                "P.PropertyName,U.SupplierID,U.SupplierName,(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM (((PropertyMaintHistory AS H INNER JOIN SecondaryCode AS S ON " & _
                "H.MaintenanceType = S.Code) INNER JOIN " & _
                "Property AS P ON H.PropertyID = P.PropertyID) LEFT JOIN " & _
                "(select Code, Description from SecondaryCode where PrimaryCode = 'MNTJOB') AS S1 ON " & _
                "H.AssignedTo = S1.Code) LEFT JOIN " & _
                "Supplier AS U ON H.AssignedTo = U.SupplierID " & _
           "WHERE S.PrimaryCode = 'MTYP' AND " & _
               "(H.ReportedFrom = 'P' OR H.ReportedFrom = 'M')"

   szSQL = szSQL & " UNION "

   szSQL = szSQL & _
           "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy," & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost,  H.AssignedIL, " & _
                "H.ReportedIS, H.RemindTime, H.Urgent, H.MaintenanceType, " & _
                "H.ReportedFrom, H.FundID, H.OverrideBudget, H.FYrID, " & _
                "H.BudgetPassed, P.PropertyID, P.ClientID, '', P.PropertyName, '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S, Units AS U, Property AS P " & _
           "WHERE S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' AND " & _
               "H.ReportedFrom = 'U' AND " & _
               "U.UnitNumber = H.PropertyID AND H.PropertyID = P.PropertyID"

   szSQL = szSQL & " UNION "

   szSQL = szSQL & _
           "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID, '', P.PropertyName , '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S, Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' AND " & _
               "H.ReportedFrom = 'L' AND " & _
               "L.Status AND " & _
               "U.UnitNumber = L.UnitNumber AND " & _
               "L.SageAccountNumber = H.PropertyID AND H.PropertyID = P.PropertyID " & _
           "ORDER BY H.ReportedDate DESC;"
'Debug.Print szSQL

   populateGridDefinedHeader conMHistory_, szSQL, flxMaintenance

   flxMaintenance.row = 0
   flxMaintenance.col = 0
End Sub

Public Sub ConfigFlxMaintenance()
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
   flxMaintenance.Clear
   flxMaintenance.Rows = 2
   flxMaintenance.Cols = 35
   flxMaintenance.RowHeight(0) = 0
   flxMaintenance.FormatString = szHeader$

   For iColumn = 1 To 11
      flxMaintenance.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   flxMaintenance.ColWidth(iColumn) = flxMaintenance.Width + flxMaintenance.Left - Label61(iColumn - 1).Left - 70

   For iColumn = 12 To flxMaintenance.Cols
      flxMaintenance.ColWidth(iColumn) = 0
   Next iColumn
End Sub

Private Sub Label44_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub lblMainUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub optAll_Click()
   If optAll.Value Then
      Label61(1).Caption = "Maintenance Type"
      Label61(2).Caption = "Date Reported"
      Label61(3).Caption = "Job No."
      Label61(4).Caption = "Job Name"
      Label61(5).Caption = "Task Owner"
'      Label61(6).Caption = "Reported by"
'      Label61(7).Caption = "Assigned to"
'      Label61(9).Caption = "Next Reminder"
'      Label61(10).Caption = "Date Completed"
'      Label61(11).Caption = "Budget Amount"
      cmdOK_Click
   End If
End Sub
Private Sub optJobs_Click()
   If optJobs.Value Then
      Label61(1).Caption = "Maintenance Type"
      Label61(2).Caption = "Date Entered"
      Label61(3).Caption = "Job No."
      Label61(4).Caption = "Job Name"
      Label61(5).Caption = "Task Owner"
'      Label61(6).Caption = "Reported by"
'      Label61(7).Caption = "Assigned to"
'      Label61(9).Caption = "Next Reminder"
'      Label61(10).Caption = "Date Completed"
     Label61(11).Caption = "Budget Amount"
      cmdOK_Click
   End If
End Sub
Private Sub optDiary_Click()
   If optDiary.Value Then
      Label61(1).Caption = "Diary Type"
      Label61(2).Caption = "Date Reported"
      Label61(3).Caption = "Diary No."
      Label61(4).Caption = "Diary Subject"
      Label61(5).Caption = "Diary Owner"
'      Label61(6).Caption = "Reported by"
'      Label61(7).Caption = "Assigned to"
'      Label61(9).Caption = "Next Reminder"
'      Label61(10).Caption = "Date Completed"
      Label61(11).Caption = "Location"
      cmdOK_Click
   End If
End Sub



Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   If TextBoxFormatDate(txtDateFrom) Then
      If txtDateTo.text <> "" Then
         If DateDiff("d", CDate(txtDateFrom.text), CDate(txtDateTo.text)) < 0 Then
            ShowMsgInTaskBar "FROM cannot be after TO date", "Y", "N"
            txtDateFrom.text = ""
            txtDateFrom.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   If TextBoxFormatDate(txtDateTo) Then
      If txtDateFrom.text <> "" Then
         If DateDiff("d", CDate(txtDateFrom.text), CDate(txtDateTo.text)) < 0 Then
            ShowMsgInTaskBar "TO cannot be before FROM date", "Y", "N"
            txtDateTo.text = ""
            txtDateTo.SetFocus
         End If
      End If
   End If
End Sub
