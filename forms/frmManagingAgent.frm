VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManagingAgent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Managing Agent"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManagingAgent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12060
   Begin TabDlg.SSTab tabMain 
      Height          =   5295
      Left            =   120
      TabIndex        =   23
      Top             =   2235
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmManagingAgent.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "cmdAgentDetailsSave"
      Tab(0).Control(3)=   "cmdAgentDetailsEdit"
      Tab(0).Control(4)=   "cmdAgentDetailsCancel"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Bank/Payment Details"
      TabPicture(1)   =   "frmManagingAgent.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraBank(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraBank(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Account History"
      TabPicture(2)   =   "frmManagingAgent.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).Control(1)=   "MSHFlexGrid1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Event History"
      TabPicture(3)   =   "frmManagingAgent.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeEventHistory"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Memo/Attachemnt"
      TabPicture(4)   =   "frmManagingAgent.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).Control(1)=   "cmdUnitMemoCancel"
      Tab(4).Control(2)=   "cmdUnitMemoSave"
      Tab(4).Control(3)=   "cmdUnitMemoEdit"
      Tab(4).Control(4)=   "txtNote"
      Tab(4).ControlCount=   5
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   120
         Top             =   4320
         Width           =   11595
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   435
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdAgentAddAtch 
            Caption         =   "&Add New"
            Height          =   435
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   435
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   124
            Top             =   360
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
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -64800
         TabIndex        =   119
         Top             =   3900
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -66360
         TabIndex        =   118
         Top             =   3900
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         Height          =   435
         Left            =   -68040
         TabIndex        =   117
         Top             =   3900
         Width           =   1350
      End
      Begin VB.TextBox txtNote 
         Height          =   3255
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Top             =   480
         Width           =   11595
      End
      Begin VB.Frame fmeEventHistory 
         Caption         =   "Property Maintenance History"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   91
         Top             =   360
         Width           =   11505
         Begin VB.CommandButton cmdSaveEvent 
            Caption         =   "&Save"
            Height          =   315
            Left            =   9360
            TabIndex        =   105
            Top             =   4410
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelEvent 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10335
            TabIndex        =   104
            Top             =   4410
            Width           =   915
         End
         Begin VB.CommandButton cmdEditEvent 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   8385
            TabIndex        =   103
            Top             =   4410
            Width           =   915
         End
         Begin VB.CommandButton cmdNewEvent 
            Caption         =   "&New"
            Height          =   315
            Left            =   7410
            TabIndex        =   102
            Top             =   4410
            Width           =   915
         End
         Begin VB.TextBox txtEventHistoryID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   270
            TabIndex        =   101
            Top             =   270
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10890
            TabIndex        =   100
            Top             =   810
            Width           =   345
         End
         Begin VB.CommandButton cmdMType 
            Caption         =   "..."
            Height          =   315
            Left            =   1980
            TabIndex        =   93
            Top             =   810
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridEventHistory 
            Height          =   3165
            Left            =   270
            TabIndex        =   106
            Top             =   1170
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   5583
            _Version        =   393216
            Cols            =   9
            BackColorFixed  =   8687183
            ForeColorFixed  =   16777215
            BackColorSel    =   13884353
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   14737632
            GridColorFixed  =   8421376
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
            _Band(0).Cols   =   9
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label59 
            Caption         =   "Description:"
            Height          =   255
            Left            =   3300
            TabIndex        =   115
            Top             =   570
            Width           =   1275
         End
         Begin VB.Label Label42 
            Caption         =   "Event Type:"
            Height          =   255
            Left            =   270
            TabIndex        =   114
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label43 
            Caption         =   "Reported   Date:"
            Height          =   435
            Left            =   2280
            TabIndex        =   113
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label Label64 
            Caption         =   "Alarm"
            Height          =   195
            Left            =   10860
            TabIndex        =   112
            Top             =   570
            Width           =   405
         End
         Begin VB.Label Label5 
            Caption         =   "Date  Actioned:"
            Height          =   435
            Left            =   5760
            TabIndex        =   111
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label Label66 
            Caption         =   "Task Owner:"
            Height          =   255
            Left            =   6870
            TabIndex        =   110
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Contact:"
            Height          =   255
            Left            =   8280
            TabIndex        =   109
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label Label2 
            Caption         =   "Remind   Date:"
            Height          =   435
            Left            =   9690
            TabIndex        =   108
            Top             =   390
            Width           =   885
         End
         Begin MSForms.ComboBox cboEventType 
            Height          =   315
            Left            =   270
            TabIndex        =   92
            Top             =   810
            Width           =   1695
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "2990;556"
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
         Begin MSForms.TextBox dtpRemindDate 
            Height          =   315
            Left            =   9690
            TabIndex        =   99
            Top             =   810
            Width           =   1125
            VariousPropertyBits=   746604571
            Size            =   "1984;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtContact 
            Height          =   315
            Left            =   8280
            TabIndex        =   98
            Top             =   810
            Width           =   1395
            VariousPropertyBits=   746604571
            Size            =   "2461;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTaskOwner 
            Height          =   315
            Left            =   6870
            TabIndex        =   97
            Top             =   810
            Width           =   1395
            VariousPropertyBits=   746604571
            Size            =   "2461;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox dtpDateCompleted 
            Height          =   315
            Left            =   5730
            TabIndex        =   96
            Top             =   810
            Width           =   1125
            VariousPropertyBits=   746604571
            Size            =   "1984;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   315
            Left            =   3300
            TabIndex        =   95
            Top             =   810
            Width           =   2415
            VariousPropertyBits=   746604571
            Size            =   "4260;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox dtpReportedDate 
            Height          =   315
            Left            =   2250
            TabIndex        =   94
            Top             =   810
            Width           =   1035
            VariousPropertyBits=   746604571
            Size            =   "1826;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEventTenantID 
            Height          =   315
            Left            =   3360
            TabIndex        =   107
            Top             =   120
            Visible         =   0   'False
            Width           =   1875
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "3307;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   1335
         Left            =   -67440
         ScaleHeight     =   1305
         ScaleWidth      =   4185
         TabIndex        =   83
         Top             =   480
         Width           =   4215
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   86
            Top             =   120
            Width           =   2000
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   85
            Top             =   480
            Width           =   2000
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   84
            Top             =   840
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   89
            Top             =   120
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Received (YTD):"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   88
            Top             =   480
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Receivable (YTD):"
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   87
            Top             =   840
            Width           =   1590
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Bank Details:"
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   5295
         Begin VB.TextBox txtBANK_NAME 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   600
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   960
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1260
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   1560
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_POST_CODE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1920
            Width           =   1395
         End
         Begin VB.TextBox txtBank_ID_ 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1920
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdNewBank 
            Caption         =   "New"
            Height          =   285
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSAdodcLib.Adodc adoBank 
            Height          =   330
            Left            =   3240
            Top             =   1920
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1920
            Width           =   750
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   555
         End
         Begin MSForms.ComboBox cboBank_ID 
            Height          =   285
            Left            =   1200
            TabIndex        =   60
            Top             =   240
            Width           =   3195
            VariousPropertyBits=   1820346399
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5636;503"
            TextColumn      =   1
            ColumnCount     =   6
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Account Details:"
         Height          =   2295
         Index           =   1
         Left            =   7080
         TabIndex        =   59
         Top             =   360
         Width           =   4575
         Begin VB.TextBox txtBank_AC_Name 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   720
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_SC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   69
            Top             =   1080
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_AC_NUM 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   1440
            Width           =   2800
         End
         Begin VB.TextBox txtBacsRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   71
            Top             =   1800
            Width           =   2800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   195
            Index           =   57
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Index           =   58
            Left            =   120
            TabIndex        =   75
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Index           =   59
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BACS REF:"
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   73
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method:"
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1335
         End
         Begin MSForms.ComboBox cboPaymentMethod 
            Height          =   285
            Left            =   1560
            TabIndex        =   67
            Top             =   240
            Width           =   2800
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4939;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   120
         TabIndex        =   52
         Top             =   2535
         Width           =   11535
         Begin VB.CommandButton cmdAddNewBank 
            Caption         =   "&Add New"
            Height          =   360
            Left            =   3720
            TabIndex        =   57
            Top             =   2265
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   360
            Left            =   6960
            TabIndex        =   56
            Top             =   2265
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteBank 
            Caption         =   "&Delete"
            Height          =   360
            Left            =   10200
            TabIndex        =   55
            Top             =   2265
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditBank 
            Caption         =   "&Edit"
            Height          =   360
            Left            =   5340
            TabIndex        =   54
            Top             =   2265
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "Canc&el"
            Enabled         =   0   'False
            Height          =   360
            Left            =   8580
            TabIndex        =   53
            Top             =   2265
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOtherBankDetails 
            Height          =   1785
            Left            =   120
            TabIndex        =   58
            Top             =   440
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3149
            _Version        =   393216
            Cols            =   7
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
            _Band(0).Cols   =   7
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Ac"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   6
            Left            =   10200
            TabIndex        =   132
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   4
            Left            =   7200
            TabIndex        =   131
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   3
            Left            =   4200
            TabIndex        =   130
            Top             =   180
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   5
            Left            =   8760
            TabIndex        =   129
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   128
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   127
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   126
            Top             =   180
            Width           =   705
         End
         Begin VB.Label lblCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   225
            Left            =   120
            TabIndex        =   125
            Top             =   180
            Width           =   11295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Managing Agent Address:"
         Height          =   4575
         Left            =   -74640
         TabIndex        =   27
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtAgentAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtAgentPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtAgentAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtAgentAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   38
            Top             =   3960
            Width           =   2655
         End
         Begin VB.TextBox txtAgentPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   37
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtAgentMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Top             =   3000
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   2565
            Width           =   2655
         End
         Begin VB.TextBox txtAgentHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   49
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   46
            Top             =   3480
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   43
            Top             =   3000
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   35
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   33
            Top             =   2520
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Alternative Address:"
         Height          =   4095
         Left            =   -68760
         TabIndex        =   24
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtAgentOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtAgentOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   26
            Top             =   2160
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAgentDetailsSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66360
         TabIndex        =   45
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentDetailsEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -68040
         TabIndex        =   44
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentDetailsCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64800
         TabIndex        =   47
         Top             =   4800
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   90
         Top             =   1920
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5636
         _Version        =   393216
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
         _Band(0).Cols   =   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   380
      Left            =   10800
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveAgent 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   380
      Left            =   4392
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteAgent 
      Caption         =   "&Delete"
      Height          =   380
      Left            =   8664
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditAgent 
      Caption         =   "&Edit"
      Height          =   380
      Left            =   2256
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelChange 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6528
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNewAgent 
      Caption         =   "&New"
      Height          =   380
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
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
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   11865
      TabIndex        =   8
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton cmdAgent 
         Caption         =   "V"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtAgentID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1725
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   120
         Width           =   2355
      End
      Begin VB.TextBox txtAgentName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2620
      End
      Begin VB.TextBox txtVATReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   9885
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1770
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   9885
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   1770
      End
      Begin MSForms.ComboBox cboAgentSageSuppAC 
         Height          =   285
         Left            =   1725
         TabIndex        =   3
         Top             =   840
         Width           =   2610
         VariousPropertyBits=   746604575
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4604;503"
         TextColumn      =   1
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         Object.Width           =   "1762;4233"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX/VAT Number:"
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
         Index           =   6
         Left            =   8400
         TabIndex        =   13
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
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
         Index           =   5
         Left            =   8400
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Supplier A/C:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Managing Agent ID:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1530
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
      Height          =   390
      Left            =   4403
      ScaleHeight     =   390
      ScaleWidth      =   3255
      TabIndex        =   20
      Top             =   3420
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   90
         Width           =   3075
      End
   End
   Begin VB.PictureBox Label3 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   100
      Left            =   -120
      ScaleHeight     =   45
      ScaleWidth      =   12195
      TabIndex        =   22
      Top             =   2040
      Width           =   12255
   End
   Begin VB.PictureBox picAgentList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1320
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   5415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgentList 
         Height          =   2240
         Left            =   15
         TabIndex        =   51
         Top             =   360
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   3942
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmManagingAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDefaultAccount As Boolean
Private szPropertyID As String
Private iRecharge As Integer
Private bGlobalData As Boolean
Private bNewEdit As Boolean
Private IMAGE_FILE_NAME_ As String
Private szaPremisisIDType() As String

Private Sub cboAgentSageSuppAC_LostFocus()
   If cboAgentSageSuppAC.text = "" Then
      MsgBox "Please choose Agent's SAGE account number.", vbCritical + vbOKOnly, "SAGE account number"
      Exit Sub
   Else
      txtAgentID.text = cboAgentSageSuppAC.text
   End If
End Sub

Private Sub cboBank_ID_Click()
   txtBANK_NAME.text = cboBank_ID.Column(1)
   txtBANK_ADDRESS1.text = cboBank_ID.Column(3)
   txtBANK_ADDRESS2.text = cboBank_ID.Column(5)
   txtBANK_ADDRESS3.text = cboBank_ID.Column(6)
   txtBANK_POST_CODE.text = cboBank_ID.Column(4)
   txtBANK_SC.text = cboBank_ID.Column(2)
End Sub

Private Sub cmdAddNewAgent_Click()
   If MsgBox("Do you wish to add a new Agent?", vbYesNo + vbQuestion, "Add New Agent") = vbNo Then Exit Sub
   If MsgBox("Have you entered the Agent's details in SAGE?", vbYesNo + vbQuestion, "Agent in SAGE") = vbNo Then Exit Sub
   bNewEdit = True

   MousePointer = vbHourglass

   SageSupplierAccCombo

   UnlockMainAgentText True
   MainCommandButtonEnable True

   txtAgentName.SetFocus

   MousePointer = vbDefault
End Sub

Private Sub cmdAgent_Click()
   Call PrepareList

   picAgentList.Top = picMain.Top + txtAgentID.Top + txtAgentID.Height + 5
   picAgentList.Left = picMain.Left + txtAgentID.Left + 5
   picAgentList.Visible = True
   picAgentList.ZOrder 0
End Sub

Private Sub cmdAgentAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachment cmbFiles, "Agent", txtAgentID.text
   MsgBox "File has been saved successfull, Thanks"
End Sub

Private Sub cmdCancelBank_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   CommandButtonEnabled True
   LockingAcText True
   NewBankText True, True
   flxOtherBankDetails_RowColChange
   cmdNewBank.Visible = False
End Sub

Private Sub cmdAgentDetailsCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText True
   CommandButtonEnable True
End Sub

Private Sub CommandButtonEnable(bEnable As Boolean)
   cmdAgentDetailsEdit.Enabled = bEnable
   cmdAgentDetailsSave.Enabled = Not bEnable
   cmdAgentDetailsCancel.Enabled = Not bEnable
End Sub

Private Sub cmdAgentDetailsEdit_Click()
   If txtAgentID.text = "" Then
      MsgBox "Please select a agent to edit.", vbCritical + vbOKOnly, "No selection"
      txtAgentID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText False
   CommandButtonEnable False
End Sub

Private Sub cmdAgentDetailsSave_Click()
   Dim conAgent As New RDO.rdoConnection
   Dim rstAgent As rdoResultset
   Dim szSQL As String

   conAgent.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conAgent.CursorDriver = rdUseIfNeeded
   conAgent.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM Agent " & _
           "WHERE AgentID = '" & txtAgentID.text & "';"
   Set rstAgent = conAgent.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   With rstAgent
      .Edit
      !AgentAddressLine1 = txtAgentAddressLine1.text
      !AgentAddressLine2 = txtAgentAddressLine2.text
      !AgentAddressLine3 = txtAgentAddressLine3.text
      !AgentPostCode = txtAgentPostCode.text
      !AgentOfficeEmail = txtAgentOfficeEmail.text
      !AgentPersonalEmail = txtAgentPersonalEmail.text
      !AgentHomeTel = txtAgentHomeTel.text
      !AgentMobile = txtAgentMobile.text
      !AgentOfficeAddressLine1 = txtAgentOfficeAddressLine1.text
      !AgentOfficeAddressLine2 = txtAgentOfficeAddressLine2.text
      !AgentOfficeAddressLine3 = txtAgentOfficeAddressLine3.text
      !AgentOfficePostCode = txtAgentOfficePostCode.text
      !AgentOfficeTel = txtAgentOfficeTel.text

      .Update
      .Close
   End With
   conAgent.Close
   Set rstAgent = Nothing
   Set conAgent = Nothing
   
   MsgBox "Data has been updated successfully", vbInformation + vbOKOnly, "Data Update"
   CommandButtonEnable True
End Sub

Private Sub cmdClinetAddAtch_Click()
End Sub

Private Sub cmdDeleteBank_Click()
   If MsgBox("Do you want to delete current account details?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
   
   flxOtherBankDetails.RemoveItem (flxOtherBankDetails.Row)

   flxOtherBankDetails_RowColChange
   NewBankText True, False
   cmdNewBank.Caption = "New"
   LockingAcText True
   MsgBox "Record has been deleted successfully.", vbInformation + vbOKOnly, "Delete"
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachment cmbFiles, cmbFiles.Column(2), txtAgentID.text, "Agent"
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdEditBank_Click()
   MousePointer = vbHourglass

   cmdNewBank.Caption = "Edit"

   cmdNewBank.Visible = True
   LockingAcText False

   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
   MousePointer = vbDefault
End Sub

Private Sub cmdAddNewBank_Click()
   If MsgBox("Is it default account?", vbQuestion + vbYesNo, "Deafult Account") = vbYes Then
      bDefaultAccount = True
   Else
      bDefaultAccount = False
   End If

   MousePointer = vbHourglass

   PopulateBank
   cmdNewBank.Caption = "New"
   cmdNewBank.Visible = True
   cboBank_ID.SetFocus

   LockingAcText False
   NewBankText True, True
   cboBank_ID.Locked = False

   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
   MousePointer = vbDefault
End Sub

Private Sub CommandButtonEnabled(bEnable As Boolean)
   cmdAddNewBank.Enabled = bEnable
   cmdEditBank.Enabled = bEnable
   cmdDeleteBank.Enabled = bEnable
   cmdSaveBank.Enabled = Not bEnable
   cmdCancelBank.Enabled = Not bEnable
   flxOtherBankDetails.Enabled = bEnable
End Sub

Public Function PopulateBank()
   Dim sSQLQuery_ As String

   adoBank.ConnectionString = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

   sSQLQuery_ = "SELECT BANK_ID, BANK_NAME, SORT_CODE, " & _
                     "BANK_ADDRESS1, BANK_POST_CODE, " & _
                     "BANK_ADDRESS2, BANK_ADDRESS3 " & _
                "FROM tlbBank"

   adoBank.RecordSource = sSQLQuery_
   adoBank.CommandType = adCmdText
   adoBank.Refresh

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoBank.Recordset.RecordCount
   TotalCol = adoBank.Recordset.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoBank.Recordset.RecordCount - 1
       For j = 0 To adoBank.Recordset.Fields.Count - 1
           Data(j, i) = adoBank.Recordset.Fields(j)
       Next j
       adoBank.Recordset.MoveNext
   Next i

   cboBank_ID.Column() = Data()
End Function

Private Sub UnlockMainAgentText(bUnlock As Boolean)
'   txtAgentID.Locked = Not bUnlock
   txtAgentName.Locked = Not bUnlock
   cboAgentSageSuppAC.Locked = Not bUnlock
'   txtAcBalance.Locked = Not bUnlock
   txtVATReg.Locked = Not bUnlock
   
   If bNewEdit Then
      txtAgentID.text = ""
      txtAgentName.text = ""
      cboAgentSageSuppAC.text = ""
'      txtAcBalance.text = ""
      txtVATReg.text = ""
   End If
End Sub

Private Sub cmdAgmntEdit_Click()
   If MsgBox("Do you want to edit the agreement?", vbQuestion + vbYesNo, "Edit Agreement") = vbNo Then Exit Sub
End Sub

Private Sub cmdAgmntSave_Click()
   If MsgBox("Are you sure to save?", vbQuestion + vbYesNo, "Data Saving") = vbNo Then Exit Sub

   MousePointer = vbHourglass

   Dim conAgr As New RDO.rdoConnection
   Dim rstAgr As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conAgr.CursorDriver = rdUseIfNeeded
   conAgr.EstablishConnection rdDriverNoPrompt

   szSQL = "DELETE * " & _
           "FROM tlbAggreement " & _
           "WHERE AGENT_ID = '" & txtAgentID.text & "' AND " & _
               "PROPERTY_ID = '" & szPropertyID & "';"
'Debug.Print szSQL
   Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   rstAgr.Close
   
   szSQL = "SELECT * " & _
           "FROM tlbAggreement"
   Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   
   With rstAgr
      .AddNew
      rstAgr!AGENT_ID = txtAgentID.text
      rstAgr!PROPERTY_ID = szPropertyID
      rstAgr!AGG_DATE = Format(Now, "DD MMMM YYYY")
      rstAgr!RECHARGES = CStr(iRecharge)
      
      .Update
      .Close
   End With
   Set rstAgr = Nothing
   
   conAgr.Close
   Set conAgr = Nothing
   MousePointer = vbDefault
   
   MsgBox "Agreement has been updated successfully.", vbInformation + vbOKOnly, "Agreement"
   Exit Sub
   
ErrorHandler:

   rstAgr.Close
   Set rstAgr = Nothing
   conAgr.Close
   Set conAgr = Nothing
   
   MsgBox ERR.Number & ERR.description & " ", vbCritical + vbOK, "PCM Error: 125"
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxAgentList
   LoadAllAgentFlxGrd
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDeleteAgent_Click()
'===========================================================================================
'This button is invisible, because user should not get facility to delete any record.
'we should give user a facility to see or remove the recode from the current list.
'===========================================================================================
   If txtAgentID.text = "" Then
      MsgBox "Please select a agent to delete.", vbInformation, "No selection"
      txtAgentID.SetFocus
      Exit Sub
   End If

   If MsgBox("Are you sure to delete current agent?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   Dim conAgent As New RDO.rdoConnection
   Dim szSQL As String

   conAgent.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conAgent.CursorDriver = rdUseIfNeeded
   conAgent.EstablishConnection rdDriverNoPrompt

   szSQL = "UPDATE AGENT " & _
           "SET InactiveAgent = TRUE, InactiveDate = '" & Format(Date, "dd mmmm yyyy") & "' " & _
           "WHERE AGENTID = '" & txtAgentID.text & "';"

   conAgent.Execute szSQL

   conAgent.Close
   Set conAgent = Nothing

   MsgBox "Agent has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"
End Sub

Private Sub cmdEditAgent_Click()
   If txtAgentID.text = "" Then
      MsgBox "Please select a agent to edit.", vbCritical + vbOKOnly, "No selection"
      cmdAgent.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to make change to the current agent?", vbYesNo + vbQuestion, "Edit Agent") = vbNo Then Exit Sub
   bNewEdit = False

   MousePointer = vbHourglass

   MainCommandButtonEnable True

   Dim szTemp As String

   If cboAgentSageSuppAC.ListCount = 0 Then
      szTemp = cboAgentSageSuppAC.text
      SageSupplierAccCombo
      cboAgentSageSuppAC.text = szTemp
   End If

   LockingAllText False
   UnlockMainAgentText True

   MousePointer = vbDefault
End Sub

Private Sub MainCommandButtonEnable(bEnabled As Boolean)
   cmdAddNewAgent.Enabled = Not bEnabled
   cmdEditAgent.Enabled = Not bEnabled
   cmdSaveAgent.Enabled = bEnabled
   cmdDeleteAgent.Enabled = Not bEnabled
   cmdCancelChange.Enabled = bEnabled
   
   cmdAgent.Enabled = Not bEnabled
End Sub

Private Sub cmdGridUnitLookup_Click()
   picAgentList.Visible = False
End Sub

Private Sub cmdHide_Click()
   picAgentList.Visible = False
End Sub

Private Sub cmdGSCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   Dim i As Integer
   
   On Error Resume Next
   For i = 0 To 67
      Label1(i).ForeColor = vbBlack
   Next i
End Sub

Private Sub cmdGSEdit_Click()
   MousePointer = vbHourglass

   MousePointer = vbDefault
End Sub

Private Sub cmdNewBank_Click()
   If cmdNewBank.Caption = "New" Then
      NewBankText False, True
      cboBank_ID.Locked = False
      cboBank_ID.Clear
      cboBank_ID.SetFocus
   Else
      NewBankText False, False
      txtBANK_NAME.SetFocus
   End If

   cmdNewBank.Enabled = False
End Sub

Private Sub NewBankText(bLock As Boolean, bNew As Boolean)
'   cboBank_ID.Locked = bLock
   txtBANK_NAME.Locked = bLock
   txtBANK_ADDRESS1.Locked = bLock
   txtBANK_ADDRESS2.Locked = bLock
   txtBANK_ADDRESS3.Locked = bLock
   txtBANK_POST_CODE.Locked = bLock

   If Not bNew Then Exit Sub
   cboBank_ID.text = ""
   txtBANK_NAME.text = ""
   txtBANK_ADDRESS1.text = ""
   txtBANK_ADDRESS2.text = ""
   txtBANK_ADDRESS3.text = ""
   txtBANK_POST_CODE.text = ""
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass
   
   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdSaveBank_Click()
   If cmdNewBank.Caption = "New" Then
      If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 1) <> "" Then flxOtherBankDetails.AddItem ""
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 1) = cboBank_ID.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 2) = txtBANK_NAME.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 3) = txtBANK_POST_CODE.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 4) = txtBank_AC_Name.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 5) = txtBANK_AC_NUM.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 6) = txtBANK_SC.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 7) = IIf(bDefaultAccount, "YES", "NO")
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 8) = txtBANK_ADDRESS1.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 9) = txtBANK_ADDRESS2.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 10) = txtBANK_ADDRESS3.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 11) = cboPaymentMethod.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 12) = txtBacsRef.text
   Else
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 1) = cboBank_ID.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 2) = txtBANK_NAME.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 3) = txtBANK_POST_CODE.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 4) = txtBank_AC_Name.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5) = txtBANK_AC_NUM.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6) = txtBANK_SC.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = IIf(bDefaultAccount, "YES", "NO")
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 8) = txtBANK_ADDRESS1.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 9) = txtBANK_ADDRESS2.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 10) = txtBANK_ADDRESS3.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 11) = cboPaymentMethod.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 12) = txtBacsRef.text
   End If

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String, szWhere As String, lSpare As Long

   On Error GoTo ErrorHandler

   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt
'
   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "New" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank;"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
      rstBank.AddNew
      rstBank!BANK_ID = cboBank_ID.text
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update
'
      NewBankText True, False
      rstBank.Close
      cmdNewBank.Visible = False
   End If

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "Edit" Then
'      Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank " & _
              "WHERE BANK_ID = '" & cboBank_ID.text & "';"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
      rstBank.Edit
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update
'
      rstBank.Close

      NewBankText True, False
      cmdNewBank.Visible = False
   End If

   If bDefaultAccount And cmdNewBank.Caption = "New" Then
      szSQL = "SELECT * " & _
              "FROM AGENT " & _
              "WHERE AGENTID = '" & txtAgentID.text & "'"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
      With rstBank
         .Edit
         !BANK_ID = cboBank_ID.text
         .Update
         .Close
      End With
   End If
   
   If cmdNewBank.Caption = "Edit" Then
      szWhere = " Where BANK_AC_NUM = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5) & "' And " & _
                     "BANK_SC = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6) & "';"
   Else
      szWhere = ""
   End If

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks" & szWhere
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   With rstBank
      If cmdNewBank.Caption = "New" Then
         .AddNew
      Else
         .Edit
      End If
      !CLIENT_ID = txtAgentID.text
      !BANK_ID = cboBank_ID.text
      !Bank_AC_Name = txtBank_AC_Name.text
      !BANK_AC_NUM = txtBANK_AC_NUM.text
      !BANK_SC = txtBANK_SC.text
      !DEFAULT_AC = bDefaultAccount
      !PaymentMethod = cboPaymentMethod.text
      !BacsRef = txtBacsRef.text
      .Update
      .MoveLast
      lSpare = CLng(!MY_ID)
   End With
   szSQL = "UPDATE tlbClientBanks " & _
           "SET Spare1 = '" & CStr(lSpare) & "' " & _
           "WHERE " & _
               "MY_ID = " & lSpare & ";"
   conBank.Execute szSQL
   If cmdNewBank.Caption = "New" Then
      MsgBox "Data has been saved successfully.", vbInformation + vbOKOnly, "Add New"
   Else
      MsgBox "Data has been updated successfully.", vbInformation + vbOKOnly, "Edit"
   End If
   CommandButtonEnabled True

NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub cmdSaveAgent_Click()
   If MsgBox("Do you want to save/update changes?", vbQuestion + vbYesNo, "Saving Data") = vbNo Then Exit Sub
   If txtAgentID.text = "" Then
      MsgBox "Please type agent id.", vbCritical + vbOKOnly, "Agent"
      txtAgentID.SetFocus
      Exit Sub
   End If
   If txtAgentName.text = "" Then
      MsgBox "Please type agent's name.", vbCritical + vbOKOnly, "Agent"
      txtAgentName.SetFocus
      Exit Sub
   End If
   If cboAgentSageSuppAC.text = "" Then
      MsgBox "Please select agent's Sage Supplier Account.", vbCritical + vbOKOnly, "Agent"
      cboAgentSageSuppAC.SetFocus
      Exit Sub
   End If
   If txtAcBalance.text = "" Then txtAcBalance.text = "0.00"

   If txtVATReg.text = "" Then
      If MsgBox("Are you registered for  VAT?" & (Chr(13) + Chr(10)) & "Press NO to continue saving.", vbQuestion + vbYesNo, "Client") = vbYes Then
         txtVATReg.SetFocus
         Exit Sub
      End If
   End If

   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   szSQL = "SELECT AgentID, AgentName, " & _
                  "AgentSageSuppAC, AcBalance, VATReg " & _
           "FROM Agent " & _
           "WHERE AgentID = '" & txtAgentID.text & "';"

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

   If PostToDBUsingADODB(Me, picMain, adoConn, szSQL, bNewEdit) Then
      MsgBox "Data has been saved succfully.", vbOKOnly, "Data Save"
   Else
      MsgBox "Data has not been saved.", vbOKOnly, "Data Save"
   End If

   UnlockMainAgentText False
   MainCommandButtonEnable False
End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveMemo("agent", "AgentMemo", txtAgentID.text, "AgentID", txtNote) Then
      MsgBox "Memo has been saved successfully.", vbInformation + vbOKOnly, "Memo"
   End If
   MemoButtonEnable False
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtNote.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub dtpDateCompleted_Change()
   MsTextBoxChangeDate dtpDateCompleted
End Sub

Private Sub dtpDateCompleted_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate dtpDateCompleted, KeyAscii
End Sub

Private Sub dtpDateCompleted_LostFocus()
   MsTextBoxFormatDate dtpDateCompleted
End Sub

Private Sub dtpRemindDate_Change()
   MsTextBoxChangeDate dtpRemindDate
End Sub

Private Sub dtpRemindDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate dtpRemindDate, KeyAscii
End Sub

Private Sub dtpRemindDate_LostFocus()
   MsTextBoxFormatDate dtpRemindDate
End Sub

Private Sub dtpReportedDate_Change()
   MsTextBoxChangeDate dtpReportedDate
End Sub

Private Sub dtpReportedDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate dtpReportedDate, KeyAscii
End Sub

Private Sub dtpReportedDate_LostFocus()
   MsTextBoxFormatDate dtpReportedDate
End Sub

Private Sub flxAgentList_Click()
   Dim sSQLQuery_ As String, sFilter As String

   txtAgentID.text = flxAgentList.TextMatrix(flxAgentList.Row, 1)

   MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   sSQLQuery_ = "SELECT * " & _
                "FROM agent " & _
                "WHERE agent.AgentID = '" & flxAgentList.TextMatrix(flxAgentList.Row, 1) & "';"
'Debug.Print sSQLQuery_
   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      MsgBox "Error in Database.", vbExclamation
   Else
      RetrieveMemo "agent", "AgentMemo", txtAgentID.text, "AgentID", txtNote
   End If

   fmeLoading.Visible = False
   MousePointer = vbDefault

   picAgentList.Visible = False
End Sub

Private Sub flxOtherBankDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxOtherBankDetails.ToolTipText = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.MouseRow, flxOtherBankDetails.MouseCol)
End Sub

Private Sub flxOtherBankDetails_RowColChange()
   Dim iCol As Integer

   MousePointer = vbHourglass

   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 10)
   cboPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 11)
   txtBacsRef.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 12)
   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50

   MousePointer = vbHourglass
   tabMain.Tab = 0
   cboPaymentMethod.AddItem "CHEQUE"
   cboPaymentMethod.AddItem "BACS"
   cboPaymentMethod.AddItem "DIRECT DEBIT"
   cboPaymentMethod.AddItem "Bank TRANSFER"
   cboPaymentMethod.AddItem "TT"
   cboPaymentMethod.AddItem "CHAPS"

   MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<AgentID|<AgentName|<AgentPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 300        'Solid column
   conFlxGrid.ColWidth(1) = 900        'agent ID
   conFlxGrid.ColWidth(2) = 3000       'agent Name
   conFlxGrid.ColWidth(3) = 800        'Post Code
   conFlxGrid.Rows = 2
'
   conFlxGrid.RowHeightMin = 300
End Sub

Private Sub imgClose_Click()
   picAgentList.Visible = False
End Sub

Private Sub LoadAllAgentFlxGrd()
   Dim conAgent As New RDO.rdoConnection
   Dim rstAgent As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgent.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conAgent.CursorDriver = rdUseIfNeeded
   conAgent.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT AgentID, AgentNAME, AgentPOSTCODE,  " & _
               "AgentSageSuppAC " & _
           "FROM Agent " & _
           "WHERE InactiveAgent = FALSE " & _
           "ORDER BY AgentNAME;"

   Set rstAgent = conAgent.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstAgent.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   While Not rstAgent.EOF
      flxAgentList.TextMatrix(iRow, 1) = rstAgent!AgentID
      flxAgentList.TextMatrix(iRow, 2) = rstAgent!AgentName
      flxAgentList.TextMatrix(iRow, 3) = IIf(IsNull(rstAgent!AgentPostCode), "", rstAgent!AgentPostCode)
      rstAgent.MoveNext
      If Not rstAgent.EOF Then flxAgentList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstAgent.Close
   conAgent.Close
   Set rstAgent = Nothing
   Set conAgent = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
   
   rstAgent.Close
   conAgent.Close
   Set rstAgent = Nothing
   Set conAgent = Nothing
End Sub

Private Sub Label13_Click()

End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
   MousePointer = vbHourglass

   Select Case tabMain.Tab
   Case 1:                    'Bank Payment details
      If cboBank_ID.text = "" Or flxOtherBankDetails.TextMatrix(1, 1) = "" Then
         LoadAllBankAC
         flxOtherBankDetails.Row = 0
         flxOtherBankDetails.Col = 0
      End If
   Case 4:                      'Attachment Files
      If txtAgentID.text <> "" Then _
            Call LoadAttachmentFiles(cmbFiles, txtAgentID.text, "Agent")
   End Select
   MousePointer = vbDefault
End Sub

Private Sub LoadAllBankAC()
   ConfigureFlxOtherBank

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT tlbClientBanks.*, tlbBank.* " & _
        "FROM tlbClientBanks, tlbBank, Agent " & _
        "WHERE Agent.AgentID = '" & txtAgentID.text & "' And " & _
            "Agent.BANK_ID = tlbBank.BANK_ID And " & _
            "tlbBank.BANK_ID = tlbClientBanks.BANK_ID"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   If Not rstBank.EOF Then
      cboBank_ID.text = rstBank!BANK_ID
      txtBANK_NAME.text = rstBank!BANK_NAME
      txtBANK_ADDRESS1.text = rstBank!BANK_ADDRESS1
      txtBANK_ADDRESS2.text = rstBank!BANK_ADDRESS2
      txtBANK_ADDRESS3.text = rstBank!BANK_ADDRESS3
      txtBANK_POST_CODE.text = rstBank!BANK_POST_CODE
      cboPaymentMethod.text = rstBank!PaymentMethod
      txtBank_AC_Name.text = rstBank!Bank_AC_Name
      txtBANK_SC.text = rstBank!BANK_SC
      txtBANK_AC_NUM.text = rstBank!BANK_AC_NUM
      txtBacsRef.text = rstBank!BacsRef
   End If
   rstBank.Close

   szSQL = "SELECT * " & _
              "FROM tlbClientBanks, tlbBank " & _
              "WHERE CLIENT_ID = '" & txtAgentID.text & "' And " & _
                  "tlbBank.BANK_ID = tlbClientBanks.BANK_ID " & _
              "ORDER BY Bank_AC_Name;"

   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   If rstBank.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstBank.EOF
      flxOtherBankDetails.TextMatrix(iRow, 1) = rstBank!BANK_ID
      flxOtherBankDetails.TextMatrix(iRow, 2) = rstBank!BANK_NAME
      flxOtherBankDetails.TextMatrix(iRow, 3) = rstBank!BANK_POST_CODE
      flxOtherBankDetails.TextMatrix(iRow, 4) = rstBank!Bank_AC_Name
      flxOtherBankDetails.TextMatrix(iRow, 5) = rstBank!BANK_AC_NUM
      flxOtherBankDetails.TextMatrix(iRow, 6) = rstBank!BANK_SC
      flxOtherBankDetails.TextMatrix(iRow, 7) = IIf(rstBank!DEFAULT_AC, "YES", "NO")
      flxOtherBankDetails.TextMatrix(iRow, 8) = rstBank!BANK_ADDRESS1
      flxOtherBankDetails.TextMatrix(iRow, 9) = rstBank!BANK_ADDRESS2
      flxOtherBankDetails.TextMatrix(iRow, 10) = rstBank!BANK_ADDRESS3
      flxOtherBankDetails.TextMatrix(iRow, 11) = rstBank!PaymentMethod
      flxOtherBankDetails.TextMatrix(iRow, 12) = rstBank!BacsRef

      rstBank.MoveNext
      If Not rstBank.EOF Then flxOtherBankDetails.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub ConfigureFlxOtherBank()
   Dim szHeader As String, i As Integer

   flxOtherBankDetails.Clear
   flxOtherBankDetails.Cols = 13
   flxOtherBankDetails.Rows = 2
   flxOtherBankDetails.RowHeight(0) = 60

   szHeader = "<BANK_ID|<BANK_NAME|<BANK_POST_CODE|<BANK_AC_NAME|<BANK_AC_NUM|<BANK_SC|<DEFAULT_AC"
   flxOtherBankDetails.FormatString = szHeader

   flxOtherBankDetails.ColWidth(0) = 0
   For i = 2 To flxOtherBankDetails.Cols - 6
      flxOtherBankDetails.ColWidth(i - 1) = Label6(i - 1).Left - Label6(i - 2).Left
   Next i
   
   flxOtherBankDetails.ColWidth(7) = flxOtherBankDetails.Width + flxOtherBankDetails.Left - Label6(6).Left - 300
   flxOtherBankDetails.ColWidth(8) = 0
   flxOtherBankDetails.ColWidth(9) = 0
   flxOtherBankDetails.ColWidth(10) = 0
   flxOtherBankDetails.ColWidth(11) = 0      'PaymentMethod
   flxOtherBankDetails.ColWidth(12) = 0      'BacsRef
End Sub

Private Sub LockingAcText(bLock As Boolean)
   txtBank_AC_Name.Locked = bLock
   txtBANK_SC.Locked = bLock
   txtBANK_AC_NUM.Locked = bLock
   txtBacsRef.Locked = bLock
   
   If cmdNewBank.Caption = "Edit" Then Exit Sub
   
   txtBank_AC_Name.text = ""
   txtBANK_SC.text = ""
   txtBANK_AC_NUM.text = ""
   txtBacsRef.text = ""
End Sub

Private Sub LockingAllText(bLock As Boolean)
   txtAgentAddressLine1.Locked = bLock
   txtAgentAddressLine2.Locked = bLock
   txtAgentAddressLine3.Locked = bLock
   txtAgentPostCode.Locked = bLock
   txtAgentHomeTel.Locked = bLock
   txtAgentOfficeTel.Locked = bLock
   txtAgentMobile.Locked = bLock
   txtAgentPersonalEmail.Locked = bLock
   txtAgentOfficeEmail.Locked = bLock
   txtAgentOfficeAddressLine1.Locked = bLock
   txtAgentOfficeAddressLine2.Locked = bLock
   txtAgentOfficeAddressLine3.Locked = bLock
   txtAgentOfficePostCode.Locked = bLock
End Sub

Private Sub SageSupplierAccCombo()
   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oPurchaseRecord As SageDataObject120.PurchaseRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oPurchaseRecord = oWS.CreateObject("PurchaseRecord")

         Dim TotalRow, TotalCol As Long
         Dim Data() As String
         Dim i As Integer
             
         TotalRow = oPurchaseRecord.Count
         TotalCol = 2
         cboAgentSageSuppAC.Clear
         
         ReDim Data(TotalCol, TotalRow) As String
         
         oPurchaseRecord.MoveFirst
         For i = 0 To TotalRow - 1
            'cboTest.AddItem adoAgent.Recordset.Fields(1)
            Data(0, i) = CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value)
            Data(1, i) = CStr(oPurchaseRecord.Fields.Item("NAME").Value)
            oPurchaseRecord.MoveNext
         Next i
         '
         cboAgentSageSuppAC.Column() = Data()
         cboAgentSageSuppAC.ColumnCount = TotalCol
         cboAgentSageSuppAC.BoundColumn = 1
         
         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oPurchaseRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "(pcm_003) The SDO generated the following error: " & oSDO.LastError.text

   Set oPurchaseRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub txtBANK_AC_NUM_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtBANK_SC_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtNOTICE_DAYS_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
