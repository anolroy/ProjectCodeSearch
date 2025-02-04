VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManagingAgent2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Managing Agent"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManagingAgent2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   14445
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   11840
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveAgent 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   4000
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditAgent 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   80
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelChange 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7920
      TabIndex        =   4
      Top             =   1200
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
      Height          =   975
      Left            =   80
      ScaleHeight     =   945
      ScaleWidth      =   12945
      TabIndex        =   9
      Top             =   120
      Width           =   12975
      Begin VB.CommandButton cmdAgent 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAgentID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   2620
      End
      Begin VB.TextBox txtAgentName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   2620
      End
      Begin VB.TextBox txtVATReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9885
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1770
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9885
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   120
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX/VAT Number:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   8400
         TabIndex        =   13
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   8400
         TabIndex        =   12
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Managing Agent ID:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   1590
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
      TabIndex        =   16
      Top             =   3060
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
         TabIndex        =   17
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
      ScaleWidth      =   13635
      TabIndex        =   18
      Top             =   1680
      Width           =   13695
   End
   Begin VB.PictureBox picAgentList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   8040
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   5415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgentList 
         Height          =   2240
         Left            =   15
         TabIndex        =   19
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
         TabIndex        =   15
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
   Begin TabDlg.SSTab tabMain 
      Height          =   5415
      Left            =   75
      TabIndex        =   22
      Top             =   1875
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmManagingAgent2.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAgentDetailsSave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAgentDetailsEdit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAgentDetailsCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Bank/Payment Details"
      TabPicture(1)   =   "frmManagingAgent2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame14"
      Tab(1).Control(1)=   "fraBank(1)"
      Tab(1).Control(2)=   "fraBank(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Account History"
      TabPicture(2)   =   "frmManagingAgent2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxACHistory"
      Tab(2).Control(1)=   "flxACHistorySplit"
      Tab(2).Control(2)=   "Label11(18)"
      Tab(2).Control(3)=   "Label11(17)"
      Tab(2).Control(4)=   "Label11(16)"
      Tab(2).Control(5)=   "Label11(15)"
      Tab(2).Control(6)=   "Label11(20)"
      Tab(2).Control(7)=   "Label11(14)"
      Tab(2).Control(8)=   "Label11(19)"
      Tab(2).Control(9)=   "Label11(21)"
      Tab(2).Control(10)=   "Label11(13)"
      Tab(2).Control(11)=   "Label11(12)"
      Tab(2).Control(12)=   "Label11(11)"
      Tab(2).Control(13)=   "Label11(10)"
      Tab(2).Control(14)=   "Label11(9)"
      Tab(2).Control(15)=   "Label11(6)"
      Tab(2).Control(16)=   "Label11(5)"
      Tab(2).Control(17)=   "Label11(1)"
      Tab(2).Control(18)=   "Label11(2)"
      Tab(2).Control(19)=   "Label11(3)"
      Tab(2).Control(20)=   "Label11(4)"
      Tab(2).Control(21)=   "Label11(7)"
      Tab(2).Control(22)=   "Label11(8)"
      Tab(2).Control(23)=   "lblGridCaption(0)"
      Tab(2).Control(24)=   "lblGridCaption(1)"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Memo/Attachemnt"
      TabPicture(3)   =   "frmManagingAgent2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame17"
      Tab(3).Control(1)=   "cmdUnitMemoCancel"
      Tab(3).Control(2)=   "cmdUnitMemoSave"
      Tab(3).Control(3)=   "cmdUnitMemoEdit"
      Tab(3).Control(4)=   "txtNote"
      Tab(3).ControlCount=   5
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   123
         Top             =   420
         Width           =   11595
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -67560
         TabIndex        =   122
         Top             =   3840
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
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
         Height          =   400
         Left            =   -65880
         TabIndex        =   121
         Top             =   3840
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
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
         Height          =   400
         Left            =   -64320
         TabIndex        =   120
         Top             =   3840
         Width           =   1350
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
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
         Height          =   855
         Left            =   -74400
         TabIndex        =   115
         Top             =   4260
         Width           =   11595
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdAgentAddAtch 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   119
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
      Begin VB.CommandButton cmdAgentDetailsCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11040
         TabIndex        =   89
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentDetailsEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7800
         TabIndex        =   88
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentDetailsSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9480
         TabIndex        =   87
         Top             =   4860
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Alternative Address:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   7080
         TabIndex        =   80
         Top             =   540
         Width           =   5295
         Begin VB.TextBox txtAgentOfficeAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   124
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   84
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtAgentOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   82
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   81
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   360
            TabIndex        =   86
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   360
            TabIndex        =   85
            Top             =   2520
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Managing Agent Address:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   480
         TabIndex        =   62
         Top             =   540
         Width           =   4575
         Begin VB.TextBox txtAgentAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   66
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtAgentHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            TabIndex        =   68
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            TabIndex        =   72
            Top             =   2565
            Width           =   2655
         End
         Begin VB.TextBox txtAgentMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            TabIndex        =   71
            Top             =   3000
            Width           =   2655
         End
         Begin VB.TextBox txtAgentPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            MaxLength       =   100
            TabIndex        =   70
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtAgentOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            MaxLength       =   100
            TabIndex        =   69
            Top             =   3960
            Width           =   2655
         End
         Begin VB.TextBox txtAgentAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   63
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtAgentAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   65
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtAgentPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtAgentAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1360
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   64
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   240
            TabIndex        =   79
            Top             =   2520
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   240
            TabIndex        =   78
            Top             =   3960
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   240
            TabIndex        =   77
            Top             =   3000
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   240
            TabIndex        =   76
            Top             =   3480
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   240
            TabIndex        =   75
            Top             =   2160
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   240
            TabIndex        =   74
            Top             =   1680
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   240
            TabIndex        =   73
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   -74280
         TabIndex        =   47
         Top             =   2475
         Width           =   11535
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "Canc&el"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8580
            TabIndex        =   52
            Top             =   2385
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditBank 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5340
            TabIndex        =   51
            Top             =   2385
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteBank 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10200
            TabIndex        =   50
            Top             =   2385
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "&Save"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6960
            TabIndex        =   49
            Top             =   2385
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddNewBank 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            TabIndex        =   48
            Top             =   2385
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOtherBankDetails 
            Height          =   1905
            Left            =   120
            TabIndex        =   53
            Top             =   435
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3360
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
               Name            =   "Calibri"
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
         Begin VB.Label lblCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H8000000F&
            Height          =   225
            Left            =   120
            TabIndex        =   61
            Top             =   180
            Width           =   11295
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   59
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   58
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   5
            Left            =   8760
            TabIndex        =   57
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   4200
            TabIndex        =   56
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   4
            Left            =   7200
            TabIndex        =   55
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Ac"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   6
            Left            =   10200
            TabIndex        =   54
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Account Details:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   1
         Left            =   -67320
         TabIndex        =   36
         Top             =   300
         Width           =   4575
         Begin VB.TextBox txtBacsRef 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            Top             =   1800
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_AC_NUM 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1440
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_SC 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   38
            Top             =   1080
            Width           =   2800
         End
         Begin VB.TextBox txtBank_AC_Name 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   720
            Width           =   2800
         End
         Begin MSForms.ComboBox cboPaymentMethod 
            Height          =   285
            Left            =   1560
            TabIndex        =   46
            Top             =   240
            Width           =   2800
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4939;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method:"
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BACS REF:"
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Index           =   59
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Index           =   58
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   195
            Index           =   57
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1065
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Bank Details:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   0
         Left            =   -74280
         TabIndex        =   23
         Top             =   300
         Width           =   5295
         Begin VB.CommandButton cmdNewBank 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtBank_ID_ 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1920
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtBANK_POST_CODE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1920
            Width           =   1395
         End
         Begin VB.TextBox txtBANK_ADDRESS3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1560
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1260
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_NAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   3195
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
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSForms.ComboBox cboBank_ID 
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   240
            Width           =   3195
            VariousPropertyBits=   1820346399
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5636;503"
            TextColumn      =   1
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1058"
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   750
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   840
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   2715
         Left            =   -74880
         TabIndex        =   90
         Top             =   540
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4789
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistorySplit 
         Height          =   1875
         Left            =   -74880
         TabIndex        =   91
         Top             =   3420
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   3307
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   18
         Left            =   -67920
         TabIndex        =   112
         Top             =   3225
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   17
         Left            =   -68880
         TabIndex        =   111
         Top             =   3225
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop No"
         Height          =   195
         Index           =   16
         Left            =   -69840
         TabIndex        =   110
         Top             =   3225
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No"
         Height          =   195
         Index           =   15
         Left            =   -70560
         TabIndex        =   109
         Top             =   3225
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   20
         Left            =   -64560
         TabIndex        =   108
         Top             =   3225
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   14
         Left            =   -71400
         TabIndex        =   107
         Top             =   3225
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   19
         Left            =   -65640
         TabIndex        =   106
         Top             =   3225
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   21
         Left            =   -63480
         TabIndex        =   105
         Top             =   3225
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   13
         Left            =   -72600
         TabIndex        =   104
         Top             =   3225
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   12
         Left            =   -73560
         TabIndex        =   103
         Top             =   3225
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   11
         Left            =   -74520
         TabIndex        =   102
         Top             =   3225
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   101
         Top             =   3225
         Width           =   240
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   9
         Left            =   -63720
         TabIndex        =   100
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   6
         Left            =   -67320
         TabIndex        =   99
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   5
         Left            =   -69960
         TabIndex        =   98
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   97
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   -74040
         TabIndex        =   96
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   -72840
         TabIndex        =   95
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   4
         Left            =   -71760
         TabIndex        =   94
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   7
         Left            =   -66120
         TabIndex        =   93
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   8
         Left            =   -64920
         TabIndex        =   92
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   113
         Top             =   300
         Width           =   12735
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   114
         Top             =   3180
         Width           =   12735
      End
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
      Left            =   0
      TabIndex        =   21
      Top             =   8760
      Width           =   1395
   End
   Begin MSForms.ComboBox cboAgentSageSuppAC__ 
      Height          =   285
      Left            =   1605
      TabIndex        =   20
      Top             =   8760
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
End
Attribute VB_Name = "frmManagingAgent2"
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
   AddNewAttachmentInCombo cmbFiles, "Agent", txtAgentID.text
   ShowMsgInTaskBar "The file has been saved successfully."
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
      ShowMsgInTaskBar "Please select a agent to edit."
      txtAgentID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText False
   CommandButtonEnable False
End Sub

Private Sub cmdAgentDetailsSave_Click()
   Dim conAgent As New ADODB.Connection
   Dim rstAgent As New ADODB.Recordset
   Dim szSQL As String

   conAgent.Open getConnectionString

   szSQL = "SELECT * " & _
           "FROM Agent " & _
           "WHERE AgentID = '" & txtAgentID.text & "';"
   rstAgent.Open szSQL, conAgent, adOpenDynamic, adLockOptimistic

   With rstAgent
      !AgentAddressLine1 = txtAgentAddressLine1.text
      !AgentAddressLine2 = txtAgentAddressLine2.text
      !AgentAddressLine3 = txtAgentAddressLine3.text
      !AgentAddressLine4 = txtAgentAddressLine4.text
      !AgentPostCode = txtAgentPostCode.text
      !AgentOfficeEmail = txtAgentOfficeEmail.text
      !AgentPersonalEmail = txtAgentPersonalEmail.text
      !AgentHomeTel = txtAgentHomeTel.text
      !AgentMobile = txtAgentMobile.text
      !AgentOfficeAddressLine1 = txtAgentOfficeAddressLine1.text
      !AgentOfficeAddressLine2 = txtAgentOfficeAddressLine2.text
      !AgentOfficeAddressLine3 = txtAgentOfficeAddressLine3.text
      !AgentOfficeAddressLine4 = txtAgentOfficeAddressLine4.text
      !AgentOfficePostCode = txtAgentOfficePostCode.text
      !AgentOfficeTel = txtAgentOfficeTel.text

      .Update
      .Close
   End With
   conAgent.Close
   Set rstAgent = Nothing
   Set conAgent = Nothing
   
   ShowMsgInTaskBar "Data has been updated successfully"
   CommandButtonEnable True
End Sub

Private Sub cmdDeleteBank_Click()
   If MsgBox("Do you want to delete current account details?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
   
   flxOtherBankDetails.RemoveItem (flxOtherBankDetails.row)

   flxOtherBankDetails_RowColChange
   NewBankText True, False
   cmdNewBank.Caption = "New"
   LockingAcText True
   ShowMsgInTaskBar "Record has been deleted successfully."
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtAgentID.text, "Agent"
   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdEditBank_Click()
   MousePointer = vbHourglass

   cmdNewBank.Caption = "Edit"

   cmdNewBank.Visible = True
   LockingAcText False

   CommandButtonEnabled False
   flxOtherBankDetails.row = flxOtherBankDetails.Rows - 1
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
   flxOtherBankDetails.row = flxOtherBankDetails.Rows - 1
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

   adoBank.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT BANK_ID, BANK_NAME, SORT_CODE, " & _
                     "BANK_ADDRESS1, BANK_POST_CODE, " & _
                     "BANK_ADDRESS2, BANK_ADDRESS3 " & _
                "FROM tlbBank"
'Debug.Print sSQLQuery_
   adoBank.RecordSource = sSQLQuery_
   adoBank.CommandType = adCmdText
   adoBank.Refresh

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoBank.Recordset.RecordCount
   TotalCol = adoBank.Recordset.Fields.count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoBank.Recordset.RecordCount - 1
       For j = 0 To adoBank.Recordset.Fields.count - 1
           Data(j, i) = IIf(IsNull(adoBank.Recordset.Fields(j).Value), "", adoBank.Recordset.Fields(j).Value)
       Next j
       adoBank.Recordset.MoveNext
   Next i

   cboBank_ID.Column() = Data()
End Function

Private Sub UnlockMainAgentText(bUnlock As Boolean)
'   txtAgentID.Locked = Not bUnlock
   txtAgentName.Locked = Not bUnlock
'   cboAgentSageSuppAC.Locked = Not bUnlock
'   txtAcBalance.Locked = Not bUnlock
   txtVATReg.Locked = Not bUnlock
   
'   If bNewEdit Then
'      txtAgentID.text = ""
'      txtAgentName.text = ""
'      cboAgentSageSuppAC.text = ""
''      txtAcBalance.text = ""
'      txtVATReg.text = ""
'   End If
End Sub

Private Sub cmdAgmntEdit_Click()
   If MsgBox("Do you want to edit the agreement?", vbQuestion + vbYesNo, "Edit Agreement") = vbNo Then Exit Sub
End Sub

Private Sub cmdAgmntSave_Click()
   If MsgBox("Are you sure to save?", vbQuestion + vbYesNo, "Data Saving") = vbNo Then Exit Sub

   MousePointer = vbHourglass

   Dim conAgr As New ADODB.Connection
   Dim rstAgr As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Open getConnectionString

   szSQL = "DELETE * " & _
           "FROM tlbAggreement " & _
           "WHERE AGENT_ID = '" & txtAgentID.text & "' AND " & _
               "PROPERTY_ID = '" & szPropertyID & "';"
   conAgr.Execute szSQL

   szSQL = "SELECT * " & _
           "FROM tlbAggreement"
   rstAgr.Open szSQL, conAgr, adOpenDynamic, adLockOptimistic

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
   
   ShowMsgInTaskBar "Agreement has been updated successfully."
   Exit Sub
   
ErrorHandler:

   rstAgr.Close
   Set rstAgr = Nothing
   conAgr.Close
   Set conAgr = Nothing
   
   ShowMsgInTaskBar ERR.Number & ERR.description & " ", , "N"
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
      ShowMsgInTaskBar "Please select a agent to delete.", , "N"
      txtAgentID.SetFocus
      Exit Sub
   End If

   If MsgBox("Are you sure to delete current agent?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   Dim conAgent As New ADODB.Connection
   Dim szSQL As String

   conAgent.Open getConnectionString

   szSQL = "UPDATE AGENT " & _
           "SET InactiveAgent = TRUE, InactiveDate = '" & Format(Date, "dd mmmm yyyy") & "' " & _
           "WHERE AGENTID = '" & txtAgentID.text & "';"

   conAgent.Execute szSQL

   conAgent.Close
   Set conAgent = Nothing

   ShowMsgInTaskBar "Agent has been deleted successfully."
End Sub

Private Sub cmdEditAgent_Click()
   If txtAgentID.text = "" Then
      ShowMsgInTaskBar "Please select a agent to edit.", , "N"
      cmdAgent.SetFocus
      txtAgentID.Locked = False
      Exit Sub
   End If

   If MsgBox("Do you want to make change to the current agent?", vbYesNo + vbQuestion, "Edit Agent") = vbNo Then Exit Sub
   bNewEdit = False

   MainCommandButtonEnable True

   Dim szTemp As String

'   If cboAgentSageSuppAC.ListCount = 0 Then
'      szTemp = cboAgentSageSuppAC.text
'      SageSupplierAccCombo
'      cboAgentSageSuppAC.text = szTemp
'   End If
'
   LockingAllText False
   UnlockMainAgentText True
End Sub

Private Sub MainCommandButtonEnable(bEnabled As Boolean)
'   cmdAddNewAgent.Enabled = Not bEnabled
   cmdEditAgent.Enabled = Not bEnabled
   cmdSaveAgent.Enabled = bEnabled
'   cmdDeleteAgent.Enabled = Not bEnabled
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
      ShowMsgInTaskBar "File has been moved from original location."

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
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1) = cboBank_ID.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 2) = txtBANK_NAME.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 3) = txtBANK_POST_CODE.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4) = txtBank_AC_Name.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5) = txtBANK_AC_NUM.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6) = txtBANK_SC.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = IIf(bDefaultAccount, "YES", "NO")
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 8) = txtBANK_ADDRESS1.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 9) = txtBANK_ADDRESS2.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 10) = txtBANK_ADDRESS3.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11) = cboPaymentMethod.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 12) = txtBacsRef.text
   End If

   Dim conBank As New ADODB.Connection
   Dim rstBank As New ADODB.Recordset
   Dim szSQL As String, szWhere As String, lSpare As Long

   On Error GoTo ErrorHandler

   conBank.Open getConnectionString

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "New" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank;"
      rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic

      rstBank.AddNew
      rstBank!BANK_ID = cboBank_ID.text
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update

      NewBankText True, False
      rstBank.Close
      cmdNewBank.Visible = False
   End If

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "Edit" Then
'      Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank " & _
              "WHERE BANK_ID = '" & cboBank_ID.text & "';"
      rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic

      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update

      rstBank.Close

      NewBankText True, False
      cmdNewBank.Visible = False
   End If

   If bDefaultAccount And cmdNewBank.Caption = "New" Then
      szSQL = "SELECT * " & _
              "FROM AGENT " & _
              "WHERE AGENTID = '" & txtAgentID.text & "'"
      rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
      With rstBank
         !BANK_ID = cboBank_ID.text
         .Update
         .Close
      End With
   End If
   
   If cmdNewBank.Caption = "Edit" Then
      szWhere = " Where BANK_AC_NUM = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5) & "' And " & _
                     "BANK_SC = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6) & "';"
   Else
      szWhere = ""
   End If

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks" & szWhere
   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
   With rstBank
      If cmdNewBank.Caption = "New" Then .AddNew

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
      ShowMsgInTaskBar "The data has been saved successfully."
   Else
      ShowMsgInTaskBar "The data has been updated successfully."
   End If
   CommandButtonEnabled True

NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub cmdSaveAgent_Click()
   If txtAgentName.text = "" Then
      ShowMsgInTaskBar "Please type agent's name.", , "N"
      txtAgentName.SetFocus
      Exit Sub
   End If
   If txtAgentID.text = "" Then
      ShowMsgInTaskBar "Please type agent id.", , "N"
      txtAgentID.SetFocus
      Exit Sub
   End If
   If txtAcBalance.text = "" Then txtAcBalance.text = "0.00"

   If txtVATReg.text = "" Then
      If MsgBox("Are you registered for  VAT?" & (Chr(13) + Chr(10)) & "Press NO to continue saving.", vbQuestion + vbYesNo, "Client") = vbYes Then
         txtVATReg.SetFocus
         Exit Sub
      End If
   End If

   If MsgBox("Do you want to save/update changes?", vbQuestion + vbYesNo, "Saving Data") = vbNo Then Exit Sub
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rst1 As New ADODB.Recordset
   Dim rst2 As New ADODB.Recordset

   szSQL = "SELECT AgentID, AgentName, " & _
                  "AgentSageSuppAC, AcBalance, VATReg " & _
           "FROM Agent " & _
           "WHERE AgentID = '" & txtAgentID.text & "';"

   adoConn.Open getConnectionString

   If PostToDBUsingADODB(Me, picMain, adoConn, szSQL, bNewEdit) Then
      ShowMsgInTaskBar "The data has been saved successfully."
   Else
      ShowMsgInTaskBar "Data has not been saved.", , "N"
   End If

   rst1.Open "SELECT * " & _
             "FROM Agent " & _
             "WHERE AgentID NOT IN (" & _
               "SELECT SupplierID FROM Supplier WHERE TYPE = 'AGENT');", adoConn, adOpenStatic, adLockReadOnly
'MsgBox "16               "
   If Not rst1.EOF Then
      rst2.Open "SELECT * FROM Supplier", adoConn, adOpenDynamic, adLockOptimistic
      While Not rst1.EOF
         rst2.AddNew
         rst2.Fields.Item("SupplierID").Value = rst1.Fields.Item("AgentID").Value
         rst2.Fields.Item("SupplierName").Value = rst1.Fields.Item("AgentName").Value
         rst2.Fields.Item("SupplierAddressLine1").Value = rst1.Fields.Item("AgentAddressLine1").Value
         rst2.Fields.Item("SupplierAddressLine2").Value = rst1.Fields.Item("AgentAddressLine2").Value
         rst2.Fields.Item("SupplierAddressLine3").Value = rst1.Fields.Item("AgentAddressLine3").Value
         rst2.Fields.Item("SupplierPostCode").Value = rst1.Fields.Item("AgentPostCode").Value
         rst2.Fields.Item("VATReg").Value = rst1.Fields.Item("VATReg").Value
         rst2.Fields.Item("TYPE").Value = "AGENT"
         rst2.Update
         rst1.MoveNext
      Wend
      rst2.Close
   End If
   rst1.Close
   Set rst1 = Nothing
   Set rst2 = Nothing
   adoConn.Close
   Set adoConn = Nothing

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
      ShowMsgInTaskBar "The memo has been saved successfully."
   End If
   MemoButtonEnable False
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtNote.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub
'
'Private Sub dtpDateCompleted_Change()
'   MsTextBoxChangeDate dtpDateCompleted
'End Sub
'
'Private Sub dtpDateCompleted_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   MsTextBoxKeyPrsDate dtpDateCompleted, KeyAscii
'End Sub
'
'Private Sub dtpDateCompleted_LostFocus()
'   MsTextBoxFormatDate dtpDateCompleted
'End Sub
'
'Private Sub dtpRemindDate_Change()
'   MsTextBoxChangeDate dtpRemindDate
'End Sub
'
'Private Sub dtpRemindDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   MsTextBoxKeyPrsDate dtpRemindDate, KeyAscii
'End Sub
'
'Private Sub dtpRemindDate_LostFocus()
'   MsTextBoxFormatDate dtpRemindDate
'End Sub
'
'Private Sub dtpReportedDate_Change()
'   MsTextBoxChangeDate dtpReportedDate
'End Sub
'
'Private Sub dtpReportedDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   MsTextBoxKeyPrsDate dtpReportedDate, KeyAscii
'End Sub
'
'Private Sub dtpReportedDate_LostFocus()
'   MsTextBoxFormatDate dtpReportedDate
'End Sub

Private Sub flxACHistory_Click()
'   If flxACHistory.TextMatrix(1, 0) = "" Then Exit Sub

   Dim iCurRowHeight As Integer, iRow As Integer
   Dim adoConn       As New ADODB.Connection
   Dim adoRst        As New ADODB.Recordset
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
   HighLightRowFlxGrid flxACHistory, flxACHistory.row

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "-" Then Exit Sub

ChildGrid:
'  Displaying the splits ************************************************************************************
   
   ConfigFlxACHistorySplit
   adoConn.Open getConnectionString

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PC" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbPayment AS P, tblPurInv AS I, tblPurInvSRec AS S " & _
              "WHERE P.PI = I.MY_ID AND " & _
                  "I.MY_ID = S.ParentID AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & " " & _
              "ORDER BY S.MY_ID;"

      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRst.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRst.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 4) = adoRst.Fields.Item("NOMINAL_CODE").Value
            .TextMatrix(iRow, 5) = adoRst.Fields.Item("JOB_ID").Value
            .TextMatrix(iRow, 6) = adoRst.Fields.Item("UNIT_ID").Value
            .TextMatrix(iRow, 7) = adoRst.Fields.Item("DEPT_ID").Value
            .TextMatrix(iRow, 8) = adoRst.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            adoRst.MoveNext
            If Not adoRst.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRst.Close
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PP" And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "PPR" Then
      szSQL = "SELECT S.*, P.ExtRef, P.UnitID, P.FundID " & _
              "FROM tlbPayment AS P, PayTransactions AS S " & _
              "WHERE P.TransactionID = S.FromTran AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & ";"
'Debug.Print szSQL
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRst.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRst.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRst.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("UnitID").Value), "", adoRst.Fields.Item("UnitID").Value)
            .TextMatrix(iRow, 7) = adoRst.Fields.Item("FundID").Value
            .TextMatrix(iRow, 8) = flxACHistory.TextMatrix(flxACHistory.row, 5)
            .TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("PaymentAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("PaymentAmount").Value, "0.00")
            adoRst.MoveNext
            If Not adoRst.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then
      szSQL = "SELECT P.* " & _
              "FROM tlbPayment AS P " & _
              "WHERE P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & ";"
'Debug.Print szSQL
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRst.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRst.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRst.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = adoRst.Fields.Item("UnitID").Value
            .TextMatrix(iRow, 7) = adoRst.Fields.Item("FundID").Value
            .TextMatrix(iRow, 8) = adoRst.Fields.Item("Details").Value
            .TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then _
               .TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Then _
               .TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
            adoRst.MoveNext
            If Not adoRst.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxAgentList_Click()
   Dim sSQLQuery_ As String, sFilter As String

   txtAgentID.text = flxAgentList.TextMatrix(flxAgentList.row, 1)

   MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoMain.ConnectionString = getConnectionString
   sSQLQuery_ = "SELECT * " & _
                "FROM agent " & _
                "WHERE agent.AgentID = '" & flxAgentList.TextMatrix(flxAgentList.row, 1) & "';"
'Debug.Print sSQLQuery_
   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      ShowMsgInTaskBar "Error in Database.", , "N"
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

   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 10)
   cboPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)
   txtBacsRef.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 12)
   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   MousePointer = vbDefault
End Sub

Private Sub LoadData()
   Dim rstAgent As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'   'Set the RDO Connections to the dataset
   adoMain.ConnectionString = getConnectionString
   
   szSQL = "SELECT * " & _
           "FROM Agent " & _
           "WHERE InactiveAgent = FALSE " & _
           "ORDER BY AgentNAME;"

'   Set rstAgent = conAgent.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   adoMain.RecordSource = szSQL
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If adoMain.Recordset.RecordCount = 0 Then
      ShowMsgInTaskBar "There is no Managing Agent record. Please create a Managing Agent record now.", , "N"
      SageSupplierAccCombo

      UnlockMainAgentText True
      MainCommandButtonEnable True
      bNewEdit = True
      txtAgentID.Locked = False
      Exit Sub
   End If
   If Not Fill_Form(Me, adoMain) Then
      ShowMsgInTaskBar "Error in Database.", , "N"
   Else
      RetrieveMemo "agent", "AgentMemo", txtAgentID.text, "AgentID", txtNote
   End If

   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.Height = 7875
   Me.Width = 13245
   Me.BackColor = MODULEBACKCOLOR
   tabMain.BackColor = MODULEBACKCOLOR

   tabMain.Tab = 0
   cboPaymentMethod.AddItem "CHEQUE"
   cboPaymentMethod.AddItem "BACS"
   cboPaymentMethod.AddItem "DIRECT DEBIT"
   cboPaymentMethod.AddItem "Bank TRANSFER"
   cboPaymentMethod.AddItem "TT"
   cboPaymentMethod.AddItem "CHAPS"

   Dim adoConn As New ADODB.Connection
   
   adoConn.Open getConnectionString

   LoadData
   LoadFlxACHistory adoConn
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadFlxACHistory(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset

   ConfigFlxACHistory

   szSQL = "SELECT P.*, TT.DESCRIPTION AS TT_DES, PI.SlNumber AS INV_REF, TT.CONSTANT " & _
           "FROM (tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON  " & _
                  "P.Type = TT.TYPE_ID) LEFT JOIN tblPurInv AS PI ON P.PI = PI.MY_ID " & _
           "WHERE  P.SageAccountNumber = '" & txtAgentID.text & "' " & _
           "ORDER BY P.TransactionID;"
'Debug.Print szSQL
   adoPty.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoPty.EOF
         If adoPty!Type = 6 Or adoPty!Type = 7 Then
            szSQL = "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type, P.SlNumber " & _
                    "FROM PayTransactions AS PT, tlbPayment AS P " & _
                    "WHERE PT.ToTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID;"
         Else
            szSQL = "SELECT SQ.*, P.SlNumber " & _
                    "FROM tlbPayment AS P, (" & _
                     "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type " & _
                     "FROM PayTransactions AS PT, tlbPayment AS P " & _
                     "WHERE PT.FromTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID) SQ " & _
                    "WHERE SQ.ToTran = P.TransactionID; "
         End If
'Debug.Print szSQL

         adoPtyDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         iChild = 0
         If adoPtyDtl.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(iKount, 0) = "+"
            iChild = iKount + 1
            While Not adoPtyDtl.EOF
               .TextMatrix(iChild, 0) = "-"
               .TextMatrix(iChild, 1) = IIf(adoPty.Fields.Item("Type").Value = 6, _
                                            adoPty.Fields.Item("INV_REF").Value, _
                                            adoPty.Fields.Item("SlNumber").Value)
               If adoPty!Type = 6 Then
                  .TextMatrix(iChild, 5) = "Payment from: PP" & adoPtyDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 5) = "Payment to: PI" & adoPtyDtl.Fields.Item("SlNumber").Value
               End If
               .TextMatrix(iChild, 6) = Format(adoPtyDtl.Fields.Item("PaymentAmount").Value, "0.00")
               .RowHeight(iChild) = 0
               iChild = iChild + 1
               adoPtyDtl.MoveNext
               If Not adoPtyDtl.EOF Then .AddItem ""
            Wend
         Else
            .TextMatrix(iKount, 0) = ""
         End If
'1:No,        2:Type,    3:Date, 4:Reference, 5:Description, 6:Amount, 7:Balance, 8:Dr, 9:Cr
         adoPtyDtl.Close
         .TextMatrix(iKount, 1) = Mid(adoPty.Fields.Item("CONSTANT").Value, 4, Len(adoPty.Fields.Item("CONSTANT").Value) - 3)
         .TextMatrix(iKount, 1) = .TextMatrix(iKount, 1) & _
                                      IIf(adoPty.Fields.Item("Type").Value = 6, _
                                      adoPty.Fields.Item("INV_REF").Value, _
                                      adoPty.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 2) = IIf(UCase(Left(adoPty.Fields.Item("TT_DES").Value, 5)) = "SALES", Mid(adoPty.Fields.Item("TT_DES").Value, 7), adoPty.Fields.Item("TT_DES").Value)
         .TextMatrix(iKount, 3) = IIf(IsNull(adoPty.Fields.Item("PDate").Value), "", adoPty.Fields.Item("PDate").Value)
         .TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("Ref").Value), "", adoPty.Fields.Item("PDate").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoPty.Fields.Item("Details").Value), "", adoPty.Fields.Item("Details").Value)
         .TextMatrix(iKount, 6) = Format(adoPty.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoPty.Fields.Item("OSAmount").Value, "0.00")
         If adoPty!Type = 6 Or adoPty!Type = 24 Then
            .TextMatrix(iKount, 8) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 9) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Credit
         End If
         .TextMatrix(iKount, 10) = Format(adoPty.Fields.Item("TransactionID").Value, "0.00")
         adoPty.MoveNext
         iKount = IIf(iChild = 0, iKount + 1, iChild)
         If Not adoPty.EOF Then .AddItem ""
      Wend
   End With

   adoPty.Close
   Set adoPty = Nothing
   Set adoPtyDtl = Nothing

   flxACHistory.row = 0
   flxACHistory.row = 0
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

   conFlxGrid.RowHeightMin = 300
End Sub

Private Sub imgClose_Click()
   picAgentList.Visible = False
End Sub

Private Sub LoadAllAgentFlxGrd()
   Dim conAgent As New ADODB.Connection
   Dim rstAgent As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgent.Open getConnectionString

   szSQL = "SELECT AgentID, AgentNAME, AgentPOSTCODE,  " & _
               "AgentSageSuppAC " & _
           "FROM Agent " & _
           "WHERE InactiveAgent = FALSE " & _
           "ORDER BY AgentNAME;"

   rstAgent.Open szSQL, conAgent, adOpenStatic, adLockReadOnly

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
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
   
   rstAgent.Close
   conAgent.Close
   Set rstAgent = Nothing
   Set conAgent = Nothing
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
   MousePointer = vbHourglass

   Select Case tabMain.Tab
   Case 1:                    'Bank Payment details
      If cboBank_ID.text = "" Or flxOtherBankDetails.TextMatrix(1, 1) = "" Then
         LoadAllBankAC
         flxOtherBankDetails.row = 0
         flxOtherBankDetails.col = 0
      End If
   Case 4:                      'Attachment Files
      If txtAgentID.text <> "" Then _
            Call LoadAttachmentFiles(cmbFiles, txtAgentID.text, "Agent")
   End Select
   MousePointer = vbDefault
End Sub

Private Sub LoadAllBankAC()
   ConfigureFlxOtherBank

   Dim conBank As New ADODB.Connection
   Dim rstBank As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conBank.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.*, tlbBank.* " & _
           "FROM tlbClientBanks, tlbBank, Agent " & _
           "WHERE Agent.AgentID = '" & txtAgentID.text & "' And " & _
             "Agent.BANK_ID = tlbBank.BANK_ID And " & _
             "tlbBank.BANK_ID = tlbClientBanks.BANK_ID"
   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic

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

   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic

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
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

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
   flxOtherBankDetails.RowHeight(0) = 0

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
   txtAgentAddressLine4.Locked = bLock
   txtAgentPostCode.Locked = bLock
   txtAgentHomeTel.Locked = bLock
   txtAgentOfficeTel.Locked = bLock
   txtAgentMobile.Locked = bLock
   txtAgentPersonalEmail.Locked = bLock
   txtAgentOfficeEmail.Locked = bLock
   txtAgentOfficeAddressLine1.Locked = bLock
   txtAgentOfficeAddressLine2.Locked = bLock
   txtAgentOfficeAddressLine3.Locked = bLock
   txtAgentOfficeAddressLine4.Locked = bLock
   txtAgentOfficePostCode.Locked = bLock
End Sub

Private Sub SageSupplierAccCombo()
'   ' Error Handler
'   On Error GoTo Error_Handler
'
'   ' Declare Objects
'   Dim oSDO As SageDataObject120.SDOEngine
'   Dim oWS As SageDataObject120.Workspace
'   Dim oPurchaseRecord As SageDataObject120.PurchaseRecord
'
'   ' Declare Variables
'   Dim szDataPath As String
'
'   ' Create the SDOEngine Object
'   Set oSDO = New SageDataObject120.SDOEngine
'
'   ' Create the Workspace
'   Set oWS = oSDO.Workspaces.Add("Prestige")
'
'   'read datapath from registr
'   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'   If szDataPath = "" Then
'      ' Select Company. The SelectCompany method takes the program install
'      ' folder as a parameter
'      szDataPath = oSDO.SelectCompany(sageDirPath)
'      'Save company name in the registry
'      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
'   Else
'     ' Try to Connect - Will Throw an Exception if it Fails
'      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'
'         Set oPurchaseRecord = oWS.CreateObject("PurchaseRecord")
'
'         Dim TotalRow, TotalCol As Long
'         Dim Data() As String
'         Dim i As Integer
'
'         TotalRow = oPurchaseRecord.Count
'         TotalCol = 2
'         cboAgentSageSuppAC.Clear
'
'         ReDim Data(TotalCol, TotalRow) As String
'
'         oPurchaseRecord.MoveFirst
'         For i = 0 To TotalRow - 1
'            Data(0, i) = CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value)
'            Data(1, i) = CStr(oPurchaseRecord.Fields.Item("NAME").Value)
'            oPurchaseRecord.MoveNext
'         Next i
'         '
'         cboAgentSageSuppAC.Column() = Data()
'         cboAgentSageSuppAC.ColumnCount = TotalCol
'         cboAgentSageSuppAC.BoundColumn = 1
'
'         'Disconnect
'         oWS.Disconnect
'      End If
'   End If
'
'   ' Destroy Objects
'   Set oPurchaseRecord = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'   MsgBox "(pcm_003) The SDO generated the following error: " & oSDO.LastError.text
'
'   Set oPurchaseRecord = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
End Sub

Private Sub txtAgentID_KeyPress(KeyAscii As Integer)
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

Private Sub txtAgentID_LostFocus()
   If txtAgentID.Locked Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL   As String
   Dim szID    As String

   adoConn.Open getConnectionString

   szID = txtAgentID.text

   If (IsAccountExist(szID, adoConn)) Then
      If (Not (txtAgentID.text = szID)) Then
         MsgBox "This ID is already in use. Possible suggestion is '" & szID & "' and you may chose different ID"
         txtAgentID.text = szID
         SelTxtInCtrl txtAgentID
      End If
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub txtAgentName_LostFocus()
'   If txtAgentName.text = "" Then Exit Sub
'
'   Dim szChoice As String, szaChoice() As String
'   Dim adoConn As New ADODB.Connection
'   Dim adoRST As New ADODB.Recordset
'   Dim szSQL As String
'
'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
'   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRST.EOF Then
'      szChoice = adoRST.Fields.Item("Value").Value
'      szaChoice = Split(szChoice, "#")
'   End If
'
'   adoRST.Close
'   Set adoRST = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
'
'   If UBound(szaChoice) > 0 Then
'      If szaChoice(5) <> "" Then
'         If InStr(szaChoice(5), "MA") > 0 Then
'            If bNewEdit And txtAgentID.text = "" Then txtAgentID.text = CreateAgentId(txtAgentName.text)
'         End If
'      Else
'         txtAgentID.Locked = False
'         txtAgentID.SetFocus
'      End If
'   End If
End Sub

Private Function CreateAgentId(szName As String) As String
   Dim szSQL As String, i As Integer, szChar As String, j As Integer

   For i = 1 To Len(szName) - 1
      szChar = UCase(Mid(szName, i, 1))
      If (szChar >= "A" And szChar <= "Z") Then
         CreateAgentId = CreateAgentId & szChar
         j = j + 1
      End If
      If j = 8 Then Exit For
   Next i
End Function

Private Sub txtAgentOfficeEmail_LostFocus()
   Dim szErrMsg As String

   If Trim(txtAgentOfficeEmail.text) <> "" Then
      If Not ValidateEmail(txtAgentOfficeEmail.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Managing Agent Email"
         SelTxtInCtrl txtAgentOfficeEmail
         txtAgentOfficeEmail.SetFocus
      End If
   End If
End Sub

Private Sub txtAgentPersonalEmail_LostFocus()
   Dim szErrMsg As String

   If Trim(txtAgentPersonalEmail.text) <> "" Then
      If Not ValidateEmail(txtAgentPersonalEmail.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Managing Agent Email"
         SelTxtInCtrl txtAgentPersonalEmail
         txtAgentPersonalEmail.SetFocus
      End If
   End If
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

Private Sub ConfigFlxACHistory()
   Dim szHeader As String, iCol As Integer

   With flxACHistory
      .Clear
      .Cols = 11
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = 230                                                       'Sign
      .ColWidth(1) = Label11(2).Left - Label11(1).Left                         'No
      .ColWidth(2) = Label11(3).Left - Label11(2).Left                         'Type
      .ColWidth(3) = Label11(4).Left - Label11(3).Left                         'Date
      .ColWidth(4) = Label11(5).Left - Label11(4).Left                         'Reference
      .ColWidth(5) = Label11(6).Left - Label11(5).Left                         'Description
      .ColWidth(6) = Label11(7).Left - Label11(6).Left                         'Amount
      .ColWidth(7) = Label11(8).Left - Label11(7).Left                         'Balance
      .ColWidth(8) = Label11(9).Left - Label11(8).Left                         'Debit
      .ColWidth(9) = .ColWidth(8)                                              'Credit
      .ColWidth(10) = 0                                                        'Transaction ID
   End With
   ConfigFlxACHistorySplit
End Sub

Private Sub ConfigFlxACHistorySplit()
   Dim szHeader As String, iCol As Integer

   With flxACHistorySplit
      .Clear
      .Cols = 12
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = Label11(11).Left - Label11(10).Left                         'No
      .ColWidth(1) = Label11(12).Left - Label11(11).Left                         'Type
      .ColWidth(2) = Label11(13).Left - Label11(12).Left                         'Date
      .ColWidth(3) = Label11(14).Left - Label11(13).Left                         'Ref
      .ColWidth(4) = Label11(15).Left - Label11(14).Left                         'N/C
      .ColWidth(5) = Label11(16).Left - Label11(15).Left                         'Job
      .ColWidth(6) = Label11(17).Left - Label11(16).Left                         'Unit
      .ColWidth(7) = Label11(18).Left - Label11(17).Left                         'Fund
      .ColWidth(8) = Label11(19).Left - Label11(18).Left                         'Desc
      .ColWidth(9) = Label11(20).Left - Label11(19).Left                         'Total
      .ColWidth(10) = Label11(21).Left - Label11(20).Left                        'Debit
      .ColWidth(11) = .Width - Label11(21).Left - 100                            'Credit
   End With
End Sub
