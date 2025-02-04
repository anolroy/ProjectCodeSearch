VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGlobal11 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7125
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   9765
   Icon            =   "Global.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9765
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
      Height          =   1515
      Left            =   4680
      TabIndex        =   87
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2672
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13553358
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      WordWrap        =   -1  'True
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame10 
      Caption         =   "Demand Notice Period"
      Height          =   855
      Left            =   5760
      TabIndex        =   94
      Top             =   2400
      Width           =   3915
      Begin VB.TextBox txt7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1360
         TabIndex        =   95
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Send Demands            Days Before Due Date"
         Height          =   195
         Left            =   180
         TabIndex        =   96
         Top             =   360
         Width           =   3240
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Bank"
      Height          =   455
      Left            =   8280
      TabIndex        =   92
      Top             =   6555
      Width           =   1335
   End
   Begin VB.Frame Frame9 
      Caption         =   "Property"
      Height          =   615
      Left            =   120
      TabIndex        =   89
      Top             =   80
      Width           =   9555
      Begin MSDataListLib.DataCombo cboProperty 
         Bindings        =   "Global.frx":08CA
         DataSource      =   "adoProperty"
         Height          =   315
         Left            =   2880
         TabIndex        =   91
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PropertyName"
         BoundColumn     =   "PropertyID"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoProperty 
         Height          =   330
         Left            =   7320
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Property:"
         Height          =   255
         Left            =   840
         TabIndex        =   90
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCompanySetup 
      Caption         =   "Co&mpany Setup"
      Height          =   455
      Left            =   6910
      TabIndex        =   88
      Top             =   6555
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2955
      Left            =   60
      TabIndex        =   43
      Top             =   3405
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   5212
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Monthly Payment Dates"
      TabPicture(0)   =   "Global.frx":08E4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Quarterly Payment Dates"
      TabPicture(1)   =   "Global.frx":0900
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Half Yearly payments"
      TabPicture(2)   =   "Global.frx":091C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Yearly payments"
      TabPicture(3)   =   "Global.frx":0938
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Yearly Payment Date"
         Height          =   2295
         Left            =   -69480
         TabIndex        =   83
         Top             =   420
         Width           =   2535
         Begin VB.ComboBox cboM7 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   85
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD7 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            TabIndex        =   84
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Half Yearly Payment Dates"
         Height          =   2295
         Left            =   -71160
         TabIndex        =   76
         Top             =   420
         Width           =   3135
         Begin VB.ComboBox cboD5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   77
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   78
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   79
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   80
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "1st"
            Height          =   195
            Left            =   240
            TabIndex        =   81
            Top             =   360
            Width           =   210
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2295
         Left            =   6420
         TabIndex        =   66
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboDay9 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboMonth9 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboDay10 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   28
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboMonth10 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   29
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboDay11 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   30
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboMonth11 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   31
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboDay12 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   32
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboMonth12 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   33
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "9th"
            Height          =   195
            Left            =   360
            TabIndex        =   70
            Top             =   420
            Width           =   225
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "10th"
            Height          =   195
            Left            =   360
            TabIndex        =   69
            Top             =   900
            Width           =   315
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "11th"
            Height          =   195
            Left            =   360
            TabIndex        =   68
            Top             =   1380
            Width           =   315
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "12th"
            Height          =   195
            Left            =   360
            TabIndex        =   67
            Top             =   1860
            Width           =   315
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2295
         Left            =   3300
         TabIndex        =   62
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboDay5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboMonth5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboDay6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   20
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboMonth6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   21
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboDay7 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   22
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboMonth7 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   23
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboDay8 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboMonth8 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   25
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "7th"
            Height          =   195
            Left            =   360
            TabIndex        =   71
            Top             =   1380
            Width           =   225
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "5th"
            Height          =   195
            Left            =   360
            TabIndex        =   65
            Top             =   420
            Width           =   225
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "6th"
            Height          =   195
            Left            =   360
            TabIndex        =   64
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "8th"
            Height          =   195
            Left            =   360
            TabIndex        =   63
            Top             =   1860
            Width           =   225
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Quarterly Payment Dates"
         Height          =   2295
         Left            =   -73080
         TabIndex        =   49
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboD1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   50
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboM1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboD2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   52
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboM2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   53
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboD3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   54
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboM3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   55
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboD4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   56
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboM4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   57
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "1st"
            Height          =   195
            Left            =   360
            TabIndex        =   61
            Top             =   360
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
            Height          =   195
            Left            =   360
            TabIndex        =   60
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "3rd"
            Height          =   195
            Left            =   360
            TabIndex        =   59
            Top             =   1320
            Width           =   225
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "4th"
            Height          =   195
            Left            =   360
            TabIndex        =   58
            Top             =   1800
            Width           =   225
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2295
         Left            =   180
         TabIndex        =   44
         Top             =   420
         Width           =   3015
         Begin VB.ComboBox cboMonth4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   17
            Top             =   1800
            Width           =   1335
         End
         Begin VB.ComboBox cboDay4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   16
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboMonth3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cboDay3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   14
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cboMonth2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboDay2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   12
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cboMonth1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboDay1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "4th"
            Height          =   195
            Left            =   360
            TabIndex        =   48
            Top             =   1860
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "3rd"
            Height          =   195
            Left            =   360
            TabIndex        =   47
            Top             =   1380
            Width           =   225
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "2nd"
            Height          =   195
            Left            =   360
            TabIndex        =   46
            Top             =   900
            Width           =   270
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "1st"
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   420
            Width           =   210
         End
      End
   End
   Begin VB.CommandButton cmdDemandChange 
      Caption         =   "Demand &Fields"
      Height          =   455
      Left            =   80
      TabIndex        =   41
      Top             =   6555
      Width           =   1335
   End
   Begin VB.CommandButton cmdDemandTypes 
      Caption         =   "&Demand/Charge Types"
      Height          =   455
      Left            =   5544
      TabIndex        =   40
      Top             =   6555
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      Height          =   455
      Left            =   4178
      TabIndex        =   34
      Top             =   6555
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   455
      Left            =   1446
      TabIndex        =   86
      Top             =   6555
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Data"
      Height          =   455
      Left            =   2812
      TabIndex        =   39
      Top             =   6555
      Width           =   1335
   End
   Begin VB.Frame Frame8 
      Caption         =   "Other Charges"
      Height          =   1380
      Left            =   5760
      TabIndex        =   72
      Top             =   800
      Width           =   3915
      Begin VB.TextBox txt5 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboVatRate 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Interest Rate:"
         Height          =   195
         Left            =   180
         TabIndex        =   75
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   2850
         TabIndex        =   74
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Rate:"
         Height          =   195
         Left            =   180
         TabIndex        =   73
         Top             =   900
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Year Budget"
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   800
      Width           =   5535
      Begin VB.TextBox txtYearlyInsurance 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdExpandBankCode 
         Caption         =   "v"
         Height          =   285
         Left            =   4680
         TabIndex        =   6
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox txt2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txt4 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txt3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtGlobalBankAccount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Yearly Insurance Charge:"
         Height          =   195
         Left            =   540
         TabIndex        =   93
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Financial Year End:"
         Height          =   195
         Left            =   540
         TabIndex        =   38
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Service Charge Per Square Foot:"
         Height          =   195
         Left            =   540
         TabIndex        =   37
         Top             =   960
         Width           =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Service Charge:"
         Height          =   195
         Left            =   540
         TabIndex        =   36
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Area (Sq feet):"
         Height          =   195
         Left            =   540
         TabIndex        =   35
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Global Bank Account:"
         Height          =   195
         Left            =   540
         TabIndex        =   42
         Top             =   1680
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmGlobal11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New RDO.rdoConnection
Dim Env As rdoEnvironment
Dim Envs As rdoEnvironments
Dim Rst As rdoResultset
Dim SQLStr As String
Dim SCperSqFoot As Double
Dim Conn1 As New RDO.rdoConnection
Dim Rst1 As rdoResultset
Dim Conn2 As New RDO.rdoConnection
Dim Rst2 As rdoResultset
Dim SQLStr1 As String
Dim OldSCPerSqFoot As Double

Dim bEditGlobalData As Boolean

Private Sub cboVatRate_LostFocus()
   Dim i, match As Integer
   match = 0
   If cboVatRate.text = "" Then Exit Sub
   For i = 0 To 12
       If cboVatRate.text = cboVatRate.List(i) Then
           match = 1
           Exit For
       End If
   Next i
   If match = 0 Then cboVatRate.text = ""
End Sub

Private Sub cboD1_LostFocus()
'
'Dim i As Integer
'Dim match As Integer
'Dim j As Integer
'
'j = cboD1.ListCount - 1
'For i = 0 To j
'    If cboD1.List(i) = cboD1.text Then
'        match = 1
'        Exit For
'    End If
'Next i
'If match = 0 Then
'    MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'    cboD1.text = ""
'    Exit Sub
'End If

End Sub

Private Sub cboD2_LostFocus()

'Dim i As Integer
'Dim j As Integer
'Dim match As Integer
'j = cboD2.ListCount - 1
'For i = 0 To j
'    If cboD2.List(i) = cboD2.text Then
'        match = 1
'        Exit For
'    End If
'Next i
'If match = 0 Then
'    MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'    cboD2.text = ""
'    Exit Sub
'End If

End Sub

Private Sub cboD3_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'    j = cboD3.ListCount - 1
'    For i = 0 To j
'        If cboD3.List(i) = cboD3.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboD3.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboD4_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'    j = cboD4.ListCount - 1
'    For i = 0 To j
'        If cboD4.List(i) = cboD4.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboD4.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboD5_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboD5.ListCount - 1
'    For i = 0 To j
'        If cboD5.List(i) = cboD5.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboD5.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboD6_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboD6.ListCount - 1
'    For i = 0 To j
'        If cboD6.List(i) = cboD6.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboD6.text = ""
'        Exit Sub
'    End If
End Sub

Private Sub cboD7_Change()
'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboD7.ListCount - 1
'    For i = 0 To j
'        If cboD7.List(i) = cboD7.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboD7.text = ""
'        Exit Sub
'    End If
End Sub


Private Sub cboM1_LostFocus()
    
'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM1.ListCount - 1
'    For i = 0 To j
'        If cboM1.List(i) = cboM1.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM1.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboM2_LostFocus()
    
'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM2.ListCount - 1
'    For i = 0 To j
'        If cboM2.List(i) = cboM2.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM2.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboM3_LostFocus()
    
'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM3.ListCount - 1
'    For i = 0 To j
'        If cboM3.List(i) = cboM3.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM3.text = ""
'        Exit Sub
'    End If
End Sub

Private Sub cboM4_LostFocus()
'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM4.ListCount - 1
'    For i = 0 To j
'        If cboM4.List(i) = cboM4.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM4.text = ""
'        Exit Sub
'    End If
'   SSTab1.SetFocus
End Sub

Private Sub cboM5_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM5.ListCount - 1
'    For i = 0 To j
'        If cboM5.List(i) = cboM5.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM5.text = ""
'        Exit Sub
'    End If

End Sub

Private Sub cboM6_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM6.ListCount - 1
'    For i = 0 To j
'        If cboM6.List(i) = cboM6.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM6.text = ""
'        Exit Sub
'    End If
    
'    SSTab1.SetFocus
End Sub

Private Sub cboM7_LostFocus()

'    Dim i As Integer
'    Dim j As Integer
'    Dim match As Integer
'
'    j = cboM7.ListCount - 1
'    For i = 0 To j
'        If cboM7.List(i) = cboM7.text Then
'            match = 1
'            Exit For
'        End If
'    Next i
'    If match = 0 Then
'        MsgBox "Invalid date selected.", vbOKOnly + vbCritical, "Invalid Date"
'        cboM7.text = ""
'        Exit Sub
'    End If
End Sub

Private Sub cboProperty_Change()
   If Not GetData Then
      If (MsgBox("There is no Global Data setup for the property " & cboProperty.text & ". Do you like to setup", vbYesNo, "No Global Data") = vbYes) Then
         Edit
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
    Call GetData
    Call DisableBoxes
End Sub

Private Sub cmdCompanySetup_Click()
   frmShoppingCentre.Show
End Sub

Private Sub cmdDemandChange_Click()
    'show form to change appearance of demand
    frmReportFields.Show
End Sub

Private Sub cmdDemandTypes_Click()

    Load frmDemandTypes
    Me.Hide
    frmDemandTypes.Show
    frmDemandTypes.SetFocus

End Sub

Private Sub cmdEdit_Click()
   If cboProperty.text = "" Then Exit Sub
   Call Edit
   bEditGlobalData = True
End Sub

Private Sub cmdExpandBankCode_Click()
   gridBankCode.Left = txtGlobalBankAccount.Left + Frame1.Left
   gridBankCode.Top = txtGlobalBankAccount.Top + txtGlobalBankAccount.Height + Frame1.Top + 5
   BankAccount
End Sub

Private Sub cmdExpandBankCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then gridBankCode.Visible = False
End Sub

Private Sub cmdSave_Click()
   'have to update the SCTotal in Lease table
   
   Dim tempdate As String
   Dim VatCode As Integer

   'where decimal point used put in relevant 0's
   If txt3.text <> "" Then txt3.text = CheckDecimal(txt3.text)
   If txt5.text <> "" Then txt5.text = CheckDecimal2(txt5.text)

   'make sure all payment dates are entered.
   If MissingDate(cboDay1.text) = True Then Exit Sub
   If MissingDate(cboDay2.text) = True Then Exit Sub
   If MissingDate(cboDay3.text) = True Then Exit Sub
   If MissingDate(cboDay4.text) = True Then Exit Sub
   If MissingDate(cboDay5.text) = True Then Exit Sub
   If MissingDate(cboDay6.text) = True Then Exit Sub
   If MissingDate(cboDay7.text) = True Then Exit Sub
   If MissingDate(cboDay8.text) = True Then Exit Sub
   If MissingDate(cboDay9.text) = True Then Exit Sub
   If MissingDate(cboDay10.text) = True Then Exit Sub
   If MissingDate(cboDay11.text) = True Then Exit Sub
   If MissingDate(cboDay12.text) = True Then Exit Sub
   
   If MissingDate(cboD1.text) = True Then Exit Sub
   If MissingDate(cboD2.text) = True Then Exit Sub
   If MissingDate(cboD3.text) = True Then Exit Sub
   If MissingDate(cboD4.text) = True Then Exit Sub
   If MissingDate(cboD5.text) = True Then Exit Sub
   If MissingDate(cboD6.text) = True Then Exit Sub
   If MissingDate(cboD7.text) = True Then Exit Sub
   
   If MissingDate(cboMonth1.text) = True Then Exit Sub
   If MissingDate(cboMonth2.text) = True Then Exit Sub
   If MissingDate(cboMonth3.text) = True Then Exit Sub
   If MissingDate(cboMonth4.text) = True Then Exit Sub
   If MissingDate(cboMonth5.text) = True Then Exit Sub
   If MissingDate(cboMonth6.text) = True Then Exit Sub
   If MissingDate(cboMonth7.text) = True Then Exit Sub
   If MissingDate(cboMonth8.text) = True Then Exit Sub
   If MissingDate(cboMonth9.text) = True Then Exit Sub
   If MissingDate(cboMonth10.text) = True Then Exit Sub
   If MissingDate(cboMonth11.text) = True Then Exit Sub
   If MissingDate(cboMonth12.text) = True Then Exit Sub
   
   If MissingDate(cboM1.text) = True Then Exit Sub
   If MissingDate(cboM2.text) = True Then Exit Sub
   If MissingDate(cboM3.text) = True Then Exit Sub
   If MissingDate(cboM4.text) = True Then Exit Sub
   If MissingDate(cboM5.text) = True Then Exit Sub
   If MissingDate(cboM6.text) = True Then Exit Sub
   If MissingDate(cboM7.text) = True Then Exit Sub
   
   'validate the dates.
   tempdate = cboMonth1.text & " " & cboDay1.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth2.text & " " & cboDay2.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth3.text & " " & cboDay3.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth4.text & " " & cboDay4.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth5.text & " " & cboDay5.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth6.text & " " & cboDay6.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth7.text & " " & cboDay7.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth8.text & " " & cboDay8.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth9.text & " " & cboDay9.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth10.text & " " & cboDay10.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth11.text & " " & cboDay11.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboMonth12.text & " " & cboDay12.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub

   tempdate = cboM1.text & " " & cboD1.text & ", 20" & Right(Date, 2)

   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM2.text & " " & cboD2.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM3.text & " " & cboD3.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM4.text & " " & cboD4.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM5.text & " " & cboD5.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM6.text & " " & cboD6.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub
   tempdate = cboM7.text & " " & cboD7.text & ", 20" & Right(Date, 2)
   If ValidDate(tempdate) = False Then Exit Sub

   Dim i As Integer

   If cboVatRate.text = "" Then
       VatCode = 1
       For i = 1 To cboVatRate.ListCount
           If Left(cboVatRate.List(i), 1) = "1" Then VatCode = cboVatRate.text = cboVatRate.List(i)
       Next i
   Else
       For i = 2 To 4
           If Mid(cboVatRate.text, i, 3) = " / " Then VatCode = Left(cboVatRate.text, i - 1)
       Next i
   End If

   Dim conn3 As New RDO.rdoConnection

'* Save records in the database
   conn3.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conn3.CursorDriver = rdUseOdbc
   conn3.EstablishConnection rdDriverNoPrompt

   Set Rst = conn3.OpenResultset("SELECT * " & _
                                 "FROM GlobalData " & _
                                 "WHERE PropertyID = '" & cboProperty.BoundText & "' ", _
                                         rdOpenDynamic, rdConcurRowVer)
   If Rst.RowCount = 0 Then
       Rst.AddNew
   Else
       Rst.MoveFirst
       Rst.Edit
   End If

   If cboProperty.text <> "" Then
       Rst!PROPERTYID = cboProperty.BoundText
   Else
       MsgBox "Please select a property to continue", vbInformation, "Save global data"
       Rst.Close
       conn3.Close
       Exit Sub
   End If

   If txt1.text <> "" Then Rst!TotalArea = CLng(txt1.text)
   If txt2.text <> "" Then Rst!TotalSC = CDbl(txt2.text)
   If txt3.text <> "" Then Rst!SCperSqFoot = CDbl(txt3.text)
   'inserted line below at 10:09am 19/02/2003
   scperfoot = CDbl(txt3.text)

   If txt4.text <> "" Then Rst!SCYearEnd = txt4.text
   If txt5.text <> "" Then Rst!BaseInterestRate = CDbl(txt5.text)

   If txtGlobalBankAccount.text <> "" Then Rst!GlobalBankCode = txtGlobalBankAccount.text

   Rst!VatRate = VatCode
   
   If Rst!MonthlyDueDate1 <> cboDay1.text & " " & cboMonth1.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate1, cboDay1.text & " " & cboMonth1.text
   Rst!MonthlyDueDate1 = cboDay1.text & " " & cboMonth1.text
   If Rst!MonthlyDueDate2 <> cboDay2.text & " " & cboMonth2.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate2, cboDay2.text & " " & cboMonth2.text
   Rst!MonthlyDueDate2 = cboDay2.text & " " & cboMonth2.text
   If Rst!MonthlyDueDate3 <> cboDay3.text & " " & cboMonth3.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate3, cboDay3.text & " " & cboMonth3.text
   Rst!MonthlyDueDate3 = cboDay3.text & " " & cboMonth3.text
   If Rst!MonthlyDueDate4 <> cboDay4.text & " " & cboMonth4.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate4, cboDay4.text & " " & cboMonth4.text
   Rst!MonthlyDueDate4 = cboDay4.text & " " & cboMonth4.text
   If Rst!MonthlyDueDate5 <> cboDay5.text & " " & cboMonth5.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate5, cboDay5.text & " " & cboMonth5.text
   Rst!MonthlyDueDate5 = cboDay5.text & " " & cboMonth5.text
   If Rst!MonthlyDueDate6 <> cboDay6.text & " " & cboMonth6.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate6, cboDay6.text & " " & cboMonth6.text
   Rst!MonthlyDueDate6 = cboDay6.text & " " & cboMonth6.text
   If Rst!MonthlyDueDate7 <> cboDay7.text & " " & cboMonth7.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate7, cboDay7.text & " " & cboMonth7.text
   Rst!MonthlyDueDate7 = cboDay7.text & " " & cboMonth7.text
   If Rst!MonthlyDueDate8 <> cboDay8.text & " " & cboMonth8.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate8, cboDay8.text & " " & cboMonth8.text
   Rst!MonthlyDueDate8 = cboDay8.text & " " & cboMonth8.text
   If Rst!MonthlyDueDate9 <> cboDay9.text & " " & cboMonth9.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate9, cboDay9.text & " " & cboMonth9.text
   Rst!MonthlyDueDate9 = cboDay9.text & " " & cboMonth9.text
   If Rst!MonthlyDueDate10 <> cboDay10.text & " " & cboMonth10.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate10, cboDay10.text & " " & cboMonth10.text
   Rst!MonthlyDueDate10 = cboDay10.text & " " & cboMonth10.text
   If Rst!MonthlyDueDate11 <> cboDay11.text & " " & cboMonth11.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate11, cboDay11.text & " " & cboMonth11.text
   Rst!MonthlyDueDate11 = cboDay11.text & " " & cboMonth11.text
   If Rst!MonthlyDueDate12 <> cboDay12.text & " " & cboMonth12.text Then UpdateLeasePaymentDate conn3, 5, Rst!MonthlyDueDate12, cboDay12.text & " " & cboMonth12.text
   Rst!MonthlyDueDate12 = cboDay12.text & " " & cboMonth12.text

   If Rst!QuarterlyDueDate1 <> cboD1.text & " " & cboM1.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate1, cboD1.text & " " & cboM1.text
   Rst!QuarterlyDueDate1 = cboD1.text & " " & cboM1.text
   If Rst!QuarterlyDueDate2 <> cboD2.text & " " & cboM2.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate2, cboD2.text & " " & cboM2.text
   Rst!QuarterlyDueDate2 = cboD2.text & " " & cboM2.text
   If Rst!QuarterlyDueDate3 <> cboD3.text & " " & cboM3.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate3, cboD3.text & " " & cboM3.text
   Rst!QuarterlyDueDate3 = cboD3.text & " " & cboM3.text
   If Rst!QuarterlyDueDate4 <> cboD4.text & " " & cboM4.text Then UpdateLeasePaymentDate conn3, 7, Rst!QuarterlyDueDate4, cboD4.text & " " & cboM4.text
   Rst!QuarterlyDueDate4 = cboD4.text & " " & cboM4.text
   
   If Rst!HalfYearlyDueDate1 <> cboD5.text & " " & cboM5.text Then UpdateLeasePaymentDate conn3, 9, Rst!HalfYearlyDueDate1, cboD5.text & " " & cboM5.text
   Rst!HalfYearlyDueDate1 = cboD5.text & " " & cboM5.text
   If Rst!HalfYearlyDueDate2 <> cboD6.text & " " & cboM6.text Then UpdateLeasePaymentDate conn3, 9, Rst!HalfYearlyDueDate2, cboD6.text & " " & cboM6.text
   Rst!HalfYearlyDueDate2 = cboD6.text & " " & cboM6.text
   
   If Rst!YearlyDueDate <> cboD7.text & " " & cboM7.text Then UpdateLeasePaymentDate conn3, 11, Rst!YearlyDueDate, cboD7.text & " " & cboM7.text
   Rst!YearlyDueDate = cboD7.text & " " & cboM7.text
   
   Rst!NoOfDaysToSendDemandsB4Due = CInt(txt7.text)
   Rst!YearlyInsurance = CDbl(IIf(txtYearlyInsurance.text = "", 0, txtYearlyInsurance.text))
   Rst.Update
   
   Rst.Close
   conn3.Close
   
   '////////////////////////////////////////////////////////////////////////////////////////
   'Try to reopen a connection to the lease and unit tables to update SCTotal in Lease Table

   Dim StringForCommon As String
   Dim TheArea As Double
   Dim TheFrequency As Integer
   Dim TheAmount As Double

   'save to record
   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseOdbc
   Conn.EstablishConnection rdDriverNoPrompt

   Set Rst = Conn.OpenResultset("SELECT * FROM LeaseDetails", rdOpenDynamic, rdConcurRowVer)

   If Rst.EOF Then GoTo NoLease

   Rst.MoveFirst
   
   While Rst.EOF = False
      StringForCommon = Rst!UnitNumber
      Set Rst1 = Conn.OpenResultset("SELECT * FROM Units WHERE UnitNumber = '" & StringForCommon & "' ", rdOpenDynamic, rdConcurRowVer)

      If Rst1.RowCount <> 0 Then

         TheArea = CDbl(Rst1!TotalArea)
         Rst.Edit

         TheAmount = Round((TheArea * scperfoot), 2)
         Rst!SCTotal = TheAmount

         '////////////////////Add in to amend amount per period///////////////////////////////

         TheFrequency = Rst!SCfrequency

         Select Case TheFrequency
             Case 1: 'weekly
                 Rst!SCAmount = Round((TheAmount / 52), 2)
             Case 2: 'weekly
                 Rst!SCAmount = Round((TheAmount / 52), 2)
             Case 3: 'Fortnightly in advance
                 Rst!SCAmount = Round((TheAmount / 26), 2)
             Case 4: 'Fortnightly in arrears
                 Rst!SCAmount = Round((TheAmount / 26), 2)
             Case 5: 'Monthly in advance
                 Rst!SCAmount = Round((TheAmount / 12), 2)
             Case 6: 'Monthly in arrears
                 Rst!SCAmount = Round((TheAmount / 12), 2)
             Case 7: 'Quarterly in advance
                 Rst!SCAmount = Round((TheAmount / 4), 2)
             Case 8: 'Quarterly in arrears
                 Rst!SCAmount = Round((TheAmount / 4), 2)
             Case 9: 'Half-yearly in advance
                 Rst!SCAmount = Round((TheAmount / 2), 2)
             Case 10: 'Half-yearly in arrears
                 Rst!SCAmount = Round((TheAmount / 2), 2)
             Case 11: 'Yearly in advance
                 Rst!SCAmount = Round(TheAmount, 2)
             Case 12: 'Yearly in arrears
                 Rst!SCAmount = Round(TheAmount, 2)
         End Select

         '////////////////////////////////////////////////////////////////////////////////////

         Rst.Update
      End If
      Rst.MoveNext
   Wend

   'Rst.Update
   Rst1.Close
NoLease:
   Rst.Close
   Conn.Close

   '////////////////////////////////////////////////////////////////////////////////////////
   'Try to reopen a connection to the lease and unit tables to update SCTotal in Lease Table

   MsgBox "Your changes have been saved.", vbOKOnly + vbInformation, "Saved"

   Call DisableBoxes
   Call GetGlobalData
End Sub

Private Sub UpdateLeasePaymentDate(dbConn As RDO.rdoConnection, iFrequency As Integer, szCurDate As String, szNewDate As String)
   If szCurDate = "" Then Exit Sub     'New Global Data Entry
   
   Dim rdoRec As rdoResultset
   Dim dtNextDueDate As Date, dtNewDueDate As Date

   dtNextDueDate = CDate(Format(szCurDate, "dd mmmm yyyy"))
   dtNewDueDate = CDate(Format(szNewDate, "dd mmmm yyyy"))

    'Service Charge
      Set rdoRec = dbConn.OpenResultset("SELECT SCNextDueDate, UnitNumber " & _
                     "FROM LeaseDetails " & _
                     "WHERE LeaseDetails.Status = True And " & _
                        "(SCFrequency = " & iFrequency & " or " & _
                        "SCFrequency = " & iFrequency + 1 & ") And " & _
                        "SCPayable = 'Y'", _
                           rdOpenDynamic, rdConcurRowVer)
      If rdoRec.EOF Then
         rdoRec.Close
         Set rdoRec = Nothing
      Else
         rdoRec.MoveFirst
         While Not rdoRec.EOF
            If Format(rdoRec!SCNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
               InThisProperty(dbConn, rdoRec!UnitNumber) Then
               rdoRec.Edit
               rdoRec!SCNextDueDate = CDate(Format(dtNewDueDate, "dd mmmm") & " " & Format(rdoRec!SCNextDueDate, "yyyy"))
               rdoRec.Update
            End If
            rdoRec.MoveNext
         Wend
         rdoRec.Close
      End If

   'Rent
      Set rdoRec = dbConn.OpenResultset("SELECT BRNextDueDate, UnitNumber " & _
                     "FROM LeaseDetails " & _
                     "WHERE LeaseDetails.Status = True And " & _
                        "(BRFrequency = " & iFrequency & " or " & _
                        "BRFrequency = " & iFrequency + 1 & ") And " & _
                        "BRPayable = 'Y'", _
                           rdOpenDynamic, rdConcurRowVer)
      If rdoRec.EOF Then
         rdoRec.Close
         Set rdoRec = Nothing
      Else
         rdoRec.MoveFirst
         While Not rdoRec.EOF
            If Format(rdoRec!BRNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
               InThisProperty(dbConn, rdoRec!UnitNumber) Then
               rdoRec.Edit
               rdoRec!BRNextDueDate = CDate(Format(dtNewDueDate, "dd mmmm") & " " & Format(rdoRec!BRNextDueDate, "yyyy"))
               rdoRec.Update
            End If
            rdoRec.MoveNext
         Wend
         rdoRec.Close
      End If

   'Insurance
      Set rdoRec = dbConn.OpenResultset("SELECT InsuranceNextDueDate, UnitNumber " & _
                     "FROM LeaseDetails " & _
                     "WHERE LeaseDetails.Status = True And " & _
                        "(InsuranceFrequency = " & iFrequency & " or " & _
                        "InsuranceFrequency = " & iFrequency + 1 & ") And " & _
                        "InsurancePayable = 'Y'", _
                           rdOpenDynamic, rdConcurRowVer)
      If rdoRec.EOF Then
         rdoRec.Close
         Set rdoRec = Nothing
      Else
         rdoRec.MoveFirst
         While Not rdoRec.EOF
            If Format(rdoRec!InsuranceNextDueDate, "dd mmmm") = Format(dtNextDueDate, "dd mmmm") And _
               InThisProperty(dbConn, rdoRec!UnitNumber) Then
               rdoRec.Edit
               rdoRec!InsuranceNextDueDate = CDate(Format(dtNewDueDate, "dd mmmm") & " " & Format(rdoRec!InsuranceNextDueDate, "yyyy"))
               rdoRec.Update
            End If
            rdoRec.MoveNext
         Wend
         rdoRec.Close
      End If
End Sub

Private Function InThisProperty(dbConn As rdoConnection, szUnitNumber As String) As Boolean
   Dim rdoRec As rdoResultset
   
   Set rdoRec = dbConn.OpenResultset("SELECT * " & _
                  "FROM Units " & _
                  "WHERE Units.UnitNumber = '" & szUnitNumber & "' And " & _
                     "Units.PropertyID = '" & cboProperty.BoundText & "'", _
                        rdOpenDynamic, rdConcurRowVer)
   If rdoRec.EOF Then
      InThisProperty = False
   Else
      InThisProperty = True
   End If
   
   rdoRec.Close
   Set rdoRec = Nothing
End Function


Private Sub Command1_Click()
   Load frmBank
   Me.Hide
   frmBank.Show
   frmBank.SetFocus
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50

   Me.Caption = "Global Data"

   Call LoadProperty
   Call FillDaysMonths
   Call GetVATRates

   bEditGlobalData = False
   SSTab1.Tab = 0
End Sub

Public Sub LoadProperty()
   Dim sSQLQuery_ As String
   adoProperty.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
     
   sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
                 "FROM PROPERTY "
                 

   adoProperty.RecordSource = sSQLQuery_
   adoProperty.CommandType = adCmdText
   adoProperty.Refresh
End Sub

Public Function GetData() As Boolean
   Dim i As Integer
   Dim a As String
   Dim b As String
   Dim c As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt
   
   SQLStr = "SELECT * FROM GlobalData WHERE PropertyID = '" & cboProperty.BoundText & "' "
   Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
   
   If Rst.RowCount = 0 Then
       Rst.Close
       Conn.Close
       GetData = False
       Exit Function
   End If
   If IsNull(Rst!TotalArea) Then txt1.text = "" Else txt1.text = Rst!TotalArea
   If IsNull(Rst!TotalSC) Then txt2.text = "" Else txt2.text = Rst!TotalSC
   If IsNull(Rst!SCperSqFoot) = False Then
       txt3.text = Rst!SCperSqFoot
       txt3.text = CheckDecimal(txt3.text)
       
       'Line added on 15/02/03 to set up a ratio system
       OldSCPerSqFoot = txt3.text
   End If
   If IsNull(Rst!SCYearEnd) Then txt4.text = "" Else txt4.text = Rst!SCYearEnd
   If IsNull(Rst!GlobalBankCode) Then txtGlobalBankAccount.text = "" Else txtGlobalBankAccount.text = Rst!GlobalBankCode
   
   If IsNull(Rst!BaseInterestRate) = False Then
       txt5.text = Rst!BaseInterestRate
       txt5.text = CheckDecimal2(txt5.text)
   End If
   For i = 0 To cboVatRate.ListCount - 1
       c = cboVatRate.List(i)
       If CInt(Left(c, 2)) = Rst!VatRate Then
           cboVatRate.text = c
       End If
   Next i
   
   If IsNull(Rst!NoOfDaysToSendDemandsB4Due) = False Then txt7.text = Rst!NoOfDaysToSendDemandsB4Due
   
   cboDay1.text = Left(Rst!MonthlyDueDate1, 2)
   cboMonth1.text = Right(Rst!MonthlyDueDate1, Len(Rst!MonthlyDueDate1) - 3)
   
   cboDay2.text = Left(Rst!MonthlyDueDate2, 2)
   cboMonth2.text = Right(Rst!MonthlyDueDate2, Len(Rst!MonthlyDueDate2) - 3)
   
   cboDay3.text = Left(Rst!MonthlyDueDate3, 2)
   cboMonth3.text = Right(Rst!MonthlyDueDate3, Len(Rst!MonthlyDueDate3) - 3)
   
   cboDay4.text = Left(Rst!MonthlyDueDate4, 2)
   cboMonth4.text = Right(Rst!MonthlyDueDate4, Len(Rst!MonthlyDueDate4) - 3)
   
   cboDay5.text = Left(Rst!MonthlyDueDate5, 2)
   cboMonth5.text = Right(Rst!MonthlyDueDate5, Len(Rst!MonthlyDueDate5) - 3)
   
   cboDay6.text = Left(Rst!MonthlyDueDate6, 2)
   cboMonth6.text = Right(Rst!MonthlyDueDate6, Len(Rst!MonthlyDueDate6) - 3)
   
   cboDay7.text = Left(Rst!MonthlyDueDate7, 2)
   cboMonth7.text = Right(Rst!MonthlyDueDate7, Len(Rst!MonthlyDueDate7) - 3)
   
   cboDay8.text = Left(Rst!MonthlyDueDate8, 2)
   cboMonth8.text = Right(Rst!MonthlyDueDate8, Len(Rst!MonthlyDueDate8) - 3)
   
   cboDay9.text = Left(Rst!MonthlyDueDate9, 2)
   cboMonth9.text = Right(Rst!MonthlyDueDate9, Len(Rst!MonthlyDueDate9) - 3)
   
   cboDay10.text = Left(Rst!MonthlyDueDate10, 2)
   cboMonth10.text = Right(Rst!MonthlyDueDate10, Len(Rst!MonthlyDueDate10) - 3)
   
   cboDay11.text = Left(Rst!MonthlyDueDate11, 2)
   cboMonth11.text = Right(Rst!MonthlyDueDate11, Len(Rst!MonthlyDueDate11) - 3)
   
   cboDay12.text = Left(Rst!MonthlyDueDate12, 2)
   cboMonth12.text = Right(Rst!MonthlyDueDate12, Len(Rst!MonthlyDueDate12) - 3)
   
   cboD2.text = Left(Rst!QuarterlyDueDate2, 2)
   cboM2.text = Right(Rst!QuarterlyDueDate2, Len(Rst!QuarterlyDueDate2) - 3)
   cboD1.text = Left(Rst!QuarterlyDueDate1, 2)
   cboM1.text = Right(Rst!QuarterlyDueDate1, Len(Rst!QuarterlyDueDate1) - 3)
   cboD3.text = Left(Rst!QuarterlyDueDate3, 2)
   cboM3.text = Right(Rst!QuarterlyDueDate3, Len(Rst!QuarterlyDueDate3) - 3)
   cboD4.text = Left(Rst!QuarterlyDueDate4, 2)
   cboM4.text = Right(Rst!QuarterlyDueDate4, Len(Rst!QuarterlyDueDate4) - 3)
   
   
   cboD5.text = Left(Rst!HalfYearlyDueDate1, 2)
   cboM5.text = Right(Rst!HalfYearlyDueDate1, Len(Rst!HalfYearlyDueDate1) - 3)
   cboD6.text = Left(Rst!HalfYearlyDueDate2, 2)
   cboM6.text = Right(Rst!HalfYearlyDueDate2, Len(Rst!HalfYearlyDueDate2) - 3)
   cboD7.text = Left(Rst!YearlyDueDate, 2)
   cboM7.text = Right(Rst!YearlyDueDate, Len(Rst!YearlyDueDate) - 3)
   txtYearlyInsurance.text = Format(IIf(IsNull(Rst!YearlyInsurance), "0.00", Rst!YearlyInsurance), "0.00")
   Rst.Close
   Conn.Close

   GetData = True
   Exit Function
   
ErrH:
   
   MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
End Function


Public Function CheckDecimal(Value As String) As String

    Dim i As Integer
    Dim char As String
    Dim a As Integer
    
    a = 0
    If Asc(Mid(Value, 1, 1)) = 46 Then Value = "0" + Value
    For i = 2 To Len(Value)
        char = Mid(Value, i, 1)
        If Asc(char) = 46 And i = Len(Value) - 2 Then a = 1
        If Asc(char) = 46 And i = Len(Value) - 1 Then
            Value = Value + "0"
            a = 1
        End If
        If Asc(char) = 46 And i = Len(Value) Then
            Value = Value + "00"
            a = 1
        End If
    Next i
    If a = 0 Then Value = Value + ".00"
    CheckDecimal = Value
        
End Function

Public Function CheckDecimal2(Value As String) As String

    Dim i As Integer
    Dim char As String
    Dim a As Integer
    
    a = 0
    If Asc(Mid(Value, 1, 1)) = 46 Then Value = "0" + Value
    For i = 2 To Len(Value)
        char = Mid(Value, i, 1)
        If Asc(char) = 46 And i = Len(Value) - 4 Then a = 1
        If Asc(char) = 46 And i = Len(Value) - 3 Then
            Value = Value + "0"
            a = 1
        End If
        If Asc(char) = 46 And i = Len(Value) - 2 Then
            Value = Value + "00"
            a = 1
        End If
        If Asc(char) = 46 And i = Len(Value) - 1 Then
            Value = Value + "000"
            a = 1
        End If
        If Asc(char) = 46 And i = Len(Value) Then
            Value = Value + "0000"
            a = 1
        End If
    Next i
    If a = 0 Then Value = Value + ".0000"
    CheckDecimal2 = Value

End Function

Public Sub DisableBoxes()
   txt1.Enabled = False
   txt2.Enabled = False
   txt3.Enabled = False
   txt4.Enabled = False
   txt5.Enabled = False
   cboVatRate.Enabled = False
   txt7.Enabled = False
   
   cboDay1.Enabled = False
   cboDay2.Enabled = False
   cboDay3.Enabled = False
   cboDay4.Enabled = False
   cboDay5.Enabled = False
   cboDay6.Enabled = False
   cboDay7.Enabled = False
   cboDay8.Enabled = False
   cboDay9.Enabled = False
   cboDay10.Enabled = False
   cboDay11.Enabled = False
   cboDay12.Enabled = False
   
   cboD1.Enabled = False
   cboD2.Enabled = False
   cboD3.Enabled = False
   cboD4.Enabled = False
   cboD5.Enabled = False
   cboD6.Enabled = False
   cboD7.Enabled = False
   
   
   cboMonth1.Enabled = False
   cboMonth2.Enabled = False
   cboMonth3.Enabled = False
   cboMonth4.Enabled = False
   cboMonth5.Enabled = False
   cboMonth6.Enabled = False
   cboMonth7.Enabled = False
   cboMonth8.Enabled = False
   cboMonth9.Enabled = False
   cboMonth10.Enabled = False
   cboMonth11.Enabled = False
   cboMonth12.Enabled = False
   
   
   cboM1.Enabled = False
   cboM2.Enabled = False
   cboM3.Enabled = False
   cboM4.Enabled = False
   cboM5.Enabled = False
   cboM6.Enabled = False
   cboM7.Enabled = False
   
   
   cmdEdit.Visible = True
   cmdSave.Visible = False
   cmdCancel.Visible = False
   '    mnuEdit.Enabled = True
   cboProperty.Enabled = True
   txtYearlyInsurance.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub gridBankCode_Click()
   Dim iRow As Integer
   iRow = gridBankCode.Row
   txtGlobalBankAccount.text = gridBankCode.TextMatrix(iRow, 0)
   gridBankCode.Visible = False
End Sub

Private Sub gridBankCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim iRow As Integer
      
      iRow = gridBankCode.Row
      txtGlobalBankAccount.text = gridBankCode.TextMatrix(iRow, 0)
      gridBankCode.Visible = False
   End If
End Sub

Private Sub mnuDemands_Click()
'    Load frmDemands1
'    Unload Me
'    frmDemands1.Show
End Sub

Private Sub mnuEdit_Click()
   Call Edit
End Sub

Public Sub Edit()
   txt1.Enabled = True
   txt2.Enabled = True
   txt3.Enabled = True
   txt4.Enabled = True
   txt5.Enabled = True
   cboVatRate.Enabled = True
   txt7.Enabled = True
   
   cboDay1.Enabled = True
   cboDay2.Enabled = True
   cboDay3.Enabled = True
   cboDay4.Enabled = True
   cboDay5.Enabled = True
   cboDay6.Enabled = True
   cboDay7.Enabled = True
   cboDay8.Enabled = True
   cboDay9.Enabled = True
   cboDay10.Enabled = True
   cboDay11.Enabled = True
   cboDay12.Enabled = True
   
   cboD1.Enabled = True
   cboD2.Enabled = True
   cboD3.Enabled = True
   cboD4.Enabled = True
   cboD5.Enabled = True
   cboD6.Enabled = True
   cboD7.Enabled = True
   
   cboMonth1.Enabled = True
   cboMonth2.Enabled = True
   cboMonth3.Enabled = True
   cboMonth4.Enabled = True
   cboMonth5.Enabled = True
   cboMonth6.Enabled = True
   cboMonth7.Enabled = True
   cboMonth8.Enabled = True
   cboMonth9.Enabled = True
   cboMonth10.Enabled = True
   cboMonth11.Enabled = True
   cboMonth12.Enabled = True
   
   cboM1.Enabled = True
   cboM2.Enabled = True
   cboM3.Enabled = True
   cboM4.Enabled = True
   cboM5.Enabled = True
   cboM6.Enabled = True
   cboM7.Enabled = True
   cmdEdit.Visible = False
   cmdSave.Visible = True
   cmdCancel.Visible = True
   cboProperty.Enabled = False
   txtYearlyInsurance.Enabled = True
End Sub

Private Sub mnuExit_Click()
   Unload frmMMain
End Sub

Private Sub mnuLease_Click()
   Load frmLease4
   Unload Me
   frmLease4.Show
End Sub

Private Sub mnuMain_Click()
   Unload Me
End Sub

Private Sub mnuShopCentre_Click()
   Load frmShoppingCentre
   Unload Me
   frmShoppingCentre.Show
End Sub

Private Sub mnuTenants_Click()

'    Load frmTenant
'    Me.Hide
'    frmTenant.Show

End Sub

Private Sub mnuUnits_Click()

'    Load frmUnit
    Unload Me
'    frmUnit.Show

End Sub

Private Sub txt1_LostFocus()

    If txt1.text <> "" Then
        If NumberCheck2(txt1.text) = False Then
            txt1.text = "0"
        End If
    Else
        txt1.text = "0"
    End If
    
    If txt2.text <> "" Then
        If NumberCheck2(txt2.text) = False Then
            txt2.text = "0"
        End If
    Else
        txt2.text = "0"
        
    End If
    
    If Val(txt2.text) > 0 And Val(txt1.text) > 0 Then
      txt3.text = CDbl(txt2.text) / CDbl(txt1.text)
    Else
      txt3.text = Format(0, "0.00")
    End If
    txt3.text = Format(txt3.text, "###0.00")
End Sub

Private Sub txt2_LostFocus()

    If txt1.text <> "" Then
        If NumberCheck2(txt1.text) = False Then
            txt1.text = "0"
        End If
    Else
        txt1.text = "0"
    End If

    If txt2.text <> "" Then
        If NumberCheck2(txt2.text) = False Then
            txt2.text = "0"
        End If
    Else
        txt2.text = "0"

    End If

    If Val(txt2.text) > 0 And Val(txt1.text) > 0 Then
      txt3.text = CDbl(txt2.text) / CDbl(txt1.text)
    Else
      txt3.text = Format(0, "0.00")
    End If
    txt3.text = Format(txt3.text, "###0.00")
End Sub

Private Sub txt3_LostFocus()

    If txt3.text <> "" Then If NumberCheck2(txt3.text) = False Then txt3.text = ""

End Sub

Private Sub txt5_LostFocus()

    If txt5.text <> "" Then If NumberCheck2(txt5.text) = False Then txt5.text = ""

End Sub

Private Sub txt7_LostFocus()

    If txt7.text <> "" Then If NumberCheck(txt7.text) = False Then txt7.text = ""

End Sub

Public Sub FillDaysMonths()

    Dim i As Integer
    Dim months(1 To 12)
    
    For i = 1 To 9
        cboD1.AddItem "0" & i
        cboD2.AddItem "0" & i
        cboD3.AddItem "0" & i
        cboD4.AddItem "0" & i
        cboD5.AddItem "0" & i
        cboD6.AddItem "0" & i
        cboD7.AddItem "0" & i
        
        cboDay1.AddItem "0" & i
        cboDay2.AddItem "0" & i
        cboDay3.AddItem "0" & i
        cboDay4.AddItem "0" & i
        cboDay5.AddItem "0" & i
        cboDay6.AddItem "0" & i
        cboDay7.AddItem "0" & i
        cboDay8.AddItem "0" & i
        cboDay9.AddItem "0" & i
        cboDay10.AddItem "0" & i
        cboDay11.AddItem "0" & i
        cboDay12.AddItem "0" & i
    Next i
    
    For i = 10 To 31
        cboD1.AddItem i
        cboD2.AddItem i
        cboD3.AddItem i
        cboD4.AddItem i
        cboD5.AddItem i
        cboD6.AddItem i
        cboD7.AddItem i
        
        cboDay1.AddItem i
        cboDay2.AddItem i
        cboDay3.AddItem i
        cboDay4.AddItem i
        cboDay5.AddItem i
        cboDay6.AddItem i
        cboDay7.AddItem i
        cboDay8.AddItem i
        cboDay9.AddItem i
        cboDay10.AddItem i
        cboDay11.AddItem i
        cboDay12.AddItem i
    Next i
    
    months(1) = "January"
    months(2) = "February"
    months(3) = "March"
    months(4) = "April"
    months(5) = "May"
    months(6) = "June"
    months(7) = "July"
    months(8) = "August"
    months(9) = "September"
    months(10) = "October"
    months(11) = "November"
    months(12) = "December"
    
    For i = 1 To 12
        cboM1.AddItem months(i)
        cboM2.AddItem months(i)
        cboM3.AddItem months(i)
        cboM4.AddItem months(i)
        cboM5.AddItem months(i)
        cboM6.AddItem months(i)
        cboM7.AddItem months(i)
        
        cboMonth1.AddItem months(i)
        cboMonth2.AddItem months(i)
        cboMonth3.AddItem months(i)
        cboMonth4.AddItem months(i)
        cboMonth5.AddItem months(i)
        cboMonth6.AddItem months(i)
        cboMonth7.AddItem months(i)
        cboMonth8.AddItem months(i)
        cboMonth9.AddItem months(i)
        cboMonth10.AddItem months(i)
        cboMonth11.AddItem months(i)
        cboMonth12.AddItem months(i)
    Next i
    
End Sub

Public Function MissingDate(text As String) As Boolean

    MissingDate = False
    If text = "" Then
        MsgBox "You must select all the payment dates", vbOKOnly + vbCritical, "Missing Payment Date"
        MissingDate = True
    End If
    
End Function

Public Function ValidDate(text As String) As Boolean

    ValidDate = True
    If IsDate(text) = False Then
        MsgBox "Invalid Date Selected.", vbOKOnly + vbCritical, "Invalid Date"
        ValidDate = False
    End If

End Function

Public Sub GetVATRates()

    'Get the VAT rates from Sage and put in cboVatRate.
    'Set the RDO Env1ironment
    Set Envs = rdoEngine.rdoEnvironments
    Set Env = Envs(0)
      
    ' Use Line100's scrolling cursor
    Env.CursorDriver = rdUseServer
    
    ' Set the RDO Conn1ection to the dataset
    Set Conn = Env.OpenConnection("", rdDriverNoPrompt, False, "DSN=" & Adsn & ";UID=;PWD=")
    
    'SQLStr = "SELECT VAT_CODE, VAT_RATE, VAT_RATE_NAME FROM SYS_VAT_FILE ORDER BY VAT_CODE"   'CHANGE TO SageLine50v12
    SQLStr = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM tlbVATCODE ORDER BY VAT_ID"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
    
    While Rst.EOF = False
        cboVatRate.AddItem Rst!VAT_ID & " / " & Rst!VAT_CODE & " / " & Rst!VAT_RATE
        Rst.MoveNext
    Wend
    
    Rst.Close
    Conn.Close
    Env.Close

End Sub

Private Sub BankAccount()
   ' Error Handler
   On Error GoTo Error_Handler

   gridBankCode.Visible = True
   Dim clsBankAC As clsArray
   Dim iBankAc As Integer
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oBankRecord As SageDataObject120.BankRecord
   Dim oNominalRecord As SageDataObject120.NominalRecord

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
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then
   
      Set oBankRecord = oWS.CreateObject("BankRecord")
   
      ' Move to the First Record
      oBankRecord.MoveFirst
      Set clsBankAC = New clsArray
      For iBankAc = 1 To oBankRecord.Count
         clsBankAC.AddItem oBankRecord.Fields.Item("ACCOUNT_REF").Value
         oBankRecord.MoveNext
      Next iBankAc
   
      Set oBankRecord = Nothing
   
      Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
      oNominalRecord.MoveFirst
   
      Dim rRow As Integer
      Dim iRec As Integer
      rRow = 1
   
      gridBankCode.TextMatrix(0, 0) = "Reference"
      gridBankCode.TextMatrix(0, 1) = "Name"
      gridBankCode.ColWidth(0) = 1200
      gridBankCode.ColWidth(1) = 2600
   
      For iRec = 1 To oNominalRecord.Count
         If clsBankAC.IsItem(CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)) Then
            gridBankCode.TextMatrix(rRow, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
            gridBankCode.TextMatrix(rRow, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
            gridBankCode.AddItem ""
            rRow = rRow + 1
         End If
         oNominalRecord.MoveNext
      Next iRec
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text

   Set oBankRecord = Nothing
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub txtGlobalBankAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      gridBankCode.SetFocus
   End If
End Sub

Private Sub txt4_Change()
   TextBoxChangeDate txt4
End Sub

Private Sub txt4_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txt4, KeyAscii
End Sub

Private Sub txt4_LostFocus()
   TextBoxFormatDate txt4
End Sub
