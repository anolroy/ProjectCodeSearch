VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUnit 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kingsgate - Unit Maintenance"
   ClientHeight    =   8955
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   11640
   Icon            =   "Unit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   11640
   Begin VB.TextBox txt5 
      Height          =   495
      Left            =   2280
      TabIndex        =   99
      Text            =   "Text4"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox txt4 
      Height          =   375
      Left            =   1080
      TabIndex        =   98
      Text            =   "Text4"
      Top             =   8160
      Width           =   615
   End
   Begin TabDlg.SSTab tabUnit 
      Height          =   6495
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   16768960
      TabCaption(0)   =   "&Unit Details"
      TabPicture(0)   =   "Unit.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "cmdNext(0)"
      Tab(0).Control(2)=   "Frame2(1)"
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(4)=   "Frame5"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Tenancy Details"
      TabPicture(1)   =   "Unit.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPre(0)"
      Tab(1).Control(1)=   "cmdNext(1)"
      Tab(1).Control(2)=   "Frame1(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Service Charges"
      TabPicture(2)   =   "Unit.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "cmdPre(1)"
      Tab(2).Control(2)=   "cmdNext(2)"
      Tab(2).Control(3)=   "Frame1(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Maintenance &History"
      TabPicture(3)   =   "Unit.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdPre(2)"
      Tab(3).Control(1)=   "cmdNext(3)"
      Tab(3).Control(2)=   "Frame1(3)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&Account's History"
      TabPicture(4)   =   "Unit.frx":093A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame1(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdNext(4)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdPre(3)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Insurance && Safety"
      TabPicture(5)   =   "Unit.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraHSInfo"
      Tab(5).Control(1)=   "cmdPre(4)"
      Tab(5).Control(2)=   "cmdNext(5)"
      Tab(5).Control(3)=   "fraInsurance"
      Tab(5).Control(4)=   "Frame9"
      Tab(5).Control(5)=   "Frame7"
      Tab(5).Control(6)=   "Frame3"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "&Health && Safety"
      TabPicture(6)   =   "Unit.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdPre(5)"
      Tab(6).ControlCount=   1
      Begin VB.Frame fraHSInfo 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Health && Safety:"
         Enabled         =   0   'False
         Height          =   4575
         Left            =   -69000
         TabIndex        =   109
         ToolTipText     =   "Click Edit to change"
         Top             =   360
         Width           =   5055
         Begin VB.CheckBox chkFireEscape 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Fire Escape"
            Height          =   255
            Left            =   1920
            TabIndex        =   113
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox chkFireAlarm 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Fire Alarm"
            Height          =   255
            Left            =   1920
            TabIndex        =   112
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox chkFireExtinguisher 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Fire Extinguisher"
            Height          =   255
            Left            =   1920
            TabIndex        =   111
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkInspection 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Inspections"
            Height          =   255
            Left            =   1920
            TabIndex        =   110
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Imanges"
         Height          =   2055
         Left            =   -69480
         TabIndex        =   105
         Top             =   3840
         Width           =   5775
         Begin VB.CommandButton cmdSaveImage 
            Caption         =   "Upload Image"
            Height          =   375
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image imgImage 
            Height          =   1935
            Left            =   3840
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Client/Landlord:"
         Enabled         =   0   'False
         Height          =   2055
         Left            =   -72360
         TabIndex        =   100
         Top             =   3840
         Width           =   7215
         Begin VB.TextBox txtClient 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtProperty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            Height          =   195
            Left            =   360
            TabIndex        =   104
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Left            =   360
            TabIndex        =   103
            Top             =   720
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   5
         Left            =   -74880
         Picture         =   "Unit.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   4
         Left            =   -74880
         Picture         =   "Unit.frx":0DD0
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   3
         Left            =   120
         Picture         =   "Unit.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   2
         Left            =   -74880
         Picture         =   "Unit.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   1
         Left            =   -74880
         Picture         =   "Unit.frx":1A96
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   0
         Left            =   -74880
         Picture         =   "Unit.frx":1ED8
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   5
         Left            =   -64440
         Picture         =   "Unit.frx":231A
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   4
         Left            =   10560
         Picture         =   "Unit.frx":275C
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   3
         Left            =   -64440
         Picture         =   "Unit.frx":2B9E
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   2
         Left            =   -64440
         Picture         =   "Unit.frx":2FE0
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   1
         Left            =   -64440
         Picture         =   "Unit.frx":3422
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE5BF&
         Height          =   375
         Index           =   0
         Left            =   -64440
         Picture         =   "Unit.frx":3864
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   6000
         Width           =   735
      End
      Begin VB.Frame fraInsurance 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Existing Insurances"
         Enabled         =   0   'False
         Height          =   4335
         Left            =   -74640
         TabIndex        =   76
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtInsuNote 
            Height          =   1935
            Left            =   120
            TabIndex        =   78
            Top             =   2280
            Width           =   4695
         End
         Begin VB.ComboBox cmbInsurance 
            Height          =   315
            Left            =   2040
            TabIndex        =   77
            Top             =   360
            Width           =   2655
         End
         Begin MSComCtl2.MonthView dptDate 
            Height          =   2370
            Left            =   2040
            TabIndex        =   79
            Top             =   960
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   16768960
            Appearance      =   1
            StartOfWeek     =   56623106
            CurrentDate     =   38621
         End
         Begin VB.Label lblInsuranceDate 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   81
            ToolTipText     =   "Click to change date"
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nature of Insurance:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Renewal Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1920
            Width           =   390
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -69720
         TabIndex        =   73
         Top             =   5400
         Width           =   4815
         Begin VB.CommandButton cmdNewSave 
            Caption         =   "Sa&ve"
            Height          =   375
            Left            =   1920
            TabIndex        =   75
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdNewClear 
            Caption         =   "Clea&r"
            Height          =   375
            Left            =   3480
            TabIndex        =   74
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Add New Insurance:"
         Height          =   4335
         Left            =   -74640
         TabIndex        =   65
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtAddNewNote 
            Height          =   1935
            Left            =   120
            TabIndex        =   68
            Top             =   2280
            Width           =   4695
         End
         Begin VB.TextBox txtAddInsuNatureInsu 
            Height          =   315
            Left            =   2160
            TabIndex        =   67
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtAddInsuRenewDt 
            Height          =   315
            Left            =   2160
            TabIndex        =   66
            Top             =   960
            Width           =   2535
         End
         Begin MSComCtl2.MonthView dptAddInsuDt 
            Height          =   2370
            Left            =   2160
            TabIndex        =   69
            Top             =   960
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   16768960
            Appearance      =   1
            StartOfWeek     =   56623106
            CurrentDate     =   38621
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Renewal Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nature of Insurance:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   1920
            Width           =   390
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         TabIndex        =   63
         Top             =   5400
         Width           =   4935
         Begin VB.CommandButton cmdInsuAddNew 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   480
            TabIndex        =   115
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditInsu 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   2040
            TabIndex        =   114
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelInsu 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   3600
            TabIndex        =   64
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Rent Details"
         Enabled         =   0   'False
         Height          =   2055
         Index           =   1
         Left            =   -74880
         TabIndex        =   56
         Top             =   3480
         Width           =   5175
         Begin VB.TextBox txt1 
            DataField       =   "Frontage"
            DataSource      =   "MSRDC1"
            Height          =   285
            Left            =   2280
            TabIndex        =   59
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txt2 
            DataField       =   "RateableValue"
            DataSource      =   "MSRDC1"
            Height          =   285
            Left            =   2280
            TabIndex        =   58
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txt3 
            DataField       =   "RatesPayable"
            DataSource      =   "MSRDC1"
            Height          =   285
            Left            =   2280
            TabIndex        =   57
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbl1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Annual Rent:"
            Height          =   195
            Left            =   960
            TabIndex        =   62
            Top             =   600
            Width           =   930
         End
         Begin VB.Label lbl2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rateable Value:"
            Height          =   195
            Left            =   960
            TabIndex        =   61
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label lbl4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rates Payable:"
            Height          =   195
            Left            =   960
            TabIndex        =   60
            Top             =   1080
            Width           =   1080
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Landscape"
         Enabled         =   0   'False
         Height          =   3375
         Left            =   -69480
         TabIndex        =   52
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton cmdRemFloor 
            Caption         =   "Remove Floor"
            Height          =   375
            Left            =   2520
            TabIndex        =   97
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddFloor 
            Caption         =   "Add Floor"
            Height          =   375
            Left            =   4320
            TabIndex        =   96
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txt6 
            BackColor       =   &H00C0FFFF&
            DataSource      =   "MSRDC1"
            Height          =   285
            Left            =   4200
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2175
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   3836
            _Version        =   393216
            BackColor       =   -2147483624
            FixedCols       =   0
            BackColorFixed  =   14737632
            BackColorBkg    =   14737632
            BackColorUnpopulated=   11576751
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Area:"
            Height          =   195
            Left            =   3360
            TabIndex        =   55
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Unit Location"
         Enabled         =   0   'False
         Height          =   3015
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtUnitName 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   108
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtUnit 
            DataField       =   "UnitNumber"
            DataSource      =   "MSRDC1"
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   48
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtAddressLine3 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   47
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtAddressLine4 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   46
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox txtAddressLine2 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   45
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txtAddressLine1 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   44
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtPostCode 
            Height          =   285
            Left            =   1440
            MaxLength       =   12
            TabIndex        =   43
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Number:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   2640
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Service Charges:"
         Height          =   3375
         Index           =   2
         Left            =   -72360
         TabIndex        =   34
         Top             =   360
         Width           =   7215
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   4920
            TabIndex        =   37
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   4920
            TabIndex        =   36
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   4920
            TabIndex        =   35
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Specific Rate:"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fixed Value:"
            Height          =   195
            Left            =   3120
            TabIndex        =   40
            Top             =   1920
            Width           =   870
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apportionment Basis:"
            Height          =   195
            Left            =   3120
            TabIndex        =   39
            Top             =   2640
            Width           =   1485
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6240
            TabIndex        =   38
            Top             =   1200
            Width           =   210
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Accounts History:"
         Height          =   5535
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   11175
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccountHistory 
            Height          =   5055
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
         End
         Begin VB.TextBox txtUnitTemp 
            Height          =   375
            Left            =   4800
            TabIndex        =   32
            Top             =   3720
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Tenancy Details"
         Height          =   5415
         Index           =   1
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   11175
         Begin VB.TextBox cboOccupied 
            Height          =   285
            Left            =   5040
            TabIndex        =   23
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtRentReDt 
            Height          =   330
            Left            =   6015
            TabIndex        =   22
            Top             =   1680
            Width           =   2000
         End
         Begin VB.ComboBox cmbBllingFreq 
            Height          =   315
            ItemData        =   "Unit.frx":3CA6
            Left            =   1365
            List            =   "Unit.frx":3CBF
            TabIndex        =   21
            Top             =   1680
            Width           =   3050
         End
         Begin VB.TextBox txtExDate 
            Height          =   330
            Left            =   6015
            TabIndex        =   20
            Top             =   1080
            Width           =   2000
         End
         Begin VB.TextBox txtStDate 
            Height          =   330
            Left            =   1365
            TabIndex        =   19
            Top             =   1080
            Width           =   3025
         End
         Begin VB.ComboBox cmbTenancyType 
            Height          =   315
            Left            =   6015
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   480
            Width           =   2000
         End
         Begin VB.CommandButton cmdTenantDetails 
            Caption         =   ">"
            Height          =   315
            Left            =   4125
            TabIndex        =   17
            Top             =   480
            Width           =   285
         End
         Begin VB.ComboBox cboTenant 
            DataSource      =   "MSRDC1"
            Height          =   315
            Left            =   1365
            TabIndex        =   16
            Text            =   "cboTenant"
            Top             =   480
            Width           =   2685
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Occupied:"
            Height          =   195
            Left            =   3720
            TabIndex        =   30
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            Height          =   195
            Left            =   4560
            TabIndex        =   29
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Billing Frequency:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   1245
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date:"
            Height          =   195
            Left            =   4560
            TabIndex        =   27
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            Height          =   195
            Left            =   4560
            TabIndex        =   25
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Tenant:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Maintenance History:"
         Height          =   5535
         Index           =   3
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   11175
      End
   End
   Begin VB.ComboBox cboUnits 
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Text            =   "cboUnits"
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton cmdMaintanance 
      Caption         =   "Maintenance History"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00FFCFBF&
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New Unit"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Unit"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame fraEdit 
      BackColor       =   &H00FFCFBF&
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   7680
      TabIndex        =   9
      Top             =   7200
      Width           =   3855
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "Save C&hanges"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel Changes"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame fraAddNew 
      BackColor       =   &H00FFCFBF&
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   4800
      TabIndex        =   6
      Top             =   7920
      Width           =   3015
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel New Unit"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New Unit"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Select Unit"
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UnitNumber As String
Public OldUnit As String
Public NewUnit As String
Public OldTenantCode As String
Public NewTenantCode As String
Public OldTenantName As String
Public NewTenantName As String
Public oldtenant As String
Dim Conn1 As New RDO.rdoConnection
Dim Env1 As rdoEnvironment
Dim Envs1 As rdoEnvironments
Dim Rst1 As rdoResultset
Dim Conn2 As New RDO.rdoConnection
Dim Env2 As rdoEnvironment
Dim Envs2 As rdoEnvironments
Dim Rst2 As rdoResultset
Dim SQLStr1 As String
Dim SQLStr2 As String

Dim szUnitID As String

Dim szInsuNature As String
Dim szInsuRenewDt As String
Dim szInsuNote As String

Private mintCurFrame As Integer ' Current Frame visible

Private Sub cboOccupied_LostFocus()

If cboOccupied <> "Yes" And cboOccupied <> "No" Then
    MsgBox "Occupied statis is invalid. It must be set to Yes or No.", vbOKOnly + vbCritical, "Invalid Occupied Status"
    If cboTenant.text <> "" Then
        cboOccupied.text = "Yes"
    Else
        cboOccupied.text = "No"
    End If
End If

End Sub

Private Sub cboTenant_LostFocus()

Dim i, j, match As Integer
match = 0

If cboTenant.text <> "" Then
    j = cboTenant.ListCount - 1
    For i = 0 To j
        If cboTenant.List(i) = cboTenant.text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Tenant selected is invalid", vbOKOnly + vbCritical, "Invalid Tenant"
        cboTenant.text = ""
    End If
End If

End Sub

Private Sub cboUnits_Click()
   Dim i, j, match As Integer
   match = 0

   If cboUnits.text = "" Then
       MsgBox "You must select a Unit to view!", vbOKOnly + vbCritical, "No unit selected"
       Exit Sub
   End If

   j = cboUnits.ListCount - 1
   For i = 0 To j
       If cboUnits.List(i) = cboUnits.text Then
           match = 1
           Exit For
       End If
   Next i
   If match = 0 Then
       MsgBox "Unit selected is invalid", vbOKOnly, vbCritical, "Invalid unit"
       cboUnits.text = ""
       Exit Sub
   End If

   Dim szUnit() As String
   szUnit = Split(cboUnits.text, " \ ")
   szUnitID = Trim(szUnit(0))
   
   'Set the RDO Connections to the dataset
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   'Get the record for selected unit
   SQLStr1 = "SELECT * FROM Units WHERE UnitNumber = '" & szUnitID & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   'Fill boxes with unit details.
   txtUnit.text = Rst1!UnitNumber
   If IsNull(Rst1!Frontage) Then txt1.text = "0" Else txt1.text = Rst1!Frontage
   If IsNull(Rst1!RateableValue) Then txt2.text = "0" Else txt2.text = Rst1!RateableValue
   If IsNull(Rst1!RatesPayable) Then txt3.text = "0" Else txt3.text = Rst1!RatesPayable
   If IsNull(Rst1!GroundFloorArea) Then txt4.text = "0" Else txt4.text = Rst1!GroundFloorArea
   If IsNull(Rst1!MezzanineArea) Then txt5.text = "0" Else txt5.text = Rst1!MezzanineArea
   If IsNull(Rst1!TotalArea) Then txt6.text = "0" Else txt6.text = Rst1!TotalArea
   If IsNull(Rst1!UnitAddressLine1) Then txtAddressLine1.text = "" Else txtAddressLine1.text = Rst1!UnitAddressLine1
   If IsNull(Rst1!UnitAddressLine2) Then txtAddressLine2.text = "" Else txtAddressLine2.text = Rst1!UnitAddressLine2
   If IsNull(Rst1!UnitAddressLine3) Then txtAddressLine3.text = "" Else txtAddressLine3.text = Rst1!UnitAddressLine3
   If IsNull(Rst1!UnitAddressLine4) Then txtAddressLine4.text = "" Else txtAddressLine4.text = Rst1!UnitAddressLine4
   If IsNull(Rst1!UnitPostCode) Then txtPostCode.text = "" Else txtPostCode.text = Rst1!UnitPostCode

   szUnit(0) = UnitLandLord(szUnitID)
   If szUnit(0) <> "ERROR" Then txtClient.text = szUnit(0)
   
   If Rst1!OCCUPIED = "Y" Then
       cboOccupied.text = "Yes"
       cboTenant.text = Rst1!SageAccountNumber & " / " & Rst1!TenantCompanyName
   Else
       cboOccupied.text = "No"
       cboTenant.text = ""
   End If
   
   Rst1.Close
   Conn1.Close
   
   OldUnit = txtUnit.text
   Call DisableBoxes

   Call AccountHistory
End Sub

Private Sub AccountHistory()
   Dim rstCN As rdoResultset, rstPI As rdoResultset, rstBP As rdoResultset
   Dim szCN As String, szPI As String, szBP As String

   'Set the RDO Connections to the dataset
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt

   'Get the record for selected unit
   szCN = "SELECT tlbCreditNote.TRAN_ID, tlbCreditNote.SUPP_AC, " & _
               "tlbCreditNote.TRAN_DATE, tlbCreditNote.TRAN_TYPE, " & _
               "tlbCreditNote.INV_NO, tlbCreditNote.DESCRIPTION, " & _
               "tlbCreditNote.NET_AMOUNT, tlbCreditNote.VAT, tlbCreditNote.PAYMENT, " & _
               "tlbCreditNote.OUT_PAYMENT, tlbCreditNote.DR_CR " & _
          "FROM tlbCreditNote tlbCreditNote " & _
          "WHERE (tlbCreditNote.TRANS='UNIT') AND " & _
               "tlbCreditNote.UNIT_ID='" & szUnitID & "';"
   
   Set rstCN = Conn1.OpenResultset(szCN, rdOpenStatic, rdConcurReadOnly)

   szPI = "SELECT tlbPurchaseInvoice.TRAN_ID, tlbPurchaseInvoice.SUPP_AC, " & _
               "tlbPurchaseInvoice.TRAN_DATE, tlbPurchaseInvoice.TRAN_TYPE, " & _
               "tlbPurchaseInvoice.INV_NO, tlbPurchaseInvoice.DESCRIPTION, " & _
               "tlbPurchaseInvoice.NET_AMOUNT, tlbPurchaseInvoice.VAT, " & _
               "tlbPurchaseInvoice.PAYMENT, " & _
               "tlbPurchaseInvoice.OUT_PAYMENT, tlbPurchaseInvoice.DR_CR " & _
          "FROM tlbPurchaseInvoice tlbPurchaseInvoice " & _
          "WHERE (tlbPurchaseInvoice.TRANS='UNIT') AND " & _
               "tlbPurchaseInvoice.UNIT_ID='" & szUnitID & "';"

   Set rstPI = Conn1.OpenResultset(szPI, rdOpenStatic, rdConcurReadOnly)

   szBP = "SELECT tlbBankPayment.TRAN_ID, tlbBankPayment.BANK_AC, " & _
               "tlbBankPayment.TRAN_DATE, tlbBankPayment.TRAN_TYPE, " & _
               "tlbBankPayment.DESCRIPTION, " & _
               "tlbBankPayment.NET_AMOUNT, tlbBankPayment.VAT " & _
          "FROM tlbBankPayment " & _
          "WHERE (tlbBankPayment.TRANS='UNIT') AND " & _
               "tlbBankPayment.UNIT_ID='" & szUnitID & "';"
   
   Set rstBP = Conn1.OpenResultset(szBP, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer

   If flxAccountHistory.Rows > 2 Then
      For iRow = 2 To flxAccountHistory.Rows - 1
         flxAccountHistory.RemoveItem 2
      Next iRow
      For iRow = 0 To 8
         flxAccountHistory.TextMatrix(1, iRow) = ""
      Next iRow
   End If
   
   iRow = 1
   While Not rstPI.EOF
      flxAccountHistory.TextMatrix(iRow, 0) = rstPI!TRAN_ID
      flxAccountHistory.TextMatrix(iRow, 1) = rstPI!SUPP_AC
      flxAccountHistory.TextMatrix(iRow, 2) = rstPI!TRAN_DATE
      flxAccountHistory.TextMatrix(iRow, 3) = rstPI!INV_NO
      flxAccountHistory.TextMatrix(iRow, 4) = rstPI!description
      flxAccountHistory.TextMatrix(iRow, 5) = rstPI!TRAN_TYPE
      flxAccountHistory.TextMatrix(iRow, 6) = rstPI!NET_AMOUNT
      flxAccountHistory.TextMatrix(iRow, 7) = rstPI!VAT
      flxAccountHistory.TextMatrix(iRow, 8) = rstPI!OUT_PAYMENT

      rstPI.MoveNext
      If Not rstPI.EOF Then flxAccountHistory.AddItem ""
      flxAccountHistory.RowHeight(flxAccountHistory.Rows - 1) = 285
      iRow = iRow + 1
   Wend
   rstPI.Close
   Set rstPI = Nothing
   
   While Not rstCN.EOF
      If Not rstCN.EOF Then flxAccountHistory.AddItem ""
      flxAccountHistory.RowHeight(flxAccountHistory.Rows - 1) = 285
      flxAccountHistory.TextMatrix(iRow, 0) = rstCN!TRAN_ID
      flxAccountHistory.TextMatrix(iRow, 1) = rstCN!SUPP_AC
      flxAccountHistory.TextMatrix(iRow, 2) = rstCN!TRAN_DATE
      flxAccountHistory.TextMatrix(iRow, 3) = rstCN!INV_NO
      flxAccountHistory.TextMatrix(iRow, 4) = rstCN!description
      flxAccountHistory.TextMatrix(iRow, 5) = rstCN!TRAN_TYPE
      flxAccountHistory.TextMatrix(iRow, 6) = rstCN!NET_AMOUNT
      flxAccountHistory.TextMatrix(iRow, 7) = rstCN!VAT
      flxAccountHistory.TextMatrix(iRow, 8) = rstCN!OUT_PAYMENT

      rstCN.MoveNext
      iRow = iRow + 1
   Wend
   
   rstCN.Close
   Set rstCN = Nothing
   
   While Not rstBP.EOF
      If Not rstBP.EOF Then flxAccountHistory.AddItem ""
      flxAccountHistory.RowHeight(flxAccountHistory.Rows - 1) = 285
      flxAccountHistory.TextMatrix(iRow, 0) = rstBP!TRAN_ID
      flxAccountHistory.TextMatrix(iRow, 1) = rstBP!BANK_AC
      flxAccountHistory.TextMatrix(iRow, 2) = rstBP!TRAN_DATE
      flxAccountHistory.TextMatrix(iRow, 4) = rstBP!description
      flxAccountHistory.TextMatrix(iRow, 5) = rstBP!TRAN_TYPE
      flxAccountHistory.TextMatrix(iRow, 6) = rstBP!NET_AMOUNT
      flxAccountHistory.TextMatrix(iRow, 7) = rstBP!VAT

      rstBP.MoveNext
      iRow = iRow + 1
   Wend
   
   rstBP.Close
   Set rstBP = Nothing
   Conn1.Close
   Set Conn1 = Nothing
   
   flxAccountHistory.Sort = flexSortGenericAscending
   flxAccountHistory.ColAlignment(0) = vbAlignLeft
   flxAccountHistory.ColAlignment(1) = vbLeftJustify
   flxAccountHistory.ColAlignment(4) = vbLeftJustify
   flxAccountHistory.ColAlignment(5) = vbAlignLeft
End Sub

Private Sub cmdAccounts_Click()

End Sub

Private Sub cmdAdd_Click()
    'Call AddNew
    Load frmUnits2
    frmUnits2.Show
End Sub

Private Sub cmdCancelEdit_Click()

OldUnit = szUnitID

'Set the RDO Connection to the db.
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseIfNeeded
Conn1.EstablishConnection rdDriverNoPrompt

'select record for selected unit
SQLStr1 = "SELECT * FROM Units WHERE UnitNumber = '" & OldUnit & "'"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

'fill boxes with details
txtUnit = OldUnit
If IsNull(Rst1!Frontage) Then txt1.text = "0" Else txt1.text = Rst1!Frontage
If IsNull(Rst1!RateableValue) Then txt2.text = "0" Else txt2.text = Rst1!RateableValue
If IsNull(Rst1!RatesPayable) Then txt3.text = "0" Else txt3.text = Rst1!RatesPayable
If IsNull(Rst1!GroundFloorArea) Then txt4.text = "0" Else txt4.text = Rst1!GroundFloorArea
If IsNull(Rst1!MezzanineArea) Then txt5.text = "0" Else txt5.text = Rst1!MezzanineArea
If IsNull(Rst1!TotalArea) Then txt6.text = "0" Else txt6.text = Rst1!TotalArea
If Rst1!OCCUPIED = "Y" Then
    cboOccupied.text = "Yes"
    cboTenant.text = Rst1!SageAccountNumber & " / " & Rst1!TenantCompanyName
Else
    cboOccupied.text = "No"
End If

Rst1.Close
Conn1.Close

Call CancelAddEdit
Call DisableBoxes

End Sub

Private Sub cmdCancelInsu_Click()
   fraInsurance.Enabled = False
   cmdEditInsu.Caption = "&Edit"
   cmbInsurance.text = szInsuNature
   lblInsuranceDate.Caption = szInsuRenewDt
   txtInsuNote.text = szInsuNote
End Sub

Private Sub cmdCancelNew_Click()
    Call CancelAddEdit
    Call ResetScreen
    Call EmptyBoxes
    Call DisableBoxes
End Sub

Private Sub cmdDelete_Click()

Call Delete

End Sub

Private Sub cmdEdit_Click()
    Call Edit
End Sub

Private Sub cmdEditInsu_Click()
   If cmdEditInsu.Caption = "&Edit" Then
      fraInsurance.Enabled = True
      cmdEditInsu.Caption = "&Update"
      szInsuNature = cmbInsurance.text
      szInsuRenewDt = lblInsuranceDate.Caption
      szInsuNote = txtInsuNote.text
   Else
      MsgBox "Under Construction"
   End If
End Sub

Private Sub cmdHSEdit_Click()
   fraHSInfo.Enabled = True
'   cmdHSEdit.Enabled = False
'   cmdHSSave.Enabled = True
End Sub

Private Sub cmdHSSave_Click()
   MsgBox "Under Construction"
'   cmdHSEdit.Enabled = True
'   cmdHSSave.Enabled = False
End Sub

Private Sub cmdMaintanance_Click()
    If cboUnits.text = "" Then
        MsgBox "Please Select Unit ID"
        Exit Sub
    End If

    Load frmMaintHistory

    frmMaintHistory.lblUnitID.Caption = cboUnits.text
    Unload Me
    frmMaintHistory.Show
End Sub

Private Sub cmdNewClear_Click()
   txtAddInsuNatureInsu.text = ""
   txtAddInsuRenewDt.text = ""
   txtAddNewNote.text = ""
End Sub

Private Sub cmdNewSave_Click()
   MsgBox "Under Construction"
End Sub

Private Sub cmdNext_Click(Index As Integer)
   tabUnit.Tab = Index + 1
End Sub

Private Sub cmdPre_Click(Index As Integer)
   tabUnit.Tab = Index
End Sub

Private Sub cmdSaveEdit_Click()
    Dim Response
    Dim match As Integer
    Dim i As Integer
    Dim j As Integer

    match = 0

    MousePointer = vbHourglass

    OldUnit = szUnitID
    NewUnit = txtUnit.text

    If OldUnit <> NewUnit Then 'unit number has been changed
        'check that another unit does not exist with new unit number.
        Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
        Conn1.CursorDriver = rdUseIfNeeded
        Conn1.EstablishConnection rdDriverNoPrompt
       
        SQLStr1 = "SELECT UnitNumber FROM Units"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

        If Rst1.EOF = False Then
            While Rst1.EOF = False
                If Rst1!UnitNumber = NewUnit Then match = 1
                Rst1.MoveNext
            Wend
        End If
        Rst1.Close
        Conn1.Close

        If match = 1 Then
            Response = MsgBox("Unit with unit number: " & NewUnit & " already exists.  Save changes with existing unit number?", vbYesNo + vbQuestion, "New unit number already exists")
            match = 0
            txtUnit.text = OldUnit
            If Response = vbNo Then
                Exit Sub
            End If
        End If
    End If

    match = 0

    If cboOccupied.text = "" Then
        MsgBox "You must select Occupied to be Yes or No", vbOKOnly + vbCritical, "Invalid Occupied Status"
        If cboTenant.text <> "" Then
            cboOccupied.text = "Yes"
        Else
            cboOccupied.text = "No"
        End If
        Exit Sub
    End If
    If cboOccupied.text = "Yes" Then
        If cboTenant.text = "" Then
            MsgBox "Occupied is set to Yes, but no tenant is selected!", vbOKOnly + vbCritical, "Can not save changes"
            Exit Sub
        Else
            For i = 2 To 9
                If Mid(cboTenant.text, i, 3) = " / " Then
                    NewTenantCode = Left(cboTenant.text, i - 1)
                    NewTenantName = Mid(cboTenant.text, i + 3, 93 - i)
                End If
            Next i
            RTrim (NewTenantName)
        End If
    Else
        If cboTenant.text <> "" Then
            MsgBox "Occupied is set to No, but a tenant is selected!", vbOKOnly + vbCritical, "Can not save changes"
            Exit Sub
        End If
    End If
    
    'Update Unit table
    'Set the RDO Connections to the dataset
    Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn2.CursorDriver = rdUseIfNeeded
    Conn2.EstablishConnection rdDriverNoPrompt
    
    'Get the record for selected unit
    SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & OldUnit & "'"
    Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
    
    Rst2.Edit
    
    Rst2!UnitNumber = NewUnit
    If cboOccupied.text = "Yes" Then
        Rst2!OCCUPIED = "Y"
        Rst2!TenantCompanyName = NewTenantName
        Rst2!SageAccountNumber = NewTenantCode
    End If
    
    If cboOccupied.text = "No" Then Rst2!OCCUPIED = "N"
    Rst2!Frontage = CLng(txt1.text)
    Rst2!RateableValue = CLng(txt2.text)
    Rst2!RatesPayable = CLng(txt3.text)
    Rst2!GroundFloorArea = CLng(txt4.text)
    Rst2!MezzanineArea = CLng(txt5.text)
    Rst2!TotalArea = CLng(txt6.text)
    
    Rst2!UnitAddressLine1 = txtAddressLine1.text
    Rst2!UnitAddressLine2 = txtAddressLine2.text
    Rst2!UnitAddressLine3 = txtAddressLine3.text
    Rst2!UnitAddressLine4 = txtAddressLine4.text
    Rst2!UnitPostCode = txtPostCode.text

    Rst2.Update
    
    Rst2.Close
    Conn2.Close
    
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt
    
    'Update the tenant table if required.
    If OldTenantCode <> NewTenantCode Then
        If OldTenantCode = "" Then
            'change new tenant record to include current rental as unit number.
            SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & NewTenantCode & "'"
            Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
            
            Rst1.Edit
            Rst1!CurrentRental = txtUnit.text
            Rst1.Update
            Rst1.Close
        End If
        
        If NewTenantCode = "" Then
            'change old tenant record to take out current rental.
            SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & OldTenantCode & "'"
            Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
            
            Rst1.Edit
            'Line below was Rst1!CurrentRental = ""
            Rst1!CurrentRental = "Prev" & OldUnit
            Rst1.Update
            Rst1.Close
        End If
        
        If OldTenantCode <> "" And NewTenantCode <> "" Then
            'change both tenant records.
            SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & OldTenantCode & "'"
            Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
            Rst1.Edit
            Rst1!CurrentRental = ""
            Rst1.Update
            Rst1.Close
            
            SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & NewTenantCode & "'"
            Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
            Rst1.Edit
            Rst1!CurrentRental = txtUnit.text
            Rst1.Update
            Rst1.Close
        End If
    End If
        
    If NewUnit <> OldUnit Then
        'change unit on tenant record.
        SQLStr1 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & NewTenantCode & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        If Rst1.EOF = False Then
            Rst1.Edit
            Rst1!CurrentRental = NewUnit
            Rst1.Update
        End If
        Rst1.Close
        'change lease details
        SQLStr1 = "SELECT UnitNumber FROM LeaseDetails WHERE SageAccountNumber = '" & NewTenantCode & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        If Rst1.EOF = False Then
            Rst1.Edit
            Rst1!UnitNumber = NewUnit
            Rst1.Update
        End If
        Rst1.Close
        'change on any demands in the demand table that have not been exported to Sage
        SQLStr1 = "SELECT UnitNumber FROM DemandRecords WHERE ExportedToSage <> 'Y'AND UnitNumber = '" & OldUnit & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        If Rst1.EOF = False Then
            While Rst1.EOF = False
                Rst1.Edit
                Rst1!UnitNumber = NewUnit
                Rst1.Update
                Rst1.MoveNext
            Wend
        End If
        Rst1.Close
    End If
    Conn1.Close
    cboUnits.text = NewUnit

    Call EmptyVars

    MsgBox "Your changes have been saved.", vbOKOnly + vbInformation, "Saved"

    fraMain.ZOrder 0
    fraEdit.ZOrder 1

    cmdAdd.Visible = True
'    cmdDelete.Visible = True
    cmdSaveNew.Visible = False
    cmdCancelNew.Visible = False
    cmdEdit.Visible = True
    cmdSaveEdit.Visible = False
    cmdCancelEdit.Visible = False
    
    Call ResetScreen
    Call DisableBoxes
    
    MousePointer = vbDefault

End Sub

Private Sub cmdSaveImage_Click()
   Dim szImageFileName As String
   szImageFileName = DoBrowse()
   If szImageFileName <> "NONE" Then
      If DoStoreInDB(szImageFileName) Then
         MsgBox "SUCCESSFUL"
      End If
   Else
      Exit Sub
   End If
   imgImage.Picture = LoadPicture(szImageFileName)
End Sub

Private Function DoStoreInDB(sTemp As String) As Boolean
   Dim mDB As Database             'open once, close upon unload
   Dim Rst As Recordset
   Dim szStr As String, msDBNameFull As String
   Dim fldLongBinary As Field
   Dim fldFileName As Field
   Dim obj As New SaveCreateFile.cStoreCreateFile
   
   msDBNameFull = szPictureDBPath
   Set mDB = OpenDatabase(msDBNameFull)
   szStr = "SELECT * FROM TLBIMAGES;"
   Set Rst = mDB.OpenRecordset(szStr)  'make sure there is at least one record and store the file name

    With Rst
      .AddNew         'add a one 'dummy' record if none
'      .Fields("MY_ID") = "xx"
      .Fields("PREMISIS_ID") = txtUnit.text
      .Fields("PREMISIS_TYPE") = "UNIT"
      .Fields("IMAGE_NAME") = txtUnit.text
      .Update         'update the table
    End With
    szStr = "SELECT * " & _
            "FROM TLBIMAGES " & _
            "WHERE IMAGE_NAME='" & txtUnit.text & "' AND " & _
            "TLBIMAGES.PREMISIS_ID='" & txtUnit.text & "';"
    Set Rst = mDB.OpenRecordset(szStr)
    With Rst
        If Not .EOF Then
            .MoveLast
            Set fldFileName = .Fields![IMAGE_PATH]   'set reference to this field
            Set fldLongBinary = .Fields![PREMISIS_IMAGE]  'set a reference to the file field
            With obj                        'now call the dll
                If .StoreFileIntoField(Rst, fldFileName, fldLongBinary, sTemp) Then  'call the dll
                    DoStoreInDB = True
                End If
            End With
        End If
        .Close
    End With
    
    Set Rst = Nothing
    Set fldFileName = Nothing
    Set fldLongBinary = Nothing
    Set obj = Nothing


' Store the file in a database field.
'   Dim Conn As New Adodc

'   Dim fldLongBinary As Field
'   Dim fldFileName   As Field
'   Dim obj As New SaveCreateFile.cStoreCreateFile
'
'   Conn.ConnectionString = getConnectionString
'   Conn.RecordSource = "SELECT * FROM TLBIMAGES;"
'   Conn.CommandType = adCmdText
'   Conn.Refresh
'
'   Set Rst = Conn.Recordset
'
'   Rst.AddNew
'   Rst.Update
'
'   szStr = "SELECT * FROM TLBIMAGES;"
'   Conn.RecordSource = szStr
'   Conn.CommandType = adCmdText
'   Conn.Refresh
'   Set Rst = Conn.Recordset
'
'   With Rst
'       If Not .EOF Then
''           .MoveLast
'           Set fldFileName = ![IMAGE_NAME]   'set reference to this field
'           Set fldLongBinary = !PREMISIS_IMAGE  'set a reference to the file field
'           With obj                        'now call the dll
'               If .StoreFileIntoField(Rst, fldFileName, fldLongBinary, sTemp) Then  'call the dll
'                   DoStoreInDB = True
'               End If
'           End With
'       End If
'       .Close
'   End With
'   Set Rst = Nothing
'   Set fldFileName = Nothing
'   Set fldLongBinary = Nothing
'   Set obj = Nothing
End Function

Private Function DoBrowse() As String
' Browse for a file
    Dim obj As New CDialog
    With obj                'find a graphics file
        .Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNFileMustExist + cdlOFNExplorer
        .Filter = "Graphics File|*.bmp;*.ico;*.emf;*.wmf;*.jpg;*.gif|All Files|*.*"
        .DialogTitle = "Select a File"
        .InitDir = Trim$(App.Path)
        .lHwnd = Me.hWnd
        .ShowOpen
        If Not .Cancelled Then
            DoBrowse = .Filename
        Else
            DoBrowse = "NONE"
        End If
    End With
    Set obj = Nothing
End Function

Private Sub cmdSaveNew_Click()
    Dim match
    Dim i
    
    'check that a unit number has been entered.
    If txtUnit.text = "" Then
            MsgBox "You must enter a unit number!", vbOKOnly + vbCritical, "Unit number required"
            Exit Sub
        Else
    
        If cboOccupied.text = "" Then
            MsgBox "You must select Yes/No"
            cboOccupied.SetFocus
            Exit Sub
        End If
        'check unit number is not already in unit table
        Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
        Conn1.CursorDriver = rdUseIfNeeded
        Conn1.EstablishConnection rdDriverNoPrompt
       
        SQLStr1 = "SELECT UnitNumber FROM Units"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
       
        If Rst1.EOF = False Then
            While Rst1.EOF = False
                If Rst1!UnitNumber = txtUnit.text Then match = 1
                Rst1.MoveNext
            Wend
        End If
        
        Rst1.Close
        Conn1.Close
         
        'If unit number does exist already tell user and empty txtunit.
        If match = 1 Then
            MsgBox "Can not save new unit details.  Unit number: " & txtUnit.text & " already exists. ", vbOKOnly + vbCritical, "Can not save"
            Exit Sub
        Else
            NewUnit = txtUnit.text
        End If
    End If
    
    'get tenant info
    If cboOccupied.text = "Yes" Then
        If cboTenant.text = "" Then
            MsgBox "Occupied is set to Yes, but no tenant is selected!", vbOKOnly + vbCritical, "Can not save changes"
            Exit Sub
        Else
            For i = 2 To 9
                If Mid(cboTenant.text, i, 3) = " / " Then
                    NewTenantCode = Left(cboTenant.text, i - 1)
                    NewTenantName = Mid(cboTenant.text, i + 3, 93 - i)
                End If
            Next i
            NewTenantName = RTrim(NewTenantName)
        End If
    Else
        If cboTenant.text <> "" Then
            MsgBox "Occupied is set to No, but a tenant is selected!", vbOKOnly + vbCritical, "Can not save changes"
            Exit Sub
        End If
    End If
    
    'If no values entered set them to zeros
    If txt1.text = "" Then txt1.text = "0"
    If txt2.text = "" Then txt2.text = "0"
    If txt3.text = "" Then txt3.text = "0"
    If txt4.text = "" Then txt4.text = "0"
    If txt5.text = "" Then txt5.text = "0"
    txt6.text = CStr(CLng(txt4.text) + CLng(txt5.text))
    
    'add new unit to unit table
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt
    
    SQLStr1 = "SELECT * FROM Units"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
    
    Rst1.AddNew
    Rst1!UnitNumber = NewUnit
    Select Case cboOccupied.text
        Case "Yes"
            Rst1!OCCUPIED = "Y"
            Rst1!TenantCompanyName = NewTenantName
            Rst1!SageAccountNumber = NewTenantCode
        Case "No"
            Rst1!OCCUPIED = "N"
    End Select
    Rst1!Frontage = CLng(txt1.text)
    Rst1!RateableValue = CLng(txt2.text)
    Rst1!RatesPayable = CLng(txt3.text)
    Rst1!GroundFloorArea = CLng(txt4.text)
    Rst1!MezzanineArea = CLng(txt5.text)
    Rst1!TotalArea = CLng(txt6.text)
    
    Rst1!UnitAddressLine1 = txtAddressLine1.text
    Rst1!UnitAddressLine2 = txtAddressLine2.text
    Rst1!UnitAddressLine3 = txtAddressLine3.text
    Rst1!UnitAddressLine4 = txtAddressLine4.text
    Rst1!UnitPostCode = txtPostCode.text
    Rst1.Update
    
    Rst1.Close
    Conn1.Close
    
    'If there is a tenant then update the tenant record.
    If cboOccupied.text = "Yes" Then
        
        Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
        Conn2.CursorDriver = rdUseIfNeeded
        Conn2.EstablishConnection rdDriverNoPrompt
            
        SQLStr2 = "SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & NewTenantCode & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
        
        Rst2.Edit
        Rst2!CurrentRental = NewUnit
        Rst2.Update
        Rst2.Close
        Conn2.Close
    
    End If
    
    Call EmptyVars
    
    MsgBox "New unit details have been saved.", vbOKOnly + vbInformation, "Saved"
    
    fraMain.ZOrder 0
    fraAddNew.ZOrder 1
    
    cmdAdd.Visible = True
'    cmdDelete.Visible = True
    cmdSaveNew.Visible = False
    cmdCancelNew.Visible = False
    cmdEdit.Visible = True
    cmdSaveEdit.Visible = False
    cmdCancelEdit.Visible = False
    
    Call DisableBoxes
    Call ResetScreen
End Sub

Private Sub cmdTenantDetails_Click()
    frmTenantNew.LodeTenant = cboTenant.text
    Load frmTenantNew
    frmTenantNew.Show
'    frmTenantNew.cboTenants_Click
    
    
End Sub

Private Sub Command1_Click()
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub dptAddInsuDt_DateClick(ByVal DateClicked As Date)
   txtAddInsuRenewDt.text = dptAddInsuDt.Value
   dptAddInsuDt.Visible = False
End Sub

Private Sub dptDate_DateClick(ByVal DateClicked As Date)
   lblInsuranceDate.Caption = Format(dptDate.Value, "dd mmmm, yyyy")
   dptDate.Visible = False
End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50

    On Error GoTo ErrH1
    Me.Caption = gCurrentShopCentreName & " - Units"
    
    cboUnits.text = ""
    cboOccupied.text = ""
    cboTenant.text = ""

    Call ResetScreen
    Call DisableBoxes

    tabUnit.Tab = 0
    
    Call FlexGridAccountHistoryConfigure

    Exit Sub
ErrH1:
        If ERR.Number = 40002 Then
            If MsgBox("Please check DSN - " & Adsn & " is set up correctly.", vbRetryCancel, "DSN Error") = vbRetry Then
                Resume
            Else
                Resume Next
            End If
        ElseIf ERR.Number <> 0 Then
            MsgBox ERR.Number & " - " & ERR.description
            Resume Next
        End If
End Sub

Private Sub FlexGridAccountHistoryConfigure()
   flxAccountHistory.ColWidth(0) = 700
   flxAccountHistory.TextMatrix(0, 0) = "No."
   
   flxAccountHistory.ColWidth(1) = 1400
   flxAccountHistory.TextMatrix(0, 1) = "A/C"
   
   flxAccountHistory.ColWidth(2) = 1000
   flxAccountHistory.TextMatrix(0, 2) = "Trans Dt"
   
   flxAccountHistory.ColWidth(3) = 1100
   flxAccountHistory.TextMatrix(0, 3) = "Inv No."
   
   flxAccountHistory.ColWidth(4) = 2000
   flxAccountHistory.TextMatrix(0, 4) = "Description"
   
   flxAccountHistory.ColWidth(5) = 1100
   flxAccountHistory.TextMatrix(0, 5) = "Type"
   
   flxAccountHistory.ColWidth(6) = 1100
   flxAccountHistory.TextMatrix(0, 6) = "Amount"
   
   flxAccountHistory.ColWidth(7) = 1100
   flxAccountHistory.TextMatrix(0, 7) = "Vat"
   
   flxAccountHistory.ColWidth(8) = 1100
   flxAccountHistory.TextMatrix(0, 8) = "Outst. Pay"

   flxAccountHistory.RowHeight(1) = 285
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub mnuAddNew_Click()

Call AddNew

End Sub

Private Sub mnuDel_Click()

Call Delete

End Sub

Private Sub mnuDemands_Click()

Unload Me
Load frmDemands

End Sub

Private Sub mnuEdit_Click()

Call Edit

End Sub

Private Sub mnuExit_Click()
   Unload frmMMain
End Sub

Public Sub ResetScreen()
   Dim temp As String
   
   cboUnits.Enabled = True
   
   temp = cboUnits.text
   cboUnits.Clear
   cboUnits.text = temp
   
   'Reset screen to show all the units in cboUnits.
   'Set the RDO Connections to the dataset
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
   
   SQLStr2 = "SELECT * FROM Units ORDER BY UnitNumber"
   Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)

   If Rst2.EOF = False Then
       While Rst2.EOF = False
           cboUnits.AddItem Rst2!UnitNumber & " \ " & Rst2!TenantCompanyName
           Rst2.MoveNext
       Wend
   End If

   Rst2.Close
   Conn2.Close

   cmdAdd.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdEdit.Visible = True
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
End Sub

Public Sub DisableBoxes()
'    Frame1(0).Enabled = False
End Sub

Public Sub EnableBoxes()
'    Frame1(0).Enabled = True
End Sub

Public Sub Edit()
    Dim i
    
    If cboUnits.text = "" Then
        MsgBox "You must select a unit to edit!", vbOKOnly + vbCritical, "No unit selected"
        Exit Sub
    End If
    oldtenant = cboTenant.text
    For i = 2 To 9
        If Mid(cboTenant.text, i, 3) = " / " Then
            OldTenantCode = Left(cboTenant.text, i - 1)
            OldTenantName = Mid(cboTenant.text, i + 3, 93 - i)
        End If
    Next i
    RTrim (OldTenantName)
    Call EnableBoxes

    fraMain.ZOrder 1
    fraEdit.ZOrder 0
    
    cmdSaveEdit.Visible = True
    cmdCancelEdit.Visible = True

    Call GetTenants
End Sub

Public Sub GetTenants()

    Dim temp
    
    If cboTenant.text = "" Then
        cboTenant.Clear
    Else
        temp = cboTenant.text
        cboTenant.Clear
    End If
    
    'Get all the tenants that do not have a current rental and put in cboTenant.
    'Set the RDO environment
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt
    
    SQLStr1 = "SELECT CompanyName, SageAccountNumber FROM Tenants WHERE CurrentRental = '' ORDER BY CompanyName"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            cboTenant.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
            Rst1.MoveNext
        Wend
    End If
    
    Rst1.Close
    Conn1.Close
    
    If temp <> "" Then
        cboTenant.AddItem temp, 0
        cboTenant.text = temp
    End If
    
    cboTenant.AddItem "", 0

End Sub

Public Sub Delete()
   Dim Response

   If cboUnits.text = "" Then
       MsgBox "You must select a unit to delete!", vbOKOnly + vbCritical, "No unit selected"
       Exit Sub
   Else
       If cboOccupied.text = "Yes" Then
           MsgBox "Can not delete unit - unit is occupied!", vbOKOnly + vbCritical, "Delete Unit"
           Exit Sub
       End If

       Response = MsgBox("Are you sure you want to delete unit: " & cboUnits.text & "?", vbYesNo + vbQuestion, "Delete unit")
       If Response = vbYes Then
           Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
           Conn1.CursorDriver = rdUseOdbc
           Conn1.EstablishConnection rdDriverNoPrompt

           SQLStr1 = "SELECT * FROM Units WHERE UnitNumber = '" & cboUnits.text & "'"
           Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

           Rst1.Delete
           Rst1.Close
           Conn1.Close

           Call EmptyBoxes
           Call ResetScreen
       End If
   End If
End Sub

Public Sub AddNew()
    cboUnits.Enabled = False

    UnitNumber = ""

'    cmdAdd.Visible = False
'    cmdDelete.Visible = False
'    cmdSaveNew.Visible = True
'    cmdCancelNew.Visible = True
'    cmdEdit.Visible = False
'    cmdSaveEdit.Visible = False
'    cmdCancelEdit.Visible = False

'    fraAddNew.Visible = True
    fraMain.ZOrder 1
    fraAddNew.ZOrder 0
    cmdSaveNew.Visible = True
    cmdCancelNew.Visible = True
    
    Call EnableBoxes
    Call EmptyBoxes
    Call GetTenants
End Sub

Public Sub EmptyBoxes()

cboUnits.text = ""
txtUnit.text = ""
txt1.text = ""
txt2.text = ""
txt3.text = ""
txt4.text = ""
txt5.text = ""
txt6.text = ""
'cboOccupied.Clear
cboTenant.Clear

End Sub

Public Sub CancelAddEdit()
    fraMain.ZOrder 0
    fraAddNew.ZOrder 1
End Sub


Private Sub mnuGlobal_Click()

Load frmGlobal
Unload Me
frmGlobal.Show

End Sub

Private Sub mnuLease_Click()

Load frmLease
Unload Me
frmLease.Show

End Sub

Private Sub mnuMain_Click()

'Load frmMain
Unload Me
'frmMain.Show

End Sub

Private Sub mnuShopCentre_Click()

Load frmShoppingCentre
Unload Me
frmShoppingCentre.Show

End Sub

Private Sub mnuTenants_Click()

'Load frmTenant
'Unload Me
'frmTenant.Show

End Sub


Private Sub Label12_Click()
'   Update
End Sub

Private Sub lblInsuranceDate_Click()
   dptDate.Top = lblInsuranceDate.Top '+ fraInsurance.Top
   dptDate.Left = lblInsuranceDate.Left '+ fraInsurance.Left
   dptDate.ZOrder 0
   dptDate.Visible = True
End Sub

Private Sub TabStrip1_Click()
'   If TabStrip1.SelectedItem.Index - 1 = mintCurFrame Then Exit Sub     ' No need to change frame.
'   ' Otherwise, hide old frame, show new.
'   Frame1(TabStrip1.SelectedItem.Index - 1).Visible = True
'   Frame1(mintCurFrame).Visible = False
'   ' Set mintCurFrame to new value.
'   mintCurFrame = TabStrip1.SelectedItem.Index - 1
'   If (TabStrip1.SelectedItem.Index - 1) = 3 Then 'Maintainance History Tab clicked
'      txtUnitTemp.Text = cboUnits.Text
''      txtUnitTemp.SetFocus
'   End If
End Sub

Private Sub tabUnit_Click(PreviousTab As Integer)
   Select Case tabUnit.Tab
   Case Is = 3
      tabUnit.TabCaption(3) = "&Maint. Address"
   Case Is = 5
      tabUnit.TabCaption(5) = "&Ins. && Safety"
   Case Else
      Select Case PreviousTab
         Case Is = 3
            tabUnit.TabCaption(3) = "Maintenance &History"
         Case Is = 5
            tabUnit.TabCaption(5) = "&Insurance && Safety"
      End Select
'      If PreviousTab = 3 Then tabUnit.TabCaption(3) = "Maintenance &History"
'      If PreviousTab = 5 Then tabUnit.TabCaption(5) = "&Insurance && Safety"
   End Select
End Sub

Private Sub txt1_LostFocus()
On Error Resume Next
If txt1.text = "" Then txt1.text = "0" Else If NumberCheck2(txt1.text) = False Then txt1.text = ""
txt1.text = Round(CDbl(txt1.text), 2)

End Sub

Private Sub txt2_LostFocus()

If txt2.text = "" Then txt2.text = "0" Else If NumberCheck2(txt2.text) = False Then txt2.text = ""
txt2.text = Round(CDbl(txt2.text), 2)

End Sub

Private Sub txt3_LostFocus()

If txt3.text = "" Then txt3.text = "0" Else If NumberCheck2(txt3.text) = False Then txt3.text = ""
txt3.text = Round(CDbl(txt3.text), 2)

End Sub

Private Sub txt4_LostFocus()

If txt4.text <> "" Then
    If NumberCheck(txt4.text) = False Then
        txt4.text = ""
        Exit Sub
    End If
    If txt5.text <> "" Then
        If NumberCheck(txt5.text) = False Then
            txt5.text = ""
            Exit Sub
        End If
        txt6.text = CLng(txt4.text) + CLng(txt5.text)
    Else
        txt6.text = CLng(txt4.text)
    End If
Else
    If txt5.text <> "" Then
        If NumberCheck(txt5.text) = False Then
            txt5.text = ""
            Exit Sub
        End If
        txt6.text = CLng(txt5.text)
    End If
End If

End Sub

Private Sub txt5_LostFocus()

If txt4.text <> "" Then
    If NumberCheck(txt4.text) = False Then
        txt4.text = ""
        Exit Sub
    End If
    If txt5.text <> "" Then
        If NumberCheck(txt5.text) = False Then
            txt5.text = ""
            Exit Sub
        End If
        txt6.text = CLng(txt4.text) + CLng(txt5.text)
    Else
        txt6.text = CLng(txt4.text)
    End If
Else
    If txt5.text <> "" Then
        If NumberCheck(txt5.text) = False Then
            txt5.text = ""
            Exit Sub
        End If
        txt6.text = CLng(txt5.text)
    End If
End If

End Sub

Public Sub EmptyVars()

OldUnit = ""
NewUnit = ""
OldTenantName = ""
NewTenantName = ""
OldTenantCode = ""
NewTenantCode = ""
oldtenant = ""

End Sub

Private Sub txtAddInsuRenewDt_GotFocus()
   dptAddInsuDt.Top = txtAddInsuRenewDt.Top
   dptAddInsuDt.Left = lblInsuranceDate.Left
   dptAddInsuDt.ZOrder 0
   dptAddInsuDt.Visible = True
End Sub

