VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTenantNew 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tenancy"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   Icon            =   "frmTenantNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11835
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Index           =   0
      Left            =   210
      TabIndex        =   36
      Top             =   0
      Width           =   10935
      Begin VB.OptionButton optHO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inv Head Office"
         Height          =   315
         Left            =   9000
         TabIndex        =   70
         Top             =   4080
         Width           =   1815
      End
      Begin VB.OptionButton optBil 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inv Billing Address"
         Height          =   315
         Left            =   9000
         TabIndex        =   69
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Frame fraAddNew 
         BackColor       =   &H00DFDDAF&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   9000
         TabIndex        =   68
         Top             =   120
         Width           =   1815
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            MaskColor       =   &H00FFC0FF&
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdSaveNew 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Save"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            MaskColor       =   &H00FFC0FF&
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1080
            Width           =   1200
         End
         Begin VB.CommandButton cmdCancelNew 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            MaskColor       =   &H00FFC0FF&
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1920
            Width           =   1200
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Billing Address:"
         Height          =   3255
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   4095
         Begin VB.TextBox txtBillTel 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   15
            Top             =   2565
            Width           =   2000
         End
         Begin VB.TextBox txtBillFax 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   16
            Top             =   2835
            Width           =   2000
         End
         Begin VB.TextBox txtBillPostCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1650
            Width           =   2000
         End
         Begin VB.TextBox txtBillDirectLine 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   14
            Top             =   2310
            Width           =   2000
         End
         Begin VB.TextBox txtBillEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   13
            Top             =   2040
            Width           =   2600
         End
         Begin VB.TextBox txtBillContact 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   7
            Top             =   240
            Width           =   2600
         End
         Begin VB.TextBox txtBillAdd1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   8
            Top             =   600
            Width           =   2600
         End
         Begin VB.TextBox txtBillAdd2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   9
            Top             =   855
            Width           =   2600
         End
         Begin VB.TextBox txtBillAdd3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   10
            Top             =   1110
            Width           =   2600
         End
         Begin VB.TextBox txtBillAdd4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   11
            Top             =   1380
            Width           =   2600
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   2520
            Width           =   510
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Left            =   180
            TabIndex        =   66
            Top             =   2760
            Width           =   300
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   180
            TabIndex        =   65
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   180
            TabIndex        =   64
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Head Office:"
         Height          =   3255
         Left            =   4200
         TabIndex        =   52
         Top             =   1200
         Width           =   4095
         Begin VB.TextBox txtHOContact 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   17
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtHOEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   24
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtHODirectLine 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   25
            Top             =   2310
            Width           =   1935
         End
         Begin VB.TextBox txtHOPostCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1665
            Width           =   1455
         End
         Begin VB.TextBox txtHOFax 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   27
            Top             =   2835
            Width           =   1935
         End
         Begin VB.TextBox txtHOTel 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   26
            Top             =   2565
            Width           =   1935
         End
         Begin VB.TextBox txtHOAdd3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   21
            Top             =   1125
            Width           =   2415
         End
         Begin VB.TextBox txtHOAdd4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   22
            Top             =   1395
            Width           =   2415
         End
         Begin VB.TextBox txtHOAdd2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   20
            Top             =   855
            Width           =   2415
         End
         Begin VB.TextBox txtHOAdd1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   18
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Left            =   180
            TabIndex        =   59
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            Height          =   195
            Left            =   180
            TabIndex        =   58
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   180
            TabIndex        =   55
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   2760
            Width           =   300
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   2520
            Width           =   510
         End
      End
      Begin VB.ComboBox cboTenants_ 
         Height          =   315
         Left            =   9360
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00DFDDAF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   960
         TabIndex        =   42
         Top             =   400
         Width           =   3240
         Begin VB.OptionButton optCurrentTenant 
            BackColor       =   &H00DFDDAF&
            Caption         =   "Current"
            Height          =   195
            Left            =   840
            TabIndex        =   45
            Top             =   80
            Width           =   975
         End
         Begin VB.OptionButton optExTenant 
            BackColor       =   &H00DFDDAF&
            Caption         =   "Ex-Tenant"
            Height          =   195
            Left            =   1920
            TabIndex        =   44
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optBoth 
            BackColor       =   &H00DFDDAF&
            Caption         =   "Both"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   80
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.ComboBox cboSage 
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   450
         Width           =   3015
      End
      Begin VB.TextBox txtTenantCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtTenantName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   840
         Width           =   3000
      End
      Begin VB.TextBox txtTenantID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   5
         Top             =   840
         Width           =   3000
      End
      Begin VB.CommandButton cmdTenantList 
         Caption         =   "v"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtTenant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   1
         Top             =   120
         Width           =   3000
      End
      Begin VB.Frame fraTenant 
         Height          =   2415
         Left            =   1500
         TabIndex        =   37
         Top             =   60
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox txtSearchAC 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   80
            TabIndex        =   38
            Top             =   120
            Width           =   1935
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTenants 
            Height          =   1935
            Left            =   80
            TabIndex        =   39
            Top             =   360
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   3413
            _Version        =   393216
            FixedCols       =   0
            GridColor       =   -2147483635
            SelectionMode   =   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblPayeeFlxConfigured 
            Caption         =   "NOT"
            Height          =   495
            Left            =   1200
            TabIndex        =   41
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblFlxPayee 
            Caption         =   "EMPTY"
            Height          =   255
            Left            =   2280
            TabIndex        =   40
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   5950
            Picture         =   "frmTenantNew.frx":030A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   240
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   195
         Left            =   4440
         TabIndex        =   50
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sage A/C:"
         Height          =   195
         Left            =   4440
         TabIndex        =   49
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   4440
         TabIndex        =   47
         Top             =   840
         Width           =   210
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   600
      Top             =   8160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraEdit 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   7880
      Width           =   10215
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   8160
         TabIndex        =   35
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Details"
         Height          =   375
         Left            =   5280
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel Changes"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "&Save Changes"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Tenant"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   75
      TabIndex        =   71
      Top             =   4680
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Tenancy Details"
      TabPicture(0)   =   "frmTenantNew.frx":074C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "gridBankCode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSageRef"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraLease"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tenant A/C History"
      TabPicture(1)   =   "frmTenantNew.frx":0768
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxTenatACHistory"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Premises"
      TabPicture(2)   =   "frmTenantNew.frx":0784
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdUnitDetails"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtClient"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtProperty"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmbUnits"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Notes/Reminder/Memo"
      TabPicture(3)   =   "frmTenantNew.frx":07A0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTenatACHistory 
         Height          =   2475
         Left            =   -74880
         TabIndex        =   109
         Top             =   480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4366
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   315
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame12 
         Caption         =   "Notes && Reminder:"
         Height          =   2295
         Left            =   -74280
         TabIndex        =   103
         Top             =   480
         Visible         =   0   'False
         Width           =   7575
         Begin VB.TextBox txtNotes 
            Height          =   1935
            Left            =   120
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Text            =   "frmTenantNew.frx":07BC
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   99
         Top             =   3600
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open Attachment"
            Height          =   375
            Left            =   7080
            TabIndex        =   102
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbFiles 
            Height          =   315
            Left            =   1800
            TabIndex        =   101
            Top             =   240
            Width           =   5055
         End
         Begin VB.CommandButton cmdAddNewFile 
            Caption         =   "&Add Attachment"
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraLease 
         BackColor       =   &H00E1E1ED&
         Caption         =   "Lease Details:"
         Enabled         =   0   'False
         Height          =   2535
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   3375
         Begin VB.ComboBox cmbTenancyType 
            Height          =   315
            Left            =   1575
            TabIndex        =   92
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtLeaseStart 
            Height          =   285
            Left            =   1560
            TabIndex        =   91
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtLeaseEnd 
            Height          =   285
            Left            =   1560
            TabIndex        =   90
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E1E1ED&
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   89
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtRentReviewDate 
            Height          =   285
            Left            =   1560
            TabIndex        =   88
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtReviewFrequency 
            Height          =   285
            Left            =   1575
            TabIndex        =   87
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Holding Over:"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Review Frequency:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   2040
            Width           =   1380
         End
      End
      Begin VB.TextBox txtSageRef 
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E1E1ED&
         Caption         =   "Accounts:"
         Height          =   2535
         Left            =   3600
         TabIndex        =   77
         Top             =   480
         Width           =   3255
         Begin VB.CommandButton cmdBankCode 
            Caption         =   "v"
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
            Left            =   2950
            TabIndex        =   81
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtBankCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtBalance 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtDeposite 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   840
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Deposit:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deposit Bank:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1800
         Width           =   1920
      End
      Begin VB.TextBox txtProperty 
         Height          =   285
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox txtClient 
         Height          =   285
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   840
         Width           =   2280
      End
      Begin VB.CommandButton cmdUnitDetails 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69600
         TabIndex        =   73
         Top             =   1800
         Width           =   325
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
         Height          =   1515
         Left            =   6840
         TabIndex        =   72
         Top             =   2040
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Ref:"
         Height          =   195
         Left            =   7200
         TabIndex        =   108
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   107
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Left            =   -72240
         TabIndex        =   106
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   -72240
         TabIndex        =   105
         Top             =   840
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmTenantNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OldUnit As String
Public NewUnit As String
Public OldSageAct As String
Public LodeTenant As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
           As Long) As Long
Const SW_SHOW = 5

Private mintCurFrame As Integer ' Current Frame visible

Dim iTenantInDB As Integer

Dim TenantCode As String
Dim TenantName As String
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

Private Sub LodeAttachedFiles()
   cmbFiles.Clear
   
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT FileName FROM AttachedFile where SageAccountNumber = '" & cboSage.text & "'"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

   While Not Rst1.EOF
       cmbFiles.AddItem Rst1!Filename
       Rst1.MoveNext
   Wend
   Rst1.Close
   Conn1.Close
End Sub

Private Sub cboSage_Click()
   Dim szaAc() As String
   
   szaAc = Split(cboSage.text, " \ ")
   
   txtTenantID.text = szaAc(0)
End Sub

Private Sub cboTenants_Click()
'   Dim i, j, match As Integer
'   match = 0
'
'   If cboTenants.text = "" Then
'       MsgBox "You must select a Tenant to view!", vbOKOnly + vbCritical, "No tenant selected"
'       Exit Sub
'   End If
'   j = cboTenants.ListCount - 1
'   For i = 0 To j
'       If cboTenants.List(i) = cboTenants.text Then
'           match = 1
'           Exit For
'       End If
'   Next i
'   If match = 0 Then
'       MsgBox "Tenant selected is invalid.", vbOKOnly + vbCritical, "Invalid Tenant"
'       cboTenants.text = ""
'       Exit Sub
'   End If
'
'   ' Set the RDO Conn1ection to the dataset
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseIfNeeded
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   For i = 2 To 10
'       If Mid(cboTenants.text, i, 3) = " / " Then
'           TenantCode = Left(cboTenants.text, i - 1)
'       End If
'   Next i
'
'   'Get record for selected Tenant.
'   SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'   'Fill text boxes with tenant details.
'   txtTenantName.text = cboTenants.text
'   cboSage.text = TenantCode
'   If IsNull(Rst1!CompanyName) Then txtTenantCompany.text = "" Else txtTenantCompany.text = Rst1!CompanyName
''   If IsNull(Rst1!CurrentRental) Then cboUnit.text = "" Else cboUnit.text = Rst1!CurrentRental
'
'   If IsNull(Rst1!Contact1) Then txtBillContact.text = "" Else txtBillContact.text = Rst1!Contact1
'   If IsNull(Rst1!Email1) Then txtBillEmail.text = "" Else txtBillEmail.text = Rst1!Email1
'   If IsNull(Rst1!DirectLine1) Then txtBillDirectLine.text = "" Else txtBillDirectLine.text = Rst1!DirectLine1
'   If IsNull(Rst1!BillAddressLine1) Then txtBillAdd1.text = "" Else txtBillAdd1.text = Rst1!BillAddressLine1
'   If IsNull(Rst1!BillAddressLine2) Then txtBillAdd2.text = "" Else txtBillAdd2.text = Rst1!BillAddressLine2
'   If IsNull(Rst1!BillAddressLine3) Then txtBillAdd3.text = "" Else txtBillAdd3.text = Rst1!BillAddressLine3
'   If IsNull(Rst1!BillAddressLine4) Then txtBillAdd4.text = "" Else txtBillAdd4.text = Rst1!BillAddressLine4
'   If IsNull(Rst1!BillPostCode) Then txtBillPostCode.text = "" Else txtBillPostCode.text = Rst1!BillPostCode
'   If IsNull(Rst1!BillTelephone) Then txtBillTel.text = "" Else txtBillTel.text = Rst1!BillTelephone
'   If IsNull(Rst1!BillFax) Then txtBillFax.text = "" Else txtBillFax.text = Rst1!BillFax
'
'   If IsNull(Rst1!HOAddressLine1) Then txtHOAdd1.text = "" Else txtHOAdd1.text = Rst1!HOAddressLine1
'   If IsNull(Rst1!HOAddressLine2) Then txtHOAdd2.text = "" Else txtHOAdd2.text = Rst1!HOAddressLine2
'   If IsNull(Rst1!HOAddressLine3) Then txtHOAdd3.text = "" Else txtHOAdd3.text = Rst1!HOAddressLine3
'   If IsNull(Rst1!HOAddressLine4) Then txtHOAdd4.text = "" Else txtHOAdd4.text = Rst1!HOAddressLine4
'   If IsNull(Rst1!HOPostCode) Then txtHOPostCode.text = "" Else txtHOPostCode.text = Rst1!HOPostCode
'   If IsNull(Rst1!HOTelephone) Then txtHOTel.text = "" Else txtHOTel.text = Rst1!HOTelephone
'   If IsNull(Rst1!HOFax) Then txtHOFax.text = "" Else txtHOFax.text = Rst1!HOFax
'   If IsNull(Rst1!DirectLine2) Then txtHODirectLine.text = "" Else txtHODirectLine.text = Rst1!DirectLine2
'   If IsNull(Rst1!Email2) Then txtHOEmail.text = "" Else txtHOEmail.text = Rst1!Email2
'
'   If IsNull(Rst1!Comments) Then txtNotes.text = "" Else txtNotes.text = Rst1!Comments
'   If Rst1!InvoiceTo = "B" Then optBil.Value = True Else optHO.Value = True
'
'   Rst1.Close
'   Conn1.Close
'
''    Call DisableTextBoxes
'
'   LodeAttachedFiles
'
''***********************************************
''delete these lines later when no longer needed
''   lblTenantName(1).Caption = cboTenants.text
''   lblTenantName(2).Caption = cboTenants.text
''   lblTenantName(3).Caption = cboTenants.text
''************************************************
End Sub

Private Sub cboSage_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboSage.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbFiles_Click()
   If cmbFiles.text = "ADD NEW FILES" Then
       AddFile
   End If
End Sub

Private Sub cmdAdd_Click()
'   cboTenants.Enabled = False
'
   TenantCode = ""
'
   cmdSaveNew.Enabled = True
   cmdCancelNew.Enabled = True
   cmdAdd.Enabled = False
   txtTenant.Enabled = False
   cmdTenantList.Enabled = False
   fraEdit(0).Enabled = False
'
   LockedTextBoxes False
   ClearTextBoxes
'    Call EmptyBoxes
   SageCustomerAccCombo
'
   txtTenantName.SetFocus
End Sub

Private Sub ClearTextBoxes()
   txtTenantName.text = ""
   txtTenantID.text = ""
   txtTenantCompany.text = ""
   cboSage.text = ""
   
   txtLeaseStart.text = ""
   txtLeaseEnd.text = ""
   txtRentReviewDate.text = ""
   cmbTenancyType.text = ""
   txtReviewFrequency.text = ""

   txtBillContact.text = ""
   txtBillAdd1.text = ""
   txtBillAdd2.text = ""
   txtBillAdd3.text = ""
   txtBillAdd4.text = ""
   txtBillPostCode.text = ""
   txtBillEmail.text = ""
   txtBillDirectLine.text = ""
   txtBillTel.text = ""
   txtBillFax.text = ""
   
   txtHOContact.text = ""
   txtHOAdd1.text = ""
   txtHOAdd2.text = ""
   txtHOAdd3.text = ""
   txtHOAdd4.text = ""
   txtHOPostCode.text = ""
   txtHOEmail.text = ""
   txtHODirectLine.text = ""
   txtHOTel.text = ""
   txtHOFax.text = ""
End Sub

Private Sub LockedTextBoxes(bTF As Boolean)
   txtTenantName.Locked = bTF
   txtTenantID.Locked = bTF
   txtTenantCompany.Locked = bTF
   cboSage.Locked = bTF

   txtBillContact.Locked = bTF
   txtBillAdd1.Locked = bTF
   txtBillAdd2.Locked = bTF
   txtBillAdd3.Locked = bTF
   txtBillAdd4.Locked = bTF
   txtBillPostCode.Locked = bTF
   txtBillEmail.Locked = bTF
   txtBillDirectLine.Locked = bTF
   txtBillTel.Locked = bTF
   txtBillFax.Locked = bTF
'
   txtHOContact.Locked = bTF
   txtHOAdd1.Locked = bTF
   txtHOAdd2.Locked = bTF
   txtHOAdd3.Locked = bTF
   txtHOAdd4.Locked = bTF
   txtHOPostCode.Locked = bTF
   txtHOEmail.Locked = bTF
   txtHODirectLine.Locked = bTF
   txtHOTel.Locked = bTF
   txtHOFax.Locked = bTF
   
   txtBankCode.Locked = bTF
   txtBalance.Locked = bTF
   txtDeposite.Locked = bTF
   
   SSTab1.TabEnabled(1) = bTF
   SSTab1.TabEnabled(2) = bTF
End Sub

Private Sub cmdAddNewFile_Click()
   AddFile
End Sub

Private Sub OpenFile()

' I CLOSED IT FOR COMPILATION*****DO NOT DELETE***USEFUL CODE, IS WORKING FINE**

'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseOdbc
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   SQLStr1 = "SELECT * FROM AttachedFile WHERE FileName= '" & cmbFiles & "' AND SageAccountNumber='" & cboSage.text & "' AND Tenant='" & cboTenants.text & "'"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
''                                          rdOpenStatic, rdConcurReadOnly)  use this options
'
'   Dim filePath As String
'   Dim Filename As String
'   filePath = Rst1!filePath
'   Filename = Rst1!Filename
'   filePath = Mid(filePath, 1, Len(filePath) - Len(Filename))
'   If Rst1.EOF Then MsgBox Rst1.RowCount
'
''    Rst1.Update
'   Rst1.Close
'   Conn1.Close
'   Dim errorLevel As Long
'
'   errorLevel = ShellExecute(hWnd, "open", Filename, vbNullString, filePath, SW_SHOW)
'   If errorLevel < 32 Then MsgBox "File has been moved from original location.", vbExclamation
'
''    ShellExecute hwnd, "open", fileName, vbNullString, "C:\EXCH", SW_SHOW
End Sub

Private Sub AddFile()

' I CLOSED IT FOR COMPILATION*****DO NOT DELETE***USEFUL CODE, IS WORKING FINE**

'   If cboTenants.text = "" Then Exit Sub
'
'   Dim ofn As OPENFILENAME, retVal As Long
'   ofn.lStructSize = Len(ofn)
'   ofn.hwndOwner = Me.hWnd
'   ofn.hInstance = App.hInstance
'   ofn.lpstrFilter = "Word Documents" & Chr(0) & "*.doc" & Chr(0) & _
'                       "Excel Documents" & Chr(0) & "*.xls"
'   ofn.lpstrFile = Space$(254)
'   ofn.nMaxFile = 255
'   ofn.lpstrFileTitle = Space$(254)
'   ofn.nMaxFileTitle = 255
'   ofn.lpstrInitialDir = CurDir
'   ofn.lpstrTitle = "Our File Open Title"
'   ofn.Flags = 0
'
'   retVal = GetOpenFileName(ofn)
'
'   If (retVal) Then
'       Dim TxtFile As String, ff As Integer
'       TxtFile = "C:\Filelist.txt"
'       ff = FreeFile
''        Open TxtFile For Append As #ff
''        Print #ff, Trim$(ofn.lpstrFile)
'       Close #ff
'   Else
'       Exit Sub
'   End If
'
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseOdbc
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   SQLStr1 = "SELECT * FROM AttachedFile"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
'
'   Rst1.AddNew
'   Rst1!filePath = Trim$(ofn.lpstrFile)
'   Rst1!Filename = Trim$(ofn.lpstrFileTitle)
'   Rst1!SageAccountNumber = cboSage.text
'   Rst1!Tenant = cboTenants.text
'
'   Rst1.Update
'   Rst1.Close
'   Conn1.Close
'
''    cmbFiles.RemoveItem (cmbFiles.ListCount - 1)
'   cmbFiles.AddItem Trim$(ofn.lpstrFileTitle)
''    cmbFiles.AddItem ("ADD NEW FILES")
''
'   MsgBox "File has been attached successfull, Thanks"
End Sub

Private Sub cmdBankCode_Click()
   MousePointer = vbHourglass
   gridBankCode.Top = Frame11.Top + 680
   gridBankCode.Left = Frame11.Left
   gridBankCode.Visible = True
   gridBankCode.ZOrder 0
'
   BankAccount
'
   MousePointer = vbDefault
End Sub

Private Sub cmdCancelEdit_Click()
   cmdSaveEdit.Enabled = False
   cmdCancelEdit.Enabled = False
   cmdEdit.Enabled = True
   fraAddNew.Enabled = True
End Sub

Private Sub cmdCancelNew_Click()
   If cmdAdd.Enabled = True Then Exit Sub
   ClearTextBoxes
   cmdAdd.Enabled = True
   cmdTenantList.Enabled = True
   cmdSaveNew.Enabled = False
   txtTenant.Enabled = True
   fraEdit(0).Enabled = True
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub SelectedFromGrid()
   Dim i As Integer
'
   ClearTextBoxes
'
   If iTenantInDB = 0 Then
      MsgBox "No Data in local database", vbInformation + vbOKOnly, "Tenant - PCM(ND623)"
      fraTenant.Visible = False
   End If
   txtTenant.text = flxTenants.TextMatrix(flxTenants.Row, 0) & " / " & flxTenants.TextMatrix(flxTenants.Row, 2)
'
'    Set the RDO Conn1ection to the dataset
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
'
'   Get record for selected Tenant.
   SQLStr1 = "SELECT Tenants.SageAccountNumber as SAC, Tenants.TenantID, " & _
               "Tenants.Name, Tenants.CompanyName, Tenants.Contact1, " & _
               "Tenants.Email1, Tenants.DirectLine1, Tenants.Contact2, " & _
               "Tenants.Email2, Tenants.DirectLine2, Tenants.HOAddressLine1, " & _
               "Tenants.HOAddressLine2, Tenants.HOAddressLine3, Tenants.HOAddressLine4, " & _
               "Tenants.HOPostCode, Tenants.HOTelephone, Tenants.HOFax, " & _
               "Tenants.BillAddressLine1, Tenants.BillAddressLine2, Tenants.BillAddressLine3, " & _
               "Tenants.BillAddressLine4, Tenants.BillPostCode, Tenants.BillTelephone, " & _
               "Tenants.BillFax, Tenants.InvoiceTo, Tenants.CurrentRental, Tenants.Comments, " & _
               "Tenants.BankCode, Tenants.Balance, Tenants.Deposite "
   SQLStr1 = SQLStr1 + "FROM Tenants "
   SQLStr1 = SQLStr1 + "WHERE Tenants.SageAccountNumber = '" & flxTenants.TextMatrix(flxTenants.Row, 1) & "';"
'
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'   Fill text boxes with tenant details.
   If Not Rst1.EOF Then
      If IsNull(Rst1!CompanyName) Then txtTenantCompany.text = "" Else txtTenantCompany.text = Rst1!CompanyName
'
      txtTenantName.text = IIf(IsNull(Rst1!Name), "", Rst1!Name)
      txtTenantCompany.text = IIf(IsNull(Rst1!CompanyName), "", Rst1!CompanyName)
      txtTenantID.text = IIf(IsNull(Rst1!TenantID), "", Rst1!TenantID)
      cboSage.text = Rst1!SAC
'
      txtBankCode.text = IIf(IsNull(Rst1!BankCode), "", Rst1!BankCode)
      txtBalance.text = IIf(IsNull(Rst1!Balance), "", Rst1!Balance)
      txtDeposite.text = IIf(IsNull(Rst1!Deposite), "", Rst1!Deposite)
'
      If IsNull(Rst1!Contact1) Then txtBillContact.text = "" Else txtBillContact.text = Rst1!Contact1
      If IsNull(Rst1!Email1) Then txtBillEmail.text = "" Else txtBillEmail.text = Rst1!Email1
      If IsNull(Rst1!DirectLine1) Then txtBillDirectLine.text = "" Else txtBillDirectLine.text = Rst1!DirectLine1
      If IsNull(Rst1!BillAddressLine1) Then txtBillAdd1.text = "" Else txtBillAdd1.text = Rst1!BillAddressLine1
      If IsNull(Rst1!BillAddressLine2) Then txtBillAdd2.text = "" Else txtBillAdd2.text = Rst1!BillAddressLine2
      If IsNull(Rst1!BillAddressLine3) Then txtBillAdd3.text = "" Else txtBillAdd3.text = Rst1!BillAddressLine3
      If IsNull(Rst1!BillAddressLine4) Then txtBillAdd4.text = "" Else txtBillAdd4.text = Rst1!BillAddressLine4
      If IsNull(Rst1!BillPostCode) Then txtBillPostCode.text = "" Else txtBillPostCode.text = Rst1!BillPostCode
      If IsNull(Rst1!BillTelephone) Then txtBillTel.text = "" Else txtBillTel.text = Rst1!BillTelephone
      If IsNull(Rst1!BillFax) Then txtBillFax.text = "" Else txtBillFax.text = Rst1!BillFax
   '
      If IsNull(Rst1!HOAddressLine1) Then txtHOAdd1.text = "" Else txtHOAdd1.text = Rst1!HOAddressLine1
      If IsNull(Rst1!HOAddressLine2) Then txtHOAdd2.text = "" Else txtHOAdd2.text = Rst1!HOAddressLine2
      If IsNull(Rst1!HOAddressLine3) Then txtHOAdd3.text = "" Else txtHOAdd3.text = Rst1!HOAddressLine3
      If IsNull(Rst1!HOAddressLine4) Then txtHOAdd4.text = "" Else txtHOAdd4.text = Rst1!HOAddressLine4
      If IsNull(Rst1!HOPostCode) Then txtHOPostCode.text = "" Else txtHOPostCode.text = Rst1!HOPostCode
      If IsNull(Rst1!HOTelephone) Then txtHOTel.text = "" Else txtHOTel.text = Rst1!HOTelephone
      If IsNull(Rst1!HOFax) Then txtHOFax.text = "" Else txtHOFax.text = Rst1!HOFax
      If IsNull(Rst1!DirectLine2) Then txtHODirectLine.text = "" Else txtHODirectLine.text = Rst1!DirectLine2
      If IsNull(Rst1!Email2) Then txtHOEmail.text = "" Else txtHOEmail.text = Rst1!Email2
   '
      If IsNull(Rst1!Comments) Then txtNotes.text = "" Else txtNotes.text = Rst1!Comments
      If Rst1!InvoiceTo = "B" Then optBil.Value = True Else optHO.Value = True
   End If
'
   Rst1.Close

   SQLStr1 = "SELECT LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.RentReviewDate, " & _
                  "LeaseDetails.UnitNumber, Property.PropertyName, Client.ClientName "
   SQLStr1 = SQLStr1 + "FROM Tenants, LeaseDetails, Units, Property, Client "
   SQLStr1 = SQLStr1 + "WHERE Tenants.SageAccountNumber = '" & flxTenants.TextMatrix(flxTenants.Row, 1) & "' AND " & _
                           "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                           "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                           "Units.PropertyID = Property.PropertyID AND " & _
                           "Property.ClientID = Client.ClientID"
'Debug.Print SQLStr1
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   If Not Rst1.EOF Then
      txtLeaseStart.text = Rst1!StartDate
      txtLeaseEnd.text = Rst1!EndDate
      cmbUnits.text = Rst1!UnitNumber
      txtProperty.text = Rst1!PropertyName
      txtClient.text = IIf(IsNull(Rst1!ClientName), "", Rst1!ClientName)
      txtRentReviewDate.text = IIf(IsNull(Rst1!RentReviewDate), "", Rst1!RentReviewDate)
   Else
      cmbUnits.text = "Not Occupied"
   End If
'
   Rst1.Close
   Conn1.Close
'
   LodeAttachedFiles
'
   fraTenant.Visible = False
'
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub cmdEdit_Click()
   cmdEdit.Enabled = False
   cmdSaveEdit.Enabled = True
   cmdCancelEdit.Enabled = True
   fraAddNew.Enabled = False
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then
       MsgBox "Select a file from list."
   Else
       OpenFile
   End If
End Sub

Private Sub cmdPrint_Click()
'   MousePointer = vbHourglass
''
'   If FileExists(App.Path & "\Tenant" & SCID & ".rpt") = False Then
'       MsgBox "Unable to Print Details", vbOKOnly + vbInformation, "Unable To Print"
'       Exit Sub
'   End If
''
'   CR1.ReportFileName = App.Path & "\Tenant" & SCID & ".rpt"
'   CR1.SelectionFormula = "{Tenants.SageAccountNumber} = '" & cboSage.text & "'"
'   CR1.printReport
''
'   MousePointer = vbDefault
End Sub

Private Function ValidityCheck() As Boolean
   ValidityCheck = True
   
   If cboSage.text = "" Then
      MsgBox "You must select a Sage Account Number!", vbOKOnly + vbCritical, "Sage Account Number required"
      ValidityCheck = False
      Exit Function
   End If
   If txtTenantName.text = "" Then
      MsgBox "You must input Tenant Name.", vbOKOnly + vbCritical, "Tenant Name required"
      ValidityCheck = False
      Exit Function
   End If
   If txtTenantID.text = "" Then
      MsgBox "You must input Tenant ID.", vbOKOnly + vbCritical, "Tenant ID required"
      ValidityCheck = False
      Exit Function
   End If
   If txtTenantCompany.text = "" Then
      MsgBox "You must input Tenant Company Name.", vbOKOnly + vbCritical, "Tenant Company required"
      ValidityCheck = False
      Exit Function
   End If
   If txtBillContact.text = "" Then
      MsgBox "You must input Tenant ID.", vbOKOnly + vbCritical, "Tenant ID required"
      ValidityCheck = False
      Exit Function
   End If
   If txtBillAdd1.text = "" Then
      MsgBox "You must input Tenant Address.", vbOKOnly + vbCritical, "Tenant Address required"
      ValidityCheck = False
      Exit Function
   End If
   If txtBillPostCode.text = "" Then
      MsgBox "You must input Tenant PostCode.", vbOKOnly + vbCritical, "Tenant PostCode required"
      ValidityCheck = False
      Exit Function
   End If
   If optHO.Value = False And optBil.Value = False Then
      MsgBox "You must Invoice option.", vbOKOnly + vbCritical, "Tenant Invoice Option"
      ValidityCheck = False
      Exit Function
   End If
   If txtBankCode.text = "" Then
      MsgBox "You must input Deposite Bank.", vbOKOnly + vbCritical, "Tenant Deposite Bank"
      ValidityCheck = False
      Exit Function
   End If
End Function

Private Sub cmdSaveEdit_Click()
   cmdSaveEdit.Enabled = False
   cmdCancelEdit.Enabled = False
   cmdEdit.Enabled = True
   fraAddNew.Enabled = True
End Sub

Private Sub cmdSaveNew_Click()
   Dim szaAc() As String
   
   If cmdAdd.Enabled = True Then Exit Sub
   If ValidityCheck = False Then Exit Sub
'
   Dim i, j, match As Integer
   match = 0
   j = cboSage.ListCount - 1
   For i = 0 To j
       If cboSage.List(i) = cboSage.text Then
           match = 1
           Exit For
       End If
   Next i
   If match = 0 Then
       MsgBox "Sage Account Number selected is invalid.", vbOKOnly + vbCritical, "Sage Account Number is invalid"
       cboSage.text = ""
       Exit Sub
   End If
'
   match = 0
'
   szaAc = Split(cboSage.text, " \ ")
   
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt
'
   SQLStr1 = "SELECT * FROM Tenants"
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
'
   Rst1.AddNew
'
   Rst1!SageAccountNumber = szaAc(0)
   Rst1!TenantID = txtTenantID.text
   Rst1!Name = txtTenantName.text
   Rst1!CompanyName = txtTenantCompany.text
   Rst1!Contact1 = txtBillContact.text
   Rst1!BillAddressLine1 = txtBillAdd1.text
   Rst1!BillAddressLine2 = txtBillAdd2.text
   Rst1!BillAddressLine3 = txtBillAdd3.text
   Rst1!BillAddressLine4 = txtBillAdd4.text
   Rst1!BillPostCode = txtBillPostCode.text
   Rst1!Email1 = txtBillEmail.text
   Rst1!DirectLine1 = txtBillDirectLine.text
   Rst1!BillTelephone = txtBillTel.text
   Rst1!BillFax = txtBillFax.text
   
   Rst1!Contact2 = txtHOContact.text
   Rst1!HOAddressLine1 = txtHOAdd1.text
   Rst1!HOAddressLine2 = txtHOAdd2.text
   Rst1!HOAddressLine3 = txtHOAdd3.text
   Rst1!HOAddressLine4 = txtHOAdd4.text
   Rst1!HOPostCode = txtHOPostCode.text
   Rst1!Email2 = txtHOEmail.text
   Rst1!DirectLine2 = txtHODirectLine.text
   Rst1!HOTelephone = txtHOTel.text
   Rst1!HOFax = txtHOFax.text
   
   Rst1!BankCode = txtBankCode.text
   Rst1!Balance = CCur(IIf(txtBalance.text = "", 0, txtBalance.text))
   Rst1!Deposite = CCur(IIf(txtDeposite.text = "", 0, txtDeposite.text))
'
   If optBil.Value = True Then Rst1!InvoiceTo = "B" Else Rst1!InvoiceTo = "H"
'
   Rst1.Update
   Rst1.Close
   Conn1.Close
'
'   cboTenants.text = cboSage.text & " / " & txtTenantCompany.text
   MsgBox "The new tenant details have been saved.", vbOKOnly + vbInformation, "New Tenant"
'
'   Call ResetScreen
   ClearTextBoxes
   LockedTextBoxes True
   
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False
   cmdTenantList.Enabled = True
   txtTenant.Enabled = True
   cmdAdd.Enabled = True
   fraEdit(0).Enabled = True
End Sub

Private Sub cmdTenantList_Click()
   MousePointer = vbHourglass
   
   fraTenant.Top = 440
   fraTenant.Left = 960
   fraTenant.Visible = True
   fraTenant.ZOrder 0
   
   ConfigureFlxGrid flxTenants
   LoadTenantInGrid
   MousePointer = vbDefault
End Sub

Private Sub LoadTenantInGrid()
   
'   Reset screen to show all Company Names and Sage Account numer in cboTenants
'    Set the RDO Conn1ection to the dataset
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
'
'  Create Query String depands on Option choice
   If optBoth.Value Then
      SQLStr1 = "SELECT SageAccountNumber, CompanyName, Name, TenantID " & _
                "FROM Tenants " & _
                "ORDER BY TenantID"
   Else
      If optCurrentTenant.Value Then
         SQLStr1 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName, " & _
                        "Tenants.Name, Tenants.TenantID " & _
                   "From Tenants, Units, LeaseDetails " & _
                   "Where (Units.Occupied = 'Y') AND " & _
                        "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                        "LeaseDetails.UnitNumber = Units.UnitNumber " & _
                   "ORDER BY Tenants.TenantID;"
      Else
         SQLStr1 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName, " & _
                        "Tenants.Name, Tenants.TenantID " & _
                   "From Tenants " & _
                   "Where " & _
                        "Tenants.CurrentRental <> '' " & _
                   "ORDER BY Tenants.TenantID;"
      End If
   End If
'
'Debug.Print SQLStr1
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
   If Rst1.EOF = False Then
      iTenantInDB = 1
      While Rst1.EOF = False
         flxTenants.TextMatrix(iTenantInDB, 0) = Rst1!TenantID
         flxTenants.TextMatrix(iTenantInDB, 1) = Rst1!SageAccountNumber
         flxTenants.TextMatrix(iTenantInDB, 2) = Rst1!Name
         flxTenants.TextMatrix(iTenantInDB, 3) = Rst1!CompanyName
         Rst1.MoveNext
         If Not Rst1.EOF Then flxTenants.AddItem ""
         iTenantInDB = iTenantInDB + 1
      Wend
   End If
'
   Rst1.Close
   Conn1.Close
'
   Exit Sub
'
ErrorTrap:
   If ERR.Number <> 0 Then
       If ERR.Number = 40002 Then
           If MsgBox("DSN -(pcm_018) " & Adsn & " does not exist.", vbRetryCancel, "DSN Error") = vbCancel Then
               Resume Next
           Else
               Resume
           End If
       Else
           MsgBox ERR.Number & " -(pcm_019) " & ERR.description
       End If
   End If
End Sub

Private Sub cmdUnitDetails_Click()
  MsgBox "Under Construction"
End Sub

Private Sub ConfigureFlxGrid(conFlxGrid As Control)
   conFlxGrid.Cols = 4

   If conFlxGrid.Rows = 2 Then
      conFlxGrid.ColWidth(0) = 1200       'ID
      conFlxGrid.ColWidth(1) = 1000       'Sage Account Number
      conFlxGrid.ColWidth(2) = 2200       'Tenant Name
      conFlxGrid.ColWidth(3) = 2200       'Tenant Company name
'      conFlxGrid.ColWidth(3) = 900        'Due Amount
   End If
   conFlxGrid.Rows = 2
   conFlxGrid.Clear
'
   conFlxGrid.TextMatrix(0, 0) = "ID"
   conFlxGrid.TextMatrix(0, 1) = "Sage A/C"
   conFlxGrid.TextMatrix(0, 2) = "Name"
   conFlxGrid.TextMatrix(0, 3) = "Company"
'
   conFlxGrid.RowHeightMin = 315
End Sub

Private Sub flxTenants_Click()
   SelectedFromGrid
   
   LoadAcHistory
End Sub

Private Sub LoadAcHistory()
   ConfigureFlxTenatACHistory flxTenatACHistory
   
   Dim szaTemp() As String, iRow As Integer
   
   szaTemp = Split(txtTenant.text, " / ")
   
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseOdbc
   Conn1.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT DemandRecords.*, DemandSplitRecords.* " & _
             "FROM DemandRecords " & _
                  "INNER JOIN DemandSplitRecords ON DemandRecords.DemandID = DemandSplitRecords.DemandID " & _
             "WHERE DemandRecords.SageAccountNumber = '" & szaTemp(0) & "' AND DemandSplitRecords.DemandStatement=True"
'Debug.Print SQLStr1
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
   
   iRow = 1
   While Not Rst1.EOF
      flxTenatACHistory.TextMatrix(iRow, 1) = Rst1!DemandId
      flxTenatACHistory.TextMatrix(iRow, 2) = Rst1!A_M
      flxTenatACHistory.TextMatrix(iRow, 3) = IIf(Rst1!TransactionType = 1, "INV", "CRN")
      flxTenatACHistory.TextMatrix(iRow, 4) = Rst1!IssueDate
      flxTenatACHistory.TextMatrix(iRow, 5) = Rst1!DueDate
      flxTenatACHistory.TextMatrix(iRow, 6) = Rst1!description
      flxTenatACHistory.TextMatrix(iRow, 7) = Rst1!SageRef
      flxTenatACHistory.TextMatrix(iRow, 8) = IIf(Rst1!TransactionType = 1, Rst1!TotalAmount, "")
      flxTenatACHistory.TextMatrix(iRow, 9) = IIf(Rst1!TransactionType = 1, "", Rst1!TotalAmount)
      flxTenatACHistory.TextMatrix(iRow, 10) = IIf(Rst1!ExportedToSage, "YES", "NO")
      iRow = iRow + 1
      Rst1.MoveNext
      If Not Rst1.EOF Then flxTenatACHistory.AddItem ""
   Wend
   Rst1.Close
   
   SQLStr1 = "SELECT tlbReceipt.* " & _
             "FROM tlbReceipt " & _
             "WHERE " & _
                  "SageAccountNumber = '" & szaTemp(0) & "' AND " & _
                  "IsSageUpdate = True" 'And " & _
                  "ReceiptView = False"
'Debug.Print SQLStr1
   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
   If Not Rst1.EOF Then flxTenatACHistory.AddItem ""
   While Not Rst1.EOF
      flxTenatACHistory.TextMatrix(iRow, 1) = Rst1!TransactionID                          'Transaction ID
      flxTenatACHistory.TextMatrix(iRow, 2) = "M"
      flxTenatACHistory.TextMatrix(iRow, 3) = "RPT"
      flxTenatACHistory.TextMatrix(iRow, 4) = Rst1!DDate
      flxTenatACHistory.TextMatrix(iRow, 5) = IIf(IsNull(Rst1!RDate), "", Rst1!RDate)
      flxTenatACHistory.TextMatrix(iRow, 6) = Rst1!Details
      flxTenatACHistory.TextMatrix(iRow, 8) = ""
      flxTenatACHistory.TextMatrix(iRow, 9) = Rst1!ReceiptAmount
      flxTenatACHistory.TextMatrix(iRow, 10) = IIf(Rst1!UpDateSage, "YES", "NO")
      
      iRow = iRow + 1
      Rst1.MoveNext
      If Not Rst1.EOF Then flxTenatACHistory.AddItem ""
   Wend
   Rst1.Close
   
   Conn1.Close
   Set Rst1 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50
   
'   On Error GoTo ErrorTrap

   Dim temp
   
'   cboTenants.Enabled = True
   iTenantInDB = 0
   SSTab1.Tab = 0
   
'   If cboTenants.text = "" Then
'       cboTenants.Clear
'   Else
'       temp = cboTenants.text
'       cboTenants.Clear
'       cboTenants.text = temp
'   End If
'
''   Reset screen to show all Company Names and Sage Account numer in cboTenants
''    Set the RDO Conn1ection to the dataset
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseIfNeeded
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   If optBoth.Value Then
'      SQLStr1 = "SELECT SageAccountNumber, CompanyName FROM Tenants ORDER BY SageAccountNumber"
'   Else
'      If optCurrentTenant.Value Then
'         SQLStr1 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName " & _
'                   "From Tenants, Units " & _
'                   "Where (Units.Occupied = 'Y') AND " & _
'                        "Tenants.SageAccountNumber = UNITS.SageAccountNumber " & _
'                   "ORDER BY Tenants.SageAccountNumber;"
'      Else
'         SQLStr1 = "SELECT Tenants.SageAccountNumber, Tenants.CompanyName " & _
'                   "From Tenants, Units " & _
'                   "Where (Units.Occupied = 'N') AND " & _
'                        "Tenants.SageAccountNumber = UNITS.SageAccountNumber " & _
'                   "ORDER BY Tenants.SageAccountNumber;"
'      End If
'   End If
'Debug.Print SQLStr1
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'   If Rst1.EOF = False Then
'       While Rst1.EOF = False
'           cboTenants.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
'           Rst1.MoveNext
'       Wend
'   End If
'
'   Rst1.Close
'   Conn1.Close
''    Call DisableTextBoxes
'   If Not LodeTenant = "" Then
'       cboTenants.text = LodeTenant
'       cboTenants_Click
'   End If
''
   ConfigureFlxTenatACHistory flxTenatACHistory
''
'ErrorTrap:
'   If ERR.Number <> 0 Then
'       If ERR.Number = 40002 Then
'           If MsgBox("DSN - " & Adsn & " does not exist.", vbRetryCancel, "DSN Error") = vbCancel Then
'               Resume Next
'           Else
'               Resume
'           End If
'       Else
'           MsgBox ERR.Number & " - " & ERR.description
'       End If
'   End If
'   Exit Sub
End Sub

Private Sub ConfigureFlxTenatACHistory(conFlxGrid As Control)
   Dim szHeader As String
   conFlxGrid.Clear
   conFlxGrid.Cols = 11
   conFlxGrid.Rows = 2

   szHeader$ = "|<ID|<A/M|<Type|<IssueDate|<DueDate|<Description|<SageRef|>Debit|>Credit|<Ex. Sage"
   conFlxGrid.FormatString = szHeader$

   conFlxGrid.ColWidth(0) = 250     'Solid column
   conFlxGrid.ColWidth(1) = 500     'ID
   conFlxGrid.ColWidth(2) = 500     'Generate Demand (A/M)
   conFlxGrid.ColWidth(3) = 500     'Type
   conFlxGrid.ColWidth(4) = 1000    'Issue Date
   conFlxGrid.ColWidth(5) = 1000    'Due Date
   conFlxGrid.ColWidth(6) = 3000    'Description
   conFlxGrid.ColWidth(7) = 1200    'SageRef
   conFlxGrid.ColWidth(8) = 1000    'Debit
   conFlxGrid.ColWidth(9) = 1000    'Credit
   conFlxGrid.ColWidth(10) = 900    'Export to sage

End Sub

Public Sub AddNew()
'   cboTenants.Enabled = False
'
'   TenantCode = ""
'
'   cmdAdd.Enabled = False
'   cmdDelete.Visible = False
'   cmdSaveNew.Visible = True
'   cmdCancelNew.Visible = True
'   cmdEdit.Visible = False
'   cmdSaveEdit.Visible = False
'   cmdCancelEdit.Visible = False
'   cmdPrint.Visible = False
''
''    Call EnableTextBoxes
''    Call EmptyBoxes
'   Call GetSageActUnits
''
'   cboUnit.text = ""
''    OldSageAct = ""
''    OldUnit = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub SageCustomerAccCombo()
   cboSage.Clear
   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oSalesRecord As SageDataObject120.SalesRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
'   oSDO.Workspaces.Clear
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

         Set oSalesRecord = oWS.CreateObject("SalesRecord")

         ' Move to the First Record
         oSalesRecord.MoveFirst
         Dim rRow As Integer
         For rRow = 1 To oSalesRecord.Count
            cboSage.AddItem CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value) & _
                                  " \ " & _
                                  CStr(oSalesRecord.Fields.Item("NAME").Value)
            oSalesRecord.MoveNext
         Next rRow

         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "(pcm_002) The SDO generated the following error: " & oSDO.LastError.text
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Public Sub GetSageActUnits()
'   Dim a As Integer
'   Dim j As Integer
'   Dim b As Integer
'   Dim k As Integer
'   Dim match As Integer
'   Dim temp1 As String
'   Dim temp2 As String
'   Dim chk1 As Boolean
'   chk1 = False
'
'   On Error GoTo ErrTrap1
'
'   If cboSage.text = "" Then
'       cboSage.Clear
'       temp1 = ""
'   Else
'       temp1 = cboSage.text
'       cboSage.Clear
'   End If
'
'   'Get all the Sage Account Numbers from Sage and put in an array.
'   ' Set the RDO Env1ironment
'   Set Envs1 = rdoEngine.rdoEnvironments
'   Set Env1 = Envs1(0)
'
'   ' Use Line100's scrolling cursor
'   Env1.CursorDriver = rdUseServer
'
'   ' Set the RDO Conn1ection to the dataset
'   Set Conn1 = Env1.OpenConnection("", rdDriverNoPrompt, False, "DSN=" & Sdsn & ";UID=" & sageUserName & ";PWD=" & sagePassword & "")
'
''    SQLStr1 = "SELECT ACCOUNT_NUMBER FROM SALES_LEDGER ORDER BY ACCOUNT_NUMBER"
'   SQLStr1 = "SELECT ACCOUNT_REF FROM SALES_LEDGER ORDER BY ACCOUNT_REF"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
''
''   Get all SageAccountNumbers from Tenants table and put in array.
'   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn2.CursorDriver = rdUseIfNeeded
'   Conn2.EstablishConnection rdDriverNoPrompt
''
'   SQLStr2 = "SELECT SageAccountNumber FROM Tenants ORDER BY SageAccountNumber"
'   Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)
''
'   If Rst1.EOF = False Then
'       If Rst2.EOF = True Then 'add all sage accounts
'           While Not Rst1.EOF
'               cboSage.AddItem Rst1!ACCOUNT_REF
'               Rst1.MoveNext
'           Wend
'       Else
'           Rst1.MoveFirst
'           While Rst1.EOF = False
'               chk1 = False
'               Rst2.MoveFirst
'               While Rst2.EOF = False
'                   If Rst2!SageAccountNumber = Rst1!ACCOUNT_REF Then chk1 = True
'                   Rst2.MoveNext
'               Wend
'               If chk1 = False Then cboSage.AddItem Rst1!ACCOUNT_REF
'               Rst1.MoveNext
'           Wend
'       End If
'   End If
''
'   Rst2.Close
'   Conn2.Close
''
'   Rst1.Close
'   Conn1.Close
'   Env1.Close
''
'   If temp1 <> "" Then
'       cboSage.text = temp1
'       cboSage.AddItem temp1, 0
'   End If
''   If cboUnit.text = "" Then
''       cboUnit.Clear
''   Else
''       temp2 = cboUnit.text
''       cboUnit.Clear
''   End If
''
''    Get all the unoccupied unit numbers and put in cboUnit.
''    Set the RDO Conn1ection to the dataset
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseIfNeeded
'   Conn1.EstablishConnection rdDriverNoPrompt
''
'   SQLStr1 = "SELECT UnitNumber FROM Units WHERE Occupied = 'N' ORDER BY UnitNumber"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
''
'       If Rst1.EOF = False Then
'           While Rst1.EOF = False
'               cboUnit.AddItem Rst1!UnitNumber
'               Rst1.MoveNext
'           Wend
'       End If
''
'   Rst1.Close
'   Conn1.Close
'
'   If temp2 <> "" Then
'       cboUnit.text = temp2
'       cboUnit.AddItem temp2, 0
'   End If
'
'   Exit Sub
'
'ErrTrap1:
'       If ERR.Number = 40002 Then
'           If MsgBox("DSN - " & Sdsn & " not found. Please check with your system adminstrator.", vbCritical, "DSN Set Up Error") = vbRetry Then
''                Resume
''            Else
'               Resume Next
'           End If
'       Else
'           If ERR.Number = 40009 Then Resume Next
'           If ERR.Number <> 0 Then
'               MsgBox ERR.Number & " - " & ERR.description
'               Resume Next
'           End If
'       End If
'
End Sub

Public Sub CancelAddEdit()

   cmdAdd.Visible = True
   cmdDelete.Visible = True
   cmdSaveNew.Visible = False
   cmdCancelNew.Visible = False
   cmdEdit.Visible = True
   cmdSaveEdit.Visible = False
   cmdCancelEdit.Visible = False
'    mnuEdit.Enabled = True
'    mnuAddNew.Enabled = True
'    mnuDel.Enabled = True
   cmdPrint.Visible = True

End Sub

Public Sub ResetScreen()
'   Dim temp
'
'   cboTenants.Enabled = True
'
'   If cboTenants.text = "" Then
'       cboTenants.Clear
'   Else
'       temp = cboTenants.text
'       cboTenants.Clear
'       cboTenants.text = temp
'   End If
'
'   'Reset screen to show all Company Names and Sage Account numer in cboTenants
'   ' Set the RDO Conn1ection to the dataset
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseIfNeeded
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   SQLStr1 = "SELECT SageAccountNumber, CompanyName FROM Tenants ORDER BY SageAccountNumber"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'       If Rst1.EOF = False Then
'           While Rst1.EOF = False
'               cboTenants.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
'               Rst1.MoveNext
'           Wend
'       End If
'
'   Rst1.Close
'   Conn1.Close
''
'   cmdAdd.Visible = True
'   cmdDelete.Visible = True
'   cmdSaveNew.Visible = False
'   cmdCancelNew.Visible = False
'   cmdEdit.Visible = True
'   cmdSaveEdit.Visible = False
'   cmdCancelEdit.Visible = False
''    mnuEdit.Enabled = True
''    mnuAddNew.Enabled = True
''    mnuDel.Enabled = True
'   cmdPrint.Visible = True
End Sub

Private Sub gridBankCode_Click()
   txtBankCode.text = gridBankCode.TextMatrix(gridBankCode.Row, 0)
   gridBankCode.Visible = False
End Sub

Private Sub Image1_Click()
   fraTenant.Visible = False
End Sub

Private Sub Label19_Click()

End Sub

Private Sub optBoth_Click()
   Form_Load
End Sub

Private Sub optCurrentTenant_Click()
   Form_Load
End Sub

Private Sub optExTenant_Click()
   Form_Load
End Sub

Private Sub BankAccount()
   ' Error Handler
   On Error GoTo Error_Handler
   
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
   Else
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

