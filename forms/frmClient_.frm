VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kingsgate - Client/Landlord"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   FillColor       =   &H00C0C000&
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11550
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSearchResult 
      Height          =   2175
      Left            =   5760
      TabIndex        =   185
      Top             =   720
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777194
      Cols            =   3
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3720
      TabIndex        =   170
      Top             =   120
      Width           =   4335
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
         TabIndex        =   186
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   255
         Left            =   1560
         TabIndex        =   178
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   285
         Left            =   1560
         TabIndex        =   177
         Top             =   120
         Width           =   2400
      End
      Begin VB.TextBox txtSageCust 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   176
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtSageSupp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   175
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   174
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtAdd3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   173
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtHomePC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   172
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFEA&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   171
         Top             =   1185
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client/Landlord ID:"
         Height          =   195
         Left            =   80
         TabIndex        =   184
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   75
         TabIndex        =   183
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Customer A/C:"
         Height          =   195
         Left            =   75
         TabIndex        =   182
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Supplier A/C:"
         Height          =   195
         Left            =   75
         TabIndex        =   181
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3480
         MouseIcon       =   "frmClient.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   180
         Top             =   1965
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   179
         Top             =   840
         Width           =   615
      End
   End
   Begin TabDlg.SSTab tabClient 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   16768960
      TabCaption(0)   =   "&Client"
      TabPicture(0)   =   "frmClient.frx":0BD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtClinetNote"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraOfficeAdd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Property"
      TabPicture(1)   =   "frmClient.frx":0BF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame11"
      Tab(1).Control(1)=   "imgList"
      Tab(1).Control(2)=   "fraType"
      Tab(1).Control(3)=   "fraOccupied"
      Tab(1).Control(4)=   "Frame4"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Agreement"
      TabPicture(2)   =   "frmClient.frx":0C0C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCommission"
      Tab(2).Control(1)=   "fraAgreement"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(3)=   "Frame16"
      Tab(2).Control(4)=   "dtDate"
      Tab(2).Control(5)=   "Frame17"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Statements"
      TabPicture(3)   =   "frmClient.frx":0C28
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "MSHFlexGrid1"
      Tab(3).Control(2)=   "Text14"
      Tab(3).Control(3)=   "Command5"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Account History"
      TabPicture(4)   =   "frmClient.frx":0C44
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSHFlexGrid2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Bank Details"
      TabPicture(5)   =   "frmClient.frx":0C60
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).Control(1)=   "Frame14"
      Tab(5).Control(2)=   "Frame15"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Add &New"
      TabPicture(6)   =   "frmClient.frx":0C7C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "tabAddNewClient"
      Tab(6).ControlCount=   1
      Begin VB.Frame fraOfficeAdd 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Details Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   4095
         Left            =   4320
         TabIndex        =   187
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtHomePh 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   198
            Top             =   3165
            Width           =   3000
         End
         Begin VB.TextBox txtPerEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   197
            Top             =   3435
            Width           =   3000
         End
         Begin VB.TextBox txtMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   196
            Top             =   2880
            Width           =   3000
         End
         Begin VB.TextBox txtOffAdd3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   195
            Top             =   1140
            Width           =   3000
         End
         Begin VB.TextBox txtOffAdd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   194
            Top             =   860
            Width           =   3000
         End
         Begin VB.TextBox txtOffEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   193
            Top             =   2080
            Width           =   3000
         End
         Begin VB.TextBox txtOffice 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1300
            TabIndex        =   192
            Top             =   240
            Width           =   3000
         End
         Begin VB.TextBox txtOffAdd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1300
            TabIndex        =   191
            Top             =   600
            Width           =   3000
         End
         Begin VB.TextBox txtOffPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1300
            TabIndex        =   190
            Top             =   1420
            Width           =   1455
         End
         Begin VB.TextBox txtOffPh 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1320
            TabIndex        =   189
            Top             =   1800
            Width           =   3000
         End
         Begin VB.TextBox txtOffPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1300
            TabIndex        =   188
            Top             =   2440
            Width           =   3000
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3120
            MouseIcon       =   "frmClient.frx":0C98
            MousePointer    =   99  'Custom
            TabIndex        =   208
            Top             =   3795
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            Height          =   195
            Left            =   120
            TabIndex        =   207
            Top             =   3165
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            Height          =   195
            Left            =   120
            TabIndex        =   206
            Top             =   3435
            Width           =   885
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   120
            TabIndex        =   205
            Top             =   2880
            Width           =   510
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   204
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lblHomeAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3960
            MouseIcon       =   "frmClient.frx":0FA2
            MousePointer    =   99  'Custom
            TabIndex        =   203
            Top             =   3800
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            Height          =   195
            Left            =   75
            TabIndex        =   202
            Top             =   2080
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   201
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            Height          =   195
            Left            =   75
            TabIndex        =   200
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position:"
            Height          =   195
            Left            =   75
            TabIndex        =   199
            Top             =   2440
            Width           =   600
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   169
         Top             =   1440
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CFFFD0&
         Caption         =   "Images"
         Enabled         =   0   'False
         Height          =   855
         Left            =   -67680
         TabIndex        =   150
         Top             =   3000
         Width           =   2295
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1000
            MousePointer    =   99  'Custom
            TabIndex        =   153
            Top             =   300
            Width           =   255
         End
         Begin VB.Label lblImage 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Next >>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1440
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblImagePreLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "<< Pre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   120
            MousePointer    =   99  'Custom
            TabIndex        =   151
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame fraOccupied 
         BackColor       =   &H00CFFFD0&
         Caption         =   "Occupied:"
         Height          =   2550
         Left            =   -71520
         TabIndex        =   141
         Top             =   1920
         Width           =   3735
         Begin VB.TextBox txtPreRentRvw 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtPreTenancyType 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedTo 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedFr 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblTenantNameLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            MouseIcon       =   "frmClient.frx":12AC
            MousePointer    =   99  'Custom
            TabIndex        =   157
            Top             =   2205
            Width           =   1095
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   2200
            Width           =   1020
         End
         Begin VB.Label lblTenantIDLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            MouseIcon       =   "frmClient.frx":15B6
            MousePointer    =   99  'Custom
            TabIndex        =   155
            Top             =   1845
            Width           =   810
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   154
            Top             =   1850
            Width           =   765
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   147
            Top             =   1440
            Width           =   1365
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   146
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Occupied To:"
            Height          =   195
            Left            =   120
            TabIndex        =   143
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Occupied From:"
            Height          =   195
            Left            =   120
            TabIndex        =   142
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame fraType 
         BackColor       =   &H00CFFFD0&
         Caption         =   "LANDLORD"
         Height          =   1575
         Left            =   -71500
         TabIndex        =   133
         Top             =   360
         Width           =   3720
         Begin VB.TextBox txtTVInfoName 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd1 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoPC 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   1220
            Width           =   1455
         End
         Begin VB.TextBox txtTVInfoAdd2 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   705
            Width           =   2655
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Left            =   80
            TabIndex        =   140
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   80
            TabIndex        =   138
            Top             =   480
            Width           =   615
         End
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   -74880
         Top             =   4680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":18C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":219A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2A74
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Generate Client Statement"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74160
         TabIndex        =   129
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00AFE5EA&
         Caption         =   "Attactment Files:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   125
         Top             =   2640
         Width           =   4935
         Begin VB.FileListBox File1 
            Height          =   1455
            Left            =   120
            TabIndex        =   132
            Top             =   720
            Width           =   3015
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open Attachment"
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox cmbFiles 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add Attachment"
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   840
            Width           =   1455
         End
      End
      Begin MSComCtl2.MonthView dtDate 
         Height          =   2370
         Left            =   -68520
         TabIndex        =   124
         Top             =   1920
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16768960
         Appearance      =   1
         StartOfWeek     =   20643842
         CurrentDate     =   38637
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00AFE5EA&
         Caption         =   "Basis Gross Rent Payments (BGRP):"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   1335
         Left            =   -69840
         TabIndex        =   121
         Top             =   360
         Width           =   4455
         Begin VB.OptionButton optBasisReceivable 
            BackColor       =   &H00AFE5EA&
            Caption         =   "Rent Receivable"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optBasisReceived 
            BackColor       =   &H00AFE5EA&
            Caption         =   "Rent Received"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Bank Details:"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtBankAdd1 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1140
            Width           =   3255
         End
         Begin VB.TextBox txtBankAdd2 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1500
            Width           =   3255
         End
         Begin VB.TextBox txtBankAdd3 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   1860
            Width           =   3255
         End
         Begin VB.TextBox txtBankName 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtBankPC 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank/Building Socity:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   660
            Width           =   1530
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   2340
            Width           =   780
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Account Details:"
         Height          =   2055
         Left            =   -69240
         TabIndex        =   71
         Top             =   960
         Width           =   3855
         Begin VB.TextBox txtBankAccount 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtBankSC 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Left            =   240
            TabIndex        =   75
            Top             =   960
            Width           =   750
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   600
            Width           =   1245
         End
      End
      Begin TabDlg.SSTab tabAddNewClient 
         Height          =   4125
         Left            =   -74920
         TabIndex        =   29
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7276
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   -2147483632
         TabCaption(0)   =   "Client"
         TabPicture(0)   =   "frmClient.frx":334E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label28"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdNewNext(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtAddNewNote"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Contact Address"
         TabPicture(1)   =   "frmClient.frx":336A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdNewBack(0)"
         Tab(1).Control(1)=   "cmdNewNext(1)"
         Tab(1).Control(2)=   "Frame13"
         Tab(1).Control(3)=   "Frame9"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Bank Details"
         TabPicture(2)   =   "frmClient.frx":3386
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraAddNewBankInfo"
         Tab(2).Control(1)=   "fraAccDetails"
         Tab(2).Control(2)=   "cmdNewBack(1)"
         Tab(2).Control(3)=   "fraBankDetails"
         Tab(2).Control(4)=   "cmdNewNext(2)"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Agreement"
         TabPicture(3)   =   "frmClient.frx":33A2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdNewSave"
         Tab(3).Control(1)=   "cmdNewBack(2)"
         Tab(3).ControlCount=   2
         Begin VB.CommandButton cmdNewBack 
            Caption         =   "<< &Back"
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
            Index           =   2
            Left            =   -68760
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   3640
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewSave 
            Caption         =   "&Save"
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
            Left            =   -66960
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   3640
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewNext 
            Caption         =   "&Next >>"
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
            Index           =   2
            Left            =   -66960
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   3640
            Width           =   1455
         End
         Begin VB.Frame fraBankDetails 
            Caption         =   "Bank Details:"
            Height          =   2895
            Left            =   -74880
            TabIndex        =   96
            Top             =   480
            Width           =   5295
            Begin VB.TextBox txtBankNewPC 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   2340
               Width           =   1335
            End
            Begin VB.TextBox txtBankNewAdd3 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   1740
               Width           =   3255
            End
            Begin VB.TextBox txtBankNewAdd2 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   88
               Top             =   1380
               Width           =   3255
            End
            Begin VB.TextBox txtBankNewAdd1 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   87
               Top             =   1020
               Width           =   3255
            End
            Begin MSForms.ComboBox cboBankList 
               Height          =   300
               Left            =   1800
               TabIndex        =   86
               Top             =   480
               Width           =   3255
               VariousPropertyBits=   746604571
               BackColor       =   16777194
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5741;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
               Height          =   195
               Left            =   120
               TabIndex        =   99
               Top             =   2340
               Width           =   780
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   120
               TabIndex        =   98
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank/Building Socity:"
               Height          =   195
               Left            =   120
               TabIndex        =   97
               Top             =   480
               Width           =   1530
            End
         End
         Begin VB.CommandButton cmdNewBack 
            Caption         =   "<< &Back"
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
            Index           =   1
            Left            =   -68760
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   3640
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewBack 
            BackColor       =   &H00CFFFD0&
            Caption         =   "<< &Back"
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
            Index           =   0
            Left            =   -68760
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   3640
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewNext 
            BackColor       =   &H00CFFFD0&
            Caption         =   "&Next >>"
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
            Index           =   1
            Left            =   -66960
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   3640
            Width           =   1455
         End
         Begin VB.Frame fraAccDetails 
            Caption         =   "Account Details:"
            Height          =   1695
            Left            =   -69360
            TabIndex        =   85
            Top             =   480
            Width           =   3855
            Begin VB.TextBox txtAddNewBankSC 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   93
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtAddNewBankAcc 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   91
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Account Number:"
               Height          =   195
               Left            =   240
               TabIndex        =   95
               Top             =   600
               Width           =   1245
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sort Code:"
               Height          =   195
               Left            =   240
               TabIndex        =   92
               Top             =   960
               Width           =   750
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00CFFFD0&
            Caption         =   "Home:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3255
            Left            =   -74880
            TabIndex        =   65
            Top             =   360
            Width           =   4575
            Begin VB.TextBox txtNewHomeAdd1 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1365
               TabIndex        =   43
               Top             =   240
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeAdd3 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1365
               TabIndex        =   45
               Top             =   960
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomePC 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1365
               TabIndex        =   46
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtNewHomeAdd2 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1365
               TabIndex        =   44
               Top             =   600
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeEmail 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   49
               Top             =   2640
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeTel 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   47
               Top             =   1920
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeMob 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   48
               Top             =   2280
               Width           =   3000
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
               Height          =   195
               Left            =   120
               TabIndex        =   69
               Top             =   1440
               Width           =   780
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Personal Email:"
               Height          =   195
               Left            =   120
               TabIndex        =   68
               Top             =   2640
               Width           =   1080
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Tel:"
               Height          =   195
               Left            =   120
               TabIndex        =   67
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile:"
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   2280
               Width           =   510
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00CFFFD0&
            Caption         =   "Office:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3255
            Left            =   -70080
            TabIndex        =   42
            Top             =   360
            Width           =   4575
            Begin VB.TextBox txtNewOffAdd3 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1320
               TabIndex        =   53
               Top             =   1320
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAdd2 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   52
               Top             =   960
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffEmail 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1320
               TabIndex        =   57
               Top             =   2760
               Width           =   3000
            End
            Begin VB.TextBox txtNewOff 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   50
               Top             =   240
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAdd1 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   51
               Top             =   600
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAddPC 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   54
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox txtNewOffTel 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   55
               Top             =   2040
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffPos 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   56
               Top             =   2400
               Width           =   3000
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Email:"
               Height          =   195
               Left            =   75
               TabIndex        =   64
               Top             =   2760
               Width           =   885
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   75
               TabIndex        =   63
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code"
               Height          =   195
               Left            =   75
               TabIndex        =   62
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Telephone:"
               Height          =   195
               Left            =   75
               TabIndex        =   61
               Top             =   2040
               Width           =   810
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Position:"
               Height          =   195
               Left            =   75
               TabIndex        =   60
               Top             =   2400
               Width           =   600
            End
         End
         Begin VB.TextBox txtAddNewNote 
            BackColor       =   &H00FFFFEA&
            Height          =   2415
            Left            =   5280
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   960
            Width           =   4215
         End
         Begin VB.CommandButton cmdNewNext 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Next >>"
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
            Index           =   0
            Left            =   8040
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   3640
            Width           =   1455
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   80
            TabIndex        =   30
            Top             =   960
            Width           =   5175
            Begin VB.ComboBox cboSageSupAcc 
               BackColor       =   &H00FFFFEA&
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   720
               Width           =   3550
            End
            Begin VB.ComboBox cboSageCustAcc 
               BackColor       =   &H00FFFFEA&
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   120
               Width           =   3550
            End
            Begin VB.TextBox txtNewClinetName 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1530
               TabIndex        =   34
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox txtNewClientID 
               BackColor       =   &H00FFFFEA&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1530
               TabIndex        =   33
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label27 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Client/Landlord ID:"
               Height          =   195
               Left            =   75
               TabIndex        =   40
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               Height          =   195
               Left            =   75
               TabIndex        =   39
               Top             =   1920
               Width           =   465
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Customer A/C:"
               Height          =   195
               Left            =   75
               TabIndex        =   38
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Supplier A/C:"
               Height          =   195
               Left            =   75
               TabIndex        =   37
               Top             =   720
               Width           =   1365
            End
         End
         Begin VB.Frame fraAddNewBankInfo 
            BackColor       =   &H80000018&
            Caption         =   "Add New Bank Info:"
            Height          =   3495
            Left            =   -74880
            TabIndex        =   100
            Top             =   480
            Visible         =   0   'False
            Width           =   5295
            Begin VB.CommandButton cmdAddNewBankCancel 
               Caption         =   "&Cancel"
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
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   113
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox txtAddNewBankID 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   101
               Top             =   480
               Width           =   1335
            End
            Begin VB.CommandButton cmdBankInfoSave 
               Caption         =   "&Save"
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
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox txtAddNewBankPC 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   106
               Top             =   2580
               Width           =   1335
            End
            Begin VB.TextBox txtAddNewBankAdd3 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   105
               Top             =   2100
               Width           =   3255
            End
            Begin VB.TextBox txtAddNewBankAdd2 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   104
               Top             =   1740
               Width           =   3255
            End
            Begin VB.TextBox txtAddNewBankAdd1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   103
               Top             =   1380
               Width           =   3255
            End
            Begin VB.TextBox txtAddNewBankName 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   102
               Top             =   840
               Width           =   3255
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "To create a new Bank info you need to enter bank name and unique ID."
               BeginProperty Font 
                  Name            =   "Garamond"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   120
               TabIndex        =   112
               Top             =   240
               Width           =   4815
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Socity ID:"
               Height          =   195
               Left            =   120
               TabIndex        =   111
               Top             =   480
               Width           =   1110
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   2580
               Width           =   780
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   120
               TabIndex        =   109
               Top             =   1380
               Width           =   615
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank/Building Socity:"
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   840
               Width           =   1530
            End
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Note:"
            Height          =   195
            Left            =   5280
            TabIndex        =   41
            Top             =   600
            Width           =   390
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1200
         TabIndex        =   26
         Top             =   3840
         Width           =   3255
         Begin VB.CommandButton cmdNoteSave 
            Caption         =   "&Update"
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
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdNoteEdit 
            Caption         =   "&Edit"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox txtClinetNote 
         BackColor       =   &H00FFFFEA&
         Height          =   1335
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00CFFFD0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -68760
         TabIndex        =   22
         Top             =   4560
         Width           =   3375
         Begin VB.CommandButton cmdContactEdit 
            Caption         =   "&Edit"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdContactSave 
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00AFE5EA&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -68760
         TabIndex        =   19
         Top             =   5160
         Width           =   3375
         Begin VB.CommandButton cmdAgSave 
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdAgEdit 
            Caption         =   "&Edit"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame fraAgreement 
         BackColor       =   &H00AFE5EA&
         Caption         =   "Dates:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   2535
         Left            =   -69840
         TabIndex        =   12
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox txtBGRPDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   2160
            Width           =   1600
         End
         Begin VB.TextBox txtAggNoticeDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   1800
            Width           =   1600
         End
         Begin VB.TextBox txtAggReviewDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   1440
            Width           =   1600
         End
         Begin VB.TextBox txtAggDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   1600
         End
         Begin VB.TextBox txtAggStDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   1600
         End
         Begin VB.TextBox txtAggEndDt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1080
            Width           =   1600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BGRP Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   130
            Top             =   2160
            Width           =   885
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Notice Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   118
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Review Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   117
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   15
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   14
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   720
            Width           =   765
         End
      End
      Begin VB.Frame fraCommission 
         BackColor       =   &H00AFE5EA&
         Caption         =   "Commission/Fees:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   4935
         Begin VB.Frame Frame5 
            Caption         =   "Letting Fees"
            Height          =   975
            Left            =   120
            TabIndex        =   161
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton optComPercReceivable 
               BackColor       =   &H00AFE5EA&
               Caption         =   "% of Rent Receivable"
               Height          =   255
               Left            =   120
               TabIndex        =   166
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton optFixed 
               BackColor       =   &H00AFE5EA&
               Caption         =   "Fixed Amount "
               Height          =   255
               Left            =   120
               TabIndex        =   163
               Top             =   600
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.TextBox txtCommissionAmt 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   162
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount:"
               Height          =   195
               Left            =   2400
               TabIndex        =   164
               Top             =   480
               Width           =   585
            End
         End
         Begin VB.OptionButton optPercReveived 
            BackColor       =   &H00AFE5EA&
            Caption         =   "% of Rent Received"
            Height          =   255
            Left            =   360
            TabIndex        =   165
            Top             =   480
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Caption         =   "Management Fees"
            Height          =   975
            Left            =   120
            TabIndex        =   158
            Top             =   1200
            Width           =   4455
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   168
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Fixed Fee"
               Height          =   375
               Left            =   120
               TabIndex        =   160
               Top             =   480
               Width           =   1815
            End
            Begin VB.OptionButton Option1 
               Caption         =   "% of Rent Received"
               Height          =   255
               Left            =   120
               TabIndex        =   159
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount:"
               Height          =   195
               Left            =   2400
               TabIndex        =   167
               Top             =   360
               Width           =   585
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -68760
         TabIndex        =   8
         Top             =   4560
         Width           =   3375
         Begin VB.CommandButton Command12 
            Caption         =   "&Edit"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command11 
            Caption         =   "&Update"
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   -68640
         TabIndex        =   7
         Text            =   "Text14"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Gross Rent payable"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70800
         TabIndex        =   5
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   6240
         TabIndex        =   2
         Top             =   4560
         Width           =   3375
         Begin VB.CommandButton cmdDelClient 
            Caption         =   "Delete"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reprint Statement"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      TabIndex        =   0
      Top             =   8520
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvwLandLord 
      Height          =   4095
      Left            =   120
      TabIndex        =   209
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7223
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image imgPremises 
      Height          =   3375
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szTextBox As Control
Dim szAddNewBankID As String
Dim bLoadAggrements As Boolean

Private Sub cboBankList_Click()
   Dim aBank() As String
   
   If cboBankList.text = "ADD NEW BANK" Then
      fraAddNewBankInfo.Top = fraBankDetails.Top
      fraAddNewBankInfo.Left = fraBankDetails.Left
      fraAddNewBankInfo.Visible = True
      fraAddNewBankInfo.ZOrder 0
      txtAddNewBankID.SetFocus
      fraAccDetails.Enabled = False
      cmdNewBack(1).Enabled = False
      cmdNewSave.Enabled = False

      Exit Sub
   End If
   
   aBank = Split(cboBankList.text, " / ")
   cboBankList.text = aBank(0)
   txtBankNewPC.text = aBank(1)
   szAddNewBankID = aBank(2)
   
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT BANK_ADDRESS1,BANK_ADDRESS2,BANK_ADDRESS3 " & _
           "FROM TLBBANK " & _
           "WHERE BANK_ID='" & aBank(2) & "';"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   
   txtBankNewAdd1.text = rstBank!BANK_ADDRESS1
   txtBankNewAdd2.text = rstBank!BANK_ADDRESS2
   txtBankNewAdd3.text = rstBank!BANK_ADDRESS3
   
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   
End Sub

Private Sub cmdAddNewBankCancel_Click()
   fraAddNewBankInfo.Visible = False
   txtAddNewBankID.text = ""
   txtAddNewBankName.text = ""
   txtAddNewBankAdd1.text = ""
   txtAddNewBankAdd2.text = ""
   txtAddNewBankAdd3.text = ""
   txtAddNewBankPC.text = ""
   fraAccDetails.Enabled = True
   cmdNewBack(1).Enabled = True
   cmdNewSave.Enabled = True
End Sub

Private Sub cmdAgEdit_Click()
   fraAgreement.Enabled = True
   fraCommission.Enabled = True
   cmdAgSave.Enabled = True
   cmdAgEdit.Enabled = False
End Sub

Private Sub cmdAgSave_Click()
   fraAgreement.Enabled = False
   fraCommission.Enabled = False
   cmdAgSave.Enabled = False
   cmdAgEdit.Enabled = True
End Sub

Private Sub cmdBankInfoSave_Click()
   If txtAddNewBankName.text = "" Then
      MsgBox "Bank name can't be empty.", vbCritical, "Bank Name Error"
      Exit Sub
   End If
   If txtAddNewBankAdd1.text = "" Then
      MsgBox "Please Provide first line of Bank Address.", vbCritical, "Bank Address Error"
      Exit Sub
   End If
   If txtAddNewBankPC.text = "" Then
      MsgBox "Bank Post Code can't be empty.", vbCritical, "Bank Post Code Error"
      Exit Sub
   End If

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM tlbBank;"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   With rstBank
      .AddNew
      !BANK_ID = txtAddNewBankID.text
      !BANK_NAME = txtAddNewBankName.text
      !BANK_ADDRESS1 = txtAddNewBankAdd1.text
      !BANK_ADDRESS2 = txtAddNewBankAdd2.text
      !BANK_ADDRESS3 = txtAddNewBankAdd3.text
      !BANK_POST_CODE = txtAddNewBankPC.text
      .Update
      .Close
   End With
   Set rstBank = Nothing
   conBank.Close
   Set conBank = Nothing
   
   cboBankList.RemoveItem cboBankList.ListCount - 1
   cboBankList.AddItem txtAddNewBankName.text & " / " & txtAddNewBankPC.text & " / " & txtAddNewBankID.text
   cboBankList.text = txtAddNewBankName.text
   txtBankNewAdd1.text = txtAddNewBankAdd1.text
   txtBankNewAdd2.text = txtAddNewBankAdd2.text
   txtBankNewAdd3.text = txtAddNewBankAdd3.text
   txtBankNewPC.text = txtAddNewBankPC.text

   fraAddNewBankInfo.Visible = False
   fraAccDetails.Enabled = True
   cmdNewBack(1).Enabled = True
   cmdNewSave.Enabled = True

   MsgBox "Bank information has been save successfully", vbOKOnly, "Success"
End Sub

Private Sub cmdClientID_Click()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT * " & _
           "FROM CLIENT, TLBBANK " & _
           "WHERE CLIENTID = '" & Trim(txtClientID.text) & "' AND " & _
               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then
      MsgBox "No such ID found in the Database", vbCritical, "Error"
      rstClient.Close
      conClient.Close
      Set rstClient = Nothing
      Set conClient = Nothing
      Exit Sub
   End If

   txtClientName.text = rstClient!ClientName
   txtSageCust.text = rstClient!LandLordSageCustAC
   txtSageSupp.text = rstClient!LandLordSageSuppAC
   txtAdd1.text = rstClient!ClientAddressLine1
   txtAdd2.text = rstClient!ClientAddressLine2
   txtAdd3.text = rstClient!ClientAddressLine3
   txtHomePC.text = rstClient!ClientPostCode
   txtHomePh.text = rstClient!ClientHomeTel
   txtMobile.text = rstClient!ClientMobile
   txtPerEmail.text = rstClient!ClientPersonalEmail
   txtOffice.text = rstClient!ClientOffice
   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
   txtOffPC.text = rstClient!ClientOfficePostCode
   txtOffPh.text = rstClient!ClientOfficeTel
   txtOffPos.text = rstClient!ClientOfficePos
   txtOffEmail.text = rstClient!ClientOfficeEmail
   txtClinetNote.text = rstClient!Note
   
   txtBankName.text = rstClient!BANK_NAME
   txtBankAdd1.text = rstClient!BANK_ADDRESS1
   txtBankAdd2.text = rstClient!BANK_ADDRESS2
   txtBankAdd3.text = rstClient!BANK_ADDRESS3
   txtBankPC.text = rstClient!BANK_POST_CODE
   txtBankAccount.text = rstClient!BANK_AC_NUM
   txtBankSC.text = rstClient!BANK_SC
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

Private Sub cmdClientName_Click()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT * " & _
           "FROM CLIENT, TLBBANK " & _
           "WHERE CLIENTNAME = '" & Trim(txtClientName.text) & "' AND " & _
               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then
      MsgBox "No such ID found in the Database", vbCritical, "Error"
      rstClient.Close
      conClient.Close
      Set rstClient = Nothing
      Set conClient = Nothing
      Exit Sub
   End If

   txtClientID.text = rstClient!ClientID
   txtSageCust.text = rstClient!LandLordSageCustAC
   txtSageSupp.text = rstClient!LandLordSageSuppAC
   txtAdd1.text = rstClient!ClientAddressLine1
   txtAdd2.text = rstClient!ClientAddressLine2
   txtAdd3.text = rstClient!ClientAddressLine3
   txtHomePC.text = rstClient!ClientPostCode
   txtHomePh.text = rstClient!ClientHomeTel
   txtMobile.text = rstClient!ClientMobile
   txtPerEmail.text = rstClient!ClientPersonalEmail
   txtOffice.text = rstClient!ClientOffice
   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
   txtOffPC.text = rstClient!ClientOfficePostCode
   txtOffPh.text = rstClient!ClientOfficeTel
   txtOffPos.text = rstClient!ClientOfficePos
   txtOffEmail.text = rstClient!ClientOfficeEmail
   txtClinetNote.text = rstClient!Note
   
   txtBankName.text = rstClient!BANK_NAME
   txtBankAdd1.text = rstClient!BANK_ADDRESS1
   txtBankAdd2.text = rstClient!BANK_ADDRESS2
   txtBankAdd3.text = rstClient!BANK_ADDRESS3
   txtBankPC.text = rstClient!BANK_POST_CODE
   txtBankAccount.text = rstClient!BANK_AC_NUM
   txtBankSC.text = rstClient!BANK_SC
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

Private Sub cmdClinetAddAtch_Click()
   Dim szImageFileName As String
   szImageFileName = DoBrowse()
   If szImageFileName <> "NONE" Then
      If DoStoreInDB(szImageFileName) Then
         MsgBox "SUCCESSFUL"
      End If
   Else
      Exit Sub
   End If
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
      .Fields("PREMISIS_ID") = txtClientID.text
      .Fields("PREMISIS_TYPE") = "CLIENT"
      .Fields("IMAGE_NAME") = cmbFiles.text
      .Update         'update the table
    End With
    szStr = "SELECT * " & _
            "FROM TLBIMAGES " & _
            "WHERE IMAGE_NAME='" & cmbFiles.text & "' AND " & _
            "TLBIMAGES.PREMISIS_ID='" & txtClientID.text & "';"
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

Private Sub cmdClose_Click()
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub cmdContactEdit_Click()
   cmdContactEdit.Enabled = False
   cmdContactSave.Enabled = True
End Sub

Private Sub cmdContactSave_Click()
   
   
   'At the end
   cmdContactEdit.Enabled = True
   cmdContactSave.Enabled = False
End Sub

Private Sub cmdNewBack_Click(index As Integer)
      tabAddNewClient.Tab = index
End Sub

Private Sub cmdNewNext_Click(index As Integer)
   tabAddNewClient.Tab = index + 1
End Sub

Private Sub cmdNewSave_Click()
   If cboBankList.text = "" Then
      MsgBox "Please choose a Bank.", vbCritical, "Bank Info Error"
      cboBankList.SetFocus
      Exit Sub
   End If
   If txtAddNewBankAcc.text = "" Then
      If MsgBox("Do you want to leave Bank Account number Blank?", vbYesNo + vbQuestion, "Bank Account") = vbNo Then
         txtAddNewBankAcc.SetFocus
         Exit Sub
      End If
   End If
   If txtAddNewBankSC.text = "" Then
      If MsgBox("Do you want to leave Bank Sort Code number Blank?", vbYesNo + vbQuestion, "Bank Account") = vbNo Then
         txtAddNewBankSC.SetFocus
         Exit Sub
      End If
   End If

   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt
   
   szSQL = "SELECT * " & _
           "FROM CLIENT;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   With rstClient
      .AddNew
      !ClientID = txtNewClientID.text
      !ClientName = txtNewClinetName.text
      !ClientAddressLine1 = txtNewHomeAdd1.text
      !ClientAddressLine2 = txtNewHomeAdd2.text
      !ClientAddressLine3 = txtNewHomeAdd3.text
      !ClientPostCode = txtNewHomePC.text
      !ClientOfficeEmail = txtNewOffEmail.text
      !ClientPersonalEmail = txtNewHomeEmail.text
      !ClientHomeTel = txtNewHomeTel.text
      !ClientMobile = txtNewHomeMob.text
      !ClientOffice = txtNewOff.text
      !ClientOfficeAddressLine1 = txtNewOffAdd1.text
      !ClientOfficeAddressLine2 = txtNewOffAdd2.text
      !ClientOfficeAddressLine3 = txtNewOffAdd3.text
      !ClientOfficePostCode = txtNewOffAddPC.text
      !ClientOfficeTel = txtNewOffTel.text
      !ClientOfficePos = txtNewOffPos.text
      !LandLordSageCustAC = cboSageCustAcc.text
      !LandLordSageSuppAC = cboSageSupAcc.text
      !Note = txtAddNewNote.text
      !BANK_ID = szAddNewBankID
      !BANK_AC_NUM = txtAddNewBankAcc.text
      !BANK_SC = txtAddNewBankSC.text
      !Files = "C:\Documents and Settings\malcolm.DOMAIN\Desktop\Prestige Property Management Software.pdf"

      .Update
      .Close
   End With
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing

   MsgBox "New customer informations have been entered successfully", vbOKOnly, "Success"

   txtNewClientID.text = ""
   txtNewClinetName.text = ""
   txtAddNewNote.text = ""
   txtNewHomeAdd1.text = ""
   txtNewHomeAdd2.text = ""
   txtNewHomeAdd3.text = ""
   txtNewHomePC.text = ""
   txtNewHomeTel.text = ""
   txtNewHomeMob.text = ""
   txtNewHomeEmail.text = ""
   txtNewOff.text = ""
   txtNewOffAdd1.text = ""
   txtNewOffAdd2.text = ""
   txtNewOffAdd3.text = ""
   txtNewOffAddPC.text = ""
   txtNewOffTel.text = ""
   txtNewOffPos.text = ""
   txtNewOffEmail.text = ""
   cboBankList.text = ""
   txtBankNewAdd1.text = ""
   txtBankNewAdd2.text = ""
   txtBankNewAdd3.text = ""
   txtBankNewPC.text = ""
   txtAddNewBankAcc.text = ""
   txtAddNewBankSC.text = ""
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then
       MsgBox "Select a file from list."
   Else
       OpenFile
   End If
End Sub

Private Sub cmdSearch_Click()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'   If optSearchID.Value Then
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTID LIKE '" & Trim(txtSearch.text) & "%'"
'   ElseIf optSearchName.Value Then
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTNAME LIKE '%" & Trim(txtSearch.text) & "%'"
'   Else
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CILENTPOSTCODE LIKE '" & Trim(txtSearch.text) & "'"
'   End If

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then
      MsgBox "No search result found in the Database", vbCritical, "Error"
      rstClient.Close
      conClient.Close
      Set rstClient = Nothing
      Set conClient = Nothing
      Exit Sub
   End If

   Dim iRow As Integer
   iRow = 1
   
   flxSearchResult.Clear
   flxSearchResult.Rows = 2
   ConfigurFlexGrid
   While Not rstClient.EOF
      flxSearchResult.TextMatrix(iRow, 0) = rstClient!ClientID
      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
      rstClient.MoveNext
      If Not rstClient.EOF Then flxSearchResult.AddItem ""
      iRow = iRow + 1
   Wend

   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
'   cmdSelected.Enabled = True
End Sub

Private Sub cmdSelected_Click()
   ResetFields
   
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT * " & _
           "FROM CLIENT, TLBBANK " & _
           "WHERE CLIENTID = '" & Trim(flxSearchResult.TextMatrix(flxSearchResult.RowSel, 0)) & "' AND " & _
               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   txtClientID.text = flxSearchResult.TextMatrix(flxSearchResult.RowSel, 0)
   txtClientName.text = rstClient!ClientName
   txtSageCust.text = rstClient!LandLordSageCustAC
   txtSageSupp.text = rstClient!LandLordSageSuppAC
   txtAdd1.text = rstClient!ClientAddressLine1
   txtAdd2.text = rstClient!ClientAddressLine2
   txtAdd3.text = rstClient!ClientAddressLine3
   txtHomePC.text = rstClient!ClientPostCode
   txtHomePh.text = rstClient!ClientHomeTel
   txtMobile.text = rstClient!ClientMobile
   txtPerEmail.text = rstClient!ClientPersonalEmail
   txtOffice.text = rstClient!ClientOffice
   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
   txtOffPC.text = rstClient!ClientOfficePostCode
   txtOffPh.text = rstClient!ClientOfficeTel
   txtOffPos.text = rstClient!ClientOfficePos
   txtOffEmail.text = rstClient!ClientOfficeEmail
   txtClinetNote.text = rstClient!Note
   
   txtBankName.text = rstClient!BANK_NAME
   txtBankAdd1.text = rstClient!BANK_ADDRESS1
   txtBankAdd2.text = rstClient!BANK_ADDRESS2
   txtBankAdd3.text = rstClient!BANK_ADDRESS3
   txtBankPC.text = rstClient!BANK_POST_CODE
   txtBankAccount.text = rstClient!BANK_AC_NUM
   txtBankSC.text = rstClient!BANK_SC
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing

   DrawLandLordTree
End Sub

Private Sub cmdTenantList_Click()
   flxSearchResult.Visible = True
   flxSearchResult.ZOrder 0
End Sub

Private Sub dtDate_DateClick(ByVal DateClicked As Date)
   szTextBox.text = dtDate.Value
   dtDate.Visible = False
End Sub

Private Sub flxSearchResult_DblClick()
   cmdSelected_Click
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50
   bLoadAggrements = False

   tabClient.Tab = 0
   ConfigurFlexGrid
   
   LoadAllClientFlxGrd
   
   PlacedAddressFrame
   
   DrawLandLordTree
End Sub

Private Sub PlacedAddressFrame()
   fraOfficeAdd.Left = flxSearchResult.Left
   fraOfficeAdd.Top = flxSearchResult.Top
End Sub

Private Sub LoadAllClientFlxGrd()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   flxSearchResult.Clear
   flxSearchResult.Rows = 2
   ConfigurFlexGrid
   While Not rstClient.EOF
      flxSearchResult.TextMatrix(iRow, 0) = rstClient!ClientID
      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
      rstClient.MoveNext
      If Not rstClient.EOF Then flxSearchResult.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
'   cmdSelected.Enabled = True
End Sub

Private Sub ConfigurFlexGrid()
   flxSearchResult.ColWidth(0) = 1400
   flxSearchResult.TextMatrix(0, 0) = "Client ID"
   
   flxSearchResult.ColWidth(1) = 2000
   flxSearchResult.TextMatrix(0, 1) = "Client Name"
   
   flxSearchResult.ColWidth(2) = 1550
   flxSearchResult.TextMatrix(0, 2) = "Post Code"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   cmdSelected.Enabled = True
End Sub

Private Sub lblAddress_Click()
   If lblAddress.Caption = "Details >>" Then
      fraOfficeAdd.Visible = True
      fraOfficeAdd.ZOrder 0
      lblAddress.Caption = "Details <<"
   Else
      fraOfficeAdd.Visible = False
      lblAddress.Caption = "Details >>"
   End If
End Sub

Private Sub lblAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblAddress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblHomeAddress_Click()
   fraOfficeAdd.Visible = False
   lblAddress.Caption = "Details >>"
End Sub

Private Sub lblHomeAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblHomeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblHomeAddress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblHomeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblOfficeAddress_Click()
   fraOfficeAdd.Visible = True
   fraOfficeAdd.ZOrder 0
End Sub

Private Sub lblOfficeAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   lblOfficeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblOfficeAddress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   lblOfficeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblSave_Click()
   MsgBox "Under Construction"
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblSave.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblSave.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblTenantIDLink_Click()
   MsgBox "under construction"
End Sub

Private Sub tabAddNewClient_Click(PreviousTab As Integer)
   If tabAddNewClient.Tab = 2 Then
      If cboBankList.ListCount = 0 Then BankAccList
   End If

   If tabAddNewClient.Tab = 1 Or tabAddNewClient.Tab = 2 Then
      If cboSageCustAcc.text = "" Then
         MsgBox "Plseas Choose Sage Customer Account ID."
         tabAddNewClient.Tab = 0
         cboSageCustAcc.SetFocus
         Exit Sub
      End If
      If cboSageSupAcc.text = "" Then
         MsgBox "Plseas Choose Sage Supplier Account ID."
         tabAddNewClient.Tab = 0
         cboSageSupAcc.SetFocus
         Exit Sub
      End If
      If txtNewClientID.text = "" Then
         MsgBox "Plseas type Client ID."
         tabAddNewClient.Tab = 0
         txtNewClientID.SetFocus
         Exit Sub
      End If
      If txtNewClinetName.text = "" Then
         MsgBox "Plseas type Client Name."
         tabAddNewClient.Tab = 0
         txtNewClinetName.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub BankAccList()
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT BANK_NAME,BANK_POST_CODE,BANK_ID " & _
           "FROM TLBBANK;"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   
   While Not rstBank.EOF
      cboBankList.AddItem rstBank!BANK_NAME & " / " & rstBank!BANK_POST_CODE & _
                           " / " & rstBank!BANK_ID
      rstBank.MoveNext
   Wend
   
   cboBankList.AddItem "ADD NEW BANK"

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub tabClient_Click(PreviousTab As Integer)
   Select Case tabClient.Tab
   Case Is = 1
      If txtSageCust.text = "" And txtSageSupp.text = "" Then
         MsgBox "Select a Customer first.", vbCritical, "No Customer"
         tabClient.Tab = 0
      End If
   Case Is = 2
      If Not bLoadAggrements And txtClientID.text <> "" Then
         LoadAggrements txtClientID.text
      End If
   Case Is = 6
      tabAddNewClient.Tab = 0
      If cboSageCustAcc.ListCount = 0 Then
         SageCustomerAccCombo
         SageSupplierAccCombo
      End If
   Case Else
      'clicked anyother tab except state above
   End Select
End Sub

Private Sub LoadAggrements(szID As String)
   On Error Resume Next
   bLoadAggrements = True

   Dim Conn As New RDO.rdoConnection
   Dim Rst As rdoResultset
   Dim szStr As String, szaTemp() As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   szStr = "SELECT * " & _
         "FROM tlbAggreement " & _
         "WHERE tlbAggreement.CLIENT_ID='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)

   If Rst!AGG_DATE <> "" Then txtAggDt.text = Format(Rst!AGG_DATE, "DD/MM/YYYY")
   If Rst!START_DATE <> "" Then txtAggStDt.text = Format(Rst!START_DATE, "DD/MM/YYYY")
   If Rst!END_DATE <> "" Then txtAggEndDt.text = Format(Rst!END_DATE, "DD/MM/YYYY")
   If Rst!REVIEW_DATE <> "" Then txtAggReviewDt.text = Format(Rst!REVIEW_DATE, "DD/MM/YYYY")
   If Rst!NOTICE_DATE <> "" Then txtAggNoticeDt.text = Format(Rst!NOTICE_DATE, "DD/MM/YYYY")
   If Rst!BGRP_DATE <> "" Then txtBGRPDt.text = Format(Rst!BGRP_DATE, "DD/MM/YYYY")
   
   Rst.Close
   Set Rst = Nothing
   
   szStr = "SELECT CommissionType,CommissionAmt,BGRPayable " & _
           "FROM CLIENT " & _
           "WHERE ClientID='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
   
   If Rst!COMMISSIONTYPE = 2 Then
      txtCommissionAmt.text = Format(Rst!COMMISSIONAMT, "0.00")
      optFixed.Value = True
   Else
      txtCommissionAmt.text = Format(Rst!COMMISSIONAMT, "0.00") & "%"
      optComPercReceivable.Value = Not Rst!COMMISSIONTYPE
      optPercReveived.Value = Rst!COMMISSIONTYPE
   End If

   Rst.Close
   Set Rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Sub DrawLandLordTree()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt
'
'   Get the record for the client id
   szSQL = "SELECT ClieNt.ClientID, Client.ClientName " & _
           "FROM Client;"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
   Dim nodX As Node   ' Declare the object variable.
   Dim i As Integer, j As Integer   ' Declare a counter variable.
   Dim szProperty As String, szProArray() As String
   Dim szUnits As String, szUtArray() As String
   Dim szaUtNmID() As String, szaProNmID() As String
   Dim szLLID As String
'
   tvwLandLord.ImageList = imgList
   While Not rstClient.EOF
      szLLID = rstClient!ClientID + "@" + "LANDLORD"
   'Landlord ID
      Set nodX = tvwLandLord.Nodes.Add(, , szLLID, rstClient!ClientName + " / " + rstClient!ClientID, 1, 1)
   '
   '   Collece all property ID
      szProperty = LLPropertyList(rstClient!ClientID)
      szProArray = Split(szProperty, " # ")
      
      If szProArray(0) <> "NULL" Then
         For i = 0 To UBound(szProArray) 'Property Loop
            szaProNmID = Split(szProArray(i), " / ")
            Set nodX = tvwLandLord.Nodes.Add(szLLID, tvwChild, szaProNmID(0) & "@" & "PROPERTY", szaProNmID(1), 2, 2)
   
            'Collect all Units for current Property
            szUnits = LLUnitList(szaProNmID(0))
            szUtArray = Split(szUnits, " # ")
            If szUtArray(0) <> "NULL" Then
               For j = 0 To UBound(szUtArray) 'Unit Loop
                  szaUtNmID = Split(szUtArray(j), " / ")
                  Set nodX = tvwLandLord.Nodes.Add(szaProNmID(0) & "@" & "PROPERTY", tvwChild, szaUtNmID(0) & "@" & "UNIT", szaUtNmID(1), 3, 3)
               Next j
            End If
         Next i
      End If
      rstClient.MoveNext
   Wend
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

Private Sub SageCustomerAccCombo()
'
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
            cboSageCustAcc.AddItem CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value) & _
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

Private Sub SageSupplierAccCombo()
'
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

         ' Move to the First Record
         oPurchaseRecord.MoveFirst
         Dim rRow As Integer
         For rRow = 1 To oPurchaseRecord.Count
            cboSageSupAcc.AddItem CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value) & _
                                 " \ " & _
                                 CStr(oPurchaseRecord.Fields.Item("NAME").Value)
            oPurchaseRecord.MoveNext
         Next rRow

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

Private Sub txtSageCustomerAC_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtSageSupplierAC_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub tvwLandLord_Click()
   Dim szaPremisisIDType() As String

   szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
   fraType.Caption = szaPremisisIDType(1)
   
   PremisisImageLoader imgPremises, szaPremisisIDType(0), szaPremisisIDType(1)
'                                   ID                         TYPE
   txtTVInfoName.text = tvwLandLord.SelectedItem.text
   
   If szaPremisisIDType(1) = "PROPERTY" Then
      PropertyDetails szaPremisisIDType(0)
   End If
   If szaPremisisIDType(1) = "UNIT" Then
      UnitDetails szaPremisisIDType(0)
   End If
End Sub

Private Sub PropertyDetails(szID As String)
   Dim Conn As New RDO.rdoConnection
   Dim Rst As rdoResultset
   Dim szStr As String, szaTemp() As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   szStr = "SELECT * " & _
         "FROM PROPERTY " & _
         "WHERE PROPERTY.PROPERTYID='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
   
   txtTVInfoName.text = Rst!PropertyName
   txtTVInfoAdd1.text = Rst!ProAddressLine1
   txtTVInfoAdd2.text = Rst!ProAddressLine2
   txtTVInfoAdd3.text = Rst!ProAddressLine3
   txtTVInfoPC.text = Rst!PROPOSTCODE
   
   fraOccupied.Enabled = False
   
   Rst.Close
   Set Rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Sub UnitDetails(szID As String)
   Dim Conn As New RDO.rdoConnection
   Dim Rst As rdoResultset
   Dim szStr As String, szaTemp() As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   szStr = "SELECT * " & _
         "FROM UNITS " & _
         "WHERE UNITS.UnitNumber='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)

   If Rst.EOF Then
      MsgBox "Error in Database, Please contact with vendor", vbCritical, "Serious Error"
   Else
      If Rst!UNITNAME <> "" Then txtTVInfoName.text = Rst!UNITNAME
      If Rst!UnitAddressLine1 <> "" Then
         txtTVInfoAdd1.text = Rst!UnitAddressLine1
      Else
         txtTVInfoAdd1.text = ""
      End If
      If Rst!UnitAddressLine2 <> "" Then
         txtTVInfoAdd2.text = Rst!UnitAddressLine2
      Else
         txtTVInfoAdd2.text = ""
      End If
      If Rst!UnitAddressLine3 <> "" Then
         txtTVInfoAdd3.text = Rst!UnitAddressLine3
      Else
         txtTVInfoAdd3.text = ""
      End If
      If Rst!UnitPostCode <> "" Then
         txtTVInfoPC.text = Rst!UnitPostCode
      Else
         txtTVInfoPC.text = ""
      End If
      If Rst!OCCUPIED = "Y" Then
         lblTenantIDLink.Caption = Rst!SageAccountNumber
         lblTenantNameLink.Caption = Rst!TenantCompanyName
         Rst.Close
         Conn.Close
         Set Rst = Nothing
         Set Conn = Nothing
         
         szStr = LeaseDetails(szID)
         If szStr = "NULL" Then
            MsgBox "Error in DATABASE, Please contact with vendor.", vbCritical, "Error"
            Exit Sub
         End If
         szaTemp = Split(szStr, " # ")
         
         txtPreOccupiedFr.text = szaTemp(0)
         txtPreOccupiedTo.text = szaTemp(1)
         txtPreTenancyType.text = szaTemp(2)
         txtPreRentRvw.text = szaTemp(3)
      Else
         lblTenantIDLink.Caption = "NOT OCCUPIED"
         lblTenantNameLink.Caption = "NOT OCCUPIED"
         fraOccupied.Enabled = False
      End If
   End If
End Sub

Private Sub lblTenantIDLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblTenantIDLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblTenantIDLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   MsgBox "hi"
   lblTenantIDLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblTenantNameLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblTenantNameLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblTenantNameLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblTenantNameLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtAdd3_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtAddNewBankID_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClientID.text = ""

   If (Len(txtNewClientID.text) = 10 And KeyAscii <> 8) Or _
      KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtAddNewBankID_LostFocus()
   If txtAddNewBankID.text = "" Then
      MsgBox "Please provide Bank ID", vbCritical, "Error"
      Exit Sub
   End If

   If IsBankID(txtAddNewBankID.text) Then
      MsgBox "Bank ID already exits, Please provide another ID.", vbCritical, "Duplicate ID"
      txtAddNewBankID.SelStart = 0
      txtAddNewBankID.SelLength = Len(txtAddNewBankID.text)
      txtAddNewBankID.SetFocus
   End If
End Sub

Private Sub txtAddNewBankName_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtAddNewBankName.text = ""
   
'   If Len(txtAddNewBankName.text) = 10 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtAggDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtAggDt
End Sub

Private Sub txtAggEndDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtAggEndDt
End Sub

Private Sub txtAggNoticeDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtAggNoticeDt
End Sub

Private Sub txtAggReviewDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtAggReviewDt
End Sub

Private Sub txtAggStDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtAggStDt
End Sub

Private Sub txtBGRPDt_Click()
   dtDate.Visible = True
   dtDate.ZOrder 0
   Set szTextBox = txtBGRPDt
End Sub

Private Sub txtClientID_Change()
'   cmdClientID.Default = True
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next
   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
              "FROM CLIENT " & _
              "WHERE CLIENTID LIKE '" & Trim(txtClientID.text) & "%' " & _
              "ORDER BY CLIENTNAME;"
           
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

   flxSearchResult.Clear
   flxSearchResult.Rows = 2
   ConfigurFlexGrid
   While Not rstClient.EOF
      flxSearchResult.TextMatrix(iRow, 0) = rstClient!ClientID
      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
      rstClient.MoveNext
      If Not rstClient.EOF Then flxSearchResult.AddItem ""
      iRow = iRow + 1
   Wend

   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
'   cmdSelected.Enabled = True
End Sub

Private Sub txtClientID_KeyPress(KeyAscii As Integer)
   Dim iTxtLen As Integer
   
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtClientID.text = ""
   iTxtLen = Len(txtClientID.text)
'   If iTxtLen = 0 Then cmdClientID.Default = False
   If (iTxtLen = 10 And KeyAscii <> 8) Or KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtClientName_Change()
'   cmdClientName.Default = True
'   cmdClientID.Default = True
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
              "FROM CLIENT " & _
              "WHERE CLIENTNAME LIKE '" & Trim(txtClientName.text) & "%' " & _
              "ORDER BY CLIENTNAME;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   Dim iRow As Integer
   iRow = 1

   flxSearchResult.Clear
   flxSearchResult.Rows = 2
   ConfigurFlexGrid
   While Not rstClient.EOF
      flxSearchResult.TextMatrix(iRow, 0) = rstClient!ClientID
      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
      rstClient.MoveNext
      If Not rstClient.EOF Then flxSearchResult.AddItem ""
      iRow = iRow + 1
   Wend

   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
'   cmdSelected.Enabled = True
End Sub

Private Sub txtClientName_KeyPress(KeyAscii As Integer)
   Dim iTxtLen As Integer
   
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtClientID.text = ""
   iTxtLen = Len(txtClientName.text)
'   If iTxtLen = 0 Then cmdClientName.Default = False
   If iTxtLen = 10 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtHomePC_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtHomePh_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtNewClientID_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClientID.text = ""
   
   If (Len(txtNewClientID.text) = 10 And KeyAscii <> 8) Or _
      KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtNewClinetName_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClinetName.text = ""
   
'   If Len(txtNewClinetName.Text) = 10 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtOffAdd1_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffAdd2_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffAdd3_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffEmail_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffice_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffPC_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffPh_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtOffPos_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtPerEmail_KeyPress(KeyAscii As Integer)
   If cmdContactEdit.Enabled Then KeyAscii = 0
End Sub

Private Sub txtSearch_Change()
'   cmdSearch.Default = True
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
   Dim iTxtLen As Integer
   
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'   If KeyAscii = 27 Then txtSearch.text = ""
'   iTxtLen = Len(txtSearch.text)
'   If iTxtLen = 0 Then cmdSearch.Default = False
   If iTxtLen = 10 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub OpenFile()
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
'   Rst1.Close
'   Conn1.Close
'   Dim errorLevel As Long
'
'   errorLevel = ShellExecute(hWnd, "open", Filename, vbNullString, filePath, SW_SHOW)
'   If errorLevel < 32 Then MsgBox "File has been moved from original location.", vbExclamation
End Sub
            
Private Sub ResetFields()
   bLoadAggrements = False
   txtCommissionAmt.text = ""
   txtAggDt.text = ""
   txtAggStDt.text = ""
   txtAggEndDt.text = ""
   txtAggReviewDt.text = ""
   txtAggNoticeDt.text = ""
   txtBGRPDt.text = ""
   
   txtBankName.text = ""
   txtBankAdd1.text = ""
   txtBankAdd2.text = ""
   txtBankAdd3.text = ""
   txtBankPC.text = ""
   txtBankAccount.text = ""
   txtBankSC.text = ""

   'REFRESH THE TREE FOR NEW LANDLORD
   tvwLandLord.Nodes.Clear
   txtTVInfoName.text = ""
   txtTVInfoAdd1.text = ""
   txtTVInfoAdd2.text = ""
   txtTVInfoAdd3.text = ""
   txtTVInfoPC.text = ""
   txtPreOccupiedFr.text = ""
   txtPreOccupiedTo.text = ""
   txtPreTenancyType.text = ""
   txtPreRentRvw.text = ""
   lblTenantIDLink.Caption = ""
   lblTenantNameLink.Caption = ""
   imgPremises.Picture = LoadPicture("")
End Sub

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
