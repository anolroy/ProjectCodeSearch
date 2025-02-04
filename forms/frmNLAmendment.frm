VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNLAmendment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nominal Ledger Amendment"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNLAmendment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleMode       =   0  'User
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   12225
      TabIndex        =   9
      Top             =   6720
      Width           =   1485
   End
   Begin TabDlg.SSTab tabNL 
      Height          =   6525
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   11509
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nominal Account Details"
      TabPicture(0)   =   "frmNLAmendment.frx":9ED32
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSave"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Nominal Account History"
      TabPicture(1)   =   "frmNLAmendment.frx":9ED4E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSearch"
      Tab(1).Control(1)=   "fraSearch"
      Tab(1).Control(2)=   "cmdPrint"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "flxNominalHistory"
      Tab(1).Control(5)=   "Label2(3)"
      Tab(1).Control(6)=   "Label2(9)"
      Tab(1).Control(7)=   "Label2(5)"
      Tab(1).Control(8)=   "Label2(6)"
      Tab(1).Control(9)=   "Label2(4)"
      Tab(1).Control(10)=   "Label3"
      Tab(1).Control(11)=   "Label1(4)"
      Tab(1).Control(12)=   "txtNLCrTotal"
      Tab(1).Control(13)=   "txtNLDrTotal"
      Tab(1).Control(14)=   "txtNLBalance"
      Tab(1).Control(15)=   "Label2(2)"
      Tab(1).Control(16)=   "Label2(7)"
      Tab(1).Control(17)=   "Label2(8)"
      Tab(1).Control(18)=   "Label2(1)"
      Tab(1).Control(19)=   "Label2(0)"
      Tab(1).Control(20)=   "lblGridCaption(0)"
      Tab(1).ControlCount=   21
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   -73290
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5895
         Width           =   1125
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Automatic Demand Generate:"
         ForeColor       =   &H00FF00FF&
         Height          =   2220
         Left            =   -70725
         TabIndex        =   53
         Top             =   3510
         Visible         =   0   'False
         Width           =   3715
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E5E5E5&
            Height          =   2100
            Index           =   0
            Left            =   40
            ScaleHeight     =   2040
            ScaleWidth      =   3555
            TabIndex        =   54
            Top             =   50
            Width           =   3615
            Begin VB.TextBox txtSearchRef 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   810
               MaxLength       =   20
               TabIndex        =   60
               Top             =   810
               Width           =   2640
            End
            Begin VB.TextBox txtDateTo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2160
               TabIndex        =   62
               Top             =   1170
               Width           =   1290
            End
            Begin VB.TextBox txtDateFrom 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   810
               TabIndex        =   61
               Top             =   1170
               Width           =   1290
            End
            Begin VB.TextBox txtxFilterNo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   810
               MaxLength       =   10
               TabIndex        =   59
               Top             =   450
               Width           =   2640
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
               TabIndex        =   65
               Top             =   0
               Width           =   255
            End
            Begin VB.CommandButton cmdSearchCancel 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   2055
               TabIndex        =   64
               Top             =   1635
               Width           =   1200
            End
            Begin VB.CommandButton cmdSearchOK 
               Caption         =   "&OK"
               Height          =   375
               Left            =   120
               TabIndex        =   63
               Top             =   1605
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
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
               Index           =   4
               Left            =   180
               TabIndex        =   58
               Top             =   1170
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Desc."
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
               TabIndex        =   57
               Top             =   810
               Width           =   420
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "No"
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
               TabIndex        =   56
               Top             =   450
               Width           =   225
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00FFC0C0&
               BorderWidth     =   3
               Height          =   1155
               Index           =   1
               Left            =   75
               Top             =   360
               Width           =   3450
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
               TabIndex        =   55
               Top             =   0
               Width           =   1200
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
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   -74865
         TabIndex        =   51
         Top             =   5895
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Report Category"
         Height          =   3735
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Width           =   13440
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   735
            Left            =   13095
            TabIndex        =   66
            Top             =   3105
            Visible         =   0   'False
            Width           =   1050
         End
         Begin MSComctlLib.ListView lstCategoryCodes 
            Height          =   3255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388736
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.CommandButton cmdAddCategory 
            Caption         =   "Add"
            Height          =   375
            Left            =   12015
            TabIndex        =   7
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12045
         TabIndex        =   8
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nominal Account History  for"
         Height          =   975
         Left            =   -74865
         TabIndex        =   30
         Top             =   360
         Width           =   13095
         Begin VB.CheckBox chkIncludeZero 
            Caption         =   "Include Zero"
            Height          =   255
            Left            =   10305
            TabIndex        =   14
            Top             =   585
            Width           =   1425
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Left            =   11880
            TabIndex        =   16
            Top             =   600
            Width           =   1125
         End
         Begin VB.CheckBox chkYtD 
            Caption         =   "Y&TD"
            Height          =   255
            Left            =   10320
            TabIndex        =   13
            Top             =   240
            Width           =   615
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
            Left            =   11880
            TabIndex        =   15
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape Shape2 
            Height          =   780
            Left            =   10170
            Top             =   180
            Width           =   2850
         End
         Begin VB.Label lblClientNAC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Client Name"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   840
            TabIndex        =   45
            Top             =   240
            Width           =   2910
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
            Height          =   195
            Index           =   66
            Left            =   7515
            TabIndex        =   40
            Top             =   240
            Width           =   360
         End
         Begin MSForms.ComboBox cmbFund 
            Height          =   285
            Left            =   8010
            TabIndex        =   12
            Top             =   240
            Width           =   1950
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3440;503"
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
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propert&y:"
            Height          =   195
            Index           =   4
            Left            =   3825
            TabIndex        =   39
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Client:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   465
         End
         Begin MSForms.ComboBox cmbProperty 
            Height          =   285
            Left            =   4545
            TabIndex        =   11
            Top             =   270
            Width           =   2865
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5054;503"
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1058"
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nominal Account Details"
         Height          =   1695
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   13440
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Type"
            Height          =   195
            Index           =   5
            Left            =   4680
            TabIndex        =   48
            Top             =   780
            Width           =   930
         End
         Begin MSForms.ComboBox cboSubType 
            Height          =   285
            Left            =   5880
            TabIndex        =   3
            Top             =   780
            Width           =   2865
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5054;503"
            TextColumn      =   2
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
            Object.Width           =   "0;1940"
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Client"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblClient 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Client Name"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1440
            TabIndex        =   46
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance (YTD)"
            Height          =   195
            Index           =   9
            Left            =   9000
            TabIndex        =   43
            Top             =   360
            Width           =   960
         End
         Begin MSForms.TextBox txtBalanceYtD 
            Height          =   285
            Left            =   10320
            TabIndex        =   10
            Top             =   360
            Width           =   2235
            VariousPropertyBits=   679495711
            MaxLength       =   100
            BorderStyle     =   1
            Size            =   "3942;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Debit/Credit"
            Height          =   195
            Index           =   8
            Left            =   9000
            TabIndex        =   42
            Top             =   780
            Width           =   1050
         End
         Begin MSForms.ComboBox cboDrCr 
            Height          =   285
            Left            =   10320
            TabIndex        =   5
            Top             =   780
            Width           =   2235
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3942;503"
            TextColumn      =   2
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
            Object.Width           =   "0;1940"
         End
         Begin MSForms.ComboBox cboType 
            Height          =   285
            Left            =   5880
            TabIndex        =   2
            Top             =   360
            Width           =   2865
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5054;503"
            TextColumn      =   2
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
            Object.Width           =   "0;1940"
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Index           =   14
            Left            =   4680
            TabIndex        =   29
            Top             =   360
            Width           =   570
         End
         Begin MSForms.TextBox txtName 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   1200
            Width           =   3105
            VariousPropertyBits=   679495707
            MaxLength       =   100
            BorderStyle     =   1
            Size            =   "5477;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtCode 
            Height          =   285
            Left            =   1440
            TabIndex        =   0
            Top             =   780
            Width           =   3105
            VariousPropertyBits=   679495711
            BackColor       =   16777215
            MaxLength       =   15
            BorderStyle     =   1
            Size            =   "5477;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Nominal Name"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   780
            Width           =   840
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Posting Allowed"
            Height          =   195
            Index           =   0
            Left            =   4680
            TabIndex        =   26
            Top             =   1200
            Width           =   1170
         End
         Begin MSForms.ComboBox cmbPosting 
            Height          =   285
            Left            =   5880
            TabIndex        =   4
            Top             =   1200
            Width           =   2865
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5054;503"
            TextColumn      =   1
            ListRows        =   20
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "881;1940"
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNominalHistory 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   18
         Top             =   1650
         Width           =   13800
         _ExtentX        =   24342
         _ExtentY        =   7223
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
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "P. Date"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   -71805
         TabIndex        =   50
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   -62400
         TabIndex        =   49
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   -69720
         TabIndex        =   44
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   -67620
         TabIndex        =   37
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C No"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   -70785
         TabIndex        =   36
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   -64440
         TabIndex        =   35
         Top             =   6135
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
         Height          =   255
         Index           =   4
         Left            =   -64440
         TabIndex        =   34
         Top             =   5775
         Width           =   615
      End
      Begin MSForms.TextBox txtNLCrTotal 
         Height          =   285
         Left            =   -62475
         TabIndex        =   33
         Top             =   5760
         Width           =   1335
         VariousPropertyBits=   679495707
         MaxLength       =   100
         Size            =   "2355;503"
         SpecialEffect   =   3
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtNLDrTotal 
         Height          =   285
         Left            =   -63825
         TabIndex        =   32
         Top             =   5760
         Width           =   1335
         VariousPropertyBits=   679495707
         MaxLength       =   100
         Size            =   "2355;503"
         SpecialEffect   =   3
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtNLBalance 
         Height          =   285
         Left            =   -62475
         TabIndex        =   31
         Top             =   6120
         Width           =   1335
         VariousPropertyBits=   679495707
         MaxLength       =   100
         Size            =   "2355;503"
         SpecialEffect   =   3
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "T. Date"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   -72795
         TabIndex        =   23
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -65970
         TabIndex        =   22
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   -63705
         TabIndex        =   21
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -74040
         TabIndex        =   20
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   19
         Top             =   1440
         Width           =   210
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
         Left            =   -74835
         TabIndex        =   24
         Top             =   1440
         Width           =   13800
      End
   End
End
Attribute VB_Name = "frmNLAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resolved By BOSL. Issue: 0000476. Modified by Asif. Date: 18 Nov 2014.
Option Explicit

Public AddNew As Boolean

Private szNName      As String
Private iDrCr        As Integer
Private iPost        As Integer
Private bFormLoaded  As Boolean
Private iType        As Integer
Private szSubType    As String
Dim szSQLNominal As String
Dim reportingDate As String
Dim sessionID As String
Private bSortingCol()      As Boolean
Dim sMode As String

Private Sub cboDrCr_LostFocus()
    If cboDrCr.text <> "" Then
            If cboDrCr.text = "Debit" Or cboDrCr.text = "Credit" Then
            Else
                MsgBox "Please select valid input from the list", vbInformation, "Incorrect input"
                cboDrCr.text = ""
            End If
       End If
End Sub

Private Sub cboSubType_LostFocus()
    Dim szSQL As String
    Dim referenceValue As String
    Dim adoConn  As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    If cboSubType.text <> "" Then
        adoConn.Open getConnectionString
        referenceValue = cboSubType.text
        referenceValue = Replace(referenceValue, "'", "''")
           szSQL = "SELECT STCode, STName   " & _
           "FROM NLSubTypes where STName ='" & referenceValue & "'" & _
           "ORDER BY STName;"
       adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
       If adoRst.EOF Then
            MsgBox "Please select valid SubType from the list", vbInformation, "Incorrect list"
            cboSubType.text = ""
       End If
       adoRst.Close
       adoConn.Close
    End If

End Sub

Private Sub cboType_LostFocus()
    Dim szSQL As String
    Dim referenceValue As String
    Dim adoConn  As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    If cboType.text <> "" Then
        adoConn.Open getConnectionString
        referenceValue = cboType.text
        referenceValue = Replace(referenceValue, "'", "''")
        szSQL = "SELECT NLTypeCode, TypeValue " & _
               "FROM NLType " & _
               "WHERE NLTypeCode > 0 AND TypeValue ='" & referenceValue & "' " & _
               "ORDER BY NLTypeCode;"
       adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
       If adoRst.EOF Then
            MsgBox "Please select valid Type from the list", vbInformation, "Incorrect list"
            cboType.text = ""
       End If
       adoRst.Close
       adoConn.Close
    End If

End Sub

Private Sub chkIncludeZero_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call GenerateNominalHistory(adoConn, 3, " ASC")
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdAddCategory_Click()
   Load frmRptCategory
   frmRptCategory.Show
   Me.Enabled = False
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

       '        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
'                        If Not ctl Is picClient Then
'                            PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
'                        Else
                            bHandled = False
'                        End If

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
Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdClear_Click()
'   cmbProperty.ListIndex = 0
'   cmbFund.ListIndex = 0
   txtDateFrom.text = ""
   txtDateTo.text = ""
   txtxFilterNo.text = ""
   cmdFilter_Click
'   Dim iRow       As Integer
'   Dim cDr        As Currency
'   Dim cCr        As Currency
'
'   cDr = 0
'   cCr = 0
'
'   flxNominalHistory.Rows = 0
''   For iRow = 1 To flxNominalHistory.Rows - 1
''      flxNominalHistory.RowHeight(iRow) = 240
''
'''Resolved By BOSL. Issue 000476. Asif. Additional column "A/C No" added in the list.
''      If flxNominalHistory.TextMatrix(iRow, 8) <> "" Then
''         cDr = cDr + CCur(flxNominalHistory.TextMatrix(iRow, 8))
''      End If
''      If flxNominalHistory.TextMatrix(iRow, 9) <> "" Then
''         cCr = cCr + CCur(flxNominalHistory.TextMatrix(iRow, 9))
''      End If
'''End
''   Next iRow
'
'
'   txtNLDrTotal.text = Format(cDr, "0.00")
'   txtNLCrTotal.text = Format(cCr, "0.00")
'   txtNLBalance.text = Format(cDr - cCr, "0,00")
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdFilter_Click()
   Dim iRow       As Integer
   Dim cDr        As Currency
   Dim cCr        As Currency

'   For iRow = 1 To flxNominalHistory.Rows - 1
'      flxNominalHistory.RowHeight(iRow) = 240
'   Next iRow
'
'   If cmbProperty.ListIndex >= 0 Then
'      For iRow = 1 To flxNominalHistory.Rows - 1
'         If flxNominalHistory.TextMatrix(iRow, 10) <> cmbProperty.Value Then
'            flxNominalHistory.RowHeight(iRow) = 0
'         End If
'      Next iRow
'   End If
'
'   If cmbFund.ListIndex >= 0 Then
'      For iRow = 1 To flxNominalHistory.Rows - 1
'         If flxNominalHistory.TextMatrix(iRow, 9) <> cmbFund.Value Then
'            flxNominalHistory.RowHeight(iRow) = 0
'         End If
'      Next iRow
'   End If
'
'   If txtDateFrom.text <> "" Then
'      For iRow = 1 To flxNominalHistory.Rows - 1
'         If flxNominalHistory.TextMatrix(iRow, 1) <> "" Then
'            If CDate(flxNominalHistory.TextMatrix(iRow, 3)) < CDate(txtDateFrom.text) Then
'               flxNominalHistory.RowHeight(iRow) = 0
'            End If
'         End If
'      Next iRow
'   End If
'
'   If txtDateTo.text <> "" Then
'      For iRow = 1 To flxNominalHistory.Rows - 1
'         If flxNominalHistory.TextMatrix(iRow, 1) <> "" Then
'            If CDate(flxNominalHistory.TextMatrix(iRow, 3)) > CDate(txtDateTo.text) Then
'               flxNominalHistory.RowHeight(iRow) = 0
'            End If
'         End If
'      Next iRow
'   End If

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   GenerateNominalHistory adoConn, 3, " ASC "
   
   adoConn.Close
   Set adoConn = Nothing
   
'   cDr = 0
'   cCr = 0
'   For iRow = 1 To flxNominalHistory.Rows - 1
'      If flxNominalHistory.RowHeight(iRow) > 0 Then
'         If flxNominalHistory.TextMatrix(iRow, 7) <> "" Then
'            cDr = cDr + CCur(flxNominalHistory.TextMatrix(iRow, 7))
'         End If
'         If flxNominalHistory.TextMatrix(iRow, 8) <> "" Then
'            cCr = cCr + CCur(flxNominalHistory.TextMatrix(iRow, 8))
'         End If
'      End If
'   Next iRow
'   txtNLDrTotal.text = Format(cDr, "0.00")
'   txtNLCrTotal.text = Format(cCr, "0.00")
'   txtNLBalance.text = Format(cDr - cCr, "0,00")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorHandler
    Dim adoConn As New ADODB.Connection
    cmdPrint.Enabled = False
    adoConn.Open getConnectionString
    createTableNLHistory adoConn
    adoConn.Execute "DELETE FROM ReportNLHIstory WHERE SessionID = '" & sessionID & "';"
    adoConn.Execute "DELETE FROM ReportNLHIstory WHERE ReportingDate < #" & reportingDate & "# ;"
    adoConn.Execute _
    "INSERT INTO ReportNLHIstory " & _
    szSQLNominal
'    adoconn.Execute "Update ReportNLHistory set TRANSACTION_REF=TRANS_ID where TRANSACTION_REF is Null"
    adoConn.Close
    Set adoConn = Nothing
    
    '******************************show report **************************************************
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NLHistory.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Report.ParameterFields(1).AddCurrentValue sessionID
   Report.ParameterFields(2).AddCurrentValue frmNominalLedger.txtClientList.text
   Report.ParameterFields(3).AddCurrentValue frmNominalLedger.txtPropertyName.text
   Report.ParameterFields(4).AddCurrentValue frmNominalLedger.txtFundName.text
   If Len(frmNominalLedger.cmbPeriodFrom.text) > 0 Then
         Report.ParameterFields(5).AddCurrentValue Format(frmNominalLedger.cmbPeriodFrom.Column(2), "dd/mm/yyyy")
   End If
   If Len(frmNominalLedger.cmbPeriodTo.text) > 0 Then
        Report.ParameterFields(6).AddCurrentValue Format(frmNominalLedger.cmbPeriodTo.Column(3), "dd/mm/yyyy")
   End If
   
   'Report.ParameterFields(4).AddCurrentValue cboClientID.Column(1)
   
'   Report.ParameterFields(6).AddCurrentValue txtFundName.text

   Load frmReport
   frmReport.LoadReportViewer Report
   cmdPrint.Enabled = True
    Exit Sub
    
ErrorHandler:
   cmdPrint.Enabled = True
    MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Could not load Balance Sheet"
'    Set adoRst = Nothing
End Sub
Private Sub createTableNLHistory(adoConn As ADODB.Connection)
       
    Dim adoRst As New ADODB.Recordset
    On Error GoTo CreateReportNLHIstory
    
       adoRst.Open "SELECT * FROM ReportNLHIstory;", adoConn, adOpenStatic, adLockReadOnly
       adoRst.Close
    
       GoTo Alreadyhavetable
    
CreateReportNLHIstory:
     adoConn.Execute _
      "CREATE TABLE ReportNLHistory " & _
         "(" & _
                "ReportingDate DateTime  NOT NULL, " & _
                "SessionID     TEXT(100) NOT NULL, " & _
                "THIS_RECORD        TEXT(50), " & _
                "TRANS_ID      TEXT(50), " & _
                "TRANSACTION_REF        TEXT(10), " & _
                "TRANSACTION_DATE   DateTime, " & _
                "POSTED_DATE     DateTime, " & _
                "UNIQUE_REFERENCE_NO   AUTOINCREMENT, " & _
                "TRANSACTION_DESCRIPTION          TEXT(255), " & _
                "AMOUNT       CURRENCY, " & _
                "FUND_ID         Integer, " & _
                "PROPERTY_ID       TEXT(4)," & _
                "TRANSACTION_TYPE   BYTE NOT NULL, " & _
                "AMOUNT_TYPE          TEXT(1), " & _
                "REFERENCE       TEXT(100), " & _
                "ACCOUNT_NUMBER      TEXT(20), " & _
                "FundName        TEXT(255), " & _
                "CONSTANT       TEXT(10), " & _
                "DESCRIPTION       TEXT(250), " & _
                "DEBIT       CURRENCY, " & _
                "CREDIT       CURRENCY " & _
         ");"
           Exit Sub
Alreadyhavetable:
End Sub
Private Sub cmdSave_Click()
   If txtCode.text = "" Then
      ShowMsgInTaskBar "Please enter the code", "Y", "N"
      txtCode.SetFocus
      Exit Sub
   End If
   If txtName.text = "" Then
      ShowMsgInTaskBar "Please enter the code name", "Y", "N"
      txtName.SetFocus
      Exit Sub
   End If
   'Resolved by BOSL
   'issue 460 The user should be warned that they cannot save a nominal record without a type.
   'Modified by anol 24 Aug 2014
   If cboType.text = "" Then
      ShowMsgInTaskBar "Please select a type", "Y", "N"
      cboType.SetFocus
      Exit Sub
   End If
   
   If IsNull(cboType.text) Then
      ShowMsgInTaskBar "Please select the type", "Y", "N"
      cboType.SetFocus
      Exit Sub
   End If
   If cboDrCr.ListIndex = -1 Then
      ShowMsgInTaskBar "Please select a Debit/Credit entry", "Y", "N"
      cboDrCr.SetFocus
      Exit Sub
   End If

   Dim adoConn       As New ADODB.Connection
   Dim adoRst        As New ADODB.Recordset
   Dim szSQL         As String
   Dim i             As Integer

   adoConn.Open getConnectionString

   If AddNew Then
     If frmNominalLedger.SSTab1.Tab = 0 Then
            szSQL = "SELECT * " & _
              "FROM   NominalLedger " & _
              "WHERE  Code = '" & txtCode.text & "' AND " & _
                     "ClientID = '" & frmNominalLedger.txtClientList.Tag & "';"
    Else
            szSQL = "SELECT * " & _
              "FROM   NominalLedger " & _
              "WHERE  Code = '" & txtCode.text & "' AND " & _
                     "ClientID = 'NONE';"
    End If
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      If Not adoRst.EOF Then
         MsgBox "This nominal code already exists.", vbInformation, "Warning"
         adoRst.Close
         Set adoRst = Nothing
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If
      adoRst.Close

      adoConn.BeginTrans
      szSQL = "SELECT * FROM NominalLedger;"
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      adoRst.AddNew
      adoRst.Fields.Item("CreatedBy").Value = User
      adoRst.Fields.Item("CreatedDate").Value = Now
      adoRst.Fields.Item("Code").Value = txtCode.text
      If frmNominalLedger.SSTab1.Tab = 0 Then
            adoRst.Fields.Item("ClientID").Value = frmNominalLedger.txtClientList.Tag 'lblClient.Tag '
      Else
            adoRst.Fields.Item("ClientID").Value = "NONE" 'lblClient.Tag '
      End If
      adoRst.Fields.Item("CAType").Value = ""
   Else
      adoConn.BeginTrans
       If frmNominalLedger.SSTab1.Tab = 0 Then
              szSQL = "SELECT * " & _
              "FROM   NominalLedger " & _
              "WHERE  Code = '" & txtCode.text & "' AND " & _
                     "ClientID = '" & frmNominalLedger.txtClientList.Tag & "';"
      Else
            szSQL = "SELECT * " & _
              "FROM   NominalLedger " & _
              "WHERE  Code = '" & txtCode.text & "' AND " & _
                     "ClientID = 'NONE';"
      End If
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   End If

   adoRst.Fields.Item("Name").Value = txtName.text
   adoRst.Fields.Item("Type").Value = cboType.Value
   adoRst.Fields.Item("Posting").Value = IIf(cmbPosting.text = "Yes", True, False)
   adoRst.Fields.Item("DrCr").Value = cboDrCr.Value
   If Not IsNull(cboSubType.Value) Then
      adoRst.Fields.Item("SubType").Value = cboSubType.Value
   End If
   adoRst.Fields.Item("LastModifiedBy").Value = User
   adoRst.Fields.Item("LastModifiedDate").Value = Now
   adoRst.Update
   adoRst.Close
   If frmNominalLedger.SSTab1.Tab = 0 Then
        adoConn.Execute "DELETE * FROM NJ_CC " & _
                   "WHERE ClientID = '" & frmNominalLedger.txtClientList.Tag & "' AND " & _
                         "Code = '" & txtCode.text & "';"
   Else
        adoConn.Execute "DELETE * FROM NJ_CC " & _
                   "WHERE ClientID = 'NONE' AND " & _
                         "Code = '" & txtCode.text & "';"
    End If

   szSQL = "SELECT * FROM NJ_CC;"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With lstCategoryCodes
      For i = 1 To .ListItems.Count
         If .ListItems(i).Checked Then
            adoRst.AddNew
            If frmNominalLedger.SSTab1.Tab = 0 Then
                    adoRst.Fields.Item("ClientID").Value = frmNominalLedger.txtClientList.Tag
            Else
                    adoRst.Fields.Item("ClientID").Value = "NONE"
            End If
            adoRst.Fields.Item("Code").Value = txtCode.text
            adoRst.Fields.Item("CC").Value = .ListItems(i)
            adoRst.Update
         End If
      Next i
   End With
   adoRst.Close

   adoConn.CommitTrans

'   frmNominalLedger.RefreshGrid adoConn

   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   cmdCancel_Click
   frmNominalLedger.cmdFilter_Click
   frmNominalLedger.LoadDefaultChartofAccounts
   ShowMsgInTaskBar "Nominal Ledger has been updated successfully", "Y", "P"
End Sub

Private Sub ConfigFlxNominalHistory()
   Dim szHeader As String
   Dim iCol As Integer

   flxNominalHistory.Clear
   flxNominalHistory.Cols = 13
   flxNominalHistory.Rows = 2
   flxNominalHistory.RowHeight(0) = 0

   szHeader$ = "THIS_RECORD|<No|<Type|<Date|<P.Date|<A/C No|<Fund|<Reference|<Description|>Dr|>Cr|FundID|PropID"

   flxNominalHistory.FormatString = szHeader$
   flxNominalHistory.ColWidth(0) = 0
   For iCol = 0 To flxNominalHistory.Cols - 5
      flxNominalHistory.ColWidth(iCol + 1) = Label2(iCol + 1).Left - Label2(iCol).Left
      Label2(iCol).Width = flxNominalHistory.ColWidth(iCol + 1)
   Next iCol
   
   flxNominalHistory.ColWidth(iCol + 1) = flxNominalHistory.Width + flxNominalHistory.Left - Label2(flxNominalHistory.Cols - 4).Left - 300 'TotalAmount
   Label2(iCol).Width = flxNominalHistory.ColWidth(iCol + 1)
   flxNominalHistory.ColWidth(iCol + 2) = 0
   flxNominalHistory.ColWidth(iCol + 3) = 0
   
   'flxNominalHistory.ColWidth(iCol + 1) = flxNominalHistory.Width + flxNominalHistory.Left - Label2(flxNominalHistory.Cols - 4).Left - 300 'TotalAmount
'
'   flxNominalHistory.ColWidth(iCol) = 1200
'   flxNominalHistory.ColWidth(iCol + 1) = 1200
'
'
'
'   Label2(iCol - 1).Width = flxNominalHistory.ColWidth(iCol)
'   Label2(iCol).Width = flxNominalHistory.ColWidth(iCol + 1) - 200
'
'   flxNominalHistory.ColWidth(iCol + 2) = 0
'   flxNominalHistory.ColWidth(iCol + 3) = 0
   flxNominalHistory.ColAlignment(7) = vbLeftJustify
   txtNLDrTotal.Left = Label2(8).Left 'flxNominalHistory.ColWidth(8)
   txtNLCrTotal.Left = Label2(9).Left
   txtNLBalance.Left = txtNLCrTotal.Left

   flxNominalHistory.row = 0
End Sub

Private Function szControlAccount(adoConn As ADODB.Connection) As String
   Dim szSQL      As String
   Dim szNC       As String
   Dim szClient   As String
   Dim adoRst     As New ADODB.Recordset

   szNC = frmNominalLedger.flxNominalCode.TextMatrix(frmNominalLedger.flxNominalCode.row, 0)
   szClient = frmNominalLedger.txtClientList.Tag

'CAName, Code AS NCode, Name AS NName, CAFixed AS Fixed, CADisOrder AS DisOrder, CAPosting AS Posting,
   szSQL = "SELECT CAType AS Type " & _
           "FROM   NominalLedger " & _
           "WHERE  NOT ISNULL(CAName) AND NOT ISNULL(CATYPE) AND " & _
                  "ClientID = '" & szClient & "' AND Code = '" & szNC & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
      szControlAccount = adoRst.Fields.Item("Type").Value
      adoRst.Close
      Set adoRst = Nothing
      Exit Function
   End If

   adoRst.Close
   Set adoRst = Nothing
   szControlAccount = "N"
End Function

'Private Sub LoadFlxNominalHistory(adoConn As ADODB.Connection)
'   Dim szSQL      As String
'   Dim iRow       As Long
'   Dim adoRst     As New ADODB.Recordset
'   Dim iTran      As Integer
'   Dim curDr      As Currency
'   Dim curCr      As Currency
'   Dim szCA       As String
''   Dim szDrCr     As String
'
'   With frmNominalLedger
'      If .flxNominalCode.TextMatrix(.flxNominalCode.row, 2) = 1 Then  'Balance Sheet
'         szSQL = "SELECT N.THIS_RECORD, R.SlNumber AS TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T, tlbReceipt AS R " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartBS, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANS_ID = CSTR(R.TransactionID) AND " & _
'                       "N.TRANSACTION_TYPE IN (1, 2, 3, 4, 23) "
'
'         szSQL = szSQL & " UNION " & _
'                     "SELECT N.THIS_RECORD, P.SlNumber AS TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T, tlbPayment AS P " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartBS, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANS_ID = CSTR(P.TransactionID) AND " & _
'                       "N.TRANSACTION_TYPE IN (6, 7, 8, 9, 24) "
'
'         szSQL = szSQL & " UNION " & _
'                     "SELECT N.THIS_RECORD, N.TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartBS, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANSACTION_TYPE IN (11, 12);"
'      Else
'         szSQL = "SELECT N.THIS_RECORD, R.SlNumber AS TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T, tlbReceipt AS R " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartPnL, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANS_ID = CSTR(R.TransactionID) AND " & _
'                       "N.TRANSACTION_TYPE IN (1, 2, 3, 4, 23) "
'
'         szSQL = szSQL & " UNION " & _
'                 "SELECT N.THIS_RECORD, P.SlNumber AS TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T, tlbPayment AS P " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartPnL, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANS_ID = CSTR(P.TransactionID) AND " & _
'                       "N.TRANSACTION_TYPE IN (6, 7, 8, 9, 24) "
'
'         szSQL = szSQL & " UNION " & _
'                 "SELECT N.THIS_RECORD, N.TRANS_ID, N.TRANSACTION_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, " & _
'                        "N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, F.FundName, T.CONSTANT, T.DESCRIPTION " & _
'                 "FROM NLPosting AS N, Fund AS F, tlbTransactionTypes AS T " & _
'                 "WHERE N.NOMINAL_CODE = '" & .flxNominalCode.TextMatrix(.flxNominalCode.row, 0) & "' AND " & _
'                       "N.TRANSACTION_DATE >= #" & Format(.dtStartPnL, "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(.dtEnd, "dd mmmm yyyy") & "# AND " & _
'                       "N.ClientID = '" & .cmbClient.Value & "' AND " & _
'                       "N.AMOUNT > 0 AND " & _
'                       "N.TRANSACTION_TYPE = T.TYPE_ID AND " & _
'                       "N.FUND_ID = F.FundID AND " & _
'                       "N.TRANSACTION_TYPE IN (11, 12);"
'
'      End If
'   End With
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   iRow = 1
'
'   While Not adoRst.EOF
'      szCA = szControlAccount(adoConn)   'S -> Sales CA, P-> Purchase CA, O-> Output VAT, I-> Input VAT; N -> NOT CA
'      flxNominalHistory.TextMatrix(iRow, 0) = adoRst.Fields.Item("THIS_RECORD").Value
'      flxNominalHistory.TextMatrix(iRow, 1) = Mid(adoRst.Fields.Item("CONSTANT").Value, 4) & adoRst.Fields.Item("TRANS_ID").Value
'      flxNominalHistory.TextMatrix(iRow, 2) = adoRst.Fields.Item("DESCRIPTION").Value
'      flxNominalHistory.TextMatrix(iRow, 3) = Format(adoRst.Fields.Item("TRANSACTION_DATE").Value, "dd/mm/yyyy")
'      flxNominalHistory.TextMatrix(iRow, 4) = adoRst.Fields.Item("FundName").Value                          'Fund Name
'      flxNominalHistory.TextMatrix(iRow, 5) = adoRst.Fields.Item("UNIQUE_REFERENCE_NO").Value
'      flxNominalHistory.TextMatrix(iRow, 6) = adoRst.Fields.Item("TRANSACTION_DESCRIPTION").Value
'
''      If frmNominalLedger.flxNominalCode.TextMatrix(frmNominalLedger.flxNominalCode.row, 2) = 1 Then     'Balance Sheet
'         If ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "2" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "3" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "4") And szCA <> "S") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "6" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "24") And szCA <> "P") Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "2" And szCA = "O") Or _
'            adoRst.Fields.Item("TRANSACTION_TYPE").Value = "15" Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "11" And _
'              (adoRst.Fields.Item("AMOUNT_TYPE").Value = "A" Or _
'               adoRst.Fields.Item("AMOUNT_TYPE").Value = "V")) Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "1" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "23") And szCA = "S") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "7" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "8" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "9") And szCA = "P") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "6" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "11") And szCA = "I") Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "12" And _
'               adoRst.Fields.Item("AMOUNT_TYPE").Value = "B") Then
'
'            flxNominalHistory.TextMatrix(iRow, 7) = Format(adoRst.Fields.Item("AMOUNT").Value, "0.00")         'Dr
'            curDr = curDr + CCur(flxNominalHistory.TextMatrix(iRow, 7))
'         End If
'         If ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "2" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "3" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "4") And szCA = "S") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "6" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "24") And szCA = "P") Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "7" And szCA = "I") Or _
'            adoRst.Fields.Item("TRANSACTION_TYPE").Value = "16" Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "12" And _
'              (adoRst.Fields.Item("AMOUNT_TYPE").Value = "A" Or _
'               adoRst.Fields.Item("AMOUNT_TYPE").Value = "V")) Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "1" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "23") And szCA <> "S") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "7" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "8" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "9") And szCA <> "P") Or _
'            ((adoRst.Fields.Item("TRANSACTION_TYPE").Value = "1" Or _
'               adoRst.Fields.Item("TRANSACTION_TYPE").Value = "12") And szCA = "O") Or _
'            (adoRst.Fields.Item("TRANSACTION_TYPE").Value = "11" And _
'               adoRst.Fields.Item("AMOUNT_TYPE").Value = "B") Then
'
'            flxNominalHistory.TextMatrix(iRow, 8) = Format(adoRst.Fields.Item("AMOUNT").Value, "0.00")         'Dr
'            curCr = curCr + CCur(flxNominalHistory.TextMatrix(iRow, 8))
'         End If
'
'      flxNominalHistory.TextMatrix(iRow, 9) = adoRst.Fields.Item("FUND_ID").Value                           'Fund ID
'      flxNominalHistory.TextMatrix(iRow, 10) = IIf(IsNull(adoRst.Fields.Item("PROPERTY_ID").Value), "", adoRst.Fields.Item("PROPERTY_ID").Value) 'Prop ID
'
'      adoRst.MoveNext
'      iRow = iRow + 1
'      If Not adoRst.EOF Then flxNominalHistory.AddItem ""
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing
'
'   txtNLDrTotal.text = Format(curDr, "0.00")
'   txtNLCrTotal.text = Format(curCr, "0.00")
'   txtNLBalance.text = Format(curDr - curCr, "0.00")
'End Sub

Public Sub LoadLstCategoryCodes(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim iRow       As Integer
   Dim adoRst     As New ADODB.Recordset
   Dim itmX       As ListItem ' Create a variable to add ListItem objects.
   Dim clmX       As ColumnHeader ' Create an object variable for the ColumnHeader object.

   lstCategoryCodes.ListItems.Clear
   lstCategoryCodes.ColumnHeaders.Clear

   ' Add ColumnHeaders.
   Set clmX = lstCategoryCodes.ColumnHeaders.Add(, , "Code", (lstCategoryCodes.Width / 3) - 150)
   Set clmX = lstCategoryCodes.ColumnHeaders.Add(, , "Category Name", (lstCategoryCodes.Width / 3) - 100)
   Set clmX = lstCategoryCodes.ColumnHeaders.Add(, , "Category Description", lstCategoryCodes.Width / 3)

   If Not AddNew Then
      szSQL = "SELECT S1.CategoryCode, S1.CategoryName, S1.CatDesc, S2.CC " & _
              "FROM   ReportCategory AS S1 LEFT JOIN NJ_CC AS S2 ON (" & _
                     "S1.ClientID = S2.ClientID AND " & _
                     "S1.CategoryCode = S2.CC AND " & _
                     "S2.Code = '" & frmNominalLedger.flxNominalCode.TextMatrix(frmNominalLedger.flxNominalCode.row, 0) & "') " & _
              "WHERE  S1.ClientID = '" & frmNominalLedger.txtClientList.Tag & "' " & _
              "ORDER BY S1.CategoryCode;"
   Else
      szSQL = "SELECT S1.CategoryCode, S1.CategoryName, S1.CatDesc " & _
              "FROM   ReportCategory AS S1 " & _
              "WHERE  S1.ClientID = '" & frmNominalLedger.txtClientList.Tag & "' " & _
              "ORDER BY S1.CategoryCode;"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRst.EOF
      Set itmX = lstCategoryCodes.ListItems.Add(, , adoRst.Fields.Item("CategoryCode").Value)
      ' Add two subitems for that item
      itmX.SubItems(1) = adoRst.Fields.Item("CategoryName").Value
      itmX.SubItems(2) = adoRst.Fields.Item("CatDesc").Value
      lstCategoryCodes.ListItems(iRow).Checked = False
      If Not AddNew Then If Not IsNull(adoRst.Fields.Item("CC").Value) Then lstCategoryCodes.ListItems(iRow).Checked = True

      iRow = iRow + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadCboType(adoConn As ADODB.Connection)
   Dim Data()     As String
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT NLTypeCode, TypeValue " & _
           "FROM NLType " & _
           "WHERE NLTypeCode > 0 " & _
           "ORDER BY NLTypeCode;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboType.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadCboSubType(adoConn As ADODB.Connection)
   Dim Data()     As String
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT STCode, STName " & _
           "FROM NLSubTypes " & _
           "ORDER BY STName;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboSubType.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadCboDrCr()
   Dim Data(1, 1)    As String

   Data(0, 0) = "Dr"
   Data(1, 0) = "Debit"
   Data(0, 1) = "Cr"
   Data(1, 1) = "Credit"

   cboDrCr.Column() = Data()
End Sub

Private Sub cmdSearch_Click()
    fraSearch.Left = 4275
    If cmdSearch.Caption = "Clear Sea&rch" Then
         sMode = "NO"
         txtxFilterNo.text = ""
         txtSearchRef.text = ""
         txtDateFrom.text = ""
         txtDateTo.text = ""
         cmdSearch.Caption = "Sea&rch"
         fraSearch.Visible = False
    Else
        If fraSearch.Visible = False Then
            fraSearch.Visible = True
            FocusControl txtxFilterNo
        Else
            fraSearch.Visible = False
        End If
    End If
End Sub

Private Sub cmdSearchCancel_Click()
    cmdFilter.SetFocus
    fraSearch.Visible = False
End Sub

Private Sub cmdSearchOK_Click()
    cmdFilter.SetFocus
    Call cmdFilter_Click
    fraSearch.Visible = False
End Sub

Private Sub Command1_Click()
    MsgBox cmbPosting.Locked
End Sub

Private Sub Form_Activate()
   txtCode.Locked = Not AddNew
   bFormLoaded = True
   If AddNew Then
      Me.Caption = "Nominal Ledger Amendment - New"
   Else
      Me.Caption = "Nominal Ledger Amendment - " & txtCode.text & ": " & txtName.text
   End If
End Sub

Private Function GenerateNominalHistory(adoConn As ADODB.Connection, j As Integer, strSort As String)
   
   On Error GoTo NewNominalCode
  
   Dim clientID As String
   Dim propertyID As String
   Dim fundID As String
   Dim fromDate As String
   Dim toDate As String
   Dim nominalCode As String
   Dim nominalType As String
   Dim sqlFilter As String
   Dim tempstr As String
   
   Dim retainedEarningsControl As String
   sessionID = GetTimeStamp
   reportingDate = Format(DateValue(Now), "dd mmmm yyyy")

   clientID = frmNominalLedger.txtClientList.Tag
   propertyID = cmbProperty.Value
   fundID = cmbFund.Value
   
   retainedEarningsControl = GetNominalCodeForControlAccount(adoConn, "Retained Earnings", clientID)
   
   nominalCode = frmNominalLedger.flxNominalCode.TextMatrix(frmNominalLedger.flxNominalCode.row, 0)
   nominalType = frmNominalLedger.flxNominalCode.TextMatrix(frmNominalLedger.flxNominalCode.row, 2)
   
   
   sqlFilter = ""
   
   sqlFilter = "AND ClientID = '" & clientID & "' " & _
               "AND NOMINAL_CODE = '" & nominalCode & "' "
               
   If propertyID <> "ALL" Then
        sqlFilter = sqlFilter & "AND PROPERTY_ID = '" & propertyID & "' "
   End If
   
   If fundID <> "ALL" Then
        sqlFilter = sqlFilter & "AND FUND_ID = " & fundID & " "
   End If
   ConfigFlxNominalHistory
   If chkYtD.Value = 1 Then
        fromDate = Format(frmNominalLedger.dtStartPnL, "dd mmmm yyyy")
        toDate = Format(frmNominalLedger.dtEnd, "dd mmmm yyyy")
        
        If retainedEarningsControl = nominalCode Then
            'sqlFilter = sqlFilter & " AND POSTED_DATE < #" & fromDate & "# "
            Dim Filter As String
            Filter = ""
            
            If propertyID <> "ALL" Then
                Filter = " AND N.PROPERTY_ID = '" & propertyID & "' "
            End If
            
            If fundID <> "ALL" Then
                Filter = Filter & " AND N.FUND_ID = " & fundID & " "
            End If
            If chkIncludeZero.Value Then
                Filter = Filter & " AND N.AMOUNT > 0 "
            End If

            szSQLNominal = GetRetainedEarningsSQL(clientID, fromDate, Filter, reportingDate, sessionID, j, strSort)
        
        ElseIf retainedEarningsControl <> nominalCode And nominalType = "1" Then
            'sqlFilter = sqlFilter & " AND POSTED_DATE <= #" & toDate & "# "
            'Modified by anol 20160518
            If Trim(txtDateFrom.text) = "" And Trim(txtDateTo.text) = "" Then
                 'sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "#  AND POSTED_DATE <= #" & toDate & "# "
                 'issue 353 Nominal Account header consistent with transaction drill down
                 'modified by anol 20170420
                 sqlFilter = sqlFilter & " AND POSTED_DATE <= #" & toDate & "# "
                 If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
                 End If
                 szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
            ElseIf Trim(txtDateFrom.text) <> "" And Trim(txtDateTo.text) <> "" Then
                fromDate = Format(txtDateFrom.text, "dd mmmm yyyy")
                toDate = Format(txtDateTo.text, "dd mmmm yyyy")
                sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "#  AND POSTED_DATE <= #" & toDate & "# "
                If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
                 End If
                szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
            ElseIf Trim(txtDateFrom.text) = "" And Trim(txtDateTo.text) <> "" Then
                fromDate = Format(txtDateFrom.text, "dd mmmm yyyy")
                toDate = Format(txtDateTo.text, "dd mmmm yyyy")
                sqlFilter = sqlFilter & " AND POSTED_DATE <= #" & toDate & "# "
                If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
                 End If
                szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
            ElseIf Trim(txtDateFrom.text) <> "" And Trim(txtDateTo.text) = "" Then
                fromDate = Format(txtDateFrom.text, "dd mmmm yyyy")
                toDate = Format(txtDateTo.text, "dd mmmm yyyy")
                If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
                 End If
                sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "# "
                szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
                
            End If
           
            
        
        ElseIf retainedEarningsControl <> nominalCode And nominalType = "2" Then
            sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "#  AND POSTED_DATE <= #" & toDate & "# "
            If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
            End If
            szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
        Else
            szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
        End If
        
   Else
   
      If txtDateFrom.text <> "" And txtDateTo.text <> "" Then
        fromDate = Format(CDate(txtDateFrom.text), "dd mmmm yyyy")
        toDate = Format(CDate(txtDateTo.text), "dd mmmm yyyy")
        
        sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "#  AND POSTED_DATE <= #" & toDate & "# "
        If Not CBool(chkIncludeZero.Value) Then
                        sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
         End If
        szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
      
      ElseIf txtDateFrom.text = "" And txtDateTo.text = "" Then
        fromDate = Format(frmNominalLedger.dtStartBS, "dd mmmm yyyy")
        toDate = Format(frmNominalLedger.dtEnd, "dd mmmm yyyy")
        
        sqlFilter = sqlFilter & " AND POSTED_DATE >= #" & fromDate & "#  AND POSTED_DATE <= #" & toDate & "# "
        If Not CBool(chkIncludeZero.Value) Then
               sqlFilter = sqlFilter & " AND N.AMOUNT <> 0 "
        End If
        szSQLNominal = GetNominalHistorySQL(sqlFilter, reportingDate, sessionID, j, strSort)
      Else
        MsgBox "You must enter both the Date From and Date To to load the periodic Nominal History"
        Exit Function
      End If
        
   End If
      

'   Debug.Print szSQLNominal
   
   Dim adoRst As New ADODB.Recordset

   adoRst.Open szSQLNominal, adoConn, adOpenStatic, adLockOptimistic
   If Len(Trim(txtxFilterNo.text)) > 0 Then
        tempstr = Replace(txtxFilterNo.text, "'", "''")
        adoRst.Filter = " INVNO Like '%" & tempstr & "%'"
   End If
   If Len(Trim(txtSearchRef.text)) > 0 Then
        tempstr = Replace(txtSearchRef.text, "'", "''")
        adoRst.Filter = " TRANSACTION_DESCRIPTION Like '%" & tempstr & "%'"
   End If
   If adoRst.RecordCount = 0 Then
        flxNominalHistory.Rows = 2
   Else
        flxNominalHistory.Rows = adoRst.RecordCount + 1
   End If

   
   If adoRst.EOF Then
       adoRst.Close
       Set adoRst = Nothing
       Exit Function
   End If

   Dim iRow As Long
   Dim debitTotal As Double, creditTotal As Double
   debitTotal = 0
   creditTotal = 0
      
   iRow = 1
   While Not adoRst.EOF
       
      flxNominalHistory.TextMatrix(iRow, 0) = adoRst.Fields.Item("THIS_RECORD").Value
      flxNominalHistory.TextMatrix(iRow, 1) = IIf(IsNull(adoRst.Fields.Item("INVNO")), "", adoRst.Fields.Item("INVNO"))
      flxNominalHistory.TextMatrix(iRow, 2) = adoRst.Fields.Item("DESCRIPTION").Value
      flxNominalHistory.TextMatrix(iRow, 3) = Format(adoRst.Fields.Item("TRANSACTION_DATE").Value, "dd/mm/yyyy")
      flxNominalHistory.TextMatrix(iRow, 4) = Format(adoRst.Fields.Item("POSTED_DATE").Value, "dd/mm/yyyy")
      flxNominalHistory.TextMatrix(iRow, 5) = IIf(IsNull(adoRst.Fields.Item("ACCOUNT_NUMBER").Value), "", adoRst.Fields.Item("ACCOUNT_NUMBER").Value) 'ACCOUNT_NUMBER
      flxNominalHistory.TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("FUNDNAME").Value), "", adoRst.Fields.Item("FUNDNAME").Value) 'Fund Name
      flxNominalHistory.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("REFERENCE").Value), "", adoRst.Fields.Item("REFERENCE").Value) 'REFERENCE
      flxNominalHistory.TextMatrix(iRow, 8) = IIf(IsNull(adoRst.Fields.Item("TRANSACTION_DESCRIPTION").Value), "", adoRst.Fields.Item("TRANSACTION_DESCRIPTION").Value) 'TRANSACTION_DESCRIPTION

      flxNominalHistory.TextMatrix(iRow, 9) = IIf(adoRst.Fields.Item("DEBIT").Value > 0, Format(adoRst.Fields.Item("DEBIT").Value, "#,###.00"), "")       'Dr
      flxNominalHistory.TextMatrix(iRow, 10) = IIf(adoRst.Fields.Item("CREDIT").Value > 0, Format(adoRst.Fields.Item("CREDIT").Value, "#,###.00"), "")       'Cr
     
      flxNominalHistory.TextMatrix(iRow, 11) = IIf(IsNull(adoRst.Fields.Item("FUND_ID").Value), "", adoRst.Fields.Item("FUND_ID").Value) 'Fund ID
      flxNominalHistory.TextMatrix(iRow, 12) = IIf(IsNull(adoRst.Fields.Item("PROPERTY_ID").Value), "", adoRst.Fields.Item("PROPERTY_ID").Value) 'Prop ID
      
      debitTotal = debitTotal + IIf(IsNull(adoRst.Fields.Item("DEBIT")), 0, adoRst.Fields.Item("DEBIT"))
      creditTotal = creditTotal + IIf(IsNull(adoRst.Fields.Item("CREDIT")), 0, adoRst.Fields.Item("CREDIT"))
      
      adoRst.MoveNext
      iRow = iRow + 1
'      If Not adoRst.EOF Then flxNominalCode.AddItem ""
    Wend

   txtNLDrTotal.text = Format(debitTotal, "#,##0.00")
   txtNLCrTotal.text = Format(creditTotal, "#,##0.00")
   
   txtNLBalance.text = Format(debitTotal - creditTotal, "#,##0.00")
   
   adoRst.Close
   Set adoRst = Nothing

   Exit Function

NewNominalCode:
   Set adoRst = Nothing
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Generating Nominal Balances"
   
End Function
Private Function GetNominalHistorySQL(Filter As String, reportingDate As String, sessionID As String, j As Integer, strSort As String) As String
    Dim szOrderby As String
    If j = 0 Then
        szOrderby = " ORDER BY Cint(N.TRANSACTION_REF) " & strSort
    End If
    If j = 1 Then
        szOrderby = " ORDER BY N.TRANSACTION_TYPE " & strSort
    End If
    If j = 2 Then
        szOrderby = " ORDER BY N.TRANSACTION_DATE " & strSort
    End If
    If j = 3 Then
        szOrderby = "ORDER BY N.POSTED_DATE " & strSort
    End If
    If j = 4 Then
        szOrderby = "ORDER BY N.ACCOUNT_NUMBER " & strSort
    End If
    If j = 5 Then
        szOrderby = "ORDER BY  N.FUND_ID " & strSort
    End If
    If j = 6 Then
        szOrderby = "ORDER BY N.UNIQUE_REFERENCE_NO " & strSort
    End If
    Dim szSQL As String
    
    szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID,N.THIS_RECORD, (Mid(CONSTANT, 4) & TRANSACTION_REF) AS INVNo, N.TRANSACTION_REF, N.TRANSACTION_DATE," & _
    "N.POSTED_DATE, N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, " & _
    "N.REFERENCE, N.ACCOUNT_NUMBER, " & _
    "(SELECT Fund.FundCode FROM Fund WHERE Fund.FundID = N.FUND_ID) as FundName, " & _
    "T.CONSTANT, T.DESCRIPTION, " & _
    "(IIf(N.AMOUNT>0,N.AMOUNT,0)) AS DEBIT, (IIf(N.AMOUNT<0,N.AMOUNT*-1,0)) AS CREDIT " & _
    "FROM NLPosting AS N, tlbTransactionTypes AS T " & _
    "WHERE N.TRANSACTION_TYPE = T.TYPE_ID " & _
    "AND N.DeleteFlag = 0 " & _
    Filter & " " & _
    szOrderby

    GetNominalHistorySQL = szSQL

End Function

Private Function GetRetainedEarningsSQL(clientID As String, asOndate As String, Filter As String, reportingDate As String, sessionID As String, j As Integer, strSort As String) As String
    Dim szOrderby As String
    If j = 0 Then
        szOrderby = " ORDER BY N.TRANS_ID " & strSort
    End If
    If j = 1 Then
        szOrderby = " ORDER BY N.TRANSACTION_TYPE " & strSort
    End If
    If j = 2 Then
        szOrderby = " ORDER BY N.TRANSACTION_DATE " & strSort
    End If
    If j = 3 Then
        szOrderby = "ORDER BY N.POSTED_DATE " & strSort
    End If
    If j = 4 Then
        szOrderby = "ORDER BY N.ACCOUNT_NUMBER " & strSort
    End If
    If j = 5 Then
        szOrderby = "ORDER BY  N.FUND_ID " & strSort
    End If
    If j = 6 Then
        szOrderby = "ORDER BY N.UNIQUE_REFERENCE_NO " & strSort
    End If
    If j = 7 Then
        Filter = Filter & " "
    End If
    
    Dim szSQL As String
    
    szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID,N.THIS_RECORD, (Mid(CONSTANT, 4) & TRANSACTION_REF) AS INVNo, N.TRANSACTION_REF, N.TRANSACTION_DATE,N.POSTED_DATE, " & _
    "N.UNIQUE_REFERENCE_NO, N.TRANSACTION_DESCRIPTION, N.AMOUNT, N.FUND_ID, N.PROPERTY_ID, N.TRANSACTION_TYPE, N.AMOUNT_TYPE, " & _
    "N.REFERENCE, N.ACCOUNT_NUMBER, " & _
    "(SELECT Fund.FundCode FROM Fund WHERE Fund.FundID = N.FUND_ID) as FundName, " & _
    "T.CONSTANT, T.DESCRIPTION, " & _
    "(IIf(N.AMOUNT>0,N.AMOUNT,0)) AS DEBIT, (IIf(N.AMOUNT<0,N.AMOUNT*-1,0)) AS CREDIT " & _
    "FROM NLPosting AS N, tlbTransactionTypes AS T, NominalLedger AS NL " & _
    "WHERE N.TRANSACTION_TYPE = T.TYPE_ID " & _
    "AND NL.CODE = N.NOMINAL_CODE " & _
    "AND NL.CLIENTID = N.CLIENTID " & _
    "AND NL.TYPE = 2 " & _
    "AND NL.ClientID = '" & clientID & "' " & _
    "AND N.DeleteFlag = 0 " & _
    "AND N.POSTED_DATE < #" & asOndate & "# " & _
    Filter & " " & _
    szOrderby
    
    GetRetainedEarningsSQL = szSQL

End Function
Private Sub Form_Load()
   bFormLoaded = False
   Me.Height = 7650
   Me.Width = 14295
   Me.Left = 100
   Me.Top = 100
   ReDim bSortingCol(9) As Boolean
   Me.BackColor = MODULEBACKCOLOR
   tabNL.BackColor = Me.BackColor

   cmbPosting.AddItem "Yes"
   cmbPosting.AddItem "No"
   cmbPosting.ListIndex = 0
   tabNL.Tab = 0
   szNName = ""
   iDrCr = -1
   iPost = -1
   iType = -1

   ConfigFlxNominalHistory
   LoadCboDrCr

   Dim adoConn As New ADODB.Connection
   Dim rsNLposting As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsNLposting.Open "Select TRANSACTION_REF from NLPOSTING where (TRANSACTION_TYPE=15 OR TRANSACTION_TYPE=16 )AND TRANSACTION_REF is NULL", adoConn, adOpenDynamic, adLockOptimistic
   If Not rsNLposting.EOF Then
        adoConn.Execute "Update NLPOSTING set TRANSACTION_REF=TRANS_ID where (TRANSACTION_TYPE=15 OR TRANSACTION_TYPE=16 )AND TRANSACTION_REF is NULL"
   End If
   
   
   LoadCboType adoConn
   LoadCboSubType adoConn
   LoadLstCategoryCodes adoConn
   
   LoadPropertyByClient adoConn
   LoadFund adoConn

   cmbProperty.Value = frmNominalLedger.txtPropertyName.Tag
    
   cmbFund.Value = frmNominalLedger.txtFundName.Tag
   chkYtD.Value = frmNominalLedger.chkYtD.Value
   
   'If Not AddNew Then LoadFlxNominalHistory adoConn
   If Not AddNew Then GenerateNominalHistory adoConn, 3, " ASC"
    Call WheelHook(Me.hWnd)
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadFund(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim szaData() As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT FundID, FundCode FROM FUND;"

  
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Funds"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cmbFund.Column() = Data()
   cmbFund.ListIndex = 0
   
NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Function LoadPropertyByClient(adoConn As ADODB.Connection) As Boolean
   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT PropertyID, PropertyName " & _
           "FROM Property " & _
           "WHERE ClientID = '" & frmNominalLedger.txtClientList.Tag & "' " & _
           "ORDER BY PropertyName;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   
   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cmbProperty.Column() = Data()
   cmbProperty.ListIndex = 0
   
   LoadPropertyByClient = True
   Exit Function

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   
   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmNominalLedger.Enabled = True
   UnLoadForm Me
End Sub

Private Sub Label2_Click(Index As Integer)
    If Index >= 0 And Index <= 6 Then
    'implementation of soring on he label click by anol
            Label2(Index).FontBold = Not Label2(Index).FontBold
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            ConfigFlxNominalHistory
            GenerateNominalHistory adoConn, Index, IIf(Label2(Index).FontBold, "DESC", "ASC")
            adoConn.Close
    End If
End Sub

Private Sub lblClient_Change()
   lblClientNAC.Caption = lblClient.Caption
End Sub

Private Sub lstCategoryCodes_Click()
   cmdSave.Enabled = True
End Sub

Private Sub tabNL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateTo.SetFocus
    End If
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   Dim X As Boolean

   X = TextBoxFormatDate(txtDateFrom)

   If X And txtDateTo.text = "" Then txtDateTo.text = txtDateFrom.text
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   TextBoxFormatDate txtDateTo
End Sub

Private Sub txtName_Change()
   If bFormLoaded And szNName <> txtName.text Then cmdSave.Enabled = True
End Sub

Private Sub txtName_GotFocus()
   If szNName = "" Then szNName = txtName.text
End Sub

Private Sub cmbPosting_Click()
   If bFormLoaded And iPost <> cmbPosting.ListIndex Then cmdSave.Enabled = True
End Sub

Private Sub cmbPosting_GotFocus()
   If iPost = -1 Then iPost = cmbPosting.ListIndex
End Sub

Private Sub cboDrCr_Click()
   If bFormLoaded And iDrCr <> cboDrCr.ListIndex Then cmdSave.Enabled = True
End Sub

Private Sub cboDrCr_GotFocus()
   If iDrCr = -1 Then iDrCr = cboDrCr.ListIndex
End Sub

Private Sub cboType_GotFocus()
   If iType = -1 Then iType = cboType.ListIndex
End Sub

Private Sub cboType_Click()
   If bFormLoaded And iType <> cboType.ListIndex Then cmdSave.Enabled = True
End Sub

Private Sub cboSubType_GotFocus()
   szSubType = IIf(IsNull(cboSubType.Value), "", cboSubType.Value)
End Sub

Private Sub cboSubType_Click()
   If bFormLoaded And szSubType <> cboSubType.Value Then cmdSave.Enabled = True
End Sub

Private Sub txtSearchRef_Change()
    Dim adoConn As New ADODB.Connection
    txtxFilterNo.text = ""
    txtDateFrom.text = ""
    txtDateTo.text = ""
    If sMode = "REF" Then
        adoConn.Open getConnectionString
        GenerateNominalHistory adoConn, 3, " ASC"
        adoConn.Close
        Set adoConn = Nothing
    End If
    
    If Len(txtSearchRef.text) > 0 Then
        cmdSearch.Caption = "Clear Sea&rch"
    Else
        cmdSearch.Caption = "Sea&rch"
    End If
End Sub

Private Sub txtSearchRef_GotFocus()
     sMode = "REF"
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateFrom.SetFocus
    End If
End Sub

Private Sub txtxFilterNo_Change()
    Dim adoConn As New ADODB.Connection
    txtSearchRef.text = ""
    txtDateFrom.text = ""
    txtDateTo.text = ""
    If sMode = "NO" Then
        adoConn.Open getConnectionString
        GenerateNominalHistory adoConn, 3, " ASC"
        adoConn.Close
        Set adoConn = Nothing
    End If
    If Len(txtxFilterNo.text) > 0 Then
        cmdSearch.Caption = "Clear Sea&rch"
    Else
        cmdSearch.Caption = "Sea&rch"
    End If
End Sub

Private Sub txtxFilterNo_GotFocus()
    sMode = "NO"
End Sub

Private Sub txtxFilterNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 39 Then
       KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        txtSearchRef.SetFocus
    End If
End Sub
