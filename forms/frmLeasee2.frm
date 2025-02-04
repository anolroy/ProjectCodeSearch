VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLeasee2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leaseholder"
   ClientHeight    =   11760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeasee2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11760
   ScaleWidth      =   14670
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   6
      Left            =   11760
      TabIndex        =   269
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descending Date Order"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   275
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintHistoryCancel 
         Caption         =   "&Cancel"
         Height          =   365
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintHistorySorted 
         Caption         =   "&OK"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ascending Date Order"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   271
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optSortingAccountHistory 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transaction Serial Number"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   270
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1215
         Index           =   18
         Left            =   120
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Sorting Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   -510
         TabIndex        =   274
         Top             =   0
         Width           =   2325
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   660
         Index           =   32
         Left            =   0
         Top             =   0
         Width           =   120
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   660
         Index           =   33
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   19
         Left            =   120
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtUnitNumber 
      Height          =   285
      Left            =   11760
      TabIndex        =   268
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2805
      Index           =   0
      Left            =   8040
      ScaleHeight     =   2775
      ScaleWidth      =   4815
      TabIndex        =   205
      Top             =   8520
      Visible         =   0   'False
      Width           =   4845
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   40
         ScaleHeight     =   3015
         ScaleWidth      =   4815
         TabIndex        =   206
         Top             =   0
         Width           =   4815
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
            Index           =   0
            Left            =   4500
            Style           =   1  'Graphical
            TabIndex        =   336
            Top             =   20
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
            Height          =   2175
            Index           =   0
            Left            =   15
            TabIndex        =   207
            Top             =   525
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   13553358
            ForeColorFixed  =   -2147483634
            BackColorSel    =   14737632
            ForeColorSel    =   -2147483630
            BackColorBkg    =   16777215
            GridColor       =   14737632
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
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
         Begin VB.Label lblFlxPayee 
            Caption         =   "EMPTY"
            Height          =   255
            Index           =   0
            Left            =   2115
            TabIndex        =   214
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblPayeeFlxConfigured 
            Caption         =   "NOT"
            Height          =   495
            Index           =   0
            Left            =   1515
            TabIndex        =   213
            Top             =   1680
            Width           =   1095
         End
         Begin MSForms.Label lblSearch0 
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   212
            Top             =   0
            Width           =   735
            VariousPropertyBits=   8388627
            Caption         =   "dynamic"
            Size            =   "1296;353"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label lblSearch1 
            Height          =   195
            Index           =   0
            Left            =   750
            TabIndex        =   211
            Top             =   15
            Width           =   735
            VariousPropertyBits=   8388627
            Caption         =   "dynamic"
            Size            =   "1296;353"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label lblSearch2 
            Height          =   195
            Index           =   0
            Left            =   1710
            TabIndex        =   210
            Top             =   15
            Width           =   735
            VariousPropertyBits=   8388627
            Caption         =   "dynamic"
            Size            =   "1296;353"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSearch1 
            Height          =   255
            Left            =   30
            TabIndex        =   209
            Top             =   240
            Width           =   975
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "1720;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSearch2 
            Height          =   255
            Left            =   1320
            TabIndex        =   208
            Top             =   240
            Width           =   1215
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2143;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Index           =   2
            Left            =   0
            Top             =   30
            Width           =   4500
         End
      End
   End
   Begin VB.PictureBox fmeTenantLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3225
      Left            =   840
      ScaleHeight     =   3195
      ScaleWidth      =   6855
      TabIndex        =   175
      Top             =   8520
      Visible         =   0   'False
      Width           =   6885
      Begin VB.CommandButton cmdGridTenantLookup 
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTenantLookup 
         Height          =   1965
         Left            =   30
         TabIndex        =   177
         Top             =   1215
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   285
         Left            =   5400
         TabIndex        =   217
         Top             =   870
         Width           =   1320
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2328;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchAddress 
         Height          =   285
         Left            =   3360
         TabIndex        =   216
         Top             =   870
         Width           =   2040
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3598;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   215
         Top             =   660
         Width           =   840
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   184
         Top             =   660
         Width           =   165
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   183
         Top             =   660
         Width           =   405
      End
      Begin VB.Label lblTenantSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   182
         Top             =   660
         Width           =   735
      End
      Begin MSForms.TextBox txtSearchTenant 
         Height          =   285
         Left            =   30
         TabIndex        =   180
         Top             =   870
         Width           =   900
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1587;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   14
         Left            =   0
         TabIndex        =   186
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   13
         Left            =   0
         TabIndex        =   185
         Top             =   50
         Width           =   465
      End
      Begin MSForms.ComboBox cboClientList 
         Height          =   285
         Left            =   720
         TabIndex        =   178
         Top             =   45
         Width           =   5805
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "10239;503"
         BoundColumn     =   0
         TextColumn      =   1
         ColumnCount     =   8
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763;7055"
      End
      Begin MSForms.ComboBox cboPropertyList 
         Height          =   285
         Left            =   720
         TabIndex        =   179
         Top             =   360
         Width           =   5805
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "10239;503"
         BoundColumn     =   0
         TextColumn      =   1
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1587;5115"
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   30
         Top             =   660
         Width           =   6780
      End
      Begin MSForms.TextBox txtSearchName 
         Height          =   285
         Left            =   960
         TabIndex        =   181
         Top             =   870
         Width           =   2385
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4207;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtClientID 
      Height          =   285
      Left            =   11640
      TabIndex        =   170
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fmeBankPaymentDetails 
      Caption         =   "Bank Payment Details"
      Height          =   5205
      Left            =   13080
      TabIndex        =   97
      Top             =   8760
      Width           =   10875
      Begin VB.CommandButton cmdSaveBank 
         Caption         =   "&Save"
         Height          =   375
         Left            =   8790
         TabIndex        =   103
         Top             =   4700
         Width           =   915
      End
      Begin VB.CommandButton cmdCancelBank 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   9765
         TabIndex        =   102
         Top             =   4700
         Width           =   915
      End
      Begin VB.CommandButton cmdEditBank 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   7815
         TabIndex        =   101
         Top             =   4700
         Width           =   915
      End
      Begin VB.CommandButton cmdNewBank 
         Caption         =   "&New"
         Height          =   375
         Left            =   6840
         TabIndex        =   100
         Top             =   4700
         Width           =   915
      End
      Begin VB.CommandButton cmdGetPaymentMethods 
         Caption         =   "..."
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
         Left            =   10380
         TabIndex        =   99
         Top             =   330
         Width           =   285
      End
      Begin VB.CheckBox chkIsDefaultAC 
         Caption         =   "Yes"
         Height          =   315
         Left            =   7410
         TabIndex        =   98
         Top             =   1980
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBank 
         Height          =   2025
         Left            =   180
         TabIndex        =   104
         Top             =   2550
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   3572
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   -2147483638
         ForeColorSel    =   0
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
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label50 
         Caption         =   "Bank:"
         Height          =   225
         Index           =   8
         Left            =   210
         TabIndex        =   132
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label50 
         Caption         =   "Address:"
         Height          =   225
         Index           =   12
         Left            =   210
         TabIndex        =   131
         Top             =   990
         Width           =   825
      End
      Begin MSForms.TextBox txtBankAddress1 
         Height          =   315
         Left            =   1380
         TabIndex        =   130
         Top             =   930
         Width           =   3285
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5794;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBankAddress2 
         Height          =   315
         Left            =   1380
         TabIndex        =   129
         Top             =   1245
         Width           =   3285
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5794;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBankPostCode 
         Height          =   315
         Left            =   1380
         TabIndex        =   128
         Top             =   1875
         Width           =   1515
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "2672;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         Caption         =   "Post Code:"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   127
         Top             =   1875
         Width           =   825
      End
      Begin MSForms.TextBox txtBankACName 
         Height          =   315
         Left            =   7410
         TabIndex        =   126
         Top             =   555
         Width           =   3255
         VariousPropertyBits=   746604571
         Size            =   "5741;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPaymentMethod 
         Height          =   315
         Left            =   7410
         TabIndex        =   125
         Top             =   210
         Width           =   2985
         VariousPropertyBits=   1820346395
         DisplayStyle    =   3
         Size            =   "5265;556"
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
      Begin VB.Label Label36 
         Caption         =   "A/C Name:"
         Height          =   225
         Left            =   5700
         TabIndex        =   124
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label Label37 
         Caption         =   "Payment Method:"
         Height          =   225
         Left            =   5700
         TabIndex        =   123
         Top             =   210
         Width           =   1305
      End
      Begin MSForms.TextBox txtBankACNumber 
         Height          =   315
         Left            =   7410
         TabIndex        =   122
         Top             =   915
         Width           =   3255
         VariousPropertyBits=   746604571
         Size            =   "5741;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBankSortCode 
         Height          =   315
         Left            =   7410
         TabIndex        =   121
         Top             =   1260
         Width           =   1545
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "2725;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label38 
         Caption         =   "Sort Code:"
         Height          =   225
         Left            =   5700
         TabIndex        =   120
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label40 
         Caption         =   "A/C Number:"
         Height          =   225
         Left            =   5700
         TabIndex        =   119
         Top             =   915
         Width           =   1035
      End
      Begin MSForms.TextBox txtBACSRef 
         Height          =   315
         Left            =   7410
         TabIndex        =   118
         Top             =   1620
         Width           =   3255
         VariousPropertyBits=   746604571
         Size            =   "5741;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label41 
         Caption         =   "BACS Ref:"
         Height          =   225
         Left            =   5700
         TabIndex        =   117
         Top             =   1620
         Width           =   1035
      End
      Begin MSForms.ComboBox cboBankId 
         Height          =   315
         Left            =   1380
         TabIndex        =   116
         Top             =   210
         Width           =   3285
         VariousPropertyBits=   1820346395
         DisplayStyle    =   3
         Size            =   "5794;556"
         BoundColumn     =   0
         TextColumn      =   1
         ColumnCount     =   8
         ListRows        =   20
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label49 
         Caption         =   "Default Account:"
         Height          =   225
         Left            =   5700
         TabIndex        =   115
         Top             =   2040
         Width           =   1425
      End
      Begin MSForms.TextBox txtBankAddress3 
         Height          =   315
         Left            =   1380
         TabIndex        =   114
         Top             =   1560
         Width           =   3285
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5794;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBranchName 
         Height          =   315
         Left            =   1380
         TabIndex        =   113
         Top             =   570
         Width           =   3285
         VariousPropertyBits=   746604575
         BackColor       =   12640511
         Size            =   "5794;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         Caption         =   "Branch:"
         Height          =   225
         Index           =   10
         Left            =   210
         TabIndex        =   112
         Top             =   615
         Width           =   825
      End
      Begin MSForms.TextBox txtBankTenantID 
         Height          =   315
         Left            =   4905
         TabIndex        =   111
         Top             =   -90
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
      Begin VB.Label Label31 
         Caption         =   "Account Number"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   110
         Top             =   2320
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Account Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   109
         Top             =   2320
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Sort Code"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   108
         Top             =   2320
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "Payment Method"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   107
         Top             =   2320
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "BacsRef"
         Height          =   255
         Index           =   4
         Left            =   7560
         TabIndex        =   106
         Top             =   2320
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "Default Account"
         Height          =   255
         Index           =   5
         Left            =   9480
         TabIndex        =   105
         Top             =   2320
         Width           =   1215
      End
   End
   Begin VB.Frame fmeTenant 
      Caption         =   "Lessee Information"
      Height          =   2595
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         TabIndex        =   345
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "C&opy Lessee"
         Height          =   385
         Left            =   4350
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         TabIndex        =   173
         Text            =   "0.00"
         Top             =   1486
         Width           =   3015
      End
      Begin VB.CommandButton cmdDeleteLessee 
         Caption         =   "&Delete Lessee"
         Height          =   385
         Left            =   8445
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Lessee"
         Height          =   385
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel Changes"
         Height          =   385
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Lessee"
         Height          =   385
         Left            =   5715
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Lessee"
         Height          =   385
         Left            =   2985
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2100
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   385
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F9EE&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "0.00"
         Top             =   1173
         Width           =   3015
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   150
         Width           =   4140
         Begin VB.OptionButton optCurrentTenant 
            Caption         =   "Current"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   60
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optExTenant 
            Caption         =   "Ex-Lessee"
            Height          =   195
            Left            =   1680
            TabIndex        =   9
            Top             =   60
            Width           =   1155
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   195
            Left            =   3480
            TabIndex        =   10
            Top             =   60
            Width           =   825
         End
      End
      Begin VB.TextBox cboSageAccountNumber 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   96
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "NiU"
         Height          =   1095
         Left            =   120
         TabIndex        =   360
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Share (%):"
         Height          =   195
         Left            =   6540
         TabIndex        =   346
         Top             =   1800
         Width           =   690
      End
      Begin MSForms.CommandButton cmdPropLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   267
         Top             =   500
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClientLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   266
         Top             =   150
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdUnitLookup 
         Height          =   255
         Left            =   10770
         TabIndex        =   265
         Top             =   860
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   264
         Top             =   1200
         Width           =   675
      End
      Begin MSForms.CheckBox chkPrintDmd 
         Height          =   255
         Left            =   400
         TabIndex        =   263
         Top             =   1560
         Width           =   1440
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2540;450"
         Value           =   "1"
         Caption         =   "Print Demands:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkEmailSC 
         Height          =   285
         Left            =   355
         TabIndex        =   262
         Top             =   1815
         Width           =   1485
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2619;503"
         Value           =   "0"
         Caption         =   "Email S. Charge:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCombEmail 
         Height          =   255
         Left            =   4020
         TabIndex        =   235
         Top             =   1815
         Width           =   1560
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2752;450"
         Value           =   "0"
         Caption         =   "Combined Email:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkEmailSt 
         Height          =   255
         Left            =   2100
         TabIndex        =   234
         Top             =   1815
         Width           =   1560
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2752;450"
         Value           =   "0"
         Caption         =   "Email Statement:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkPrintSt 
         Height          =   255
         Left            =   4005
         TabIndex        =   232
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;450"
         Value           =   "0"
         Caption         =   "Print Statements:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CheckBox chkEmailDmd 
         Height          =   255
         Left            =   2220
         TabIndex        =   231
         Top             =   1560
         Width           =   1440
         VariousPropertyBits=   746596371
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "2540;450"
         Value           =   "0"
         Caption         =   "Email Demands:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label lblLeaseChanged 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Left            =   5640
         TabIndex        =   199
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Held (£):"
         Height          =   195
         Left            =   6540
         TabIndex        =   174
         Top             =   1486
         Width           =   1170
      End
      Begin MSForms.TextBox txtUnit 
         Height          =   315
         Left            =   8040
         TabIndex        =   18
         Top             =   822
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   315
         Left            =   8040
         TabIndex        =   17
         Top             =   471
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClient 
         Height          =   315
         Left            =   8040
         TabIndex        =   16
         Top             =   120
         Width           =   3015
         VariousPropertyBits=   679495711
         BackColor       =   12640511
         Size            =   "5318;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance (£):"
         Height          =   195
         Index           =   6
         Left            =   6540
         TabIndex        =   15
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   4
         Left            =   6540
         TabIndex        =   14
         Top             =   460
         Width           =   645
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   3
         Left            =   6540
         TabIndex        =   13
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lessee ID:"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   12
         Top             =   825
         Width           =   720
      End
      Begin MSForms.TextBox txtCompanyName 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   1200
         Width           =   3945
         VariousPropertyBits=   746604571
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtName 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   460
         Width           =   3945
         VariousPropertyBits=   747653147
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdTenantLookup 
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Top             =   860
         Width           =   255
         Caption         =   """"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   2
         Top             =   450
         Width           =   435
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Index           =   5
         Left            =   6540
         TabIndex        =   1
         Top             =   830
         Width           =   330
      End
      Begin MSForms.TextBox txtTenantID 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   830
         Width           =   3945
         VariousPropertyBits=   746604575
         BackColor       =   15858158
         MaxLength       =   8
         Size            =   "6959;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
      Height          =   315
      Left            =   5235
      ScaleHeight     =   315
      ScaleWidth      =   2655
      TabIndex        =   60
      Top             =   8745
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
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
         Height          =   195
         Left            =   90
         TabIndex        =   61
         Top             =   60
         Width           =   2475
      End
   End
   Begin TabDlg.SSTab tabTenant 
      Height          =   5625
      Left            =   75
      TabIndex        =   11
      Top             =   2760
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   8
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Lessee Details"
      TabPicture(0)   =   "frmLeasee2.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblGridCaption(2)"
      Tab(0).Control(1)=   "lblContacts(6)"
      Tab(0).Control(2)=   "lblContacts(4)"
      Tab(0).Control(3)=   "lblContacts(5)"
      Tab(0).Control(4)=   "lblContacts(3)"
      Tab(0).Control(5)=   "lblContacts(2)"
      Tab(0).Control(6)=   "lblContacts(0)"
      Tab(0).Control(7)=   "lblContacts(1)"
      Tab(0).Control(8)=   "flxContacts"
      Tab(0).Control(9)=   "fmeTenantAddress"
      Tab(0).Control(10)=   "txtAddress"
      Tab(0).Control(11)=   "cmdEditContacts"
      Tab(0).Control(12)=   "cmdAddNewContacts"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&Defaults"
      TabPicture(1)   =   "frmLeasee2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(4)"
      Tab(1).Control(1)=   "txtSLControlName"
      Tab(1).Control(2)=   "Shape1(5)"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label18(0)"
      Tab(1).Control(5)=   "lblVatCode(0)"
      Tab(1).Control(6)=   "txtNominalCodeName"
      Tab(1).Control(7)=   "Label1(0)"
      Tab(1).Control(8)=   "txtNominalCode"
      Tab(1).Control(9)=   "cmdSLC"
      Tab(1).Control(10)=   "txtSLControl"
      Tab(1).Control(11)=   "txtCodeVat"
      Tab(1).Control(12)=   "cmdTaxList"
      Tab(1).Control(13)=   "cmdNC"
      Tab(1).Control(14)=   "txtDefault(0)"
      Tab(1).Control(15)=   "cmdCancelDefaults"
      Tab(1).Control(16)=   "cmdEditDefaults"
      Tab(1).Control(17)=   "cmdSaveDefaults"
      Tab(1).Control(18)=   "txtDefault(1)"
      Tab(1).Control(19)=   "txtDefault(2)"
      Tab(1).Control(20)=   "txtDefault(3)"
      Tab(1).Control(21)=   "txtDefault(4)"
      Tab(1).Control(22)=   "txtDefault(5)"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "De&posit Held"
      TabPicture(2)   =   "frmLeasee2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape5"
      Tab(2).Control(1)=   "Label50(4)"
      Tab(2).Control(2)=   "Label50(11)"
      Tab(2).Control(3)=   "Label50(5)"
      Tab(2).Control(4)=   "Label50(6)"
      Tab(2).Control(5)=   "Label50(7)"
      Tab(2).Control(6)=   "Label50(9)"
      Tab(2).Control(7)=   "Label6(0)"
      Tab(2).Control(8)=   "Label6(1)"
      Tab(2).Control(9)=   "Label6(2)"
      Tab(2).Control(10)=   "Label6(3)"
      Tab(2).Control(11)=   "Label6(5)"
      Tab(2).Control(12)=   "Label6(6)"
      Tab(2).Control(13)=   "Label6(7)"
      Tab(2).Control(14)=   "Label6(8)"
      Tab(2).Control(15)=   "Label6(9)"
      Tab(2).Control(16)=   "Label50(3)"
      Tab(2).Control(17)=   "cmbBank"
      Tab(2).Control(18)=   "Label50(1)"
      Tab(2).Control(19)=   "cboDepositType"
      Tab(2).Control(20)=   "Label6(4)"
      Tab(2).Control(21)=   "cmbDptAmtType"
      Tab(2).Control(22)=   "Label6(10)"
      Tab(2).Control(23)=   "txtDNC(1)"
      Tab(2).Control(24)=   "Label19(34)"
      Tab(2).Control(25)=   "cboFund"
      Tab(2).Control(26)=   "txtDNC(0)"
      Tab(2).Control(27)=   "flxDeposit"
      Tab(2).Control(28)=   "txtDptDetails"
      Tab(2).Control(29)=   "txtDptAmount"
      Tab(2).Control(30)=   "cmdSetAmtType"
      Tab(2).Control(31)=   "cmdDptDelete"
      Tab(2).Control(32)=   "cmdDptNew"
      Tab(2).Control(33)=   "cmdDptSave"
      Tab(2).Control(34)=   "cmdDptEdit"
      Tab(2).Control(35)=   "cmdDptRefund"
      Tab(2).Control(36)=   "cmdDptPrint"
      Tab(2).Control(37)=   "cmdDptCancel"
      Tab(2).Control(38)=   "txtDate"
      Tab(2).Control(39)=   "cmdSetDptType"
      Tab(2).Control(40)=   "txtOSDpt"
      Tab(2).Control(41)=   "Frame1(1)"
      Tab(2).Control(42)=   "cmdDptExpenses"
      Tab(2).Control(43)=   "cmdNCList"
      Tab(2).ControlCount=   44
      TabCaption(3)   =   "&Maintenance Entry"
      TabPicture(3)   =   "frmLeasee2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(0)"
      Tab(3).Control(1)=   "fmeEventHistory"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lease &Agreement"
      TabPicture(4)   =   "frmLeasee2.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeTenancyDetails"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Account &History"
      TabPicture(5)   =   "frmLeasee2.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label72(1)"
      Tab(5).Control(1)=   "Label1(8)"
      Tab(5).Control(2)=   "Label1(7)"
      Tab(5).Control(3)=   "Label1(1)"
      Tab(5).Control(4)=   "Label1(2)"
      Tab(5).Control(5)=   "Label1(5)"
      Tab(5).Control(6)=   "Label1(6)"
      Tab(5).Control(7)=   "Label1(3)"
      Tab(5).Control(8)=   "Label1(4)"
      Tab(5).Control(9)=   "Label1(35)"
      Tab(5).Control(10)=   "Label1(33)"
      Tab(5).Control(11)=   "Label1(39)"
      Tab(5).Control(12)=   "Label1(32)"
      Tab(5).Control(13)=   "Label1(31)"
      Tab(5).Control(14)=   "Label1(40)"
      Tab(5).Control(15)=   "Label1(41)"
      Tab(5).Control(16)=   "Label1(34)"
      Tab(5).Control(17)=   "Label1(37)"
      Tab(5).Control(18)=   "Label1(36)"
      Tab(5).Control(19)=   "Label1(38)"
      Tab(5).Control(20)=   "Label1(42)"
      Tab(5).Control(21)=   "flxACHistorySplit"
      Tab(5).Control(22)=   "flxACHistory"
      Tab(5).Control(23)=   "cmdCopyReceipt"
      Tab(5).Control(24)=   "cmdSentStByEmail"
      Tab(5).Control(25)=   "cmdPrintHistory"
      Tab(5).Control(26)=   "cmdPrintStatement"
      Tab(5).Control(27)=   "cmdPrintReceipt"
      Tab(5).ControlCount=   28
      TabCaption(6)   =   "Letters && Email"
      TabPicture(6)   =   "frmLeasee2.frx":0972
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label50(41)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label50(42)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "flxEmails"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "flxLetters"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdViewLetter"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "cmdEmail"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "cmdDelLetter"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "cmdResendEmail"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "cmdPrintEmail"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "cmdPrintWord"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).ControlCount=   10
      TabCaption(7)   =   "&Memo/Attachments"
      TabPicture(7)   =   "frmLeasee2.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame2"
      Tab(7).Control(1)=   "Frame17"
      Tab(7).Control(2)=   "Frame8"
      Tab(7).ControlCount=   3
      Begin VB.CommandButton cmdAddNewContacts 
         Caption         =   "&New"
         Height          =   345
         Left            =   -66600
         TabIndex        =   359
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditContacts 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   -65040
         TabIndex        =   358
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   -68760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   348
         Text            =   "frmLeasee2.frx":09AA
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrintWord 
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
         Left            =   9516
         Picture         =   "frmLeasee2.frx":09B2
         Style           =   1  'Graphical
         TabIndex        =   347
         Top             =   2880
         Width           =   410
      End
      Begin VB.CommandButton cmdPrintReceipt 
         Caption         =   "Print Receipt"
         Height          =   385
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   344
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdPrintEmail 
         Caption         =   "&Print Email"
         Height          =   315
         Left            =   9900
         TabIndex        =   339
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton cmdResendEmail 
         Caption         =   "Resend Email"
         Height          =   315
         Left            =   7950
         TabIndex        =   338
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   -72120
         TabIndex        =   337
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   -72120
         TabIndex        =   335
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   -72840
         TabIndex        =   334
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   -72840
         TabIndex        =   333
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   -72120
         TabIndex        =   332
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveDefaults 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66960
         TabIndex        =   324
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditDefaults 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -69000
         TabIndex        =   323
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelDefaults 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64920
         TabIndex        =   322
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtDefault 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   -72840
         TabIndex        =   321
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNC 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   320
         Top             =   1320
         Width           =   320
      End
      Begin VB.CommandButton cmdTaxList 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         TabIndex        =   319
         Top             =   1800
         Width           =   320
      End
      Begin VB.TextBox txtCodeVat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68565
         Locked          =   -1  'True
         TabIndex        =   318
         Top             =   1800
         Width           =   1080
      End
      Begin VB.TextBox txtSLControl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   317
         Top             =   2280
         Width           =   1080
      End
      Begin VB.CommandButton cmdSLC 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   316
         Top             =   2280
         Width           =   320
      End
      Begin VB.Frame fmeTenancyDetails 
         Caption         =   "Lease Details"
         Height          =   3315
         Left            =   -74880
         TabIndex        =   294
         Top             =   1200
         Width           =   11115
         Begin MSForms.TextBox txtSERVICECHARGEFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   315
            Top             =   2250
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSCPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   314
            Top             =   1860
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtBASERENTFREQ 
            Height          =   315
            Left            =   7980
            TabIndex        =   313
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBRPayable 
            Height          =   315
            Left            =   7980
            TabIndex        =   312
            Top             =   690
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent:"
            Height          =   195
            Index           =   21
            Left            =   6045
            TabIndex        =   311
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Frequency:"
            Height          =   195
            Index           =   22
            Left            =   6045
            TabIndex        =   310
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/C Frequency:"
            Height          =   195
            Index           =   25
            Left            =   6045
            TabIndex        =   309
            Top             =   2250
            Width           =   1065
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Charge:"
            Height          =   195
            Index           =   24
            Left            =   6045
            TabIndex        =   308
            Top             =   1860
            Width           =   1110
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            Height          =   195
            Index           =   18
            Left            =   315
            TabIndex        =   307
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Termination Date:"
            Height          =   195
            Index           =   20
            Left            =   315
            TabIndex        =   306
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Index           =   17
            Left            =   315
            TabIndex        =   305
            Top             =   720
            Width           =   750
         End
         Begin MSForms.TextBox txtStartDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   304
            Top             =   720
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEndDate 
            Height          =   315
            Left            =   2340
            TabIndex        =   303
            Top             =   1080
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            MaxLength       =   10
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TextBox13 
            Height          =   315
            Left            =   2340
            TabIndex        =   302
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtRentReviewDate 
            Height          =   315
            Left            =   7980
            TabIndex        =   301
            Top             =   1470
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TextBox15 
            Height          =   315
            Left            =   2340
            TabIndex        =   300
            Top             =   2250
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Type:"
            Height          =   195
            Index           =   19
            Left            =   315
            TabIndex        =   299
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Rent Review Date:"
            Height          =   195
            Index           =   23
            Left            =   6045
            TabIndex        =   298
            Top             =   1470
            Width           =   1680
         End
         Begin MSForms.TextBox txtLeaseId 
            Height          =   315
            Left            =   7995
            TabIndex        =   297
            Top             =   2850
            Width           =   2985
            VariousPropertyBits=   746604575
            BackColor       =   16448250
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lease Ref:"
            Height          =   195
            Index           =   16
            Left            =   5955
            TabIndex        =   296
            Top             =   2910
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSForms.CheckBox chkHoldingOver 
            Height          =   345
            Left            =   285
            TabIndex        =   295
            Top             =   1470
            Width           =   2280
            VariousPropertyBits=   1015031839
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4022;609"
            Value           =   "0"
            Caption         =   "Holding Over"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Memo"
         Height          =   3705
         Left            =   -74880
         TabIndex        =   289
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtUnitMemo 
            Height          =   2955
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   293
            Top             =   210
            Width           =   10750
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7500
            TabIndex        =   292
            Top             =   3270
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   291
            Top             =   3270
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   290
            Top             =   3270
            Width           =   1125
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -74880
         TabIndex        =   284
         Top             =   4080
         Width           =   11175
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   315
            Left            =   9870
            Style           =   1  'Graphical
            TabIndex        =   287
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   315
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   286
            Top             =   240
            Width           =   1110
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   315
            Left            =   8685
            Style           =   1  'Graphical
            TabIndex        =   285
            Top             =   240
            Width           =   1110
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   180
            TabIndex        =   288
            Top             =   240
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
      Begin VB.Frame Frame2 
         Caption         =   "Comments:"
         Height          =   735
         Left            =   -74880
         TabIndex        =   279
         Top             =   4800
         Width           =   11175
         Begin VB.CommandButton cmdRCCCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   282
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   281
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdRCCEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7440
            TabIndex        =   280
            Top             =   240
            Width           =   1125
         End
         Begin MSForms.TextBox txtRCCComments 
            Height          =   315
            Left            =   180
            TabIndex        =   283
            Top             =   240
            Width           =   4890
            VariousPropertyBits=   746604575
            Size            =   "8625;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdDelLetter 
         Caption         =   "&Delete Letter"
         Height          =   315
         Left            =   6000
         TabIndex        =   276
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Email Letter"
         Height          =   315
         Left            =   7758
         TabIndex        =   261
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame fmeEventHistory 
         BackColor       =   &H0000FFFF&
         Caption         =   "Maintenance History"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   11265
         Begin VB.TextBox dtpReportedDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   45
            Top             =   520
            Width           =   960
         End
         Begin VB.TextBox dtpRemindDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   9380
            MaxLength       =   10
            TabIndex        =   50
            Top             =   520
            Width           =   960
         End
         Begin VB.TextBox dtpDateCompleted 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   47
            Top             =   520
            Width           =   960
         End
         Begin VB.CommandButton cmdMType 
            Caption         =   "..."
            Height          =   315
            Left            =   1680
            TabIndex        =   44
            Top             =   520
            Width           =   255
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10360
            TabIndex        =   51
            Top             =   520
            Width           =   225
         End
         Begin VB.TextBox txtEventHistoryID 
            Appearance      =   0  'Flat
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
            Left            =   7440
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CommandButton cmdNewEvent 
            Caption         =   "&New"
            Height          =   315
            Left            =   6120
            TabIndex        =   38
            Top             =   2760
            Width           =   915
         End
         Begin VB.CommandButton cmdEditEvent 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7470
            TabIndex        =   39
            Top             =   2775
            Width           =   915
         End
         Begin VB.CommandButton cmdCancelEvent 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10170
            TabIndex        =   41
            Top             =   2775
            Width           =   915
         End
         Begin VB.CommandButton cmdSaveEvent 
            Caption         =   "&Save"
            Height          =   315
            Left            =   8820
            TabIndex        =   40
            Top             =   2775
            Width           =   915
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridEventHistory 
            Height          =   1845
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   3254
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   15329508
            ForeColorSel    =   0
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
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.TextBox txtEventTenantID 
            Height          =   315
            Left            =   3960
            TabIndex        =   86
            Top             =   120
            Visible         =   0   'False
            Width           =   1275
            VariousPropertyBits=   746604575
            BackColor       =   12640511
            Size            =   "2249;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   315
            Left            =   2880
            TabIndex        =   46
            Top             =   510
            Width           =   2655
            VariousPropertyBits=   746604571
            Size            =   "4683;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTaskOwner 
            Height          =   315
            Left            =   6480
            TabIndex        =   48
            Top             =   510
            Width           =   1460
            VariousPropertyBits=   746604571
            Size            =   "2575;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtContact 
            Height          =   315
            Left            =   7920
            TabIndex        =   49
            Top             =   510
            Width           =   1460
            VariousPropertyBits=   746604571
            Size            =   "2575;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboEventType 
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   505
            Width           =   1575
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "2778;556"
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
         Begin VB.Label Label42 
            Caption         =   "Remind   Date:"
            Height          =   435
            Index           =   6
            Left            =   9380
            TabIndex        =   59
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label42 
            Caption         =   "Contact:"
            Height          =   255
            Index           =   5
            Left            =   7920
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label Label42 
            Caption         =   "Task Owner:"
            Height          =   255
            Index           =   4
            Left            =   6480
            TabIndex        =   57
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Date  Actioned:"
            Height          =   435
            Index           =   3
            Left            =   5520
            TabIndex        =   56
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label42 
            Caption         =   "Alarm"
            Height          =   195
            Index           =   7
            Left            =   10360
            TabIndex        =   55
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label42 
            Caption         =   "Reported Datet:"
            Height          =   435
            Index           =   1
            Left            =   1920
            TabIndex        =   54
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label42 
            Caption         =   "Event Type:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Description:"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   52
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Property Maintenance"
         Height          =   5085
         Index           =   0
         Left            =   -75000
         TabIndex        =   239
         Top             =   360
         Width           =   11505
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "Add &Diary Entry"
            Height          =   355
            Left            =   5160
            TabIndex        =   247
            Top             =   4560
            Width           =   1395
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   9960
            TabIndex        =   246
            Top             =   4560
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   7680
            TabIndex        =   245
            Top             =   4560
            Width           =   1395
         End
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "&Add Job"
            Height          =   355
            Left            =   3600
            TabIndex        =   244
            Top             =   4560
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   240
            Top             =   4455
            Width           =   3375
            Begin VB.OptionButton optDiary 
               Caption         =   "View Diary"
               Height          =   255
               Left            =   2160
               TabIndex        =   243
               Top             =   160
               Width           =   1095
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "View Jobs"
               Height          =   255
               Left            =   1080
               TabIndex        =   242
               Top             =   160
               Width           =   1095
            End
            Begin VB.OptionButton optAll 
               Caption         =   "View All"
               Height          =   255
               Left            =   120
               TabIndex        =   241
               Top             =   160
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   3765
            Left            =   120
            TabIndex        =   248
            Top             =   690
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   6641
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
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   9120
            TabIndex        =   259
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   3000
            TabIndex        =   258
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Reported"
            Height          =   480
            Index           =   2
            Left            =   2145
            TabIndex        =   257
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   256
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   255
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9120
            TabIndex        =   254
            Top             =   255
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   8040
            TabIndex        =   253
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Item / Dairy Entry"
            Height          =   495
            Index           =   4
            Left            =   4200
            TabIndex        =   252
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   5400
            TabIndex        =   251
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   6840
            TabIndex        =   250
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   10200
            TabIndex        =   249
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdPrintStatement 
         Caption         =   "Print Statement"
         Height          =   385
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdPrintHistory 
         Caption         =   "Print Account History"
         Height          =   385
         Left            =   -74925
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   5160
         Width           =   1755
      End
      Begin VB.CommandButton cmdSentStByEmail 
         Caption         =   "Email Statement"
         Height          =   385
         Left            =   -68125
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   5160
         Width           =   1515
      End
      Begin VB.CommandButton cmdCopyReceipt 
         Caption         =   "Copy"
         Height          =   385
         Left            =   -64845
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   5160
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdNCList 
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
         Height          =   260
         Left            =   -70880
         TabIndex        =   135
         Top             =   850
         Width           =   255
      End
      Begin VB.CommandButton cmdDptExpenses 
         Caption         =   "Expenses"
         Height          =   315
         Left            =   -70142
         Style           =   1  'Graphical
         TabIndex        =   203
         ToolTipText     =   "Refund or Expenses"
         Top             =   5180
         Width           =   1080
      End
      Begin VB.CommandButton cmdViewLetter 
         Caption         =   "&Print Letter"
         Height          =   315
         Left            =   10140
         TabIndex        =   201
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Grouping:"
         Height          =   780
         Index           =   1
         Left            =   -66330
         TabIndex        =   197
         Top             =   1140
         Width           =   2775
         Begin VB.OptionButton optExitingGroup 
            Caption         =   "Add to Existing Group"
            Height          =   255
            Left            =   40
            TabIndex        =   146
            Top             =   420
            Width           =   1840
         End
         Begin VB.OptionButton optNewGroup 
            Caption         =   "Create a New Group"
            Height          =   255
            Left            =   40
            TabIndex        =   145
            Top             =   200
            Value           =   -1  'True
            Width           =   1815
         End
         Begin MSForms.ComboBox cboGroup 
            Height          =   285
            Left            =   1920
            TabIndex        =   147
            Top             =   420
            Width           =   780
            VariousPropertyBits=   679495709
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1376;503"
            TextColumn      =   2
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "987;5000"
         End
      End
      Begin VB.TextBox txtOSDpt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EDEDED&
         Height          =   285
         Left            =   -65775
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   840
         Width           =   2220
      End
      Begin VB.CommandButton cmdSetDptType 
         Caption         =   "- -"
         Height          =   285
         Left            =   -70890
         TabIndex        =   138
         Top             =   1560
         Width           =   285
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73260
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   136
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdDptCancel 
         Caption         =   "Cancel Deposit"
         Height          =   315
         Left            =   -64900
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   5180
         Width           =   1345
      End
      Begin VB.CommandButton cmdDptPrint 
         Caption         =   "Print Receipt"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   5180
         Width           =   1275
      End
      Begin VB.CommandButton cmdDptRefund 
         Caption         =   "Refund"
         Height          =   315
         Left            =   -71331
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Refund or Expenses"
         Top             =   5180
         Width           =   1080
      End
      Begin VB.CommandButton cmdDptEdit 
         Caption         =   "Edit Deposit"
         Height          =   315
         Left            =   -68953
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   5180
         Width           =   1145
      End
      Begin VB.CommandButton cmdDptSave 
         Caption         =   "Save Deposit"
         Height          =   315
         Left            =   -67699
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   5180
         Width           =   1230
      End
      Begin VB.CommandButton cmdDptNew 
         Caption         =   "New Deposit"
         Height          =   315
         Left            =   -72600
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   5180
         Width           =   1160
      End
      Begin VB.CommandButton cmdDptDelete 
         Caption         =   "Delete Deposit"
         Height          =   315
         Left            =   -66360
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   5180
         Width           =   1350
      End
      Begin VB.CommandButton cmdSetAmtType 
         Caption         =   "- -"
         Height          =   285
         Left            =   -66765
         TabIndex        =   143
         Top             =   1200
         Width           =   285
      End
      Begin VB.TextBox txtDptAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68340
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   840
         Width           =   1860
      End
      Begin VB.TextBox txtDptDetails 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68340
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   139
         Top             =   480
         Width           =   4785
      End
      Begin VB.Frame fmeTenantAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         Left            =   -66480
         TabIndex        =   19
         Top             =   4440
         Width           =   10995
         Begin VB.CommandButton cmdCancelTenantAddress 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   9750
            TabIndex        =   82
            Top             =   4710
            Width           =   1035
         End
         Begin VB.CommandButton cmdSaveTenantAddress 
            Caption         =   "&Save"
            Height          =   315
            Left            =   7935
            TabIndex        =   81
            Top             =   4710
            Width           =   1035
         End
         Begin VB.CommandButton cmdEditTenantAddress 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   6120
            TabIndex        =   80
            Top             =   4710
            Width           =   1035
         End
         Begin MSForms.CommandButton cmdSendMail 
            Height          =   330
            Index           =   1
            Left            =   10200
            TabIndex        =   278
            Top             =   3120
            Width           =   420
            Size            =   "732;591"
            Picture         =   "frmLeasee2.frx":0B6C
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdSendMail 
            Height          =   330
            Index           =   0
            Left            =   3960
            TabIndex        =   277
            Top             =   3000
            Width           =   420
            Size            =   "732;591"
            Picture         =   "frmLeasee2.frx":678E
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   " Alternative Address:"
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
            Index           =   34
            Left            =   6300
            TabIndex        =   85
            Top             =   615
            Width           =   1560
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   " Lessee Address:"
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
            Index           =   33
            Left            =   240
            TabIndex        =   84
            Top             =   495
            Width           =   1230
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice To:"
            Height          =   195
            Index           =   15
            Left            =   3075
            TabIndex        =   83
            Top             =   210
            Width           =   765
         End
         Begin MSForms.ComboBox cboInvoiceTo 
            Height          =   315
            Left            =   4080
            TabIndex        =   79
            Top             =   210
            Width           =   3000
            VariousPropertyBits=   1820346395
            DisplayStyle    =   3
            Size            =   "5292;556"
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "0"
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   36
            Left            =   6300
            TabIndex        =   78
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   37
            Left            =   6300
            TabIndex        =   77
            Top             =   2790
            Width           =   750
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Index           =   35
            Left            =   6300
            TabIndex        =   76
            Top             =   960
            Width           =   585
         End
         Begin MSForms.TextBox txtContact2 
            Height          =   315
            Left            =   7680
            TabIndex        =   66
            Top             =   900
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine1 
            Height          =   315
            Left            =   7680
            TabIndex        =   67
            Top             =   1290
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   68
            Top             =   1620
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine3 
            Height          =   315
            Left            =   7680
            TabIndex        =   69
            Top             =   1950
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillAddressLine4 
            Height          =   315
            Left            =   7680
            TabIndex        =   70
            Top             =   2280
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   40
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillPostCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   71
            Top             =   2730
            Width           =   1815
            VariousPropertyBits=   746604571
            MaxLength       =   12
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillFax 
            Height          =   315
            Left            =   7680
            TabIndex        =   75
            Top             =   4290
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBillTelephone 
            Height          =   315
            Left            =   7680
            TabIndex        =   74
            Top             =   3900
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine2 
            Height          =   315
            Left            =   7680
            TabIndex        =   73
            Top             =   3510
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail2 
            Height          =   315
            Left            =   7680
            TabIndex        =   72
            Top             =   3120
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Index           =   0
            Left            =   6300
            TabIndex        =   65
            Top             =   3210
            Width           =   405
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   40
            Left            =   6300
            TabIndex        =   64
            Top             =   3585
            Width           =   795
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   39
            Left            =   6300
            TabIndex        =   63
            Top             =   4320
            Width           =   285
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   38
            Left            =   6300
            TabIndex        =   62
            Top             =   3945
            Width           =   525
         End
         Begin MSForms.TextBox txtHOFax 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            Top             =   4200
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOTelephone 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   3810
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDirectLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   3420
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtEmail1 
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   3030
            Width           =   2445
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "4313;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   36
            Top             =   3060
            Width           =   405
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   30
            Left            =   240
            TabIndex        =   35
            Top             =   3435
            Width           =   795
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   34
            Top             =   4170
            Width           =   285
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   33
            Top             =   3795
            Width           =   525
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   31
            Top             =   2640
            Width           =   750
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   585
         End
         Begin MSForms.TextBox txtContact1 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   810
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   100
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine1 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   1200
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine2 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   1530
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine3 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   1860
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOAddressLine4 
            Height          =   315
            Left            =   1440
            TabIndex        =   24
            Top             =   2190
            Width           =   2985
            VariousPropertyBits=   746604571
            MaxLength       =   70
            Size            =   "5265;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtHOPostCode 
            Height          =   315
            Left            =   1440
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
            VariousPropertyBits=   746604571
            MaxLength       =   12
            Size            =   "3201;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   4065
            Index           =   0
            Left            =   120
            Top             =   600
            Width           =   4665
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   3945
            Left            =   6120
            Top             =   720
            Width           =   4665
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   2100
         Left            =   -74925
         TabIndex        =   89
         Top             =   555
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3704
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDeposit 
         Height          =   2805
         Left            =   -74880
         TabIndex        =   196
         Top             =   2205
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   4948
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
         BackColorSel    =   15329508
         ForeColorSel    =   0
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
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLetters 
         Height          =   2265
         Left            =   75
         TabIndex        =   202
         Top             =   600
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   3995
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistorySplit 
         Height          =   2100
         Left            =   -74925
         TabIndex        =   218
         Top             =   3000
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3704
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmails 
         Height          =   2025
         Left            =   75
         TabIndex        =   340
         Top             =   3240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   3572
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxContacts 
         Height          =   4395
         Left            =   -74880
         TabIndex        =   349
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   7
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblContacts 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         Height          =   195
         Index           =   1
         Left            =   -73560
         TabIndex        =   357
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblContacts 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address Name"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   355
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label lblContacts 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   -71880
         TabIndex        =   354
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblContacts 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel"
         Height          =   195
         Index           =   3
         Left            =   -70080
         TabIndex        =   353
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblContacts 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice to Email"
         Height          =   195
         Index           =   5
         Left            =   -68160
         TabIndex        =   352
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label lblContacts 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         Height          =   195
         Index           =   4
         Left            =   -69240
         TabIndex        =   351
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblContacts 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Email"
         Height          =   195
         Index           =   6
         Left            =   -65880
         TabIndex        =   350
         Top             =   480
         Width           =   585
      End
      Begin MSForms.TextBox txtDNC 
         Height          =   285
         Index           =   0
         Left            =   -70560
         TabIndex        =   343
         Top             =   840
         Visible         =   0   'False
         Width           =   375
         VariousPropertyBits=   746604575
         BackColor       =   -2147483648
         BorderStyle     =   1
         Size            =   "661;503"
         BorderColor     =   -2147483648
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emails:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   42
         Left            =   75
         TabIndex        =   342
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letters:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   41
         Left            =   75
         TabIndex        =   341
         Top             =   360
         Width           =   585
      End
      Begin MSForms.TextBox txtNominalCode 
         Height          =   285
         Left            =   -69960
         TabIndex        =   331
         Top             =   1320
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
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
         Caption         =   "Nominal Code:"
         Height          =   195
         Index           =   0
         Left            =   -72315
         TabIndex        =   330
         Top             =   1320
         Width           =   1020
      End
      Begin MSForms.TextBox txtNominalCodeName 
         Height          =   285
         Left            =   -68565
         TabIndex        =   329
         Top             =   1320
         Width           =   2055
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3625;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVatCode 
         Height          =   285
         Index           =   0
         Left            =   -69960
         TabIndex        =   328
         Top             =   1800
         Width           =   1080
         BackColor       =   16777215
         Size            =   "1905;503"
         BorderStyle     =   1
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         Height          =   195
         Index           =   0
         Left            =   -72315
         TabIndex        =   327
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Sales LedgerControl Account:"
         Height          =   195
         Left            =   -72315
         TabIndex        =   326
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         Height          =   1935
         Index           =   5
         Left            =   -72840
         Top             =   960
         Width           =   6975
      End
      Begin MSForms.TextBox txtSLControlName 
         Height          =   285
         Left            =   -68565
         TabIndex        =   325
         Top             =   2280
         Width           =   2055
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3625;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFund 
         Height          =   285
         Left            =   -68340
         TabIndex        =   144
         Top             =   1560
         Width           =   1860
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         TextColumn      =   2
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;7055"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         Height          =   195
         Index           =   34
         Left            =   -69480
         TabIndex        =   260
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   42
         Left            =   -64920
         TabIndex        =   230
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   38
         Left            =   -69960
         TabIndex        =   229
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Index           =   36
         Left            =   -71880
         TabIndex        =   228
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   37
         Left            =   -70920
         TabIndex        =   227
         Top             =   2760
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   34
         Left            =   -73920
         TabIndex        =   226
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   41
         Left            =   -65880
         TabIndex        =   225
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   40
         Left            =   -66960
         TabIndex        =   224
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   31
         Left            =   -74925
         TabIndex        =   223
         Top             =   2760
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   32
         Left            =   -74925
         TabIndex        =   222
         Top             =   2760
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   39
         Left            =   -68040
         TabIndex        =   221
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   33
         Left            =   -74925
         TabIndex        =   220
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   35
         Left            =   -72600
         TabIndex        =   219
         Top             =   2760
         Width           =   585
      End
      Begin MSForms.TextBox txtDNC 
         Height          =   285
         Index           =   1
         Left            =   -73260
         TabIndex        =   204
         Top             =   840
         Width           =   2655
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "4683;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   10
         Left            =   -64200
         TabIndex        =   198
         Top             =   1995
         Width           =   735
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Group"
         Size            =   "1296;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   4
         Left            =   -71280
         TabIndex        =   195
         Top             =   360
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   -72360
         TabIndex        =   194
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   6
         Left            =   -67440
         TabIndex        =   193
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   5
         Left            =   -68640
         TabIndex        =   192
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   -74040
         TabIndex        =   191
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   190
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   7
         Left            =   -66240
         TabIndex        =   189
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   8
         Left            =   -65040
         TabIndex        =   188
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label72 
         Height          =   195
         Index           =   1
         Left            =   -74955
         TabIndex        =   187
         Top             =   360
         Width           =   11400
      End
      Begin MSForms.ComboBox cmbDptAmtType 
         Height          =   285
         Left            =   -68340
         TabIndex        =   142
         Top             =   1200
         Width           =   1575
         VariousPropertyBits=   1753237535
         BackColor       =   -2147483628
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2787;503"
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
         Object.Width           =   "776;5000"
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   4
         Left            =   -72240
         TabIndex        =   172
         Top             =   1995
         Width           =   975
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Deposit Type"
         Size            =   "1720;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDepositType 
         Height          =   285
         Left            =   -73260
         TabIndex        =   137
         Top             =   1560
         Width           =   2355
         VariousPropertyBits=   1753237535
         BackColor       =   -2147483628
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4154;503"
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
         Object.Width           =   "987;5000"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   195
         Index           =   1
         Left            =   -74775
         TabIndex        =   171
         Top             =   1560
         Width           =   375
      End
      Begin MSForms.ComboBox cmbBank 
         Height          =   285
         Left            =   -73260
         TabIndex        =   134
         Top             =   480
         Width           =   2655
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4683;503"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "987;5000"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account:"
         Height          =   195
         Index           =   3
         Left            =   -74775
         TabIndex        =   169
         Top             =   480
         Width           =   975
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   9
         Left            =   -65130
         TabIndex        =   168
         Top             =   1995
         Width           =   855
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Refund/Exp"
         Size            =   "1508;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   8
         Left            =   -66000
         TabIndex        =   167
         Top             =   1995
         Width           =   495
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "O/S"
         Size            =   "873;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   7
         Left            =   -66960
         TabIndex        =   166
         Top             =   1995
         Width           =   615
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Amount"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   6
         Left            =   -70200
         TabIndex        =   165
         Top             =   1995
         Width           =   615
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Details"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   5
         Left            =   -70920
         TabIndex        =   164
         Top             =   1995
         Width           =   495
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Type"
         Size            =   "873;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   3
         Left            =   -73080
         TabIndex        =   163
         Top             =   1995
         Width           =   495
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Date"
         Size            =   "873;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   2
         Left            =   -73560
         TabIndex        =   162
         Top             =   1995
         Width           =   255
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Ref"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   161
         Top             =   1995
         Width           =   375
         ForeColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Bank"
         Size            =   "661;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   195
         Index           =   0
         Left            =   -74835
         TabIndex        =   160
         Top             =   1995
         Visible         =   0   'False
         Width           =   180
         ForeColor       =   0
         VariousPropertyBits=   276824083
         Caption         =   "ID"
         Size            =   "317;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O/S:"
         Height          =   195
         Index           =   9
         Left            =   -66360
         TabIndex        =   153
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         Height          =   195
         Index           =   7
         Left            =   -69480
         TabIndex        =   152
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details:"
         Height          =   195
         Index           =   6
         Left            =   -69480
         TabIndex        =   151
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code:"
         Height          =   195
         Index           =   5
         Left            =   -74775
         TabIndex        =   150
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Type:"
         Height          =   195
         Index           =   11
         Left            =   -69480
         TabIndex        =   149
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   4
         Left            =   -74775
         TabIndex        =   133
         Top             =   1200
         Width           =   375
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00000000&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   -74895
         Top             =   1995
         Width           =   11340
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00004040&
         BorderWidth     =   3
         Height          =   1935
         Index           =   4
         Left            =   -72840
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   -74880
         TabIndex        =   356
         Top             =   480
         Width           =   11295
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      Height          =   15
      Index           =   1
      Left            =   0
      Top             =   2670
      Width           =   12015
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Height          =   15
      Index           =   0
      Left            =   0
      Top             =   2670
      Width           =   12015
   End
End
Attribute VB_Name = "frmLeasee2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_OUT_STANDING_REFUND = 8
Private Const COL_IS_REFUND = 9
Private Const COL_BANK_CODE = 11
Private Const COL_NOMINAL_CODE = 12
Private Const COL_AMOUNT_TYPE_CODE = 13
Private Const COL_DEPOSIT_TYPE_CODE = 14
Private Const COL_NC_NAME = 15
Private Const COL_FUND = 16
Private Const COL_RECON = 17

Private GROUP_NO     As String

Public LOAD_TENANT_TENANTID As String
Private CLIENT_ID As String
Private PROPERTY_ID As String

Dim NEWMODE_ As Boolean
Dim COPYMODE_ As Boolean
Dim SEARCHTenantMODE_ As Boolean
Dim M_HISTORY_NEW_ENTRY_ As Boolean
Dim BANK_PAYMENT_NEW_ENTRY_ As Boolean
Dim IMAGE_FILE_NAME_ As String
Dim bDEPOSIT_HELD As Boolean
Dim yDEPOSIT As Byte                   'using for deposit(0-2), refund(3)
Dim bLeaseSetup As Boolean, cCurDepAmt As Currency, iGridRow As Integer
Dim bSortingCol1 As Boolean, bSortingCol2 As Boolean, bSortingCol3 As Boolean
Dim szSel      As String

Dim szaTenantBalance()  As String
Dim cOriRAmt            As Currency
Dim dataProperty()      As String

Private Type SendDemandByEmail
   szLesseeID    As String
   szLesseeEmail As String
   colAtt        As Collection
End Type
Private uLessee   As SendDemandByEmail

Private Sub cboBankId_Change()
   If Not IsNull(cboBankId) And cboBankId <> "" Then
      txtBankSortCode.text = cboBankId.Column(2)
      txtBranchName.text = cboBankId.Column(3)
      txtBankAddress1.text = cboBankId.Column(4)
      txtBankAddress2.text = cboBankId.Column(5)
      txtBankAddress3.text = cboBankId.Column(6)
      txtBankPostCode.text = cboBankId.Column(7)
   End If
End Sub

Private Sub cboClientList_Click()
   If fmeTenantLookup.Visible = False Then Exit Sub

   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer

   If cboClientList.Column(0) <> "ALL" Then
'     Filter properties

'     Get the count of the properties of the client
      For i = 1 To UBound(dataProperty, 2)
         If dataProperty(4, i) = cboClientList.Column(0) Then
            TotalRow = TotalRow + 1
         End If
      Next i

      TotalCol = UBound(dataProperty, 1)
      ReDim Data(TotalCol, TotalRow) As String

'     Load the properties in the combo
      Data(0, 0) = "ALL"
      Data(1, 0) = "All Properties"
      j = 1
      For i = 1 To UBound(dataProperty, 2)
         If dataProperty(4, i) = cboClientList.Column(0) Then
            For K = 0 To TotalCol - 1
               Data(K, j) = dataProperty(K, i)
            Next K
            j = j + 1
         End If
      Next i
      
      cboPropertyList.Column() = Data()
      
'     Filter the Lessee
   Else
'     Reload all Properties in the combo
      Data = dataProperty
      cboPropertyList.Column() = Data()
   End If
   
   Exit Sub
   
   Dim adoConn    As New ADODB.Connection
   Dim adoRST     As New ADODB.Recordset
   Dim szSQL      As String

   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

   If cboClientList.Column(0) = "ALL" Then
      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboClientList.Column(0) & "' " & _
              "ORDER BY PropertyID;"
   End If

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboPropertyList.Column() = Data()
'*************************************************************************************************
'  All tenants
   If optBoth.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                      "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "WHERE IQ.PropertyID = '" & cboPropertyList.Column(0) & "' AND "
         Else
            szSQL = szSQL + "WHERE "
         End If
         szSQL = szSQL + "IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "WHERE IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If
                 
      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"

      PopulateTenantLookup szSQL, adoConn
   End If

'  Current tenants Only
   If optCurrentTenant.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                      "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
         szSQL = szSQL + "AND IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If

      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"

      PopulateTenantLookup szSQL, adoConn
   End If

'  Deleted tenants only
   If optExTenant.Value Then
       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                    "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
               "FROM Tenants AS T LEFT JOIN " & _
                    "[" & _
                    "SELECT U.UnitName, L.SageAccountNumber " & _
                    "From Units AS U, LeaseDetails AS L, Property AS P, P.PropertyID, P.ClientID " & _
                    "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                       "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE (T.Comments) IS NOT NULL OR T.Comments<>'' "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
         szSQL = szSQL + "AND IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If

      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"

      PopulateTenantLookup szSQL, adoConn
   End If
'**************************************************************************************************

NoRes:
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cboClientList_GotFocus()
   SelTxtInCtrl cboClientList
End Sub

Private Sub cboPropertyList_Click()
   If fmeTenantLookup.Visible = False Then Exit Sub

   Dim i          As Integer

'  Reset the grid - open all hidden rows
   For i = 1 To gridTenantLookup.Rows - 1
      gridTenantLookup.RowHeight(i) = 240
   Next i
   
   If cboPropertyList.Column(0) <> "ALL" Then
      For i = 1 To gridTenantLookup.Rows - 1
         If gridTenantLookup.TextMatrix(i, 5) <> cboPropertyList.Column(0) Then
            gridTenantLookup.RowHeight(i) = 0
         End If
      Next i
   Else              'Check client
      If cboClientList.Column(0) <> "ALL" Then
         For i = 1 To gridTenantLookup.Rows - 1
            If gridTenantLookup.TextMatrix(i, 6) <> cboClientList.Column(0) Then
               gridTenantLookup.RowHeight(i) = 0
            End If
         Next i
      End If
   End If
   
   
   Exit Sub
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

'  All tenants
   If optBoth.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                      "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "WHERE IQ.PropertyID = '" & cboPropertyList.Column(0) & "' AND "
         Else
            szSQL = szSQL + "WHERE "
         End If
         szSQL = szSQL + "IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "WHERE IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If

      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"

      PopulateTenantLookup szSQL, adoConn
   End If

'  Current tenants Only
   If optCurrentTenant.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, P.PropertyID, P.ClientID " & _
                   "From Units AS U, LeaseDetails AS L, Property AS P " & _
                   "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                      "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
         szSQL = szSQL + "AND IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If

      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"
'Debug.Print szSQL
      PopulateTenantLookup szSQL, adoConn
   End If

'  Deleted tenants only
   If optExTenant.Value Then
       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                    "iif(isnull(Comments),'CURRENT','DELETED') as Notes " & _
               "FROM Tenants AS T LEFT JOIN " & _
                    "[" & _
                    "SELECT U.UnitName, L.SageAccountNumber " & _
                    "From Units AS U, LeaseDetails AS L, Property AS P, P.PropertyID, P.ClientID " & _
                    "Where U.UnitNumber = L.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                       "L.Status = TRUE "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE (T.Comments) IS NOT NULL OR T.Comments<>'' "

      If cboClientList.Column(0) <> "ALL" Then
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
         szSQL = szSQL + "AND IQ.ClientID = '" & cboClientList.Column(0) & "' "
      Else
         If cboPropertyList.Column(0) <> "ALL" Then
            szSQL = szSQL + "AND IQ.PropertyID = '" & cboPropertyList.Column(0) & "' "
         End If
      End If

      szSQL = szSQL + "ORDER BY T.SageAccountNumber;"

      PopulateTenantLookup szSQL, adoConn
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cboPropertyList_GotFocus()
   SelTxtInCtrl cboPropertyList
End Sub

Private Sub chkCombEmail_Click()
   If chkCombEmail.Value Then
      chkEmailDmd.Value = True
      chkEmailSt.Value = True
   End If
End Sub

Private Sub chkEmailDmd_Click()
   If chkEmailDmd.Value = False Then
      chkEmailSC.Value = False
'      chkCombEmail.Value = False
      chkEmailSt.Value = False
   End If
End Sub

Private Sub chkEmailSC_Click()
   If chkEmailSC.Value Then
      chkEmailDmd.Value = True
   End If
End Sub

Private Sub chkEmailSt_Click()
   If chkEmailSt.Value = False Then
      chkCombEmail.Value = False
   End If
End Sub

Private Sub cmbBank_Click()
   cmdNCList.SetFocus
End Sub

Private Sub cmdAddDiary_Click()
   If txtTenantID.text = "" Then Exit Sub

   With frmMaintananceDairy
      .CallingForm = "L"          'Calling from lessee form
      .isEdit = False
      .RecordType = "D"
      .lblJobName.Caption = "Diary Name"
      .Label1.Caption = "Diary Entry No."
      Load frmMaintananceDairy
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With

   Me.Enabled = False
End Sub

Private Sub cmdCancel_Click()
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
'   If NEWMODE_ Then SageCustomerAccCombo cboSageAccountNumber

   NEWMODE_ = False
   SEARCHTenantMODE_ = True

   txtName.Enabled = True
   cmdTenantLookup.Enabled = True
   cmdTenantLookup.Visible = True

   If txtTenantID.text = "" Then Exit Sub

   tabTenant.Enabled = True

   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth

   ComponentInFrameClearMode Me, fmeTenant, ClearOnlyTextBoxes
   txtTenantID.BackColor = &HF1F9EE
   txtTenantID.Locked = True
   cmdTenantLookup.Visible = True
End Sub

Private Sub cmdCancelBank_Click()
   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
End Sub

Private Sub cmdCancelTenantAddress_Click()
   ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
End Sub

Private Sub cmdClientLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmClientNew4
   frmClientNew4.LOAD_CLINT_CLIENTID = txtClientID.text
   frmClientNew4.Show
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachmentInCombo cmbFiles, "Tenants", txtTenantID.text
   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdCopy_Click()
   If txtTenantID.text = "" Then
      ShowMsgInTaskBar "Please select a Lessee to continue."
      Exit Sub
   End If

   COPYMODE_ = True
   NEWMODE_ = True
   SEARCHTenantMODE_ = False

   txtClient.text = ""
   txtProperty.text = ""
   txtUnit.text = ""
   txtDeposit.text = ""

   txtClient.Enabled = True
   txtProperty.Enabled = True
   txtUnit.Enabled = True
   txtDeposit.Enabled = True

   txtTenantID.Enabled = True
   txtTenantID.Locked = False
   txtName.Enabled = True
   txtCompanyName.Enabled = True

   cmdNew.Enabled = False
   cmdEdit.Enabled = False
   cmdCopy.Enabled = False
   cmdDeleteLessee.Enabled = False

   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   cmdClose.Enabled = True

'   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
   ComponentInFrameClearMode Me, fmeEventHistory, ClearBoth
'   ConfigurFlxEventHistory
   ComponentInFrameClearMode Me, Frame8, ClearBoth
   ComponentInFrameClearMode Me, Frame17, ClearBoth
End Sub

Private Sub cmdCopyReceipt_Click()
   If flxACHistory.row < 1 Then
      MsgBox "Please select a transaction from the grid.", vbInformation + vbOKOnly, "Selection"
      flxACHistory.SetFocus
      Exit Sub
   End If
'
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + cmdCopyReceipt.Top + _
'                      Me.Top + tabTenant.Top + 1160
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + _
'                       tabTenant.Left + cmdCopyReceipt.Left + 80

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "SRR" Then
      frmPopUpMenu.CallingFrom "LESSEE_SRR"
   Else
      frmPopUpMenu.CallingFrom "LESSEE_" & Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2)
      
      If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI" Or _
            Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SC" Then
         Dim adoConn As New ADODB.Connection
         Dim adoRST As New ADODB.Recordset
         Dim szSQL As String

         adoConn.Open getConnectionString
         
         szSQL = "SELECT DemandID " & _
                 "FROM DemandRecords " & _
                 "WHERE DmdSlNo = " & _
                     StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & " AND " & _
                    "TransactionType = " & _
                     IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", 1, 2) & ";"
         adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         frmPopUpMenu.TRANSACITON_ID_LONG = CLng(adoRST.Fields.Item(0).Value)
         frmPopUpMenu.TRANS_TYPE = IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", "INV", "CRN")

         adoRST.Close
         Set adoRST = Nothing
         adoConn.Close
         Set adoConn = Nothing
      End If
      If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR" Or _
            Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SA" Then
         frmPopUpMenu.CallingFrom "LESSEE_" & Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2)
      End If
   End If

   frmPopUpMenu.Show
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtTenantID.text, "Tenants"
   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdDeleteLessee_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a Lessee to continue.", vbInformation, "Delete Lessee"
      Exit Sub
   End If

   If MsgBox("Do you sure to delete this Lessee?", vbQuestion + vbYesNo, "Delete Lessee") = vbNo Then Exit Sub

'Check for is this Lessee occupying any unit or not
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szNewID As String

   adoConn.Open getConnectionString

'------------------------------------LEASE RECORDS-------------------------------------------------------------
   szSQL = "SELECT Comments " & _
           "FROM Tenants, LeaseDetails " & _
           "WHERE Tenants.SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      MsgBox "This lessee cannot be deleted because a lease agreement exists for the lessee." & Chr(10) & _
             "You must terminated the lease agreement before you delete this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"

      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------DEMAND RECORDS-------------------------------------------------------------
   szSQL = "SELECT DemandID " & _
           "FROM DemandRecords " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "UPDATE_SAGE = True;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      MsgBox "This lessee cannot be deleted because demands exits that have not been updated to SAGE." & Chr(10) & _
             "You must update these demands into SAGE before you can delete this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------RECEIPT TABLE-------------------------------------------------------------
   szSQL = "SELECT * " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "(OSAmount > 0 OR " & _
               "IsSageUpdate = FALSE OR " & _
               "ReceiptView = TRUE);"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      If adoRST!OSAmount > 0 Or adoRST!ReceiptView Then
         MsgBox "This lessee cannot be deleted because are some outstanding amounts in the receipt." & Chr(10) & _
                "You must clear down all outstanding amount of this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"
      ElseIf Not adoRST!IsSageUpdate Then
         MsgBox "This lessee cannot be deleted because receipts exits that have not been updated to SAGE." & Chr(10) & _
                "You must update these receipts into SAGE before you can delete this lessee.", vbExclamation + vbOKOnly, "Deleting Lessee"
      End If
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoRST.Close

'------------------------------------DELETE LESSEE-------------------------------------------------------------
   szSQL = "SELECT Comments " & _
           "FROM Tenants " & _
           "WHERE Tenants.SageAccountNumber = '" & txtTenantID.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If InStr("DELETED", adoRST.Fields.Item("Comments").Value) > 0 Then
      MsgBox "The lessee has been delete already.", vbExclamation + vbOKOnly, "Deleting Lessee"
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

'  Create New id for the Lessee to release the SAGE id
   szNewID = DelTenantID(adoConn)

   UpdateTenantComment adoConn

'  SWAP Lessee's old SAGE id with New ID
'  First swap the id into Tenants table
   SwapID "Tenants", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "DemandRecords", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "LeaseDetails", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "tlbReceipt", "SageAccountNumber", txtTenantID.text, szNewID, adoConn
   SwapID "tblPoA", "SageAccountNumber", txtTenantID.text, szNewID, adoConn

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   MsgBox "The Lessee has been deleted successfully.", vbInformation + vbOKOnly, "Deleting Lessee"

   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
'   SageCustomerAccCombo txtTenantID

   tabTenant.Enabled = False
   cmdTenantLookup.Enabled = True
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
End Sub

Private Sub UpdateTenantComment(adoConn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE Tenants " & _
           "SET Comments ='DELETED' " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "';"
   adoConn.Execute szSQL
End Sub

Private Sub SwapID(szTableName As String, szTableField As String, szOldID As String, szNewID As String, adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT " & szTableField & " " & _
           "FROM " & szTableName & " " & _
           "WHERE " & szTableField & " = '" & szOldID & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   While Not adoRST.EOF
      adoRST.Fields.Item(szTableField).Value = szNewID
      adoRST.Update
      adoRST.MoveNext
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub cmdDelLetter_Click()
   If MsgBox("Do you wish to delete the letter?", vbQuestion + vbYesNo, "Deleting letters") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim iRow As Integer

   adoConn.Open getConnectionString
   
   For iRow = 1 To flxLetters.Rows - 1
      If flxLetters.TextMatrix(iRow, 0) = "X" Then
         adoConn.Execute "DELETE * FROM tlbLetterReports WHERE Id = " & Val(flxLetters.TextMatrix(iRow, 1)) & ";"
      End If
   Next iRow

   LoadFlxLetter adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdDptCancel_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE (PrimaryCode.Value = 'DEPOSIT TYPE' OR " & _
                  "PrimaryCode.Value = 'REFUND TYPE' OR " & _
                  "PrimaryCode.Value = 'EXPENSES TYPE') AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!SC
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cboDepositType.Clear
   cboDepositType.Column() = Data()

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   ButtonHanlding DefaultMode
End Sub

Private Sub cmdDptDelete_Click()
   Dim connCon As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   If MsgBox("Are you sure to delete current deposit?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   connCon.Open getConnectionString

   ButtonHanlding DefaultMode
   populateGroupCombo connCon

   szSQL = "SELECT * " & _
           "FROM TenantDeposit " & _
           "WHERE RefundRef = " & flxDeposit.TextMatrix(iGridRow, 0) & ";"

   adoRST.Open szSQL, connCon, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
         MsgBox "This deposit could not be deleted. This deposit has refund deposit.", vbCritical + vbOKOnly, "Delete Not Possible"
         adoRST.Close
         Set adoRST = Nothing
         connCon.Close
         Set connCon = Nothing
         Exit Sub
   End If

   adoRST.Close

   connCon.Execute "UPDATE TenantDeposit SET Status = false, Deleted = true WHERE DepositID = " & flxDeposit.TextMatrix(flxDeposit.row, 0) & " OR RefundRef = " & flxDeposit.TextMatrix(iGridRow, 0) & ";"
   LoadFlxDeposit connCon

   Set adoRST = Nothing
   connCon.Close
   Set connCon = Nothing

   MsgBox "Deposit has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"
End Sub

Private Sub cmdDptEdit_Click()
   If Val(txtDptAmount.text) <> Val(txtOSDpt.text) And flxDeposit.TextMatrix(flxDeposit.row, 5) = "Deposit" Then
      MsgBox "You cannot edit this transaction. Refund has been booked against this deposit.", vbCritical + vbOKOnly, "Edit Deposit"
      Exit Sub
   End If
'  Check the bank reconciliation
   If flxDeposit.TextMatrix(flxDeposit.row, COL_RECON) <> "" Then
      MsgBox "You cannot edit this transaction. Bank receipt of the deposit has been reconciled.", vbCritical + vbOKOnly, "Edit Deposit"
      Exit Sub
   End If

   cCurDepAmt = CCur(txtDptAmount.text)
   ButtonHanlding EditMode

   If flxDeposit.TextMatrix(flxDeposit.row, 5) = "Deposit" Then yDEPOSIT = 2
   If flxDeposit.TextMatrix(flxDeposit.row, 5) = "Refund" Then
      cOriRAmt = CCur(txtDptAmount.text)
      yDEPOSIT = 31
   End If
   If flxDeposit.TextMatrix(flxDeposit.row, 5) = "Expenses" Then
      cOriRAmt = CCur(txtDptAmount.text)
      yDEPOSIT = 41
   End If

   iGridRow = flxDeposit.row
End Sub

Private Sub cmdDptRefund_Click()
   If flxDeposit.TextMatrix(flxDeposit.row, 5) <> "Deposit" Then
      MsgBox "Select a deposit statement to make a refund.", vbInformation + vbOKOnly, "Refund - Deposit"
      Exit Sub
   End If

   If Val(flxDeposit.TextMatrix(iGridRow, COL_OUT_STANDING_REFUND)) <= 0 Then
      MsgBox "Deposit has been refunded fully.", vbExclamation + vbOKOnly, "Refund - Deposit"
      Exit Sub
   End If

   txtDptAmount.text = txtOSDpt.text

   ButtonHanlding RefundMode

   yDEPOSIT = 3

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   adoConn.Open getConnectionString

   cboDepositType.Clear
   szSQL = "SELECT SC.Code as C, SC.Value as V " & _
           "FROM SecondaryCode AS SC " & _
           "WHERE SC.PrimaryCode = 'RTYP' " & _
           "ORDER BY SC.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!c
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cboDepositType.Column() = Data()

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdDptExpenses_Click()
   If flxDeposit.TextMatrix(flxDeposit.row, 5) <> "Deposit" Then
      MsgBox "Select a deposit statement to book an expenses.", vbInformation + vbOKOnly, "Expenses - Deposit"
      Exit Sub
   End If

   txtDptAmount.text = txtOSDpt.text

   ButtonHanlding ExpensesMode

   yDEPOSIT = 4

   'Load default Bank Code and Nominal Code
   If Not bLeaseSetup Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   adoConn.Open getConnectionString

   cboDepositType.Clear
   szSQL = "SELECT SC.Code as C, SC.Value as V " & _
           "FROM SecondaryCode AS SC " & _
           "WHERE SC.PrimaryCode = 'EXPTYP' " & _
           "ORDER BY SC.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!c
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cboDepositType.Column() = Data()

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   cmbBank.SetFocus
End Sub

Private Sub cmdDptNew_Click()
   ButtonHanlding NewEntryMode
   optNewGroup.Value = True
   cmbBank.SetFocus

   yDEPOSIT = 1

   'Load default Bank Code and Nominal Code
   If Not bLeaseSetup Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT C.spare1, C.spare2, N.Name " & _
           "FROM Client AS C LEFT OUTER JOIN NominalLedger AS N ON C.spare2 = N.Code " & _
           "WHERE C.ClientID = '" & txtClientID.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   cmbBank.Value = IIf(IsNull(adoRST.Fields.Item("spare1").Value), "", adoRST.Fields.Item("spare1").Value)
   txtDNC(1).ToolTipText = IIf(IsNull(adoRST.Fields.Item("spare2").Value), "", adoRST.Fields.Item("spare2").Value)
   txtDNC(1).text = IIf(IsNull(adoRST.Fields.Item("Name").Value), "", adoRST.Fields.Item("Name").Value)

   adoRST.Close

   cboDepositType.Clear
   szSQL = "SELECT SC.Code as C, SC.Value as V " & _
           "FROM SecondaryCode AS SC " & _
           "WHERE SC.PrimaryCode = 'DPTYP' " & _
           "ORDER BY SC.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!c
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cboDepositType.Column() = Data()

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   cmbBank.SetFocus
End Sub

Private Sub cmdDptPrint_Click()
   If flxDeposit.TextMatrix(iGridRow, 0) = "" Or flxDeposit.TextMatrix(iGridRow, COL_IS_REFUND) <> "" Then
      If flxDeposit.TextMatrix(iGridRow, COL_IS_REFUND) = "" Then
'         MsgBox "This is a refund.", vbInformation + vbOKOnly, "Print - Deposit"
'      Else
         MsgBox "Please select a deposit transaction from the grid.", vbCritical + vbOKOnly, "Print - Deposit"
      End If
      Exit Sub
   End If
   ' Passing the from and to date values to Crystal Reports
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\DepositReceipt.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue flxDeposit.TextMatrix(iGridRow, 0)
   Report.ParameterFields(2).AddCurrentValue CInt(flxDeposit.TextMatrix(iGridRow, 10))
   Report.ParameterFields(3).AddCurrentValue txtTenantID.text

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdDptSave_Click()
   If yDEPOSIT = 0 Then Exit Sub

   If cmbBank.text = "" Then
      MsgBox "Please select a Bank Code from the drop down list.", vbCritical + vbOKOnly, "Deposit"
      cmbBank.SetFocus
      Exit Sub
   End If
   If txtDNC(0).text = "" Then
      MsgBox "Please select a Nominal Code from the drop down list.", vbCritical + vbOKOnly, "Deposit"
      cmdNCList.SetFocus
      Exit Sub
   End If
   If txtDate.text = "" Then
      MsgBox "Please enter the Date.", vbCritical + vbOKOnly, "Deposit"
      txtDate.SetFocus
      Exit Sub
   End If
   If cmbDptAmtType.text = "" Then
      MsgBox "Please select a Amount Type from the list", vbCritical + vbOKOnly, "Deposit"
      cmbDptAmtType.SetFocus
      Exit Sub
   End If
   If txtDptAmount.text = "" Or Val(txtDptAmount.text) <= 0 Then
      MsgBox "Please enter correct Amount.", vbCritical + vbOKOnly, "Deposit"
      txtDptAmount.SetFocus
      Exit Sub
   End If
   If cboDepositType.text = "" Then
      MsgBox "Please select type.", vbCritical + vbOKOnly, "Deposit Type"
      cboDepositType.SetFocus
      Exit Sub
   End If
   If optExitingGroup.Value = True And cboGroup.text = "" Then
      MsgBox "Type the group number.", vbCritical + vbOKOnly, "Group Number"
      cboGroup.SetFocus
      Exit Sub
   End If
   If cboFund.text = "" Then
      MsgBox "Select the fund.", vbCritical + vbOKOnly, "Group Number"
      cboFund.SetFocus
      Exit Sub
   End If
   
   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
   Dim szSQL         As String
   Dim btYesNo       As Byte
   Dim szID          As String

   adoConn.Open getConnectionString
   btYesNo = 0
'///////////////////////////////////// ADD NEW DEPOSIT /////////////////////////////////////////////////////
   If yDEPOSIT = 1 Then
      Dim GP_ID As Integer

      szSQL = "SELECT MAX(GROUPNO) AS GPNO FROM TENANTDEPOSIT WHERE TenantID = '" & cboSageAccountNumber.text & "';"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      If (optNewGroup.Value = True) Then
         GP_ID = CInt(IIf(IsNull(adoRST.Fields.Item("GPNO").Value), 0, adoRST.Fields.Item("GPNO").Value))
         GP_ID = GP_ID + 1
      Else
         GP_ID = CInt(cboGroup.text)
      End If
      adoRST.Close

      szSQL = "SELECT * FROM TENANTDEPOSIT;"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         .AddNew
         szID = UniqueID()
         .Fields.Item("DepositID").Value = szID
         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = Format(txtDate.text, "dd mmmm yyyy")
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)
         .Fields.Item("DptVatRate").Value = 0
         .Fields.Item("DptRefund").Value = False
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("BankTransaction").Value = False               '? what is this for?
         .Fields.Item("OSRefund").Value = CCur(txtDptAmount.text)
         .Fields.Item("GroupNo").Value = GP_ID
         .Fields.Item("TransactionID").Value = "D" & CStr(TransID("D") + 1)
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With
      Set adoRST = Nothing
      btYesNo = MsgBox("Deposit has been saved successfully." & Chr(13) & "Do you want to print the receipt now?", vbInformation + vbYesNo, "Deposit")

'     Export the transaction as Bank Recipt
      Export2BankReceipt szID, adoConn
   End If
'////////////////////////////////////// EDIT DEPOSIT ///////////////////////////////////////////////////////
   If yDEPOSIT = 2 Then
      Dim RemovedGrpNo As String
      If (GROUP_NO <> cboGroup.text) Then
         RemovedGrpNo = GROUP_NO
         GROUP_NO = cboGroup.text
      End If

      szSQL = "SELECT * FROM TENANTDEPOSIT " & _
              "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 0) & "';"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         If Left(flxDeposit.TextMatrix(iGridRow, 2), 1) <> "D" Then
            adoConn.Execute "UPDATE TenantDeposit " & _
                            "SET OSRefund = OSRefund + " & .Fields.Item("DptAmount").Value & " - " & _
                                 Val(txtDptAmount.text) & " " & _
                                 "WHERE DepositID = " & .Fields.Item("RefundRef").Value & ";"
         End If
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = Format(txtDate.text, "dd mmmm yyyy")
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)
         .Fields.Item("DptVatRate").Value = 0
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("OSRefund").Value = CCur(IIf(txtOSDpt.text = "", 0, txtOSDpt.text))
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With

      If RemovedGrpNo <> "" Then
         szSQL = "SELECT GROUPNO FROM TENANTDEPOSIT " & _
                 "WHERE GROUPNO = " & RemovedGrpNo & " AND " & _
                     "TenantID = '" & txtTenantID.text & "' AND DptRefund = FALSE;"
         adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

         Dim isUpdate As Boolean
         isUpdate = IIf(adoRST.EOF, True, False)
         adoRST.Close

         If isUpdate Then
            szSQL = "SELECT GROUPNO FROM TENANTDEPOSIT " & _
                    "WHERE GROUPNO > " & RemovedGrpNo & " AND TenantID = '" & txtTenantID.text & "';"
            adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
            While Not adoRST.EOF
               adoRST.Fields.Item("GROUPNO").Value = CInt(adoRST.Fields.Item("GROUPNO").Value) - 1
               adoRST.MoveNext
            Wend
            adoRST.Close
         End If
      End If

      Set adoRST = Nothing
      ShowMsgInTaskBar "The Modifications have been saved successfully."
      btYesNo = vbNo

      UpdateBankReceipt flxDeposit.TextMatrix(iGridRow, 0), adoConn
   End If
'////////////////////////////////////// REFUND DEPOSIT /////////////////////////////////////////////////////
   If yDEPOSIT = 3 Then
      szSQL = "SELECT * FROM TENANTDEPOSIT;"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         .AddNew
         szID = UniqueID()
         .Fields.Item("DepositID").Value = szID
         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value                     'TRANSACTION AMT TYPE
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
         .Fields.Item("RefundRef").Value = flxDeposit.TextMatrix(iGridRow, 0) 'Deposit ID
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("GroupNo").Value = cboGroup.text
         .Fields.Item("DptRefund").Value = True
         .Fields.Item("TransactionID").Value = "R" & CStr(TransID("R") + 1)
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With
      Set adoRST = Nothing
      adoConn.Execute "UPDATE TenantDeposit " & _
                      "SET OSRefund = " & _
                           Val(flxDeposit.TextMatrix(iGridRow, COL_OUT_STANDING_REFUND)) - Val(txtDptAmount.text) & " " & _
                           "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 0) & "';"
      ShowMsgInTaskBar "Refund has been saved successfully.", "Y", "P"
      btYesNo = vbNo

'     Export the refund transaction as Bank Payment
      Export2BankPayment szID, adoConn
   End If
'////////////////////////////////////// EDIT REFUND /////////////////////////////////////////////////////
   If yDEPOSIT = 31 Then
      szSQL = "SELECT * FROM TENANTDEPOSIT " & _
              "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 0) & "';"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value                     'TRANSACTION AMT TYPE
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
         szID = .Fields.Item("RefundRef").Value                                     'Deposit ID
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("GroupNo").Value = cboGroup.text
         .Fields.Item("DptRefund").Value = True
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With
      Set adoRST = Nothing

      adoConn.Execute "UPDATE TenantDeposit " & _
                      "SET OSRefund = OSRefund + " & _
                           cOriRAmt - Val(txtDptAmount.text) & " " & _
                           "WHERE DepositID = '" & szID & "';"
      ShowMsgInTaskBar "Refund has been saved successfully.", "Y", "P"

      UpdateBankPayment flxDeposit.TextMatrix(iGridRow, 0), adoConn
   End If

'////////////////////////////////////// EXPENSES  ///////////////////////////////////////////////////////////
   If yDEPOSIT = 4 Then
      szSQL = "SELECT * FROM TENANTDEPOSIT;"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         .AddNew
         szID = UniqueID()
         .Fields.Item("DepositID").Value = szID
         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value                     'TRANSACTION AMT TYPE
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'EXPENSES
         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
         .Fields.Item("RefundRef").Value = flxDeposit.TextMatrix(iGridRow, 0) 'Deposit ID
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("GroupNo").Value = cboGroup.text
         .Fields.Item("DptRefund").Value = True
         .Fields.Item("TransactionID").Value = "E" & CStr(TransID("E") + 1)
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With
      Set adoRST = Nothing
      adoConn.Execute "UPDATE TenantDeposit " & _
                      "SET OSRefund = " & _
                           Val(flxDeposit.TextMatrix(iGridRow, COL_OUT_STANDING_REFUND)) - Val(txtDptAmount.text) & " " & _
                           "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 0) & "';"
      ShowMsgInTaskBar "The Expenses have been saved successfully.", "Y", "P"
      btYesNo = vbNo

'     Export the refund transaction as Bank Payment
      Export2BankPayment szID, adoConn
   End If
'////////////////////////////////////// EDIT EXPENSES /////////////////////////////////////////////////////
   If yDEPOSIT = 41 Then
      szSQL = "SELECT * FROM TENANTDEPOSIT " & _
              "WHERE DepositID = '" & flxDeposit.TextMatrix(iGridRow, 0) & "';"
      adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With adoRST
         .Fields.Item("TenantID").Value = cboSageAccountNumber.text
         .Fields.Item("BankCode").Value = cmbBank.Value
         .Fields.Item("NominalCode").Value = txtDNC(0).text
         .Fields.Item("DepositDate").Value = txtDate.text                           'TRANSACTION DATE
         .Fields.Item("DptType").Value = cboDepositType.Value
         .Fields.Item("DptAmtType").Value = cmbDptAmtType.Value                     'TRANSACTION AMT TYPE
         .Fields.Item("DptDetails").Value = txtDptDetails.text
         .Fields.Item("DptAmount").Value = CCur(txtDptAmount.text)                  'REFUND TOTAL
         .Fields.Item("DptVatRate").Value = 0                                       'ALWAYS 0
         szID = .Fields.Item("RefundRef").Value                                     'Deposit ID
         .Fields.Item("BCNAME").Value = cmbBank.text
         .Fields.Item("NCNAME").Value = txtDNC(1).text
         .Fields.Item("GroupNo").Value = cboGroup.text
         .Fields.Item("DptRefund").Value = True
         .Fields.Item("FundID").Value = cboFund.Value
         .Update
         .Close
      End With
      Set adoRST = Nothing

      adoConn.Execute "UPDATE TenantDeposit " & _
                      "SET OSRefund = OSRefund + " & _
                           cOriRAmt - Val(txtDptAmount.text) & " " & _
                           "WHERE DepositID = '" & szID & "';"
      ShowMsgInTaskBar "Expenses has been updated successfully.", "Y", "P"

      UpdateBankPayment flxDeposit.TextMatrix(iGridRow, 0), adoConn
   End If
'///////////////////////////////////////////////////////////////////////////////////////////////////////////
   LoadFlxDeposit adoConn
   LoadComboes adoConn
   populateGroupCombo adoConn

   If yDEPOSIT = 2 And btYesNo = 6 Then cmdDptPrint_Click           'Edit Deposit

   If yDEPOSIT = 1 And btYesNo = 6 Then          'Deposit
      flxDeposit.row = flxDeposit.Rows - 1
      iGridRow = flxDeposit.Rows - 1

      cmdDptPrint_Click
   End If

   adoConn.Close
   Set adoConn = Nothing

   ButtonHanlding DefaultMode
End Sub

Private Sub Export2BankPayment(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset
   
   szStr = "SELECT * FROM tlbBankPayment;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = txtClientID.text
      .Fields.Item("BANK_AC").Value = cmbBank.Value
      .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
      .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
      .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
      .Fields.Item("PROJ_REF").Value = cboDepositType.Value
      .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
      .Fields.Item("DEPT_ID").Value = cboFund.Value
      .Fields.Item("TRAN_DATE").Value = txtDate.text
      .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
      .Fields.Item("TAX_CODE").Value = "T9"
      .Fields.Item("VAT").Value = 0
      .Fields.Item("TransactionType").Value = 11
      .Fields.Item("TRANS").Value = "BP"
      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Fields.Item("TenantDeposit").Value = szID

      If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
         .Fields.Item("CT").Value = "C"
      End If

      .Update

      If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = cmbBank.Value
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = cboDepositType.Value
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = cboFund.Value
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("CT").Value = "C"
         .Fields.Item("TenantDeposit").Value = szID
         .Update
      End If

      .Close
   End With

   Set adoRST = Nothing
End Sub

Private Sub Export2BankReceipt(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset

   szStr = "SELECT * FROM tlbBankPayment;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("ClientID").Value = txtClientID.text
      .Fields.Item("BANK_AC").Value = cmbBank.Value
      .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
      .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
      .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
      .Fields.Item("PROJ_REF").Value = cboDepositType.Value
      .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
      .Fields.Item("DEPT_ID").Value = cboFund.Value
      .Fields.Item("TRAN_DATE").Value = txtDate.text
      .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
      .Fields.Item("TAX_CODE").Value = "T9"
      .Fields.Item("VAT").Value = 0
      .Fields.Item("TransactionType").Value = 12
      .Fields.Item("TRANS").Value = "BR"
      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
      .Fields.Item("TenantDeposit").Value = szID

      If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
         .Fields.Item("CT").Value = "C"
      End If

      .Update

      If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = cmbBank.Value
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = cboDepositType.Value
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = cboFund.Value
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("CT").Value = "C"
         .Fields.Item("TenantDeposit").Value = szID
         .Update
      End If

      .Close
   End With

   Set adoRST = Nothing
End Sub

Private Sub UpdateBankPayment(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset
   
   szStr = "SELECT * FROM tlbBankPayment " & _
           "WHERE TenantDeposit = '" & szID & "' AND " & _
                 "TransactionType = 11;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      
      If Not adoRST.EOF Then
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = cmbBank.Value
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = cboDepositType.Value
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = cboFund.Value
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("TenantDeposit").Value = szID
      
         If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
            .Fields.Item("CT").Value = "C"
         End If
   
         .Update
         .Close
      End If
   
      szStr = "SELECT * FROM tlbBankPayment " & _
              "WHERE TenantDeposit = '" & szID & "' AND " & _
                    "TransactionType = 12;"

      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      If Not adoRST.EOF Then                                         'Contra Transation FOUND
         If .Fields.Item("BANK_AC").Value = .Fields.Item("NOMINAL_CODE").Value Then
            If cmbBank.Value <> txtDNC(0).text Then
               adoConn.Execute "DELETE * FROM tlbBankPayment " & _
                               "WHERE TenantDeposit = '" & szID & "' AND " & _
                                     "TransactionType = 11;"
               .Close
            Else
               .Fields.Item("ClientID").Value = txtClientID.text
               .Fields.Item("BANK_AC").Value = cmbBank.Value
               .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
               .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
               .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
               .Fields.Item("PROJ_REF").Value = cboDepositType.Value
               .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
               .Fields.Item("DEPT_ID").Value = cboFund.Value
               .Fields.Item("TRAN_DATE").Value = txtDate.text
               .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
               .Fields.Item("TAX_CODE").Value = "T9"
               .Fields.Item("VAT").Value = 0
               .Fields.Item("TransactionType").Value = 12
               .Fields.Item("TRANS").Value = "BR"
               .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
               .Fields.Item("CT").Value = "C"
               .Fields.Item("TenantDeposit").Value = szID
               .Update
               .Close
            End If
         End If
      Else              'IF CONTRA NOT FOUND
         If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
            szStr = "SELECT * FROM tlbBankPayment;"

            .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
            
            .AddNew
            .Fields.Item("MY_ID").Value = UniqueID()
            .Fields.Item("ClientID").Value = txtClientID.text
            .Fields.Item("BANK_AC").Value = cmbBank.Value
            .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
            .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
            .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
            .Fields.Item("PROJ_REF").Value = cboDepositType.Value
            .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
            .Fields.Item("DEPT_ID").Value = cboFund.Value
            .Fields.Item("TRAN_DATE").Value = txtDate.text
            .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
            .Fields.Item("TAX_CODE").Value = "T9"
            .Fields.Item("VAT").Value = 0
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
            .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
            .Fields.Item("CT").Value = "C"
            .Fields.Item("TenantDeposit").Value = szID
            .Update
            .Close
         End If
      End If
   End With

   Set adoRST = Nothing
End Sub

Private Sub UpdateBankReceipt(szID As String, adoConn As ADODB.Connection)
   Dim szStr   As String
   Dim adoRST  As New ADODB.Recordset

   szStr = "SELECT * FROM tlbBankPayment " & _
           "WHERE TenantDeposit = '" & szID & "' AND " & _
                 "TransactionType = 12;"

   With adoRST
      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

      If Not adoRST.EOF Then
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientID.text
         .Fields.Item("BANK_AC").Value = cmbBank.Value
         .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
         .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
         .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
         .Fields.Item("PROJ_REF").Value = cboDepositType.Value
         .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
         .Fields.Item("DEPT_ID").Value = cboFund.Value
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Fields.Item("TenantDeposit").Value = szID

         If cmbBank.Value = txtDNC(0).text Then      'Contra Transation
            .Fields.Item("CT").Value = "C"
         End If

         .Update
         .Close
      End If

      szStr = "SELECT * FROM tlbBankPayment " & _
              "WHERE TenantDeposit = '" & szID & "' AND " & _
                    "TransactionType = 11;"

      .Open szStr, adoConn, adOpenDynamic, adLockOptimistic
      If Not adoRST.EOF Then                                         'Contra Transation FOUND
         If .Fields.Item("BANK_AC").Value = .Fields.Item("NOMINAL_CODE").Value Then
            If cmbBank.Value <> txtDNC(0).text Then
               adoConn.Execute "DELETE * FROM tlbBankPayment " & _
                               "WHERE TenantDeposit = '" & szID & "' AND " & _
                                     "TransactionType = 11;"
               .Close
            Else
               .Fields.Item("ClientID").Value = txtClientID.text
               .Fields.Item("BANK_AC").Value = cmbBank.Value
               .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
               .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
               .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
               .Fields.Item("PROJ_REF").Value = cboDepositType.Value
               .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
               .Fields.Item("DEPT_ID").Value = cboFund.Value
               .Fields.Item("TRAN_DATE").Value = txtDate.text
               .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
               .Fields.Item("TAX_CODE").Value = "T9"
               .Fields.Item("VAT").Value = 0
               .Fields.Item("TransactionType").Value = 11
               .Fields.Item("TRANS").Value = "BP"
               .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
               .Fields.Item("CT").Value = "C"
               .Fields.Item("TenantDeposit").Value = szID
               .Update
               .Close
            End If
         End If
      Else              'IF CONTRA NOT FOUND
         If cmbBank.Value = txtDNC(0).text Then
            szStr = "SELECT * FROM tlbBankPayment;"

            .Open szStr, adoConn, adOpenDynamic, adLockOptimistic

            .AddNew
            .Fields.Item("MY_ID").Value = UniqueID()
            .Fields.Item("ClientID").Value = txtClientID.text
            .Fields.Item("BANK_AC").Value = cmbBank.Value
            .Fields.Item("PropertyID").Value = Left(txtProperty.text, 4)
            .Fields.Item("UNIT_ID").Value = Left(txtUnit.text, 8)
            .Fields.Item("DESCRIPTION").Value = txtDptDetails.text
            .Fields.Item("PROJ_REF").Value = cboDepositType.Value
            .Fields.Item("NOMINAL_CODE").Value = txtDNC(0).text
            .Fields.Item("DEPT_ID").Value = cboFund.Value
            .Fields.Item("TRAN_DATE").Value = txtDate.text
            .Fields.Item("NET_AMOUNT").Value = Val(txtDptAmount.text)
            .Fields.Item("TAX_CODE").Value = "T9"
            .Fields.Item("VAT").Value = 0
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
            .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
            .Fields.Item("CT").Value = "C"
            .Fields.Item("TenantDeposit").Value = szID
            .Update
            .Close
         End If
      End If
   End With

   Set adoRST = Nothing
End Sub

Private Function TransID(DRE As String) As Integer
   Dim iRow As Integer

   TransID = 0
   For iRow = 1 To flxDeposit.Rows - 1
      If Left(flxDeposit.TextMatrix(iRow, 2), 1) = DRE Then
         If Val(Mid(flxDeposit.TextMatrix(iRow, 2), 2, Len(flxDeposit.TextMatrix(iRow, 2)))) > TransID Then _
            TransID = Val(Mid(flxDeposit.TextMatrix(iRow, 2), 2, Len(flxDeposit.TextMatrix(iRow, 2))))
      End If
   Next iRow
End Function

Private Sub cmdEdit_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a Lessee to continue.", vbInformation, "Edit Lessee"
      Exit Sub
   End If
   NEWMODE_ = False
   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, fmeTenant, EditMode

   txtName.SetFocus
   tabTenant.Enabled = False
   cmdTenantLookup.Enabled = False
End Sub

Private Sub cmdEditBank_Click()
   BANK_PAYMENT_NEW_ENTRY_ = False
   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, EditMode
   cboBankId.Locked = True
   txtBankACNumber.Locked = True
End Sub

Private Sub cmdCancelDefaults_Click()
   txtNominalCode.text = txtDefault(0).text
   txtNominalCodeName.text = txtDefault(1).text
   txtCodeVat.text = txtDefault(2).text
   txtSLControl.text = txtDefault(3).text
   txtSLControlName.text = txtDefault(4).text
   lblVatCode(0).Caption = txtDefault(5).text

   cmdNC.Enabled = False
   cmdTaxList.Enabled = False
   cmdSLC.Enabled = False
   fmeTenant.Enabled = True

   cmdSaveDefaults.Enabled = False
   cmdCancelDefaults.Enabled = False
   cmdEditDefaults.Enabled = True
End Sub

Private Sub cmdEditDefaults_Click()
   txtDefault(0).text = txtNominalCode.text
   txtDefault(1).text = txtNominalCodeName.text
   txtDefault(2).text = txtCodeVat.text
   txtDefault(3).text = txtSLControl.text
   txtDefault(4).text = txtSLControlName.text
   txtDefault(5).text = lblVatCode(0).Caption
   
   cmdNC.Enabled = True
   cmdTaxList.Enabled = True
   cmdSLC.Enabled = True

   cmdSaveDefaults.Enabled = True
   cmdCancelDefaults.Enabled = True
   cmdEditDefaults.Enabled = False

   fmeTenant.Enabled = False
End Sub

Private Sub cmdEditMHistory_Click()
   If gridMaintenanceHistory.TextMatrix(1, 0) = "" Then Exit Sub

   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      frmMaintenanceJob.isEdit = True
      frmMaintenanceJob.CallingForm = "L"                         'Lessee
      frmMaintenanceJob.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintenanceJob
      frmMaintenanceJob.ZOrder 0
      frmMaintenanceJob.Show
   Else
      frmMaintananceDairy.isEdit = True
      frmMaintananceDairy.CallingForm = "L"                        'Lessee
      frmMaintananceDairy.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintananceDairy
      frmMaintananceDairy.ZOrder 0
      frmMaintananceDairy.Show
   End If
   Me.Enabled = False
End Sub

Private Sub cmdEditTenantAddress_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a lessee to continue.", vbCritical + vbOKOnly, "Lessee - Address"
      cmdTenantLookup.SetFocus
      Exit Sub
   End If
   ComponentInFrameEnableMode Me, fmeTenantAddress, EditMode
End Sub

Private Sub cmdEmail_Click()
   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub
   
   If szFromEmail = "" Or szSMTPserver = "" Then
      ShowMsgInTaskBar "Company email or SMTP server IP has not been setup.", "Y", "N"
      Exit Sub
   End If
   If cboInvoiceTo.Value = "B" Then
      If txtEmail2.text = "" Then
         ShowMsgInTaskBar "Lessee's email address is missing.", "Y", "N"
         Exit Sub
      End If
   Else
      If txtEmail1.text = "" Then
         ShowMsgInTaskBar "Lessee's email address is missing.", "Y", "N"
         Exit Sub
      End If
   End If
   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
   Dim szTemp As String
   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
   If FileExists(szTemp) Then
      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If

   Dim adoConn       As New ADODB.Connection
   Dim szID          As String
   Dim bEmailResult  As Boolean
   Dim szSQL         As String

   uLessee.szLesseeID = txtTenantID.text
   If cboInvoiceTo.Value = "B" Then
      uLessee.szLesseeEmail = txtEmail2.text
   Else
      uLessee.szLesseeEmail = txtEmail1.text
   End If

   szID = SelectedIDs

   If Len(szID) > 1 Then
      adoConn.Open getConnectionString

      adoConn.Execute "UPDATE tlbLetterReports SET isPrint = '';"

      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE Id IN (" & szID & ");"
      adoConn.Execute szSQL

      adoConn.Close
      Set adoConn = Nothing
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   szSQL = txtTenantID.text & "_" & UniqueID() & ".pdf"

   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing

'                    Attach the PDF in the email
   SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL

   bEmailResult = SendDemandByE_Mail("General Letter", "Please find the letter in the attachment.", "General Letter")

   If bEmailResult Then
      ShowMsgInTaskBar "Email sent.", "Y", "P"
   Else
      ShowMsgInTaskBar "No email sent.", "Y", "N"
   End If
End Sub

Private Sub SaveAttachment(szFile As String)
   Set uLessee.colAtt = New Collection
   uLessee.colAtt.Add szFile
End Sub

Private Function SelectedIDs() As String
   Dim i As Integer

   For i = 1 To flxLetters.Rows - 1
      If (flxLetters.TextMatrix(i, 0) = "X") Then
        SelectedIDs = SelectedIDs & flxLetters.TextMatrix(i, 1) & ", "
      End If
   Next i
   
   If Len(SelectedIDs) > 2 Then SelectedIDs = Left(SelectedIDs, Len(SelectedIDs) - 2)
End Function

Private Sub cmdGridTenantLookup_Click()
   fmeTenantLookup.Visible = False
End Sub

Private Sub cmdGridUnitLookup_Click(Index As Integer)
   fraList(0).Visible = False

   tabTenant.Enabled = True
End Sub

Private Sub cmdNC_Click()
   Dim szSQL As String

   szSQL = "SELECT S.Code, S.Name " & _
              "FROM NominalLedger AS S " & _
              "WHERE S.Posting AND " & _
                    "S.ClientID = '" & txtClientID.text & "';"
   
   szSel = "NC"
   LoadNC szSQL

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Left = tabTenant.Left + txtNominalCode.Left
   fraList(0).Top = tabTenant.Top + txtNominalCode.Top + txtNominalCode.Height + 10
   fraList(0).Width = 3520
   Picture1.Width = fraList(0).Width - 80
   flxSupplier(0).Width = Picture1.Width - 40
   fraList(0).Height = 2805
   Picture1.Height = fraList(0).Height - 80
   flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
End Sub

Private Sub LoadNC(szSQL As String)
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "NAME"
   flxSupplier(0).RowHeight(0) = 0

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1900
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40
   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)


   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "NAME"
   lblSearch2(0).Visible = False

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   flxSupplier(0).Sort = 1
   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdNCList_Click()
   LoadNominalCode

   tabTenant.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Width = 4815
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = 4695
   fraList(0).Left = tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "DH"         'Deposit Held
End Sub

Private Sub cmdNew_Click()
   NEWMODE_ = True
   SEARCHTenantMODE_ = False
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   
   adoConn.Open getConnectionString
   
   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If
   
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   tabTenant.Enabled = False
   ComponentInFrameEnableMode Me, fmeTenant, NewEntryMode
   chkPrintDmd.Value = True
   
   txtTenantID.Locked = False
   txtTenantID.BackColor = &H80000005
   cmdTenantLookup.Visible = False

   If UBound(szaChoice) > 0 Then
      If szaChoice(2) <> "" Then
         If InStr(szaChoice(2), "L") > 0 Then
            txtName.SetFocus
         End If
      Else
         txtTenantID.Locked = False
         txtTenantID.SetFocus
      End If
   End If

   ComponentInFrameClearMode Me, fmeTenantAddress, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fmeTenancyDetails, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth
'   ConfigurFlxEventHistory
   ComponentInFrameClearMode Me, fmeEventHistory, ClearBoth
   ComponentInFrameClearMode Me, Frame8, ClearBoth
   ComponentInFrameClearMode Me, Frame17, ClearBoth
End Sub

Private Sub ClearLesseeAddress(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
      txtTenantID.text = ""
      txtName.text = ""
      txtCompanyName.text = ""
'      cboSageAccountNumber.text = ""
      txtClient.text = ""
      txtProperty.text = ""
      txtUnit.text = ""
'      txtBankCode.text = ""
      txtDeposit.text = "0.00"
      txtBalance.text = "0.00"
      txtContact1.text = ""
      txtHOAddressLine1.text = ""
      txtHOAddressLine2.text = ""
      txtHOAddressLine3.text = ""
      txtHOAddressLine4.text = ""
      txtHOPostCode.text = ""
      txtEmail1.text = ""
      txtDirectLine1.text = ""
      txtHOTelephone.text = ""
      txtHOFax.text = ""
      txtContact.text = ""
      txtBillAddressLine1.text = ""
      txtBillAddressLine2.text = ""
      txtBillAddressLine3.text = ""
      txtBillAddressLine4.text = ""
      txtBillPostCode.text = ""
      txtEmail2.text = ""
      txtDirectLine2.text = ""
      txtBillTelephone.text = ""
      txtBillFax.text = ""
   
   Case CearEntryComponents.ClearOnlyComboBoxes
   
   Case CearEntryComponents.ClearBoth
      ClearLesseeAddress ClearOnlyTextBoxes
      ClearLesseeAddress ClearOnlyComboBoxes
   End Select
End Sub

Private Sub cmdNewBank_Click()
   BANK_PAYMENT_NEW_ENTRY_ = True
   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, NewEntryMode
   cboBankId.Locked = False
   txtBankACNumber.Locked = False
End Sub
'
'Private Sub cmdNewEvent_Click()
'   M_HISTORY_NEW_ENTRY_ = True
'   ComponentInFrameEnableMode me, fmeEventHistory, NewEntryMode
'End Sub

Private Sub cmdNewMHistory_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmMaintenanceJob
   With frmMaintenanceJob
      .CallingForm = "L"          'Calling from lessee form
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

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub cmdPrintEmail_Click()
   If flxEmails.TextMatrix(flxEmails.row, 6) = "" Then Exit Sub
   If flxEmails.TextMatrix(flxEmails.row, 0) = "" Then Exit Sub
      
   Dim szPath        As String
   Dim bLesseeEmail  As Boolean
   Dim szaDateTime() As String
   Dim szBody        As String
   Dim szSubject     As String
   Dim szAddress     As String
   Dim szLine        As String

   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
   bLesseeEmail = FileExists(szPath)

   If bLesseeEmail Then
      Open szPath For Input As #2

      Do Until EOF(2)
         Line Input #2, szLine
         If InStr(szLine, "Email sent on:") > 0 Then
            szLine = Mid(szLine, 15)
            szaDateTime = Split(szLine, "#")

            If UBound(szaDateTime) > 1 Then
               If flxEmails.TextMatrix(flxEmails.row, 6) = szaDateTime(2) Then
                  Line Input #2, szLine
                  szAddress = Mid(szLine, 15)            'Address
                  Line Input #2, szLine
                  szSubject = Mid(szLine, 15)            'Subject

                  Close #2
                  szBody = GetTheBody()
                  Exit Do
               End If
            End If
         End If
      Loop
      
      Close #2
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Email_Print.rpt")

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue szAddress
   Report.ParameterFields(2).AddCurrentValue szaDateTime(0)
   Report.ParameterFields(3).AddCurrentValue szaDateTime(1)
   Report.ParameterFields(4).AddCurrentValue szSubject
   Report.ParameterFields(5).AddCurrentValue szBody

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdPrintHistory_Click()
'   Frame5(6).Left = tabTenant.Left + 40
'   Frame5(6).Top = tabTenant.Top + tabTenant.Height - Frame5(6).Height - 40
'   CreateBorder4Frame Frame5(6), Shape4(32), Shape4(33)
'   Frame5(6).Visible = True
'   Exit Sub
   
   
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As New frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtTenantID.text
   Report.ParameterFields(2).AddCurrentValue txtName.text
   Report.ParameterFields(3).AddCurrentValue txtClient.text
   Report.ParameterFields(4).AddCurrentValue txtProperty.text
   Report.ParameterFields(5).AddCurrentValue txtUnit.text
   Report.ParameterFields(6).AddCurrentValue CDbl(txtBalance.text)

'MsgBox CCur(txtAcBal.text)
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintHistorySorted_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As New frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtTenantID.text
   Report.ParameterFields(2).AddCurrentValue txtName.text
   Report.ParameterFields(3).AddCurrentValue txtClient.text
   Report.ParameterFields(4).AddCurrentValue txtProperty.text
   Report.ParameterFields(5).AddCurrentValue txtUnit.text
   Report.ParameterFields(6).AddCurrentValue CDbl(txtBalance.text)
   Report.ParameterFields(7).AddCurrentValue IIf(optSortingAccountHistory(0), 0, IIf(optSortingAccountHistory(1), 1, 2))

'MsgBox CCur(txtAcBal.text)
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintJobSheet_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
      
   Report.ParameterFields(1).AddCurrentValue gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3)
   
   If (gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB") Then
      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "JOB SHEET"
   Else
      Report.ParameterFields(2).AddCurrentValue "Diary Entry"
      Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"
   End If
      
   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdPrintReceipt_Click()
   If flxACHistory.row > 0 And _
         InStr(flxACHistorySplit.TextMatrix(flxACHistorySplit.row, 2), "Receipt") > 0 Then
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report
      Dim rep As frmReport

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Receipt.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      If Report.HasSavedData Then Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue CLng(flxACHistorySplit.TextMatrix(flxACHistorySplit.row, 1))

      Set rep = New frmReport
      Load rep
      rep.LoadReportViewer Report
   End If
End Sub

Private Sub cmdPrintStatement_Click()
   If Val(txtBalance.text) = 0 Then
      ShowMsgInTaskBar "Statement will not be printed as account balance is 0.", "Y", "N"
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim szSQL   As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   adoConn.Open getConnectionString

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = '' " & _
           "WHERE  spare2 = 'Y';"
   adoConn.Execute szSQL

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = 'Y' " & _
           "WHERE  SageAccountNumber = ('" & txtTenantID.text & "');"
   adoConn.Execute szSQL

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False
   Report.ParameterFields(2).AddCurrentValue False
   Report.ParameterFields(3).AddCurrentValue "1"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdPrintWord_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String
   Dim rstReport As ADODB.Recordset
   Dim rstRst As New ADODB.Recordset

   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub

   adoConn.Open getConnectionString
   
   szSQL = "UPDATE tlbLetterReports SET isPrint = '';"
   adoConn.Execute szSQL
   
   For i = 1 To flxLetters.Rows - 1
    If (flxLetters.TextMatrix(i, 0) = "X") Then
      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE id = " & flxLetters.TextMatrix(i, 1) & ";"
      adoConn.Execute szSQL
    End If
   Next i

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   szSQL = txtTenantID.text & "_" & UniqueID() & ".DOC"
   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTWordForWindows
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing

   OpenFile szSQL, DB_PATH & "\AllStuff\Temp\"

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdPropLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmProperty2
   frmProperty2.LOAD_PROPERTY_PROPERTYID = Left(txtProperty, 4)
   frmProperty2.CLIENT_NAME = Trim(txtClient.text)
   frmProperty2.Show
End Sub

Private Sub cmdRCCCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   txtRCCComments.Locked = True
   cmdRCCEdit.Enabled = True
   cmdRCCSave.Enabled = False
   cmdRCCCancel.Enabled = False
End Sub

Private Sub cmdRCCEdit_Click()
   txtRCCComments.Locked = False
   cmdRCCEdit.Enabled = False
   cmdRCCSave.Enabled = True
   cmdRCCCancel.Enabled = True
End Sub

Private Sub cmdRCCSave_Click()
   If SaveComments("Tenants", "RCCComments", txtRCCComments.text, "SageAccountNumber", txtTenantID.text) Then
      ShowMsgInTaskBar "The comments have been saved successfully."
   End If
   txtRCCComments.Locked = True
   cmdRCCEdit.Enabled = True
   cmdRCCSave.Enabled = False
   cmdRCCCancel.Enabled = False
End Sub

Private Function GetTheBody() As String
   Dim szLine     As String
   Dim szPath     As String
   Dim szaDateTime() As String
                  
   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
   Open szPath For Input As #2

   Do Until EOF(2)
      Line Input #2, szLine
      If InStr(szLine, "Email sent on:") > 0 Then
         szLine = Mid(szLine, 15)
         szaDateTime = Split(szLine, "#")
         
         If UBound(szaDateTime) > 1 Then
            If flxEmails.TextMatrix(flxEmails.row, 6) = szaDateTime(2) Then
               Line Input #2, szLine
               Line Input #2, szLine
               Line Input #2, szLine
               While Not EOF(2) And InStr(szLine, "*****") = 0
            'Debug.Print szLine
                  GetTheBody = GetTheBody + szLine + (Chr(10) & Chr(13))
                  Line Input #2, szLine
               Wend
               If InStr(szLine, "*****") > 0 Then
                  GetTheBody = GetTheBody + Mid(szLine, 1, Len(szLine) - 5)
               End If
               Exit Do
            End If
         End If
      End If
   Loop
   
   Close #2
End Function

Private Sub cmdResendEmail_Click()
   If flxEmails.TextMatrix(flxEmails.row, 6) = "" Then Exit Sub
   If flxEmails.TextMatrix(flxEmails.row, 0) = "" Then Exit Sub
   If MsgBox("Would you like to send the email again?", vbQuestion + vbYesNo, "Email") = vbNo Then Exit Sub

   Dim bEmailResult     As Boolean
   Dim szBody           As String

   szBody = GetTheBody()
   bEmailResult = SendEmail(szFromEmail, flxEmails.TextMatrix(flxEmails.row, 3), _
                           flxEmails.TextMatrix(flxEmails.row, 4), _
                           szBody, , , _
                           , Me.txtTenantID.text, "Just Email")
   If bEmailResult Then
      ShowMsgInTaskBar "Email sent.", "Y", "P"

      SavingEmailInformation szBody
      LoadFlxEmails
   Else
      ShowMsgInTaskBar "No email sent.", "Y", "N"
   End If

End Sub

Private Sub SavingEmailInformation(szBody As String)
   Dim szLine  As String
   Dim szPath  As String

   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"
'  even the file does not exists, system will create the file and start adding text at the bottom of the file
   Open szPath For Append As #1

   szLine = "Email sent on:" & Format(Now, "dd/mm/yyyy") & "#" & Format(Now, "hh:nn") & "#" & UniqueID() & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Address:" & flxEmails.TextMatrix(flxEmails.row, 3) & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Subject:" + flxEmails.TextMatrix(flxEmails.row, 4) & Chr$(13) & Chr$(10)
   szLine = szLine + szBody
   szLine = szLine + "*****"

   Print #1, szLine
   Close #1
End Sub

Private Sub cmdSave_Click()
   If txtTenantID.text = "" Then
      MsgBox "Please select a lessee to continue.", vbExclamation, "No Lessee Selected"
      Exit Sub

   ElseIf txtName.text = "" Then
      MsgBox "Please enter a Lessee Name to continue.", vbExclamation, "No Lessee Name"
      txtName.SetFocus
      txtName.text = ""
      Exit Sub
   End If

   If chkCombEmail.Value And (Not chkEmailDmd.Value Or Not chkEmailSt.Value) Then
      MsgBox "Email Demand and Email Statement both have to be selected.", vbInformation + vbOKOnly, "Combined Email"
      If Not chkEmailDmd.Value Then chkEmailDmd.SetFocus
      If Not chkEmailSt.Value Then chkEmailSt.SetFocus
      Exit Sub
   End If

   If txtTenantID.text <> cboSageAccountNumber.text Then cboSageAccountNumber.text = txtTenantID.text

   If txtCompanyName.text = "" Then txtCompanyName.text = txtName.text

   If txtDeposit.text = "" Then txtDeposit.text = "0.00"
   If txtBalance.text = "" Then txtBalance.text = "0.00"

   If SaveTenantInformation Then
      If COPYMODE_ Then
         cmdSaveTenantAddress_Click
         COPYMODE_ = False
      End If
      ShowMsgInTaskBar "The lessee record has been saved successfully."
      NEWMODE_ = False
      ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
      SEARCHTenantMODE_ = True
      tabTenant.Enabled = True
      cmdTenantLookup.Enabled = True
      cmdTenantLookup.Visible = True
   Else
      txtTenantID.SetFocus
   End If

   If Not NEWMODE_ Then lblLeaseChanged.Caption = txtTenantID.text
End Sub

Public Function PopulateTenantLookup(ByVal sSQLQuery_ As String, adoConn As ADODB.Connection)
   Dim iRow As Integer

   iRow = 1

   ConfigurFlexGrid

   populateGrid adoConn, sSQLQuery_, gridTenantLookup
End Function

Private Sub cmdSaveBank_Click()
   Dim sSQLQuery As String, sWhere As String
   Dim adoConn As New ADODB.Connection
   Dim oResultSet As New ADODB.Recordset

   If cboBankId.text = "" Then
      MsgBox "Please select a bank from the drop down list.", vbCritical + vbOKOnly, "Bank Payment Details"
      cboBankId.SetFocus
      Exit Sub
   End If
   If cboPaymentMethod.text = "" Then
      MsgBox "Please select a payment method from the drop down list.", vbCritical + vbOKOnly, "Bank Payment Details"
      cboPaymentMethod.SetFocus
      Exit Sub
   End If
   If txtBankACName.text = "" Then
      MsgBox "Please enter bank account name.", vbCritical + vbOKOnly, "Bank Payment Details"
      txtBankACName.SetFocus
      Exit Sub
   End If
   If txtBankACNumber.text = "" Then
      MsgBox "Please enter bank account number.", vbCritical + vbOKOnly, "Bank Payment Details"
      txtBankACNumber.SetFocus
      Exit Sub
   End If
   
   adoConn.Open getConnectionString

   txtBankTenantID.text = txtTenantID.text

   sSQLQuery = "SELECT BankTenantID, BankID, BankACNumber, BankACName, " & _
                  "BankSortCode, IsDefaultAC, PaymentMethod, BacsRef " & _
               "FROM TenantBankDetails "
   sWhere = " WHERE BankTenantID = '" & txtTenantID.text & "' AND " & _
              "BankID = '" & cboBankId.Value & "' AND " & _
              "BankACNumber = '" & txtBankACNumber.text & "'"

   oResultSet.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   With oResultSet
      If BANK_PAYMENT_NEW_ENTRY_ And .EOF Then .AddNew
'else it is updating data
      !BankTenantID = txtBankTenantID.text
      !BankID = cboBankId.Value
      !BankACNumber = txtBankACNumber.text
      !BankACName = txtBankACName.text
      !BankSortCode = txtBankSortCode.text
      !IsDefaultAC = chkIsDefaultAC.Value
      !PaymentMethod = cboPaymentMethod.Value
      !BacsRef = txtBACSRef.text
      .Update
      .Close
   End With

   Set oResultSet = Nothing

   populateGrid adoConn, "SELECT BankTenantID, BankID, BankACNumber, BankACName, " & _
                           "BankSortCode, IsDefaultAC, PaymentMethod, BacsRef " & _
                         "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'", gridBank
   adoConn.Close
   Set adoConn = Nothing

   ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
End Sub
'
'Private Sub cmdSaveEvent_Click()
'   Dim adoConn As New ADODB.Connection
'   Dim oResultSet As New ADODB.Recordset
'   Dim szHeader As String, sSQLQuery As String
'
'   If cboEventType.Value = "" Then
'      MsgBox "Please select the event type.", vbCritical + vbOKOnly, "Event History"
'      cboEventType.SetFocus
'      Exit Sub
'   End If
'   If dtpReportedDate.text = "" Then
'      MsgBox "Please enter the reported date.", vbCritical + vbOKOnly, "Event History"
'      dtpReportedDate.SetFocus
'      Exit Sub
'   End If
'   If txtDescription.text = "" Then
'      MsgBox "Please enter the description of the event.", vbCritical + vbOKOnly, "Event History"
'      txtDescription.SetFocus
'      Exit Sub
'   End If
'   If txtTaskOwner.text = "" Then
'      MsgBox "Please enter the name of the task owner of the event.", vbCritical + vbOKOnly, "Event History"
'      txtTaskOwner.SetFocus
'      Exit Sub
'   End If
'   If txtContact.text = "" Then
'      MsgBox "Please enter the name of the contact of the event.", vbCritical + vbOKOnly, "Event History"
'      txtContact.SetFocus
'      Exit Sub
'   End If
'   adoConn.Open getConnectionString
'
'   ' Event Type
'   txtEventHistoryID.text = txtTenantID.text & "-" & cboEventType.Value & "-" & dtpReportedDate.text
'   txtEventTenantID.text = txtTenantID.text
'
'   sSQLQuery = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                      "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm, REMINDER_ID " & _
'               "FROM TenantEventHistory " & _
'               "WHERE EventHistoryID = '" & txtEventHistoryID.text & "'"
'
'   oResultSet.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
'
'   With oResultSet
'      If M_HISTORY_NEW_ENTRY_ And .EOF Then .AddNew
''else it is updating data
'      !EventHistoryID = txtEventHistoryID.text
'      !EventTenantID = txtTenantID.text
'      !EventType = cboEventType.Value
'      !ReportedDate = dtpReportedDate.text
'      !description = txtDescription.text
'      !DateCompleted = IIf(dtpDateCompleted.text = "", Null, dtpDateCompleted.text)
'      !TaskOwner = txtTaskOwner.text
'      !Contact = txtContact.text
'      !RemindDate = IIf(dtpRemindDate.text = "", Null, dtpRemindDate.text)
'
'      !Alarm = IIf(chkAlarm.Value = 1, True, False)
'      If chkAlarm.Value = 1 Then
'         If M_HISTORY_NEW_ENTRY_ Then
'            !Reminder_ID = NewReminder(Format(CDate(!RemindDate), "YYYYMMDD"), "083000", txtDescription.text, "TenantEventHistory", txtEventHistoryID.text)
'         Else
'            UpdateReminder !Reminder_ID, Format(CDate(!RemindDate), "YYYYMMDD"), "083000", txtDescription.text
'         End If
'      Else
'         ClearReminder !Reminder_ID
'      End If
'
'      .Update
'      .Close
'   End With
'
'   Set oResultSet = Nothing
'
''   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
''   sSQLQuery = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
''                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
''                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
''   populateGridSimply adoConn, sSQLQuery, gridEventHistory, szHeader
'   LoadGridMaintenanceHistory adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'   ComponentInFrameEnableMode me, fmeEventHistory, DefaultMode
'End Sub

Private Sub cmdSaveDefaults_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset
   Dim szSQL   As String

   adoConn.Open getConnectionString

   szSQL = "SELECT SLControl, DefaultNC, VAT_CODE " & _
           "FROM  Tenants " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "';"
   
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With adoRST
      .Fields.Item("DefaultNC").Value = IIf(IsNull(txtNominalCode.text), "", txtNominalCode.text)
      .Fields.Item("VAT_CODE").Value = IIf(IsNull(lblVatCode(0).Caption), "", lblVatCode(0).Caption)
      .Fields.Item("SLControl").Value = IIf(IsNull(txtSLControl.text), "", txtSLControl.text)
      .Update
      .Close
   End With

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   cmdNC.Enabled = False
   cmdTaxList.Enabled = False
   cmdSLC.Enabled = False
   fmeTenant.Enabled = True

   cmdSaveDefaults.Enabled = False
   cmdCancelDefaults.Enabled = False
   cmdEditDefaults.Enabled = True

   ShowMsgInTaskBar "It has been saved successfully", "Y", "P"
End Sub

Private Sub cmdSaveTenantAddress_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
   sSQLQuery = "SELECT SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                  "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                  "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                  "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                  "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                  "TenantMemo, Balance, Deposite, BankCode, spare1, EmailDmd " & _
                "FROM TENANTS " & _
                "WHERE TenantID = '" & txtTenantID.text & "'"

   If PostToDBUsingADODB(Me, fmeTenantAddress, adoConn, sSQLQuery, False) Then
      If Not COPYMODE_ Then MsgBox "The contact details of the tessee has been updated successfully", vbInformation
'       UpdateSAGECustomerAddress
   Else
       MsgBox "Error occured while updating the contact information", vbInformation
   End If

   adoConn.Close
   Set adoConn = Nothing

   ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
End Sub

Private Function SendDemandByE_Mail(szSub, szBody, szType) As Boolean
   Dim i As Integer

   SendDemandByE_Mail = SendEmail(szFromEmail, Trim(uLessee.szLesseeEmail), _
                                  szSub, _
                                  szBody, , , _
                                  uLessee.colAtt, uLessee.szLesseeID, szType)
End Function

Private Sub cmdSendMail_Click(Index As Integer)
   If Index = 0 Then
      If txtEmail1.text = "" Then Exit Sub
      Load frmSendMail
      frmSendMail.lblRecipientAddress.Caption = txtEmail1.text
   End If
   If Index = 1 Then
      If txtEmail2.text = "" Then Exit Sub
      Load frmSendMail
      frmSendMail.lblRecipientAddress.Caption = txtEmail2.text
   End If

   frmSendMail.Show
   Me.Enabled = False
End Sub

Private Sub cmdSentStByEmail_Click()
   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
   Dim szTemp As String
   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
   If FileExists(szTemp) Then
      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If

   uLessee.szLesseeID = txtTenantID.text
   If cboInvoiceTo.Value = "B" Then
      uLessee.szLesseeEmail = txtEmail2.text
   Else
      uLessee.szLesseeEmail = txtEmail1.text
   End If

   Dim szPath        As String
   Dim i             As Integer
   Dim bEmailResult  As Boolean

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = '' " & _
                   "WHERE  spare2 = 'Y';"

   adoConn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = 'Y' " & _
                   "WHERE  SageAccountNumber = ('" & txtTenantID.text & "');"

   adoConn.Close
   Set adoConn = Nothing

   '  Create the pdf file name of the statement
   szPath = txtTenantID.text & "_" & UniqueID() & ".pdf"

   CreatePDF_Statement txtTenantID.text, DB_PATH & "\AllStuff\Temp\" & szPath

   '  Attaching the statement to the email
   SaveAttachment DB_PATH & "\AllStuff\Temp\" & szPath

   bEmailResult = SendDemandByE_Mail("Account Statement", "Please find attachment your account statement.", "Account Statement")

   If bEmailResult Then
      MsgBox "Email sent.", vbInformation + vbOKOnly, "Lessee Statement"
   Else
      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Lessee Statement"
   End If
Exit Sub
'
'
'
'
'   Dim szEmailRecepent  As String
'   Dim colAtt           As New Collection
'
'   If cboInvoiceTo.Value = "B" Then
'      szEmailRecepent = txtEmail2.text
'   Else
'      szEmailRecepent = txtEmail1.text
'   End If
'
'   Dim szPath        As String
'   Dim bEmailResult  As Boolean
'   Dim adoConn       As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
''   CreateNonExistsFolder App.Path & "\Temp"
''   szPath = App.Path & "\Temp\" & txtTenantID.text & "_" & UniqueID() & ".pdf"
'   szPath = DB_PATH & "\AllStuff\Temp\" & txtTenantID.text & "_" & UniqueID() & ".pdf"
'
'   CreatePDF_Statement txtTenantID.text, szPath, adoConn
'
'   '  Attaching the demand to the email
'   If colAtt.count > 0 Then colAtt.Remove (1)
'   colAtt.Add szPath
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   bEmailResult = SendEmail(szFromEmail, szEmailRecepent, _
'                  "Account Statement", "Please see the attached of your account statement for details.", , , _
'                  colAtt, txtTenantID.text, "Account Statement")
'
'   MousePointer = vbDefault
'   If bEmailResult Then
'      MsgBox "Email sent.", vbInformation + vbOKOnly, "Demands"
'   Else
'      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Demands"
'   End If
End Sub

Private Sub CreatePDF_Statement(szLessee As String, szFileName As String)
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False          'Only outstanding statement
   'Report.ParameterFields(2).AddCurrentValue False          'specific lessee
   'Modified by anol 11 April 2016
   Report.ParameterFields(2).AddCurrentValue True          'specific lessee
   Report.ParameterFields(3).AddCurrentValue szLessee       'Lessee id

'   Transfer report into PDF file
   Report.ExportOptions.DiskFileName = szFileName
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing
End Sub
'
'Private Sub CreatePDF_Statement(szLessee As String, szFileName As String, adoConn As ADODB.Connection)
'   Dim szSQL      As String
'   Dim dLesBal    As Double
'   Dim adoRst     As New ADODB.Recordset
'   Dim reportApp  As New CRAXDRT.Application
'   Dim Report     As CRAXDRT.Report
'
'   szSQL = "SELECT T.Name, U.UnitName, P.PropertyName, C.ClientName " & _
'           "FROM Tenants AS T,  LeaseDetails AS L, Units AS U, Property AS P, Client AS C " & _
'           "WHERE Status = TRUE AND T.SageAccountNumber = '" & szLessee & "' AND " & _
'               "T.SageAccountNumber = L.SageAccountNumber AND " & _
'               "L.UnitNumber = U.UnitNumber AND " & _
'               "U.PropertyID = P.PropertyID AND " & _
'               "P.ClientID = C.ClientID;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      Set adoRst = Nothing
'      Exit Sub
'   End If
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   If Report.HasSavedData Then Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue szLessee
'   Report.ParameterFields(2).AddCurrentValue adoRst.Fields.Item("Name").Value
'   Report.ParameterFields(3).AddCurrentValue adoRst.Fields.Item("ClientName").Value
'   Report.ParameterFields(4).AddCurrentValue adoRst.Fields.Item("PropertyName").Value
'   Report.ParameterFields(5).AddCurrentValue adoRst.Fields.Item("UnitName").Value
'   dLesBal = CDbl(LesseeAccountBalance(adoConn, szLessee))
'   Report.ParameterFields(6).AddCurrentValue dLesBal
'
'   adoRst.Close
'   Set adoRst = Nothing
'
'   'Transfer report into PDF file
'   Report.ExportOptions.DiskFileName = szFileName
'   Report.ExportOptions.DestinationType = crEDTDiskFile
'   Report.ExportOptions.FormatType = crEFTPortableDocFormat
'   Report.ExportOptions.PDFExportAllPages = True
'   Report.Export False
'   Set Report = Nothing
'End Sub

Private Sub cmdSetAmtType_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   LoadRptAmtType cmbDptAmtType, "RECEIPT AMOUNT TYPE", adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdSetDptType_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If yDEPOSIT = 1 Then
      frmSecondaryCode.PRIMARY_CODE_SHOW = "DPTYP"
   Else
      If yDEPOSIT = 3 Then _
         frmSecondaryCode.PRIMARY_CODE_SHOW = "RTYP"
      If yDEPOSIT = 4 Then _
         frmSecondaryCode.PRIMARY_CODE_SHOW = "EXPTYP"
   End If
   Load frmSecondaryCode
   frmSecondaryCode.Show

   If yDEPOSIT = 1 Then
      LoadRptAmtType cboDepositType, "DEPOSIT TYPE", adoConn
   Else
      If yDEPOSIT = 3 Then _
         LoadRptAmtType cboDepositType, "REFUND TYPE", adoConn
      If yDEPOSIT = 4 Then _
         LoadRptAmtType cboDepositType, "EXPENSES TYPE", adoConn
   End If
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadRptAmtType(conComboBox As Control, szValue As String, adoConn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   conComboBox.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   conComboBox.Column() = szaData()
End Sub

Private Sub cmdSLC_Click()
   If Not bLeaseSetup Then
      ShowMsgInTaskBar "No lease has been setup", "Y", "P"
      Exit Sub
   End If
   
   Dim szSQL As String
   
   szSQL = "SELECT Code AS NCode, CAName " & _
           "FROM NominalLedger " & _
           "WHERE NOT CAFixed AND " & _
                 "ClientID = '" & txtClientID.text & "' AND " & _
                 "NOT ISNULL(CAName) AND CAType = 'S';"

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList(0).Left = tabTenant.Left + txtSLControl.Left
   fraList(0).Top = tabTenant.Top + txtSLControl.Top + txtSLControl.Height + 10
   fraList(0).Width = 3520
   Shape4(2).Width = fraList(0).Width - 200
   Picture1.Width = fraList(0).Width - 80
   flxSupplier(0).Width = Picture1.Width - 40
   fraList(0).Height = 1965
   Picture1.Height = fraList(0).Height - 80
   flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
   szSel = "SLC"

   LoadControlAccount szSQL
End Sub

Private Sub LoadControlAccount(szSQL As String)
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
'   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "NAME"
'   flxSupplier(0).TextMatrix(0, 2) = "Control Account"
   flxSupplier(0).RowHeight(0) = 0

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1900
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40
   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "NAME"
   lblSearch2(0).Visible = False

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
'         flxSupplier(0).TextMatrix(rRow, 2) = rstRec.Fields.Item(2).Value
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   flxSupplier(0).Sort = 1
   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdTaxList_Click()
   szSel = "TAX"
   LoadVAT

   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList(0).Left = lblVatCode(0).Left + tabTenant.Left
   fraList(0).Top = cmdTaxList.Top + cmdTaxList.Height + tabTenant.Top
   fraList(0).Width = 3520
   Picture1.Width = fraList(0).Width - 80
   flxSupplier(0).Width = Picture1.Width - 40
   fraList(0).Height = 2805
   Picture1.Height = fraList(0).Height - 80
   flxSupplier(0).Height = Picture1.Height - flxSupplier(0).Top
   cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width - 60
   Shape4(2).Width = flxSupplier(0).Width - cmdGridUnitLookup(0).Width - 100

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1.SetFocus
End Sub

Private Sub LoadVAT()
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "RATE"

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1000
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40

   txtSearch2.Width = 1000
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "RATE"
   lblSearch2(0).Visible = False

   flxSupplier(0).RowHeight(0) = 0

   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   rstRec.Open "SELECT VAT_CODE, VAT_RATE FROM tlbVatCode;", adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      flxSupplier(0).TextMatrix(0, 0) = "VAT Code"
      flxSupplier(0).TextMatrix(0, 1) = "VAT Rate"

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!VAT_CODE
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec!VAT_RATE
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close

   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdTenantLookup_Click()
   If cmdSave.Enabled Then Exit Sub

   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   PrepareList adoConn             'prepare the list of clients and properties in dwopdown comboes

'  All tenants
   If optBoth.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                   "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE OCCUPIDE_ = FALSE " & _
                 "ORDER BY T.SageAccountNumber;"
   End If

'  Current tenants Only
   If optCurrentTenant.Value Then
      szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                   "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                   "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber  " & _
              "FROM Tenants AS T LEFT JOIN " & _
                   "[" & _
                   "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                   "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE " & _
                 "ORDER BY T.SageAccountNumber;"
   End If

'  Deleted tenants only
   If optExTenant.Value Then
'   "SELECT SageAccountNumber, CompanyName " & _
'             "FROM Tenants " & _
'             "WHERE Tenants.SageAccountNumber NOT IN " & _
'                 "(SELECT LeaseDetails.SageAccountNumber " & _
'                 "FROM LeaseDetails " & _
'                 "WHERE Status=True) AND " & _
'                 "(Tenants.Comments IS NULL OR Tenants.Comments = '') " & _
'             "ORDER BY SageAccountNumber"
       szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitName, " & _
                   "'' AS Balance, " & _
                    "IIF(ISNULL(T.Comments),'CURRENT','DELETED') AS Notes, " & _
                    "IQ.PropertyID, IQ.ClientID, IQ.UnitNumber " & _
               "FROM Tenants AS T INNER JOIN " & _
                    "[" & _
                    "SELECT U.UnitName, L.SageAccountNumber, " & _
                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
                    "FROM Units AS U, LeaseDetails AS L, " & _
                        "Property AS P " & _
                    "WHERE U.UnitNumber = L.UnitNumber AND " & _
                        "L.Status = FALSE AND U.PropertyID = P.PropertyID "

      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
                 "ORDER BY T.SageAccountNumber;"
   End If

'Debug.Print szSQL
   PopulateTenantLookup szSQL, adoConn

   UpdateBalance

   adoConn.Close
   Set adoConn = Nothing

   fmeTenantLookup.Left = txtName.Left
   fmeTenantLookup.Top = optCurrentTenant.Top + 80
   fmeTenantLookup.Visible = True
   fmeTenantLookup.ZOrder 0
   cboClientList.SetFocus
   txtSearchTenant.text = ""
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer
   
   For i = 1 To gridTenantLookup.Rows - 1
      For j = 0 To UBound(szaTenantBalance, 2) - 1
         If gridTenantLookup.TextMatrix(i, 0) = szaTenantBalance(0, j) Then
            gridTenantLookup.TextMatrix(i, 3) = Format(szaTenantBalance(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaTenantBalance, 2) Then gridTenantLookup.TextMatrix(i, 3) = ""
   Next i
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   
   On Error GoTo ErrorHandler

'*************************************** CLIENT ********************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   Dim Data() As String
   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboClientList.Column() = Data()
   cboClientList.ListIndex = 0
   adoRST.Close

'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode, ClientID " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   ReDim dataProperty(TotalCol, TotalRow) As String

   dataProperty(0, 0) = "ALL"
   dataProperty(1, 0) = "All Properties"
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         dataProperty(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   cboPropertyList.Column() = dataProperty()
   cboPropertyList.ListIndex = 0

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub cmdUnitLookup_Click()
   If txtTenantID.text = "" Then Exit Sub

   Load frmUnits2
   frmUnits2.LOAD_UNIT_UNITID = txtUnitNumber.text
   frmUnits2.Show
End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtUnitMemo.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveMemo("Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo) Then
      ShowMsgInTaskBar "The Memo has been saved successfully."
   End If
   MemoButtonEnable False
End Sub

Private Sub cmdViewLetter_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String
   Dim rstReport As ADODB.Recordset
   Dim rstRst As New ADODB.Recordset

   If flxLetters.row < 1 Then Exit Sub
   If flxLetters.TextMatrix(flxLetters.row, 1) = "" Then Exit Sub

   adoConn.Open getConnectionString
   
   szSQL = "UPDATE tlbLetterReports SET isPrint = '';"
   adoConn.Execute szSQL
   
   For i = 1 To flxLetters.Rows - 1
    If (flxLetters.TextMatrix(i, 0) = "X") Then
      szSQL = "UPDATE tlbLetterReports SET isPrint = 'Y' " & _
              "WHERE id = " & flxLetters.TextMatrix(i, 1) & ";"
      adoConn.Execute szSQL
    End If
   Next i

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArchiveLetterTemplate.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub dtpDateCompleted_Change()
   TextBoxChangeDate dtpDateCompleted
End Sub

Private Sub dtpDateCompleted_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpDateCompleted, KeyAscii
End Sub

Private Sub dtpDateCompleted_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpDateCompleted
End Sub

Private Sub dtpRemindDate_Change()
   TextBoxChangeDate dtpRemindDate
End Sub

Private Sub dtpRemindDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpRemindDate, KeyAscii
End Sub

Private Sub dtpRemindDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpRemindDate
End Sub

Private Sub dtpReportedDate_Change()
   TextBoxChangeDate dtpReportedDate
End Sub

Private Sub dtpReportedDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate dtpReportedDate, KeyAscii
End Sub

Private Sub dtpReportedDate_LostFocus()
   If dtpReportedDate.text <> "" Then TextBoxFormatDate dtpReportedDate
End Sub

Private Sub flxACHistory_Click()
   Dim iCurRowHeight As Integer
   Dim iRow          As Integer
   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
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

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SC" Then
      szSQL = "SELECT S.*, F.FundCode " & _
              "FROM tlbReceipt AS R, DemandRecords AS D, DemandSplitRecords AS S, Fund AS F " & _
              "WHERE R.DemandRef = D.DemandID AND S.SageDepartment = F.FundID AND " & _
                  "D.SageAccountNumber = '" & Trim(txtTenantID.text) & "' AND " & _
                  "D.DemandID = S.DemandID AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", 1, 2) & " AND " & _
                  "R.SlNumber = " & StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & " " & _
              "ORDER BY S.SplitID;"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = adoRST.Fields.Item("SplitID").Value
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("DueDate").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundCode").Value
            .TextMatrix(iRow, 5) = adoRST.Fields.Item("NominalCodeforAmount").Value
            .TextMatrix(iRow, 6) = adoRST.Fields.Item("DateFrom").Value
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("DateTo").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRST.Close
      End With
   End If

   If (Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SA") And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "SRR" Then
      szSQL = "SELECT S.*, R.SlNumber, F.FundCode " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S, Fund AS F " & _
              "WHERE R.TransactionID = S.RptHeader AND S.FundID = F.FundID AND " & _
                  "R.SageAccountNumber = '" & Trim(txtTenantID.text) & "' AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR", 3, 4) & " AND " & _
                  "R.SlNumber = " & StrDigitVal(flxACHistory.TextMatrix(flxACHistory.row, 1)) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = adoRST.Fields.Item("SlNumber").Value
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 3) = IIf(IsNull(adoRST.Fields.Item("DueDate").Value), "", adoRST.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = ""
            .TextMatrix(iRow, 7) = ""
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "SRR" Then
      szSQL = "SELECT S.*, F.FundCode " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S, Fund AS F " & _
              "WHERE R.TransactionID = S.RptHeader AND S.FundID = F.FundID AND " & _
                  "R.Type = 23 AND " & _
                  "R.SlNumber = " & Mid(flxACHistory.TextMatrix(flxACHistory.row, 1), 4, _
                                    Len(flxACHistory.TextMatrix(flxACHistory.row, 1)) - 3) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 1) = ""
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 3) = IIf(IsNull(adoRST.Fields.Item("DueDate").Value), "", adoRST.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("FundCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = ""
            .TextMatrix(iRow, 7) = ""
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 12) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   adoConn.Close
   Set adoConn = Nothing
'
'   flxACHistorySplit.ColAlignment(0) = vbAlignLeft
End Sub

Private Sub flxDeposit_Click()
   flxDeposit_RowColChange
End Sub

Private Sub flxDeposit_RowColChange()
   On Error Resume Next

   iGridRow = flxDeposit.row

   ButtonHanlding GridRowOnSelection

   cmbBank.Value = flxDeposit.TextMatrix(flxDeposit.row, COL_BANK_CODE)
   txtDNC(0).text = flxDeposit.TextMatrix(flxDeposit.row, COL_NOMINAL_CODE)
   txtDNC(1).text = flxDeposit.TextMatrix(flxDeposit.row, COL_NC_NAME)
   txtDate.text = flxDeposit.TextMatrix(flxDeposit.row, 3)
   cboDepositType.Value = flxDeposit.TextMatrix(flxDeposit.row, COL_DEPOSIT_TYPE_CODE)
   txtDptDetails.text = flxDeposit.TextMatrix(flxDeposit.row, 6)
   If flxDeposit.TextMatrix(flxDeposit.row, 7) <> "" Then
      txtDptAmount.text = flxDeposit.TextMatrix(flxDeposit.row, 7)
   Else
      txtDptAmount.text = flxDeposit.TextMatrix(flxDeposit.row, 9)
   End If
   cmbDptAmtType.Value = flxDeposit.TextMatrix(flxDeposit.row, COL_AMOUNT_TYPE_CODE)
   txtOSDpt.text = flxDeposit.TextMatrix(flxDeposit.row, COL_OUT_STANDING_REFUND)
   cboGroup.text = flxDeposit.TextMatrix(flxDeposit.row, 10)
   GROUP_NO = flxDeposit.TextMatrix(flxDeposit.row, 10)
   cboFund.Value = flxDeposit.TextMatrix(flxDeposit.row, COL_FUND)
End Sub

Private Sub flxEmails_RowColChange()
   SelectOnly1RowFlxGrid flxEmails, flxEmails.row
End Sub

Private Sub flxLetters_Click()
    Dim i As Integer
    i = Select1RowFlxGrid(flxLetters, flxLetters.row, 0)
End Sub

Private Sub flxSupplier_Click(Index As Integer)
   If szSel = "TAX" Then
      lblVatCode(0).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtCodeVat.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   If szSel = "NC" Then
      txtNominalCode.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtNominalCodeName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   If szSel = "SLC" Then
      txtSLControl.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtSLControlName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
   End If
   If szSel = "DH" Then
      txtDNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtDNC(1).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      tabTenant.Enabled = True
      txtDate.SetFocus
   End If

   flxSupplier(0).Clear
   fraList(0).Visible = False
End Sub

Private Sub fmeTenant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Activate()
   Dim iRow As Integer

   If LOAD_TENANT_TENANTID <> "" Then LoadTenantByTenantID
End Sub

Private Sub Form_Load()
   MousePointer = vbHourglass

   Me.Height = tabTenant.Top + tabTenant.Height + 555
   Me.Width = 11775
   Me.Top = 0
   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   fmeTenant.BackColor = Me.BackColor
   tabTenant.BackColor = Me.BackColor
   optCurrentTenant.BackColor = fmeTenant.BackColor
   optExTenant.BackColor = fmeTenant.BackColor
   optBoth.BackColor = fmeTenant.BackColor
   Frame13.BackColor = Me.BackColor
   optCurrentTenant.BackColor = Me.BackColor
   optExTenant.BackColor = Me.BackColor
   optBoth.BackColor = Me.BackColor
   fmeLoading.Top = 4305
   fmeLoading.Left = 4508
   Me.Caption = "Lessee"
   tabTenant.Tab = 0
   ComponentInFrameEnableMode Me, fmeTenant, DefaultMode
   ComponentInFrameEnableMode Me, fmeTenancyDetails, DefaultMode
   ConfigureFlxAsLable flxContacts, lblContacts, 0, 6

   txtSearchTenant.Enabled = True
   COPYMODE_ = False
   NEWMODE_ = False
   SEARCHTenantMODE_ = True
   bDEPOSIT_HELD = False

'   SageCustomerAccCombo cboSageAccountNumber
   TenantTabEnabled False

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

'    Populate the codes
   PopulateCodes adoConn

'  Load all tenants with balance in gridTenantLookup
   TenantAccountBalance adoConn

   adoConn.Close
   Set adoConn = Nothing

'   Call WheelHook(Me.hWnd)
   MousePointer = vbDefault
End Sub

Private Sub PopulateLessee(adoConn As ADODB.Connection)
'   Dim szSQL As String
'   Dim rstSQL As New ADODB.Recordset
End Sub

Private Sub ConfigFlxDeposit()
   Dim szHeader As String, iCol As Integer

   flxDeposit.Clear
   flxDeposit.Cols = 18
   flxDeposit.Rows = 2
   flxDeposit.RowHeight(0) = 0

   For iCol = 1 To flxDeposit.Cols - 8
      flxDeposit.ColWidth(iCol - 1) = Label6(iCol).Left - Label6(iCol - 1).Left
   Next iCol
   flxDeposit.ColWidth(iCol - 1) = flxDeposit.Width + flxDeposit.Left - Label6(iCol - 1).Left - 280   'col = 9
   flxDeposit.ColWidth(COL_BANK_CODE) = 0
   flxDeposit.ColWidth(COL_NOMINAL_CODE) = 0
   flxDeposit.ColWidth(COL_AMOUNT_TYPE_CODE) = 0
   flxDeposit.ColWidth(COL_DEPOSIT_TYPE_CODE) = 0
   flxDeposit.ColWidth(COL_NC_NAME) = 0
   flxDeposit.ColWidth(COL_FUND) = 0
   flxDeposit.ColWidth(COL_RECON) = 0
End Sub

Private Sub ConfigFlxACHistorySplit()
   Dim szHeader As String, iCol As Integer

   With flxACHistorySplit
      .Clear
      .Cols = 13
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = 0
      For iCol = 1 To .Cols - 2
         .ColWidth(iCol) = Label1(30 + iCol + 1).Left - Label1(30 + iCol).Left
      Next iCol

      .ColWidth(iCol) = .Width + .Left - Label1(42).Left - 360
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   LOAD_TENANT_TENANTID = ""
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub ConfigurFlexGrid()
   Dim szHeader   As String
   Dim i          As Integer

   gridTenantLookup.Clear
   gridTenantLookup.Rows = 2
   gridTenantLookup.Cols = 8
   
   gridTenantLookup.Visible = True
   gridTenantLookup.RowHeight(0) = 240
   gridTenantLookup.row = 0

   For i = 0 To gridTenantLookup.Cols - 1
      gridTenantLookup.col = i
      gridTenantLookup.CellFontBold = True
   Next i

   szHeader$ = "<Sage A/C|<Name|<Address|>Balance||Client|Property|Unit"
   gridTenantLookup.FormatString = szHeader$

   gridTenantLookup.ColWidth(0) = lblTenantSort(1).Left - lblTenantSort(0).Left
   gridTenantLookup.ColWidth(1) = lblTenantSort(2).Left - lblTenantSort(1).Left
   gridTenantLookup.ColWidth(2) = lblTenantSort(3).Left - lblTenantSort(2).Left
   gridTenantLookup.ColWidth(3) = 1200
'   gridTenantLookup.TextMatrix(0, 0) = "Sage A/C"
'   gridTenantLookup.ColAlignment(0) = vbLeftJustify
'   gridTenantLookup.TextMatrix(0, 1) = "Name"
'   gridTenantLookup.ColAlignment(1) = vbLeftJustify
'   gridTenantLookup.TextMatrix(0, 2) = "Address"
'   gridTenantLookup.ColAlignment(2) = vbLeftJustify
'   gridTenantLookup.TextMatrix(0, 2) = "Balance"
'   gridTenantLookup.ColAlignment(2) = vbRightJustify

   gridTenantLookup.ColWidth(4) = 0
   gridTenantLookup.ColWidth(5) = 0
   gridTenantLookup.ColWidth(6) = 0
   gridTenantLookup.ColWidth(7) = 0
End Sub

Private Sub gridBank_Click()
   populateControl Me, gridBank
End Sub

Private Sub LoadTenantByTenantID()
   Dim sSQLQuery_ As String, szHeader As String

   SEARCHTenantMODE_ = False
   fmeTenantLookup.Visible = False

   fmeLoading.Visible = True
   fmeLoading.Refresh

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   'Populate the Tenant Header
   PopulateTenantInformation adoConn, LOAD_TENANT_TENANTID
   lblLeaseChanged.Caption = LOAD_TENANT_TENANTID

   ' Populate Bank Details
   ConfigureFlxBank
   szHeader$ = "<BankTenantID|<BankID|<BankACName|<BankSortCode|<BankACNumber|<PaymentMethod|<BacsRef|<IsDefaultAC"
   sSQLQuery_ = "SELECT BankTenantID, BankID, BankACName, BankSortCode, BankACNumber, " & _
                           "PaymentMethod, BacsRef, IsDefaultAC " & _
                         "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'"
   populateGridSimply adoConn, sSQLQuery_, gridBank, szHeader

'   ' Populate Event History
'   ConfigurFlxEventHistory
'   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
'   sSQLQuery_ = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
'                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
'   populateGridSimply adoConn, sSQLQuery_, gridEventHistory, szHeader
   LoadGridMaintenanceHistory adoConn
   LoadFlxACHistory adoConn

   '' LOAD Tenant DETAIL INFORMATION
   RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo

   fmeLoading.Visible = False
   adoConn.Close
   Set adoConn = Nothing

   ' SET OTHERS
   SEARCHTenantMODE_ = True
   TenantTabEnabled True
End Sub

Private Sub gridMaintenanceHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridMaintenanceHistory.ToolTipText = gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.MouseRow, gridMaintenanceHistory.MouseCol)
End Sub

Private Sub gridMaintenanceHistory_RowColChange()
   populateControl Me, gridMaintenanceHistory
End Sub

Private Sub gridMaintenanceHistory_Click()
   If (gridMaintenanceHistory.row > 0 And gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) <> "") Then
      cmdEditMHistory.Enabled = True
   Else
      cmdEditMHistory.Enabled = False
   End If
End Sub

Private Sub gridTenantLookup_Click()
   Dim szSQL As String, szHeader As String

   SEARCHTenantMODE_ = False
   fmeTenantLookup.Visible = False

   fmeLoading.Visible = True
   fmeLoading.Refresh

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   adoConn.Open getConnectionString

'   Populate the Tenant Header
   If gridTenantLookup.TextMatrix(gridTenantLookup.row, 4) = "CURRENT" And optCurrentTenant.Value Then
      PopulateTenantInformation adoConn, gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      lblLeaseChanged.Caption = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
   Else
      txtTenantID.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      lblLeaseChanged.Caption = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
      txtName.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 1)
      txtCompanyName.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 1)
      txtProperty.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 5)
      txtClient.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 6)
      txtUnit.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 2)
   End If

   szSQL = "SELECT N.Code, N.Name, S.CAName, S.Code AS NCode, V.VAT_CODE, V.VAT_RATE " & _
           "FROM   ((Tenants AS T LEFT JOIN NominalLedger AS N ON T.DefaultNC = N.Code) LEFT JOIN " & _
                  "( SELECT Code, CAName " & _
                   " FROM   NominalLedger AS N " & _
                   " WHERE  N.ClientID = '" & txtClientID.text & "') AS S ON T.SLControl = S.Code) " & _
                  "LEFT JOIN tlbVatCode AS V ON T.VAT_CODE = V.VAT_CODE " & _
           "WHERE  N.ClientID = '" & txtClientID.text & "' AND " & _
                  "T.SageAccountNumber = '" & txtTenantID.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      txtNominalCode.text = IIf(IsNull(adoRST.Fields.Item("Code").Value), "", adoRST.Fields.Item("Code").Value)
      txtNominalCodeName.text = IIf(IsNull(adoRST.Fields.Item("Name").Value), "", adoRST.Fields.Item("Name").Value)
      lblVatCode(0).Caption = IIf(IsNull(adoRST.Fields.Item("VAT_CODE").Value), "", adoRST.Fields.Item("VAT_CODE").Value)
      txtCodeVat.text = IIf(IsNull(adoRST.Fields.Item("VAT_RATE").Value), "", adoRST.Fields.Item("VAT_RATE").Value)
      txtSLControl.text = IIf(IsNull(adoRST.Fields.Item("NCode").Value), "", adoRST.Fields.Item("NCode").Value)
      txtSLControlName.text = IIf(IsNull(adoRST.Fields.Item("CAName").Value), "", adoRST.Fields.Item("CAName").Value)
   Else
      txtNominalCode.text = ""
      txtNominalCodeName.text = ""
      lblVatCode(0).Caption = ""
      txtCodeVat.text = ""
      txtSLControl.text = ""
      txtSLControlName.text = ""
   End If

   adoRST.Close

'    Populate Bank Details
   ConfigureFlxBank
   szHeader$ = "<BankTenantID|<BankID|<BankACName|<BankSortCode|<BankACNumber|<PaymentMethod|<BacsRef|<IsDefaultAC"
   szSQL = "SELECT BankTenantID, BankID, BankACName, BankSortCode, BankACNumber, " & _
              "PaymentMethod, BacsRef, IsDefaultAC " & _
           "FROM TenantBankDetails WHERE BankTenantID = '" & txtTenantID.text & "'"
   populateGridSimply adoConn, szSQL, gridBank, szHeader

'  Populate Lessee Account history
   LoadFlxACHistory adoConn
   LoadFlxLetter adoConn

   txtBalance.text = Format(AccountBalance(adoConn), "0.00")

''    Populate Event History
'   ConfigurFlxEventHistory
'   szHeader$ = "<EventHistoryID|<EventTenantID|<EventType|<ReportedDate|<Description|<DateCompleted|<TaskOwner|<Contact|<RemindDate|<Alarm"
'   szSQL = "SELECT EventHistoryID, EventTenantID, EventType, ReportedDate, " & _
'                  "Description, DateCompleted, TaskOwner, Contact, RemindDate, Alarm " & _
'                "FROM TenantEventHistory WHERE EventTenantID = '" & txtTenantID.text & "'"
'   populateGridSimply adoConn, szSQL, gridEventHistory, szHeader
   LoadGridMaintenanceHistory adoConn

'    LOAD Tenant DETAIL INFORMATION
   RetrieveMemo "Tenants", "TenantMemo", txtTenantID.text, "SageAccountNumber", txtUnitMemo

   szSQL = "SELECT TenantDeposit.TenantID, (SUM(DptAmount) - " & _
                 "iif(isnull(DepositRefund.TotalDepositRefund),0,DepositRefund.TotalDepositRefund)) as TotalDepositHeld " & _
           "FROM TenantDeposit LEFT JOIN [SELECT SUM(DptAmount) AS TotalDepositRefund, TenantID " & _
              "FROM TenantDeposit " & _
              "WHERE DptRefund = True and Status = true and Deleted = false " & _
              "GROUP BY tenantid]. AS DepositRefund ON  TenantDeposit.TenantID = DepositRefund.TenantID " & _
           "WHERE TenantDeposit.TenantID = '" & txtTenantID.text & "' AND " & _
                 "TenantDeposit.Deleted = FALSE AND DptRefund = FALSE " & _
           "GROUP BY TenantDeposit.TenantID, DepositRefund.TotalDepositRefund;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      txtDeposit.text = "0.00"
   Else
      txtDeposit.text = Format(IIf(IsNull(adoRST.Fields.Item("TotalDepositHeld").Value), 0, adoRST.Fields.Item("TotalDepositHeld").Value), "0.00")
   End If

   adoRST.Close
   Set adoRST = Nothing

   fmeLoading.Visible = False

'    SET OTHERS
   SEARCHTenantMODE_ = True
   TenantTabEnabled True
   tabTenant.Tab = 0

   DeletedMode gridTenantLookup.TextMatrix(gridTenantLookup.row, gridTenantLookup.Cols - 1)

   populateGroupCombo adoConn

   adoConn.Close
   Set adoConn = Nothing

   LoadFlxEmails
End Sub

Public Sub LoadGridMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection)
   Dim rstMHistory_ As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.ReportedBy, " & _
                "H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent, " & _
                "H.MaintenanceType, H.ReportedFrom " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
           "WHERE H.PropertyID = '" & txtTenantID.text & "' AND " & _
               "S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' AND " & _
               "H.ReportedFrom = 'L' " & _
           "ORDER BY H.ReportedDate DESC;"
'Debug.Print szSQL

   rstMHistory_.Open szSQL, conMHistory_, adOpenStatic, adLockReadOnly

   ConfigGridMaintenanceHistory rstMHistory_

   If rstMHistory_.EOF Then
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
      Exit Sub
   Else
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
   End If

   populateGridDefinedHeader conMHistory_, szSQL, gridMaintenanceHistory

   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.col = 0
End Sub

Private Function AccountBalance(ByVal adoConn As ADODB.Connection) As Currency
   Dim iRow As Integer, cDr As Currency, cCr As Currency

   For iRow = 1 To flxACHistory.Rows - 1
'If Val(flxACHistory.TextMatrix(iRow, 7)) <> 0 Then
'MsgBox ""
'End If
      cDr = cDr + Round(CCur(IIf(flxACHistory.TextMatrix(iRow, 7) = "", 0, flxACHistory.TextMatrix(iRow, 7))), 2)
      cCr = cCr + Round(CCur(IIf(flxACHistory.TextMatrix(iRow, 8) = "", 0, flxACHistory.TextMatrix(iRow, 8))), 2)
'Debug.Print flxACHistory.TextMatrix(iRow, 7) + "   -" + flxACHistory.TextMatrix(iRow, 8)
   Next iRow

   AccountBalance = cDr - cCr
   Exit Function

'  Calculating Lessee's balance from the history grid. thats why this function does not execute the following code.

'  NOT IN USE THE FOLLWING CODE -------->>>>>>>>>>

'--------------------------------------------------------------------------------------------------------------------
   Dim szSQL As String, cAllOSInv As Currency, cAllPoA As Currency, cAllCredit As Currency
   Dim adoRST As New ADODB.Recordset

   szSQL = "SELECT SageAccountNumber, SUM(OSAmount) AS OSI " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "ReceiptView = TRUE AND TYPE <> 2 " & _
           "GROUP BY SageAccountNumber;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllOSInv = CCur(IIf(IsNull(adoRST.Fields.Item("OSI").Value), 0, adoRST.Fields.Item("OSI").Value))
   adoRST.Close

   szSQL = "SELECT SageAccountNumber, SUM(OSAllocation) AS OSA " & _
           "FROM tblPoA " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "PoAView = TRUE " & _
           "GROUP BY SageAccountNumber;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllPoA = CCur(IIf(IsNull(adoRST.Fields.Item("OSA").Value), 0, adoRST.Fields.Item("OSA").Value))
   adoRST.Close

   szSQL = "SELECT SageAccountNumber, SUM(OSAmount) AS OSI " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & txtTenantID.text & "' AND " & _
               "ReceiptView = TRUE AND TYPE = 2 " & _
           "GROUP BY SageAccountNumber;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then _
      cAllCredit = CCur(IIf(IsNull(adoRST.Fields.Item("OSI").Value), 0, adoRST.Fields.Item("OSI").Value))
   adoRST.Close
   Set adoRST = Nothing

   cAllCredit = cAllCredit + cAllPoA

   AccountBalance = cAllOSInv - cAllCredit
End Function

Private Sub DeletedMode(szMode As String)
   If szMode = "DELETED" Then
      cmdDeleteLessee.Enabled = False
      cmdEdit.Enabled = False
      cmdCopy.Enabled = False
      cmdEditTenantAddress.Enabled = False
      cmdNewBank.Enabled = False
      cmdEditBank.Enabled = False
      cmdNewEvent.Enabled = False
      cmdEditEvent.Enabled = False
   End If

   If szMode = "CURRENT" Then TenantTabEnabled True

End Sub

Private Sub lblTenantSort_Click(Index As Integer)
   If Index = 0 Then                               ' Sort Tenant ID
      SortingGrid gridTenantLookup, Index, bSortingCol1
      bSortingCol1 = IIf(bSortingCol1, False, True)
      lblTenantSort(0).FontBold = True
      lblTenantSort(1).FontBold = False
      lblTenantSort(2).FontBold = False
   End If

   If Index = 1 Then                               ' Sort Tenant Name
      SortingGrid gridTenantLookup, Index, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      lblTenantSort(0).FontBold = False
      lblTenantSort(1).FontBold = True
      lblTenantSort(2).FontBold = False
   End If

   If Index = 2 Then                               ' Sort Unit Name
      SortingGrid gridTenantLookup, Index, bSortingCol3
      bSortingCol3 = IIf(bSortingCol3, False, True)
      lblTenantSort(0).FontBold = False
      lblTenantSort(1).FontBold = False
      lblTenantSort(2).FontBold = True
   End If
End Sub

Private Sub optAll_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget / Location"
   Label61(3).Caption = "Ref"
   Label61(1).Caption = "Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
End Sub

Private Sub optDiary_Click()
   Dim i As Integer

   Label61(10).Caption = "Location"
   Label61(3).Caption = "Diary No"
   Label61(1).Caption = "Event Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) = "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optJobs_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget"
   Label61(3).Caption = "Job No"
   Label61(1).Caption = "Maintenance Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) <> "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optExitingGroup_Click()
   cboGroup.Enabled = True
   cboGroup.Locked = False
End Sub

Private Sub optNewGroup_Click()
   cboGroup.Enabled = False
   cboGroup.Locked = True
End Sub

Private Sub tabTenant_Click(PreviousTab As Integer)
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Select Case tabTenant.Tab
   Case 2:
      If Not bDEPOSIT_HELD Then
         ConfigFlxDeposit
         LoadComboes adoConn

         bDEPOSIT_HELD = True
      End If

      If txtTenantID.text <> "" Then LoadFlxDeposit adoConn      'Filling the deposit grid

      ButtonHanlding DefaultMode
      cmbBank.SetFocus
   Case 4:
      If txtTenantID.text <> "" Then _
         Call LoadAttachmentFiles(cmbFiles, txtTenantID.text, "Tenants")

   End Select

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadFlxDeposit(ByVal adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, i As Integer

   szSQL = "SELECT D.*, B.ReconNow " & _
           "FROM TenantDeposit AS D LEFT JOIN tlbBankPayment AS B ON D.DepositID = B.TenantDeposit " & _
           "Where " & _
               "D.TenantID = '" & txtTenantID.text & "' AND " & _
               "D.Deleted = False;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With adoRST
      iRow = 1
      flxDeposit.Clear
      flxDeposit.Rows = 1
      While Not .EOF
         If Not .EOF Then flxDeposit.AddItem ""
         flxDeposit.TextMatrix(iRow, 0) = .Fields.Item("DepositID").Value
         flxDeposit.TextMatrix(iRow, 1) = .Fields.Item("BCName").Value
         flxDeposit.TextMatrix(iRow, 2) = IIf(IsNull(.Fields.Item("TransactionID").Value), "", .Fields.Item("TransactionID").Value)
         flxDeposit.TextMatrix(iRow, 3) = Format(.Fields.Item("DepositDate").Value, "dd/mm/yyyy")
         If Left(.Fields.Item("TransactionID").Value, 1) = "D" Then _
            flxDeposit.TextMatrix(iRow, 5) = "Deposit"
         If Left(.Fields.Item("TransactionID").Value, 1) = "R" Then _
            flxDeposit.TextMatrix(iRow, 5) = "Refund"
         If Left(.Fields.Item("TransactionID").Value, 1) = "E" Then _
            flxDeposit.TextMatrix(iRow, 5) = "Expenses"

         flxDeposit.TextMatrix(iRow, 6) = .Fields.Item("DptDetails").Value
         If flxDeposit.TextMatrix(iRow, 6) <> "" Then flxDeposit.ColAlignment(6) = vbLeftJustify

         If Not .Fields.Item("DptRefund").Value Then
            flxDeposit.TextMatrix(iRow, 4) = Value_SecondaryCode("DPTYP", .Fields.Item("DptType").Value, adoConn)   'Deposit Type code name
            flxDeposit.TextMatrix(iRow, 7) = Format(Val(.Fields.Item("DptAmount").Value), "0.00")
            flxDeposit.TextMatrix(iRow, COL_OUT_STANDING_REFUND) = Format(Val(.Fields.Item("OSRefund").Value), "0.00")
            flxDeposit.TextMatrix(iRow, COL_IS_REFUND) = ""
         Else
            If .Fields.Item("DptType").Value = "DR" Then
               flxDeposit.TextMatrix(iRow, 4) = Value_SecondaryCode("RTYP", .Fields.Item("DptType").Value, adoConn)   'Deposit Type code name
            Else
               flxDeposit.TextMatrix(iRow, 4) = Value_SecondaryCode("EXPTYP", .Fields.Item("DptType").Value, adoConn)   'Deposit Type code name
            End If
            flxDeposit.TextMatrix(iRow, COL_IS_REFUND) = Format(Val(.Fields.Item("DptAmount").Value), "0.00")
         End If

         flxDeposit.TextMatrix(iRow, 10) = .Fields.Item("GroupNo").Value
         flxDeposit.TextMatrix(iRow, COL_BANK_CODE) = .Fields.Item("BankCode").Value
         flxDeposit.TextMatrix(iRow, COL_NOMINAL_CODE) = IIf(IsNull(.Fields.Item("NominalCode")), "", .Fields.Item("NominalCode").Value)
         flxDeposit.TextMatrix(iRow, COL_NC_NAME) = IIf(IsNull(.Fields.Item("NCName")), "", .Fields.Item("NCName").Value)
         flxDeposit.TextMatrix(iRow, COL_AMOUNT_TYPE_CODE) = .Fields.Item("DptAmtType").Value
         flxDeposit.TextMatrix(iRow, COL_DEPOSIT_TYPE_CODE) = .Fields.Item("DptType").Value
         flxDeposit.TextMatrix(iRow, COL_FUND) = .Fields.Item("FundID").Value
         flxDeposit.TextMatrix(iRow, COL_RECON) = IIf(IsNull(.Fields.Item("ReconNow").Value), "", .Fields.Item("ReconNow").Value)
         iRow = iRow + 1
         .MoveNext
      Wend
      .Close
   End With
   Set adoRST = Nothing
   flxDeposit.row = 0
End Sub

Private Sub LoadComboes(ByVal adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim Data() As String, i As Integer

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code " & _
           "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup bank account for the client."
   Else
      ReDim Data(1, adoRST.RecordCount - 1) As String
      i = 0
      While Not adoRST.EOF
         Data(0, i) = adoRST.Fields.Item("BNC").Value
         Data(1, i) = adoRST.Fields.Item("BNN").Value
         i = i + 1
         adoRST.MoveNext
      Wend
      cmbBank.Clear
      cmbBank.Column() = Data()
   End If

   adoRST.Close

   szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!SC
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cmbDptAmtType.Clear
   cmbDptAmtType.Column() = Data()

   adoRST.Close
'////////////////////////////////// DEPOSIT TYPE //////////////////////////////////////////////////////
   szSQL = "SELECT SecondaryCode.Code as SC, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE (PrimaryCode.Value = 'DEPOSIT TYPE' OR " & _
                  "PrimaryCode.Value = 'REFUND TYPE' OR " & _
                  "PrimaryCode.Value = 'EXPENSES TYPE') AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim Data(1, adoRST.RecordCount - 1) As String

   i = 0
   While Not adoRST.EOF
      Data(0, i) = adoRST!SC
      Data(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend

   cboDepositType.Clear
   cboDepositType.Column() = Data()

   adoRST.Close
'////////////////////////////////// DEPOSIT TYPE //////////////////////////////////////////////////////
   szSQL = "SELECT FundID, FundName FROM FUND;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount) As String

   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST.Fields.Item("FundID").Value
      szaData(1, i) = adoRST.Fields.Item("FundName").Value
      i = i + 1
      adoRST.MoveNext
   Wend

   cboFund.Clear
   cboFund.Column() = szaData()

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigFlxACHistory()
   Dim szHeader As String, iCol As Integer

   flxACHistory.Clear
   flxACHistory.Cols = 11
   flxACHistory.Rows = 2
   flxACHistory.RowHeight(0) = 0
'   szHeader$ = "|<REF|<TYPE|<UnitNumber|<IssueDate|>AMOUNT|>OS|>DEBIT|>CREDIT"
'   flxACHistory.FormatString = szHeader$

   flxACHistory.ColWidth(0) = Label1(1).Left - flxACHistory.Left
   For iCol = 1 To flxACHistory.Cols - 4
      flxACHistory.ColWidth(iCol) = Label1(iCol + 1).Left - Label1(iCol).Left
   Next iCol

   flxACHistory.ColAlignment(4) = vbLeftJustify
   flxACHistory.ColWidth(8) = flxACHistory.Width + flxACHistory.Left - Label1(8).Left - 360
   flxACHistory.ColWidth(9) = 0
   flxACHistory.ColWidth(10) = 0
End Sub

'  Build up lessee's Account History
Private Sub LoadFlxACHistory(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoRpt As New ADODB.Recordset, adoRptDtl As New ADODB.Recordset

   ConfigFlxACHistory
   ConfigFlxACHistorySplit

   szSQL = "SELECT Rpt.*, TT.DESCRIPTION AS TT_DES, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
           "FROM tlbReceipt AS Rpt, tlbTransactionTypes AS TT " & _
           "WHERE Rpt.SageAccountNumber = '" & txtTenantID.text & "' And " & _
               "Rpt.Type = TT.TYPE_ID " & _
           "ORDER BY Rpt.RDate;"
'Debug.Print szSQL
   adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoRpt.EOF                                           '//1
         If adoRpt!Type = 1 Then                                                             '//2
            szSQL = "SELECT MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, SQ.* " & _
                    "FROM (tlbReceipt AS R INNER JOIN " & _
                    "(" & _
                        "SELECT RT.SlNumber, RT.FromTran, RT.ReceiptAmount, R.DemandRef " & _
                        "FROM (RptTransactions AS RT INNER JOIN " & _
                              "tlbReceipt AS R ON RT.ToTran = R.TransactionID) INNER JOIN " & _
                              "tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                        "Where RT.ToTran = " & adoRpt.Fields.Item("TransactionID").Value & " " & _
                    ") AS SQ ON R.TransactionID = SQ.FromTran) " & _
                        "INNER JOIN tlbTransactionTypes AS T ON R.Type = T.TYPE_ID;"
         Else
            szSQL = "SELECT RT.*, R.SlNumber AS RefID, " & _
                        "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
                    "FROM (RptTransactions AS RT INNER JOIN tlbReceipt AS R ON RT.ToTran = R.TransactionID) " & _
                        "INNER JOIN tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                    "WHERE RT.FromTran = " & adoRpt.Fields.Item("TransactionID").Value & ";"
         End If
'Debug.Print szSQL
         adoRptDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         iChild = 0
         If adoRptDtl.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(iKount, 0) = "+"
            iChild = iKount + 1
            While Not adoRptDtl.EOF
               .TextMatrix(iChild, 0) = "-"
               If adoRpt!Type = 1 Then
                  .TextMatrix(iChild, 4) = "Receipt from " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 4) = "Receipt to " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("RefID").Value
               End If
               .TextMatrix(iChild, 5) = Format(adoRptDtl.Fields.Item("ReceiptAmount").Value, "0.00")
               .RowHeight(iChild) = 0
               iChild = iChild + 1
               adoRptDtl.MoveNext
               If Not adoRptDtl.EOF Then .AddItem ""
            Wend
         Else
            .TextMatrix(iKount, 0) = ""
         End If
'1:DemandRef, 2:Invoice, 3:Date, 4:Details, 5:Amount, 6:Amount (OS), 7:Amount (Dr)
         adoRptDtl.Close
'*************
         .TextMatrix(iKount, 1) = adoRpt.Fields.Item("PF").Value & adoRpt.Fields.Item("SlNumber").Value
         .TextMatrix(iKount, 2) = IIf(UCase(Left(adoRpt.Fields.Item("TT_DES").Value, 5)) = "SALES", Mid(adoRpt.Fields.Item("TT_DES").Value, 7), adoRpt.Fields.Item("TT_DES").Value)
         .TextMatrix(iKount, 3) = IIf(IsNull(adoRpt.Fields.Item("RDate").Value), "", _
                                             adoRpt.Fields.Item("RDate").Value)

         If adoRpt.Fields.Item("Type").Value = 3 Or _
               adoRpt.Fields.Item("Type").Value = 4 Or _
               adoRpt.Fields.Item("Type").Value = 23 Then
            .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item("ExtRef").Value), "", _
                                                adoRpt.Fields.Item("ExtRef").Value)
         Else
            .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item("Details").Value), "", _
                                                adoRpt.Fields.Item("Details").Value)
         End If

         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item("OSAmount").Value, "0.00")
         If adoRpt!Type = 1 Or adoRpt!Type = 23 Then
            .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 8) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Credit
         End If
         adoRpt.MoveNext
         iKount = IIf(iChild = 0, iKount + 1, iChild)
         If Not adoRpt.EOF Then .AddItem ""
      Wend

      adoRpt.Close
'############################## if there any unposted demand in the demand then its picking up here ##################
      szSQL = "SELECT D.DemandID, D.IssueDate, D.Details, SUM(DS.TotalAmount) " & _
              "FROM DemandRecords AS D, DemandSplitRecords AS DS " & _
              "WHERE D.DemandID = DS.DemandID AND DS.TrfReceipt = FALSE AND " & _
                  "D.SageAccountNumber = '" & txtTenantID.text & "' " & _
              "GROUP BY D.DemandID, D.Details, D.IssueDate;"

      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
   End With

   Set adoRpt = Nothing
   Set adoRptDtl = Nothing
   flxACHistory.row = 0
   flxACHistory.row = 0
End Sub

'  Build up lessee's Account Balance
Private Sub TenantAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber), 2 " & _
           "From " & _
            "(" & _
             "SELECT tlbReceipt.SageAccountNumber  " & _
             "From tlbReceipt " & _
             "GROUP BY tlbReceipt.SageAccountNumber" & _
            ");"
   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRptDr.EOF Then
      adoRptDr.Close
      Set adoRptDr = Nothing
      Exit Sub
   End If

   ReDim szaTenantBalance(1, adoRptDr.Fields.Item(0).Value) As String
   adoRptDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbReceipt AS Rpt " & _
           "WHERE Type = 1 OR Type = 23 " & _
           "GROUP BY SageAccountNumber " & _
           "ORDER BY SageAccountNumber;"
'Debug.Print szSQL
   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoRptDr.EOF
      szaTenantBalance(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
'If adoRptDr.Fields.Item("SageAccountNumber").Value = "Payden01" Then
'MsgBox ""
'End If
      szaTenantBalance(1, iIndex) = RoundingNumber(adoRptDr.Fields.Item("Dr").Value, 2)
      iIndex = iIndex + 1
      adoRptDr.MoveNext
   Wend

   adoRptDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbReceipt AS Rpt " & _
           "WHERE Type <> 1 AND Type <> 23 " & _
           "GROUP BY SageAccountNumber;"
'Debug.Print szSQL
   adoRptCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRptCr.EOF
      For i = 0 To iIndex - 1
         If szaTenantBalance(0, i) = adoRptCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
'If adoRptCr.Fields.Item("SageAccountNumber").Value = "Payden01" Then
'MsgBox ""
'End If
      If i < iIndex Then
         szaTenantBalance(1, i) = Val(szaTenantBalance(1, i)) - RoundingNumber(adoRptCr.Fields.Item("Cr").Value, 2)
      Else
         iIndex = iIndex + 1
         szaTenantBalance(0, iIndex) = RoundingNumber(adoRptCr.Fields.Item("Cr").Value, 2)
      End If
      adoRptCr.MoveNext
   Wend

   adoRptCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

Private Sub tabTenant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtBankAddress1_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankAddress2_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankAddress3_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankPostCode_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBankSortCode_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBranchName_DblClick(Cancel As MSForms.ReturnBoolean)
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtCompanyName_GotFocus()
   If txtCompanyName.text = "" Then Exit Sub

   SelTxtInCtrl txtCompanyName
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
End Sub

Private Sub txtDate_GotFocus()
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   TextBoxFormatDate txtDate
End Sub

Private Sub txtDptAmount_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtDptAmount, KeyAscii
End Sub

Private Sub txtDptAmount_LostFocus()
   If txtDptAmount.text = "" Then Exit Sub

'  if user changes the deposit amount then system should change the O/S amount as the same time
   If yDEPOSIT = 2 And CCur(txtDptAmount.text) - cCurDepAmt <> 0 And txtOSDpt.text <> "" Then
      txtOSDpt.text = Format(CCur(txtOSDpt.text) + CCur(txtDptAmount.text) - cCurDepAmt, "0.00")
   End If

'  system should check if the refund amount is greater than the O/S amount then system will warn user
   If (yDEPOSIT = 3 Or yDEPOSIT = 4) And Val(txtOSDpt.text) < Val(txtDptAmount.text) Then                'Refund
      MsgBox "Maximum refund/expenses amount could be £" & Format(txtOSDpt.text, "0.00"), vbCritical + vbOKOnly, "Refund/Expenses - Deposit"
      txtDptAmount.text = Format(txtOSDpt.text, "0.00")
      SelTxtInCtrl txtDptAmount
      txtDptAmount.SetFocus
   End If
End Sub

Private Sub txtEmail1_LostFocus()
   Dim szErrMsg As String

   If Trim(txtEmail1.text) <> "" Then
      If Not ValidateEmail(txtEmail1.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Lessee Email"
         SelTxtInCtrl txtEmail1
         txtEmail1.SetFocus
      End If
   End If
End Sub

Private Sub txtEmail2_LostFocus()
   Dim szErrMsg As String

   If Trim(txtEmail2.text) <> "" Then
      If Not ValidateEmail(txtEmail2.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Supplier Email"
         SelTxtInCtrl txtEmail2
         txtEmail2.SetFocus
      End If
   End If
End Sub
'
'Private Sub txtLeaseId_Change()
'MsgBox ""
'End Sub

Private Sub txtName_LostFocus()
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   If UBound(szaChoice) > 0 Then
      If szaChoice(2) <> "" Then
         If InStr(szaChoice(2), "L") > 0 Then
            If NEWMODE_ And txtTenantID.text = "" And Trim(txtName.text) <> "" Then
'               txtTenantID.text = CreateTenantId(txtName.text)
               cboSageAccountNumber.text = txtTenantID.text
            End If
         End If
      End If
   End If
   txtCompanyName.text = txtName.text
End Sub

Private Function CreateTenantId(szName As String) As String
   Dim szSQL As String, i As Integer, szChar As String, j As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   For i = 1 To Len(szName) - 1
      szChar = UCase(Mid(szName, i, 1))
      If (szChar >= "A" And szChar <= "Z") Then
         CreateTenantId = CreateTenantId & szChar
         j = j + 1
      End If
      If j = 8 Then Exit For
   Next i

   If j < 8 Then CreateTenantId = Left(CreateTenantId & "01234567", 8)

   adoConn.Open getConnectionString
   
   szSQL = "SELECT SageAccountNumber " & _
           "FROM Tenants " & _
           "WHERE Tenants.SageAccountNumber = '" & CreateTenantId & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   j = 1
   Do
      If adoRST.EOF Then Exit Do
      adoRST.Close
      CreateTenantId = Left(CreateTenantId & "01234567", 6) & Format(j, "00")
      szSQL = "SELECT SageAccountNumber " & _
              "FROM Tenants " & _
              "WHERE Tenants.SageAccountNumber = '" & CreateTenantId & "';"
   
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      
      j = j + 1
   Loop

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function

Private Sub txtSearch1_Change()
   Dim i As Integer

   If Len(txtSearch1.text) > 0 Then
      txtSearch2.text = ""
   End If

   For i = 1 To flxSupplier(0).Rows - 1
      flxSupplier(0).RowHeight(i) = 240
      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtSearch1.text))) <> UCase(txtSearch1.text) Then
         flxSupplier(0).RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearch2_Change()
   Dim i As Integer

   If Len(txtSearch2.text) > 0 Then
      txtSearch1.text = ""
   End If

   For i = 1 To flxSupplier(0).Rows - 1
      flxSupplier(0).RowHeight(i) = 240
      If UCase(Left(flxSupplier(0).TextMatrix(i, 1), Len(txtSearch2.text))) <> UCase(txtSearch2.text) Then
         flxSupplier(0).RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchAddress_Change()
   Dim i As Integer

   If Len(txtSearchAddress.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
   End If

   For i = 1 To gridTenantLookup.Rows - 1
      gridTenantLookup.RowHeight(i) = 240
      If UCase(Left(gridTenantLookup.TextMatrix(i, 2), Len(txtSearchAddress.text))) <> UCase(txtSearchAddress.text) Then
         gridTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchAddress_GotFocus()
   SelTxtInCtrl txtSearchAddress
End Sub

Private Sub txtSearchName_Change()
   Dim i As Integer

   If Len(txtSearchName.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchAddress.text = ""
   End If

   For i = 1 To gridTenantLookup.Rows - 1
      gridTenantLookup.RowHeight(i) = 240
      If UCase(Left(gridTenantLookup.TextMatrix(i, 1), Len(txtSearchName.text))) <> UCase(txtSearchName.text) Then
         gridTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchName_GotFocus()
   SelTxtInCtrl txtSearchName
End Sub

Private Sub txtSearchTenant_Change()
   Dim i As Integer

   If Len(txtSearchTenant.text) > 0 Then
      txtSearchName.text = ""
      txtSearchAddress.text = ""
   End If

   For i = 1 To gridTenantLookup.Rows - 1
      gridTenantLookup.RowHeight(i) = 240
      If UCase(Left(gridTenantLookup.TextMatrix(i, 0), Len(txtSearchTenant.text))) <> UCase(txtSearchTenant.text) Then
         gridTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Public Function PopulateTenantInformation(ByVal adoConn As ADODB.Connection, ByVal sTenantSageAC As String) As Boolean
   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT TENANTS.*, LEASEINFO.* " _
       & " FROM TENANTS LEFT JOIN (" _
       & "  SELECT " _
       & "     CLIENT.CLIENTNAME AS CLIENT, CLIENT.ClientID, " _
       & "     PROPERTY.PROPERTYID + '-' + PROPERTY.PROPERTYNAME AS PROPERTY, " _
       & "     UNITS.UNITNUMBER + '-'+ UNITS.UNITNAME AS UNIT, " _
       & "     LeaseDetails.LeaseID,LeaseDetails.SageAccountNumber as LeaseSAGEAC, " _
       & "     LeaseDetails.HoldingOver, UNITS.UNITNUMBER, " _
       & "     LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.RentReviewDate " _
       & "  From " _
       & "     LEASEDETAILS, UNITS , CLIENT, TENANTS, PROPERTY " _
       & "  Where LEASEDETAILS.STATUS = TRUE AND " _
       & "     LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " _
       & "     UNITS.PROPERTYID = PROPERTY.PROPERTYID AND " _
       & "     Property.CLIENTID = CLIENT.CLIENTID AND " _
       & "     LEASEDETAILS.SageAccountNumber=TENANTS.SageAccountNumber " _
       & " )AS LEASEINFO ON TENANTS.SAGEACCOUNTNUMBER = LEASEINFO.LeaseSAGEAC " _
       & " WHERE SageAccountNumber = '" & sTenantSageAC & "'"
'Debug.Print sSQLQuery_
   If Not FillFormUsingADODB(Me, adoConn, sSQLQuery_) Then
      MsgBox "WARNING !! No information found for the specified Lessee.", vbExclamation
   End If

   bLeaseSetup = True
   If txtLeaseId.text = "" Then
       MsgBox "WARNING !! There is no Lease setup for this Lessee.", vbExclamation
       bLeaseSetup = False
   End If
End Function

Public Sub PopulateCodes(adoConn As ADODB.Connection)
   Dim sSQLQuery As String

   ' Event Type
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'MTYP'"

   populateCombo adoConn, sSQLQuery, cboEventType

   ' Invoice Address
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'INVADD'"

   populateCombo adoConn, sSQLQuery, cboInvoiceTo

   ' Banks
   sSQLQuery = "SELECT BANK_ID, BANK_NAME, SORT_CODE, BANK_BRANCH, BANK_ADDRESS1, BANK_ADDRESS2, BANK_ADDRESS3, BANK_POST_CODE " & _
                 "FROM tlbBank "

   populateCombo adoConn, sSQLQuery, cboBankId
   
   ' Payment Method
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'PM'"

   populateCombo adoConn, sSQLQuery, cboPaymentMethod
End Sub

Public Sub populateGroupCombo(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRST As ADODB.Recordset
   Set adoRST = New ADODB.Recordset
   
   szSQL = "SELECT Distinct  TenantDeposit.GroupNo " & _
            "FROM TenantDeposit " & _
            "WHERE TenantDeposit.TenantID = '" & txtTenantID.text & "';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   Dim TotalRow As Long, TotalCol As Long

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   Dim Data() As String
   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   Dim i As Integer, j As Integer

   For i = 0 To adoRST.RecordCount - 1
      For j = 0 To adoRST.Fields.Count - 1
         Data(j, i) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
      Next j
      adoRST.MoveNext
   Next i

   cboGroup.Clear
   cboGroup.ColumnCount = TotalCol
   cboGroup.Column() = Data()
   cboGroup.BoundColumn = 1
   cboGroup.TextColumn = 1

   adoRST.Close
   Set adoRST = Nothing

   Exit Sub
'     Error Handling Code
Error_Handler:
   MsgBox "An Error occurred while populating the Group ID"
End Sub

Public Sub EventHistoryButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control
   Select Case mode

   Case ComponentMode.DefaultMode
      cmdNewEvent.Enabled = True
      cmdEditEvent.Enabled = False
      cmdSaveEvent.Enabled = False
      cmdCancelEvent.Enabled = False
'      gridEventHistory.Enabled = True

      cboEventType.Enabled = False
      dtpReportedDate.Enabled = False
      txtDescription.Enabled = False
      dtpDateCompleted.Enabled = False
      txtTaskOwner.Enabled = False
      txtContact.Enabled = False
      dtpRemindDate.Enabled = False
      chkAlarm.Enabled = False

   Case ComponentMode.GridRowOnSelection
      cmdNewEvent.Enabled = True
      cmdEditEvent.Enabled = True
      cmdSaveEvent.Enabled = False
      cmdCancelEvent.Enabled = False
'      gridEventHistory.Enabled = True

   Case ComponentMode.NewEntryMode
      cmdNewEvent.Enabled = False
      cmdEditEvent.Enabled = False
      cmdSaveEvent.Enabled = True
      cmdCancelEvent.Enabled = True
'      gridEventHistory.Enabled = False

      cboEventType.Enabled = True
      dtpReportedDate.Enabled = True
      txtDescription.Enabled = True
      txtDescription.text = ""
      dtpDateCompleted.Enabled = True
      txtTaskOwner.Enabled = True
      txtTaskOwner.text = ""
      txtContact.Enabled = True
      txtContact.text = ""
      dtpRemindDate.Enabled = True
      chkAlarm.Enabled = True
      chkAlarm.Value = 0

   Case ComponentMode.EditMode
      cmdNewEvent.Enabled = False
      cmdEditEvent.Enabled = False
      cmdSaveEvent.Enabled = True
      cmdCancelEvent.Enabled = True
'      gridEventHistory.Enabled = False

      cboEventType.Enabled = True
      dtpReportedDate.Enabled = True
      txtDescription.Enabled = True
      dtpDateCompleted.Enabled = True
      txtTaskOwner.Enabled = True
      txtContact.Enabled = True
      dtpRemindDate.Enabled = True
      chkAlarm.Enabled = True
   End Select
End Sub

Public Sub ConfigureFlxBank()
   With gridBank
      .Clear
      .Rows = 2
      .Cols = 7

      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = Label31(1).Left - Label31(0).Left
      .ColWidth(3) = Label31(2).Left - Label31(1).Left
      .ColWidth(4) = Label31(3).Left - Label31(2).Left
      .ColWidth(5) = Label31(4).Left - Label31(3).Left
      .ColWidth(6) = Label31(5).Left - Label31(4).Left
      .ColWidth(7) = .Width + .Left - Label31(5).Left - 200
   End With
End Sub

Public Function SaveTenantInformation() As Boolean
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String

   If Not NEWMODE_ Then
      sSQLQuery = "SELECT  SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                "TenantMemo, Balance, Deposite, BankCode, spare1, " & _
                "EmailDmd, EmailSt, CombEmail, EmailSC " & _
             "FROM TENANTS " & _
             "Where TenantID = '" & lblLeaseChanged.Caption & "'"

      If PostToDBUsingADODB(Me, fmeTenant, adoConn, sSQLQuery, False) Then
         SaveTenantInformation = True

         UpdateLesseeRec lblLeaseChanged.Caption, txtTenantID.text, adoConn
      Else
         SaveTenantInformation = False
      End If
   Else
      sSQLQuery = "SELECT  SageAccountNumber, TenantID, Name, CompanyName, Contact1, " & _
                  "Email1, DirectLine1, Contact2, Email2, DirectLine2, HOAddressLine1, " & _
                  "HOAddressLine2, HOAddressLine3, HOAddressLine4, HOPostCode, HOTelephone, " & _
                  "HOFax, BillAddressLine1, BillAddressLine2, BillAddressLine3, BillAddressLine4, " & _
                  "BillPostCode, BillTelephone, BillFax, InvoiceTo, " & _
                  "TenantMemo, Balance, Deposite, BankCode, spare1, " & _
                "EmailDmd, EmailSt, CombEmail, EmailSC " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & txtTenantID.text & "'"
       If PostToDBUsingADODB(Me, fmeTenant, adoConn, sSQLQuery, True) Then
           SaveTenantInformation = True
       Else
           SaveTenantInformation = False
       End If
   End If

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub UpdateLesseeRec(szOldID As String, szNewID As String, adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szTemp  As String

   szTemp = Replace(txtCompanyName.text, "'", "''")

   szSQL = "UPDATE DemandRecords " & _
           "SET TenantCompanyName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "';"
'Debug.Print szSQL
   adoConn.Execute szSQL

   szSQL = "UPDATE LeaseDetails " & _
           "SET CompanyName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "';"
   adoConn.Execute szSQL

   szTemp = Replace(txtName.text, "'", "''")
   szSQL = "UPDATE tlbLetterReports " & _
           "SET LesseeName = '" & szTemp & "' " & _
           "WHERE SageAccountNumber = '" & szOldID & "';"
   adoConn.Execute szSQL
End Sub

Private Sub TenantTabEnabled(ByVal IsEnabled As Boolean)
   tabTenant.Enabled = IsEnabled

   If IsEnabled Then
       ComponentInFrameEnableMode Me, fmeTenantAddress, DefaultMode
'       ComponentInFrameEnableMode Me, fmeTenancyDetails, EditMode
'       ComponentInFrameEnableMode Me, fmeBankPaymentDetails, DefaultMode
       ComponentInFrameEnableMode Me, fmeEventHistory, DefaultMode
   End If
End Sub

Private Sub LoadNominalCode()
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1400
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "N/C"
   lblSearch1(0).Caption = "Name"
   lblSearch2(0).Visible = False
   
   flxSupplier(0).RowHeight(0) = 0

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger " & _
           "WHERE ClientID = '" & txtClientID.text & "' " & _
           "ORDER BY Code;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("Code").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("Name").Value
      If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub ButtonHanlding(ByVal mode As ComponentMode)
   Dim X As String

   Select Case mode

   Case ComponentMode.DefaultMode
      cmdDptNew.Enabled = True
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = False
      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = False
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Deposit"
      cmdDptCancel.Caption = "Cancel Deposit"
      cmdDptEdit.Caption = "Edit Deposit"
      cmdDptDelete.Caption = "Delete Deposit"

      flxDeposit.Enabled = True

      cmbBank.Locked = True
      txtDate.Locked = True
      cmbDptAmtType.Locked = True
      cboDepositType.Locked = True
      cmdSetAmtType.Enabled = False
      cmdSetDptType.Enabled = False
      txtDptDetails.Locked = True
      txtDptAmount.Locked = True
      txtOSDpt.Locked = True

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = False
      cboGroup.Enabled = False
      cboGroup.Locked = True

      cmbBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = ""
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

      yDEPOSIT = 0

   Case ComponentMode.NewEntryMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False
      cmdDptExpenses.Enabled = False

      cmdDptSave.Caption = "Save Deposit"
      cmdDptCancel.Caption = "Cancel Deposit"
      cmdDptEdit.Caption = "Edit Deposit"
      cmdDptDelete.Caption = "Delete Deposit"

      flxDeposit.Enabled = False

      cmbBank.Locked = False
      txtDate.Locked = False

      cmbDptAmtType.Locked = False
      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False

      optNewGroup.Enabled = True
      optExitingGroup.Enabled = True
      cboGroup.Enabled = False
      cboGroup.Locked = True

      cmbBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = Format(Now, "dd/mm/yyyy")
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

   Case ComponentMode.EditMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save"
      cmdDptCancel.Caption = "Cancel"
      cmdDptEdit.Caption = "Edit"
      cmdDptDelete.Caption = "Delete"

      flxDeposit.Enabled = False

      cmbBank.Locked = False
      txtDate.Locked = False
      cmbDptAmtType.Locked = False
      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = True
      optExitingGroup.Value = True
      cboGroup.Enabled = True
      cboGroup.Locked = False

   Case ComponentMode.GridRowOnSelection
      cmdDptNew.Enabled = True
      cmdDptRefund.Enabled = True
      cmdDptExpenses.Enabled = True
      cmdDptEdit.Enabled = True
      cmdDptSave.Enabled = False
      cmdDptDelete.Enabled = True
      cmdDptCancel.Enabled = False
      cmdDptPrint.Enabled = True

      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "D" Then X = "Deposit"
      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "R" Then X = "Refund"
      If Left(flxDeposit.TextMatrix(flxDeposit.row, 2), 1) = "E" Then X = "Expenses"
      cmdDptSave.Caption = "Save " & X
      cmdDptEdit.Caption = "Edit " & X
      cmdDptDelete.Caption = "Delete " & X
      cmdDptCancel.Caption = "Cancel " & X

      cmbBank.Locked = True
      txtDate.Locked = True
      cmbDptAmtType.Locked = True
      cboDepositType.Locked = True
      cmdSetAmtType.Enabled = False
      cmdSetDptType.Enabled = False
      txtDptDetails.Locked = True
      txtDptAmount.Locked = True
      txtOSDpt.Locked = True

      optNewGroup.Enabled = False
      optExitingGroup.Enabled = False
      optExitingGroup.Value = True
      cboGroup.Enabled = False
      cboGroup.Locked = True

      yDEPOSIT = 0
      cmbBank.text = ""
      txtDNC(0).text = ""
      txtDNC(1).text = ""
      txtDate.text = ""
      txtDptDetails.text = ""
      txtDptAmount.text = ""
      txtOSDpt.text = ""

   Case ComponentMode.RefundMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Refund"
      cmdDptCancel.Caption = "Cancel Refund"
      txtDate.text = ""

      flxDeposit.Enabled = False

      cmbBank.Locked = False
      txtDate.Locked = False
      cmbDptAmtType.Locked = False
      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False
   
   Case ComponentMode.ExpensesMode
      cmdDptNew.Enabled = False
      cmdDptRefund.Enabled = False
      cmdDptExpenses.Enabled = False
      cmdDptEdit.Enabled = False
      cmdDptSave.Enabled = True
      cmdDptDelete.Enabled = False
      cmdDptCancel.Enabled = True
      cmdDptPrint.Enabled = False

      cmdDptSave.Caption = "Save Expenses"
      cmdDptCancel.Caption = "Cancel Expenses"

      flxDeposit.Enabled = False
      txtDptAmount.text = txtOSDpt.text
      txtDate.text = ""

      cmbBank.Locked = False
      txtDate.Locked = False
      cmbDptAmtType.Locked = False
      cboDepositType.Locked = False
      cmdSetAmtType.Enabled = True
      cmdSetDptType.Enabled = True
      txtDptDetails.Locked = False
      txtDptAmount.Locked = False
   End Select
End Sub

Private Sub ConfigFlxLetter()
   Dim szHeader As String, iCol As Integer

   flxLetters.Clear
   flxLetters.Cols = 6
   flxLetters.Rows = 2
   flxLetters.RowHeight(0) = 260
   szHeader$ = "|<ID|<TemplateName|<Subject|<Unit No|<Print Date"
   '           0  1     2           3         4        5
   flxLetters.FormatString = szHeader$
   flxLetters.ColWidth(0) = 200
   flxLetters.ColWidth(1) = 0
   flxLetters.ColWidth(2) = 3000
   flxLetters.ColWidth(3) = 5000
   flxLetters.ColWidth(4) = 0
   flxLetters.ColWidth(5) = 1600
End Sub

Private Sub ConfigFlxEmails()
   Dim szHeader As String, iCol As Integer

   flxEmails.Clear
   flxEmails.Cols = 7
   flxEmails.Rows = 2
   flxEmails.RowHeight(0) = 260
   szHeader$ = "|<Date|<Time|<Email|<Subject|<Body|ID"
   '           0    1    2      3       4       5   6
   flxEmails.FormatString = szHeader$
   flxEmails.ColWidth(0) = 200
   flxEmails.ColWidth(1) = 1000
   flxEmails.ColWidth(2) = 700
   flxEmails.ColWidth(3) = 2600
   flxEmails.ColWidth(4) = 3450
   flxEmails.ColWidth(5) = 3100
   flxEmails.ColWidth(6) = 0
End Sub

Public Sub LoadFlxEmails()
   Dim bLesseeEmail  As Boolean
   Dim bFlag         As Boolean
   Dim szPath        As String
   Dim szLine        As String
   Dim szaDateTime() As String
   Dim iRow          As Integer

   ConfigFlxEmails

   szPath = DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & txtTenantID.text & ".dat"

   bLesseeEmail = FileExists(szPath)

   If bLesseeEmail Then
      Open szPath For Input As #2

      iRow = 1
      While Not EOF(2)
         Line Input #2, szLine
         bFlag = False
'                                szHeader$ = "|<Date|<Time|<Email|<Subject|<Body|ID"
         If InStr(szLine, "Email sent on:") > 0 Then
            szLine = Mid(szLine, 15)
            szaDateTime = Split(szLine, "#")
            flxEmails.TextMatrix(iRow, 1) = szaDateTime(0)
            If UBound(szaDateTime) > 0 Then
               flxEmails.TextMatrix(iRow, 2) = szaDateTime(1)
            End If
            If UBound(szaDateTime) > 1 Then
               flxEmails.TextMatrix(iRow, 6) = szaDateTime(2)           'ID
            End If
            bFlag = True
         End If
         If InStr(szLine, "Email Address:") > 0 Then
            szLine = Mid(szLine, 15)
            flxEmails.TextMatrix(iRow, 3) = szLine
            bFlag = True
         End If
         If InStr(szLine, "Email Subject:") > 0 Then
            szLine = Mid(szLine, 15)
            flxEmails.TextMatrix(iRow, 4) = szLine
            bFlag = True
         End If
         If Not bFlag Then
            If InStr(szLine, "*****") = 0 Then
               flxEmails.TextMatrix(iRow, 5) = szLine
               While Not EOF(2) And InStr(szLine, "*****") = 0
                  Line Input #2, szLine
               Wend
            Else
               flxEmails.TextMatrix(iRow, 5) = Mid(szLine, 1, Len(szLine) - 5)
            End If
            iRow = iRow + 1
            If Not EOF(2) Then flxEmails.AddItem ""
         End If
      Wend
      Close #2
   End If

   flxEmails.col = 1
   flxEmails.ColSel = 2
   flxEmails.Sort = flexSortGenericDescending
   flxEmails.row = 0
End Sub

Private Sub LoadFlxLetter(conUnit_ As ADODB.Connection)
   Dim rstReport As New ADODB.Recordset
   Dim sSQLQuery_ As String, iRow As Integer

   ConfigFlxLetter

   sSQLQuery_ = "SELECT L.Id, T.TemplateName, " & _
                     "T.Description AS Subject, L.UnitNo, L.printDate " & _
                "FROM tlbLetterReports AS L, Template AS T " & _
                "WHERE L.sageAccountNumber = '" & txtTenantID.text & "' AND " & _
                     " L.TemplateID = T.TemplateID;"

   rstReport.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

'   szHeader$ = "|<ID|<TemplateName|<Subject|<Unit No|<Print Date"
   iRow = 1
   With flxLetters
      While Not rstReport.EOF
         .TextMatrix(iRow, 1) = IIf(IsNull(rstReport!Id), "", rstReport!Id)
         .TextMatrix(iRow, 2) = IIf(IsNull(rstReport!TemplateName), "", rstReport!TemplateName)
         .TextMatrix(iRow, 3) = IIf(IsNull(rstReport!Subject), "", rstReport!Subject)
         .TextMatrix(iRow, 4) = IIf(IsNull(rstReport!UnitNo), "", rstReport!UnitNo)
         .TextMatrix(iRow, 5) = IIf(IsNull(rstReport!PrintDate), "", rstReport!PrintDate)
         iRow = iRow + 1
         rstReport.MoveNext
         If Not rstReport.EOF Then .AddItem ""
      Wend
   End With
   rstReport.Close
   Set rstReport = Nothing
End Sub

Private Sub txtSearchTenant_GotFocus()
   SelTxtInCtrl txtSearchTenant
End Sub

Private Sub txtTenantID_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtTenantID_LostFocus()
   If txtTenantID.Locked Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim szID As String

   adoConn.Open getConnectionString

   szID = txtTenantID.text

   If (IsAccountExist(szID, adoConn)) Then
      If (Not (txtTenantID.text = szID)) Then
         MsgBox "The ID is already in use. Possible suggestion is '" & szID & "' and you may chose different ID"
         txtTenantID.text = szID
         SelTxtInCtrl txtTenantID
      End If
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub ConfigGridMaintenanceHistory(ByVal rstMHistory_ As ADODB.Recordset)
   Dim iColumn As Integer
   Dim oColumn As ADODB.Field

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = rstMHistory_.Fields.Count + 1

   For iColumn = 1 To 10
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn - 1).Left - 70

   For iColumn = 11 To rstMHistory_.Fields.Count
      gridMaintenanceHistory.ColWidth(iColumn) = 0
   Next iColumn

   iColumn = 0
   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.RowHeight(0) = 0
   For Each oColumn In rstMHistory_.Fields
      gridMaintenanceHistory.TextMatrix(0, iColumn) = oColumn.Name
      gridMaintenanceHistory.col = iColumn
      gridMaintenanceHistory.CellFontBold = True
      iColumn = iColumn + 1
   Next oColumn
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
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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
