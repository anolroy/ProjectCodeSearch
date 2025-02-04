VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BB5807FE-DBD2-11D3-87C1-4C980CC10374}#1.0#0"; "MyHover.ocx"
Begin VB.Form frmPurchaseExpense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase & Expenses"
   ClientHeight    =   13320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   23595
   Icon            =   "frmPurchaseExpense.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13320
   ScaleWidth      =   23595
   Begin VB.CommandButton cmdSPClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   355
      Left            =   14490
      Style           =   1  'Graphical
      TabIndex        =   398
      Top             =   10170
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.PictureBox picPurchaseHistory 
      BackColor       =   &H00E5E5E5&
      Height          =   1605
      Left            =   10485
      ScaleHeight     =   1545
      ScaleWidth      =   3600
      TabIndex        =   386
      Top             =   11430
      Visible         =   0   'False
      Width           =   3660
      Begin VB.CommandButton cmdPrintHistOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   390
         Top             =   1065
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrintHistCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2055
         TabIndex        =   392
         Top             =   1095
         Width           =   1200
      End
      Begin VB.TextBox txtStartDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         MaxLength       =   80
         TabIndex        =   388
         Top             =   450
         Width           =   1200
      End
      Begin VB.TextBox txtEndDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         MaxLength       =   80
         TabIndex        =   389
         Top             =   450
         Width           =   1200
      End
      Begin VB.CommandButton cmdClosePrintHIst 
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
         TabIndex        =   387
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   55
         Left            =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   30
         Left            =   0
         Top             =   260
         Width           =   3855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase History Print"
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
         Index           =   10
         Left            =   75
         TabIndex        =   393
         Top             =   0
         Width           =   1740
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   615
         Index           =   21
         Left            =   75
         Top             =   360
         Width           =   3450
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   660
         Index           =   20
         Left            =   75
         Top             =   360
         Width           =   3450
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
         Index           =   5
         Left            =   180
         TabIndex        =   391
         Top             =   495
         Width           =   360
      End
   End
   Begin VB.PictureBox picSupplierType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   19440
      ScaleHeight     =   2145
      ScaleWidth      =   2235
      TabIndex        =   382
      Top             =   9945
      Visible         =   0   'False
      Width           =   2265
      Begin VB.CommandButton cmdSupplierTypeClose 
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
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   383
         Top             =   90
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierType 
         Height          =   1725
         Left            =   45
         TabIndex        =   384
         Top             =   405
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   3043
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Category"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   385
         Top             =   135
         Width           =   1260
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   19
         Left            =   40
         Top             =   90
         Width           =   1800
      End
   End
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   6975
      ScaleHeight     =   450
      ScaleWidth      =   3195
      TabIndex        =   374
      Top             =   6345
      Visible         =   0   'False
      Width           =   3195
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading......"
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
         Height          =   330
         Left            =   405
         TabIndex        =   375
         Top             =   135
         Width           =   4590
      End
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   2220
      Left            =   5715
      TabIndex        =   348
      Top             =   11070
      Visible         =   0   'False
      Width           =   4530
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         Height          =   2100
         Left            =   45
         ScaleHeight     =   2040
         ScaleWidth      =   4410
         TabIndex        =   349
         Top             =   45
         Width           =   4470
         Begin VB.CommandButton cmdClearSearch 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2205
            TabIndex        =   360
            Top             =   1620
            Width           =   1110
         End
         Begin VB.CommandButton cmdOkSearch 
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   135
            TabIndex        =   358
            Top             =   1620
            Width           =   930
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
            Left            =   4095
            Style           =   1  'Graphical
            TabIndex        =   362
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2565
            MaxLength       =   80
            TabIndex        =   357
            Top             =   1125
            Width           =   1200
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            MaxLength       =   80
            TabIndex        =   356
            Top             =   1125
            Width           =   1200
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   355
            ToolTipText     =   "Press Enter Key to run search"
            Top             =   810
            Width           =   2460
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   354
            ToolTipText     =   "Press Enter Key to run search"
            Top             =   450
            Width           =   2460
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1125
            TabIndex        =   359
            Top             =   1620
            Width           =   1020
         End
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3375
            TabIndex        =   361
            Top             =   1620
            Width           =   975
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
            Left            =   450
            TabIndex        =   353
            Top             =   1170
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
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
            Left            =   450
            TabIndex        =   352
            Top             =   810
            Width           =   645
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
            Left            =   450
            TabIndex        =   351
            Top             =   450
            Width           =   225
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   1155
            Index           =   18
            Left            =   75
            Top             =   360
            Width           =   4260
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   1155
            Index           =   17
            Left            =   75
            Top             =   360
            Width           =   4215
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
            TabIndex        =   350
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0FFFF&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   30
            Left            =   0
            Top             =   255
            Width           =   4350
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   60
            Left            =   0
            Top             =   240
            Width           =   4350
         End
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   16965
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   325
      Top             =   4365
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
         TabIndex        =   326
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   334
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
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   332
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
         TabIndex        =   331
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
         Left            =   1875
         TabIndex        =   330
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
         TabIndex        =   329
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
         TabIndex        =   328
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   327
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   15
         Left            =   0
         Top             =   80
         Width           =   5355
      End
   End
   Begin VB.PictureBox picAccList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3220
      Left            =   17775
      ScaleHeight     =   3195
      ScaleWidth      =   5895
      TabIndex        =   199
      Top             =   3690
      Visible         =   0   'False
      Width           =   5925
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
         Index           =   2
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2535
         Index           =   2
         Left            =   15
         TabIndex        =   201
         Top             =   645
         Width           =   5860
         _ExtentX        =   10345
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   210
         Top             =   375
         Width           =   2415
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4260;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   4
         Left            =   930
         TabIndex        =   209
         Top             =   375
         Width           =   1470
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2602;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   208
         Top             =   375
         Width           =   900
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1587;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   207
         Top             =   120
         Width           =   1335
         VariousPropertyBits=   8388627
         Caption         =   "Account Name"
         Size            =   "2355;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   206
         Top             =   120
         Width           =   1095
         VariousPropertyBits=   8388627
         Caption         =   "Account ID"
         Size            =   "1931;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   6
         Left            =   30
         TabIndex        =   205
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "A/C type"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   3
         Left            =   1515
         TabIndex        =   204
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   3
         Left            =   2115
         TabIndex        =   203
         Top             =   1200
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   202
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "NotLoaded"
         Size            =   "1296;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   14
         Left            =   0
         Top             =   90
         Width           =   5580
      End
   End
   Begin VB.PictureBox picAccounts 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   17100
      ScaleHeight     =   3330
      ScaleWidth      =   6390
      TabIndex        =   187
      Top             =   600
      Visible         =   0   'False
      Width           =   6420
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
         Index           =   1
         Left            =   6090
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2445
         Index           =   1
         Left            =   45
         TabIndex        =   193
         Top             =   855
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4313
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   9
         Left            =   5085
         TabIndex        =   347
         Top             =   315
         Width           =   1545
         VariousPropertyBits=   8388627
         Caption         =   "This Client"
         Size            =   "2725;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   265
         Index           =   7
         Left            =   5085
         TabIndex        =   346
         Top             =   540
         Width           =   1065
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1879;467"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   6
         Left            =   4050
         TabIndex        =   345
         Top             =   540
         Width           =   1020
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1799;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   4
         Left            =   4155
         TabIndex        =   198
         Top             =   300
         Width           =   870
         VariousPropertyBits=   8388627
         Caption         =   "A/C Balance"
         Size            =   "1535;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   1
         Left            =   2115
         TabIndex        =   197
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   1
         Left            =   1515
         TabIndex        =   196
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   195
         Top             =   300
         Visible         =   0   'False
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "A/C type"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   194
         Top             =   300
         Width           =   1095
         VariousPropertyBits=   8388627
         Caption         =   "Account ID"
         Size            =   "1931;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   3
         Left            =   1395
         TabIndex        =   189
         Top             =   300
         Width           =   1335
         VariousPropertyBits=   8388627
         Caption         =   "Account Name"
         Size            =   "2355;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   265
         Index           =   0
         Left            =   30
         TabIndex        =   190
         Top             =   540
         Visible         =   0   'False
         Width           =   900
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1587;467"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   1
         Left            =   225
         TabIndex        =   191
         Top             =   540
         Width           =   1155
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2037;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountSearch 
         Height          =   270
         Index           =   2
         Left            =   1410
         TabIndex        =   192
         Top             =   540
         Width           =   2595
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4577;476"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   13
         Left            =   5400
         Top             =   0
         Visible         =   0   'False
         Width           =   5580
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   16
         Left            =   0
         Top             =   255
         Width           =   6345
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2115
      Index           =   0
      Left            =   525
      TabIndex        =   169
      Top             =   11145
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtChqNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   480
         TabIndex        =   154
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdChqRemittNo 
         Caption         =   "&No"
         Height          =   365
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdChqRemittYes 
         Caption         =   "&Yes"
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optChqRemitt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cheque with Remittance"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   171
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optRemittanceOnly 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remittance Only"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   170
         Top             =   320
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   480
         TabIndex        =   175
         Top             =   960
         Width           =   825
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1335
         Index           =   11
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Do you wish to print?"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   174
         Top             =   20
         Width           =   1680
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H0080FF80&
         BorderWidth     =   3
         Height          =   1335
         Index           =   12
         Left            =   120
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   17010
      ScaleHeight     =   4290
      ScaleWidth      =   5310
      TabIndex        =   52
      Top             =   8595
      Visible         =   0   'False
      Width           =   5340
      Begin VB.CheckBox chkShowBal 
         BackColor       =   &H80000009&
         Caption         =   "Show Bal"
         Height          =   195
         Left            =   4050
         TabIndex        =   379
         Top             =   405
         Visible         =   0   'False
         Width           =   1140
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
         Index           =   0
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3570
         Index           =   0
         Left            =   45
         TabIndex        =   216
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6297
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
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
         TabIndex        =   220
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   219
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   217
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch1 
         Height          =   195
         Left            =   1875
         TabIndex        =   215
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Left            =   3855
         TabIndex        =   213
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   212
         Top             =   375
         Width           =   1305
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2302;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1350
         TabIndex        =   214
         Top             =   375
         Width           =   1215
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2143;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   75
         Width           =   5355
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1875
      Index           =   1
      Left            =   13695
      TabIndex        =   135
      Top             =   11010
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkSettleAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Settle All"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   140
         Top             =   1140
         Width           =   1095
      End
      Begin VB.OptionButton optOIF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "The Oldest invoice first"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRIF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "The Recent invoices first"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdAutoAllocSel 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAutoAllocSelCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   855
         Index           =   7
         Left            =   120
         Top             =   260
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C00000&
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Height          =   855
         Index           =   8
         Left            =   120
         Top             =   260
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Option:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   141
         Top             =   20
         Width           =   1005
      End
   End
   Begin VB.Frame fraInvCrChoice 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   1815
      Left            =   3390
      TabIndex        =   54
      Top             =   11505
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E5E5E5&
         Height          =   1695
         Left            =   40
         ScaleHeight     =   1635
         ScaleWidth      =   3555
         TabIndex        =   55
         Top             =   50
         Width           =   3615
         Begin VB.OptionButton optManualAdjInv 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Adjustment Invoice"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   1335
            TabIndex        =   61
            Top             =   390
            Width           =   1695
         End
         Begin VB.OptionButton optManualAdjCrNote 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Adjustment Credit Note"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   1335
            TabIndex        =   60
            Top             =   735
            Width           =   2055
         End
         Begin VB.OptionButton optManualCrNote 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Credit Note"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   59
            Top             =   735
            Width           =   1200
         End
         Begin VB.OptionButton optManualInv 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Invoice"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   405
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdManualDmdOk 
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   1170
            Width           =   1200
         End
         Begin VB.CommandButton cmdManualDmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2235
            TabIndex        =   56
            Top             =   1180
            Width           =   1200
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   55
            Left            =   0
            Top             =   240
            Width           =   3855
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00C0FFFF&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   30
            Left            =   0
            Top             =   260
            Width           =   3855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Add a Transaction"
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
            Index           =   26
            Left            =   80
            TabIndex        =   62
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   5
            Left            =   75
            Top             =   360
            Width           =   3360
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   4
            Left            =   75
            Top             =   360
            Width           =   3360
         End
      End
   End
   Begin TabDlg.SSTab tabPurExp 
      Height          =   10875
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   16830
      _ExtentX        =   29686
      _ExtentY        =   19182
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Purchase Invoices && Credit Notes"
      TabPicture(0)   =   "frmPurchaseExpense.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraEditDemand"
      Tab(0).Control(1)=   "fraTab0"
      Tab(0).Control(2)=   "fraLay(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Payments"
      TabPicture(1)   =   "frmPurchaseExpense.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabPayment"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Purchase History"
      TabPicture(2)   =   "frmPurchaseExpense.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTab2"
      Tab(2).Control(1)=   "txtDisplayMaxPurchaseHist"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Purchase Payment History"
      TabPicture(3)   =   "frmPurchaseExpense.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraTab3"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraLay 
         BackColor       =   &H00DFDFDF&
         Height          =   10440
         Index           =   0
         Left            =   -72390
         TabIndex        =   12
         Top             =   495
         Width           =   16680
         Begin VB.CommandButton cmdaddnewline 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add a &New Line"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   14445
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   394
            Top             =   1215
            Width           =   1560
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
            Index           =   3
            Left            =   16380
            Style           =   1  'Graphical
            TabIndex        =   381
            Top             =   135
            Width           =   255
         End
         Begin VB.CommandButton cmdOpenFileView 
            Caption         =   "&Open File"
            Height          =   360
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   324
            Top             =   8550
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.CommandButton cmdviewMenu 
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
            Height          =   290
            Left            =   8460
            TabIndex        =   320
            Top             =   8550
            Width           =   255
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Attach Files:"
            ForeColor       =   &H00000000&
            Height          =   1320
            Left            =   8820
            TabIndex        =   315
            Top             =   8550
            Visible         =   0   'False
            Width           =   1545
            Begin VB.CommandButton cmdCross 
               Caption         =   "X"
               Height          =   195
               Left            =   1260
               TabIndex        =   322
               Top             =   90
               Width           =   240
            End
            Begin VB.CommandButton cmdOpenFile 
               Caption         =   "&Open File"
               Height          =   360
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   318
               Top             =   585
               Width           =   1110
            End
            Begin VB.CommandButton cmdClinetAddAtch 
               Caption         =   "&Add New"
               Height          =   315
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   317
               Top             =   270
               Width           =   1110
            End
            Begin VB.CommandButton cmdDeleteFile 
               Caption         =   "&Delete File"
               Height          =   315
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   316
               Top             =   945
               Width           =   1110
            End
         End
         Begin VB.TextBox txtDept 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1980
            Width           =   1650
         End
         Begin VB.TextBox txtDept 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6030
            Locked          =   -1  'True
            TabIndex        =   312
            Top             =   1980
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox chkRecover 
            Appearance      =   0  'Flat
            BackColor       =   &H00DFDFDF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1575
            TabIndex        =   32
            Top             =   2595
            Width           =   255
         End
         Begin VB.CommandButton cmdNCList 
            Appearance      =   0  'Flat
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
            Height          =   290
            Left            =   3555
            TabIndex        =   16
            Top             =   2295
            Width           =   285
         End
         Begin VB.TextBox txtNC 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   2295
            Width           =   1650
         End
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   13695
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   2640
            Width           =   1120
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit the Line"
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
            Height          =   355
            Index           =   0
            Left            =   2925
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   8505
            Width           =   1215
         End
         Begin VB.Frame fraLay 
            BackColor       =   &H00DFDFDF&
            Height          =   1020
            Index           =   1
            Left            =   180
            TabIndex        =   143
            Top             =   135
            Width           =   16140
            Begin VB.CheckBox chkIsMgtFee 
               Height          =   255
               Left            =   11610
               TabIndex        =   150
               Top             =   225
               Width           =   255
            End
            Begin VB.CommandButton cmdClientSerc 
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
               Height          =   285
               Left            =   15795
               TabIndex        =   151
               Top             =   180
               Width           =   285
            End
            Begin VB.TextBox txtDueDate 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8520
               TabIndex        =   149
               Top             =   585
               Width           =   1260
            End
            Begin VB.CommandButton cmdTypeList 
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
               Height          =   285
               Left            =   15780
               TabIndex        =   152
               Top             =   620
               Width           =   285
            End
            Begin VB.TextBox txtProperty 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   13260
               TabIndex        =   163
               Top             =   600
               Width           =   2520
            End
            Begin VB.TextBox txtDate 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8520
               TabIndex        =   148
               Top             =   240
               Width           =   1005
            End
            Begin VB.TextBox txtTransType 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   144
               Top             =   240
               Width           =   1275
            End
            Begin VB.CommandButton cmdACList 
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
               Height          =   285
               Index           =   0
               Left            =   4965
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   240
               Width           =   285
            End
            Begin VB.TextBox txtSupplierID 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3465
               Locked          =   -1  'True
               TabIndex        =   153
               Top             =   240
               Width           =   1485
            End
            Begin VB.TextBox txtReference 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3840
               MaxLength       =   20
               TabIndex        =   147
               Top             =   600
               Width           =   3705
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Is Mgt Fee"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   10755
               TabIndex        =   400
               Top             =   270
               Width           =   735
            End
            Begin MSForms.TextBox txtClientID 
               Height          =   285
               Left            =   13275
               TabIndex        =   333
               Top             =   180
               Width           =   2520
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "4445;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label lblPostingDate 
               Height          =   285
               Left            =   9555
               TabIndex        =   307
               Top             =   240
               Width           =   210
               ForeColor       =   8421504
               BackColor       =   16761024
               Caption         =   " P"
               Size            =   "379;503"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.ComboBox cmbSC 
               Height          =   285
               Left            =   1500
               TabIndex        =   145
               Top             =   600
               Width           =   1275
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2249;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Client:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   12585
               TabIndex        =   211
               Top             =   240
               Width           =   465
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Due Date:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   7740
               TabIndex        =   183
               Top             =   600
               Width           =   705
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Property:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   12585
               TabIndex        =   164
               Top             =   600
               Width           =   645
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/C:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   3045
               TabIndex        =   160
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   7740
               TabIndex        =   159
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trans Type:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
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
               TabIndex        =   158
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   3000
               TabIndex        =   157
               Top             =   600
               Width           =   750
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Account Category:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   156
               Top             =   600
               Width           =   1455
            End
            Begin MSForms.TextBox txtSupplierName 
               Height          =   285
               Left            =   5265
               TabIndex        =   155
               Top             =   240
               Width           =   2280
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "4022;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&OK"
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
            Height          =   375
            Index           =   1
            Left            =   14865
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2640
            Width           =   1120
         End
         Begin VB.Frame fraCmds 
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   30
            TabIndex        =   39
            Top             =   9225
            Width           =   16665
            Begin VB.CommandButton cmdSavePIRef 
               Caption         =   "&Save"
               Height          =   400
               Left            =   3510
               TabIndex        =   408
               Top             =   270
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton cmdSavePI 
               Caption         =   "&Save"
               Height          =   400
               Left            =   4950
               TabIndex        =   26
               Top             =   270
               Width           =   1455
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "C&lose"
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
               Index           =   0
               Left            =   14010
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   255
               Width           =   1450
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
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
               Index           =   0
               Left            =   6435
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   255
               Width           =   1450
            End
            Begin MyHoverButton.Button cmdNew 
               Height          =   405
               Index           =   0
               Left            =   8505
               TabIndex        =   305
               Top             =   270
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   714
               HoverBackColor  =   15066597
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmPurchaseExpense.frx":093A
               HoverPicture    =   "frmPurchaseExpense.frx":0956
               DisabledPicture =   "frmPurchaseExpense.frx":0972
               DownPicture     =   "frmPurchaseExpense.frx":098E
               MouseIcon       =   "frmPurchaseExpense.frx":09AA
               Caption         =   "Add New"
               HoverCaption    =   "Add New"
               DownCaption     =   ""
            End
         End
         Begin VB.TextBox txtPICNVat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   14340
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "vat"
            Top             =   8520
            Width           =   735
         End
         Begin VB.TextBox txtPICNNet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13350
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "net"
            Top             =   8520
            Width           =   975
         End
         Begin VB.TextBox txtDetails_ 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   8205
            MaxLength       =   80
            TabIndex        =   19
            Top             =   2280
            Width           =   4215
         End
         Begin VB.TextBox txtNet_ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   14115
            MaxLength       =   11
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   1725
            Width           =   1875
         End
         Begin VB.TextBox txtVat_ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   14775
            MaxLength       =   11
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   2025
            Width           =   1215
         End
         Begin VB.TextBox txtPICNTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   15105
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "total"
            Top             =   8520
            Width           =   975
         End
         Begin VB.CommandButton cmdTaxList 
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
            Height          =   285
            Index           =   0
            Left            =   14475
            TabIndex        =   21
            Top             =   2025
            Width           =   285
         End
         Begin VB.CommandButton cmdUnitList 
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
            Height          =   285
            Left            =   5655
            TabIndex        =   13
            Top             =   1680
            Width           =   285
         End
         Begin VB.CommandButton cmdDeptList 
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
            Height          =   290
            Left            =   3555
            TabIndex        =   14
            Top             =   1980
            Width           =   285
         End
         Begin VB.CommandButton cmdSchedules 
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
            Height          =   285
            Index           =   0
            Left            =   12150
            TabIndex        =   18
            Top             =   1980
            Width           =   285
         End
         Begin VB.CommandButton cmdJobNo 
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
            Height          =   285
            Index           =   0
            Left            =   12150
            TabIndex        =   17
            Top             =   1665
            Width           =   285
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPI 
            Height          =   4920
            Left            =   120
            TabIndex        =   25
            Top             =   3525
            Width           =   16530
            _ExtentX        =   29157
            _ExtentY        =   8678
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   12648447
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
         Begin VB.TextBox txtUnit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1680
            Width           =   3765
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete the Line"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   8505
            Width           =   1450
         End
         Begin MSForms.TextBox txtJobNo 
            Height          =   285
            Left            =   8190
            TabIndex        =   44
            Top             =   1680
            Width           =   3945
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "6959;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attachments:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   4410
            TabIndex        =   321
            Top             =   8595
            Width           =   930
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   5400
            TabIndex        =   319
            Top             =   8550
            Width           =   3000
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5292;503"
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
         Begin MSForms.CommandButton cmdPO 
            Height          =   375
            Left            =   10485
            TabIndex        =   31
            Top             =   8550
            Width           =   1455
            ForeColor       =   12582912
            BackColor       =   16761024
            Caption         =   "Purchase Order"
            Size            =   "2566;661"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2415
            TabIndex        =   185
            Top             =   2670
            Width           =   135
         End
         Begin MSForms.TextBox txtRecoverable 
            Height          =   255
            Index           =   0
            Left            =   1875
            TabIndex        =   33
            Top             =   2595
            Width           =   495
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "873;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recoverable:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   184
            Top             =   2595
            Width           =   915
         End
         Begin MSForms.TextBox txtSchedules 
            Height          =   285
            Left            =   8190
            TabIndex        =   166
            Top             =   1980
            Width           =   3945
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "6959;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtNCName 
            Height          =   285
            Left            =   3885
            TabIndex        =   165
            Top             =   2295
            Width           =   2040
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3598;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   12870
            TabIndex        =   142
            Top             =   8520
            Width           =   390
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sch ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   9045
            TabIndex        =   74
            Top             =   2700
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund Code"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   2835
            TabIndex        =   73
            Top             =   3240
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SchID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   7515
            TabIndex        =   72
            Tag             =   "new one"
            Top             =   2700
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   40
            Left            =   15375
            TabIndex        =   71
            Top             =   3240
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   14535
            TabIndex        =   70
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   13890
            TabIndex        =   69
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Net"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   12975
            TabIndex        =   68
            Top             =   3240
            Width           =   270
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   5460
            TabIndex        =   67
            Top             =   3240
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job No"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   4470
            TabIndex        =   66
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NC"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   2115
            TabIndex        =   65
            Top             =   3240
            Width           =   225
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit No"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   855
            TabIndex        =   64
            Top             =   3240
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   285
            TabIndex        =   63
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   13470
            TabIndex        =   51
            Top             =   2055
            Width           =   300
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/C:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   50
            Top             =   2295
            Width           =   315
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   49
            Top             =   1995
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   7260
            TabIndex        =   48
            Top             =   2280
            Width           =   870
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Net:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   13470
            TabIndex        =   47
            Top             =   1740
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Name:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   255
            TabIndex        =   46
            Top             =   1665
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job No:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   7260
            TabIndex        =   45
            Top             =   1680
            Width           =   510
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule ID:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   7260
            TabIndex        =   43
            Top             =   1980
            Width           =   885
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   13470
            TabIndex        =   42
            Top             =   2340
            Width           =   390
         End
         Begin MSForms.TextBox txtTotal 
            Height          =   285
            Left            =   14115
            TabIndex        =   23
            Top             =   2325
            Width           =   1875
            VariousPropertyBits=   679495705
            BorderStyle     =   1
            Size            =   "3307;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.Label lblVatCode 
            Height          =   165
            Index           =   0
            Left            =   14085
            TabIndex        =   41
            Top             =   2070
            Width           =   330
            VariousPropertyBits=   8388627
            Size            =   "582;291"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtPFName 
            Height          =   285
            Left            =   3885
            TabIndex        =   40
            Top             =   1980
            Width           =   2040
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3598;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Index           =   49
            Left            =   120
            TabIndex        =   176
            Top             =   3195
            Width           =   16530
         End
      End
      Begin VB.TextBox txtDisplayMaxPurchaseHist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -61140
         MaxLength       =   80
         TabIndex        =   368
         Top             =   10215
         Width           =   1065
      End
      Begin VB.Frame fraTab0 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   10320
         Left            =   -73605
         TabIndex        =   274
         Top             =   480
         Width           =   15210
         Begin VB.CheckBox chkProperty 
            Caption         =   "Excl."
            Height          =   195
            Left            =   8640
            TabIndex        =   376
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox txtSupplier 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10935
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   373
            Top             =   315
            Width           =   1635
         End
         Begin VB.TextBox txtPropID 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6075
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   342
            Top             =   315
            Width           =   2130
         End
         Begin VB.TextBox txtIDClient 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   340
            Top             =   315
            Width           =   2175
         End
         Begin VB.CheckBox chkSelectAllDemands 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   0
            TabIndex        =   277
            Top             =   825
            Width           =   215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchaseSplit 
            Height          =   1635
            Left            =   45
            TabIndex        =   276
            Top             =   8730
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   2884
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   10
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchase 
            Height          =   7215
            Left            =   0
            TabIndex        =   275
            Top             =   1080
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   12726
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   12
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Client ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   1035
            TabIndex        =   397
            Top             =   8370
            Width           =   630
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Controls count"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   12690
            TabIndex        =   395
            Top             =   90
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   5355
            TabIndex        =   378
            Top             =   315
            Width           =   615
         End
         Begin MSForms.CommandButton cmdOpProperty 
            Height          =   285
            Left            =   8235
            TabIndex        =   300
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdOpClient 
            Height          =   285
            Left            =   2835
            TabIndex        =   299
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   10260
            TabIndex        =   302
            Top             =   315
            Width           =   855
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   3
            Left            =   0
            Top             =   90
            Width           =   15030
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   301
            Top             =   315
            Width           =   705
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   7815
            TabIndex        =   298
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   4515
            TabIndex        =   297
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   6735
            TabIndex        =   296
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   12450
            TabIndex        =   295
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   294
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "A/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   2700
            TabIndex        =   293
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   1740
            TabIndex        =   292
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   945
            TabIndex        =   291
            Top             =   840
            Width           =   345
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   2310
            TabIndex        =   290
            Top             =   8385
            Width           =   690
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   3165
            TabIndex        =   289
            Top             =   8385
            Width           =   1125
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "N/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   4755
            TabIndex        =   288
            Top             =   8385
            Width           =   285
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   287
            Top             =   8385
            Width           =   210
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   10515
            TabIndex        =   286
            Top             =   8385
            Width           =   675
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Job No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   6165
            TabIndex        =   285
            Top             =   8385
            Width           =   510
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   5415
            TabIndex        =   284
            Top             =   8385
            Width           =   360
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   7095
            TabIndex        =   283
            Top             =   8385
            Width           =   840
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   9780
            TabIndex        =   282
            Top             =   8385
            Width           =   360
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Net £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   8835
            TabIndex        =   281
            Top             =   8385
            Width           =   390
         End
         Begin VB.Label lblPurchaseSplit 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Recoverable"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   11550
            TabIndex        =   280
            Top             =   8385
            Width           =   885
         End
         Begin MSForms.CommandButton cmdAccSel 
            Height          =   285
            Left            =   12585
            TabIndex        =   279
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   13620
            TabIndex        =   278
            Top             =   840
            Width           =   1005
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   6
            Left            =   45
            Top             =   120
            Width           =   14985
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   195
            Index           =   19
            Left            =   0
            TabIndex        =   303
            Top             =   855
            Width           =   15090
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   285
            Index           =   270
            Left            =   45
            TabIndex        =   304
            Top             =   8355
            Width           =   14895
         End
      End
      Begin VB.Frame fraTab2 
         Height          =   10395
         Left            =   -74970
         TabIndex        =   246
         Top             =   360
         Width           =   16665
         Begin VB.TextBox txtPropertyIDHist 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   338
            Text            =   "ALL"
            Top             =   315
            Width           =   1545
         End
         Begin VB.CheckBox chkPropertyHist 
            Caption         =   "Excl."
            Height          =   195
            Left            =   6390
            TabIndex        =   337
            Top             =   315
            Width           =   1185
         End
         Begin VB.CheckBox chkAllPurchaseHistory 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   215
            Left            =   135
            TabIndex        =   367
            Top             =   945
            Width           =   215
         End
         Begin VB.CommandButton cmdSearchPurchaseHistory 
            Caption         =   "Sea&rch"
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
            Left            =   3375
            Style           =   1  'Graphical
            TabIndex        =   366
            Top             =   9720
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpenSupp 
            Caption         =   ". ."
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
            Left            =   10215
            TabIndex        =   343
            Top             =   315
            Width           =   315
         End
         Begin VB.CommandButton cmdOClientList 
            Caption         =   ". ."
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
            Left            =   2700
            TabIndex        =   336
            Top             =   285
            Width           =   315
         End
         Begin VB.CommandButton cmdRevHistory 
            Caption         =   "Reverse History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   313
            Top             =   9720
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrintListHistory 
            Caption         =   "Print List"
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
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   247
            Top             =   9720
            Width           =   1440
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistory 
            Height          =   5400
            Left            =   120
            TabIndex        =   248
            Top             =   1185
            Width           =   16365
            _ExtentX        =   28866
            _ExtentY        =   9525
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   12
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistorySplit 
            Height          =   2565
            Left            =   45
            TabIndex        =   249
            Top             =   7110
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   4524
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   12
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3555
            TabIndex        =   377
            Top             =   315
            Width           =   615
         End
         Begin MSForms.CommandButton cmdOpPropertyHist 
            Height          =   285
            Left            =   5985
            TabIndex        =   339
            Top             =   315
            Width           =   315
            Caption         =   "; ;"
            Size            =   "556;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label lblDisplay 
            Caption         =   "Display : "
            Height          =   195
            Left            =   13005
            TabIndex        =   369
            Top             =   9900
            Width           =   690
         End
         Begin MSForms.TextBox txtSupplierSearc 
            Height          =   255
            Left            =   8730
            TabIndex        =   341
            Top             =   315
            Width           =   1485
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "2619;450"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtClientIdlist 
            Height          =   255
            Left            =   1170
            TabIndex        =   335
            Top             =   285
            Width           =   1530
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "2699;450"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   271
            Top             =   975
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2580
            TabIndex        =   270
            Top             =   975
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier A/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3540
            TabIndex        =   269
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   268
            Top             =   975
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   14850
            TabIndex        =   267
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   8370
            TabIndex        =   266
            Top             =   945
            Width           =   255
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   630
            TabIndex        =   264
            Top             =   315
            Width           =   465
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   2
            Left            =   120
            Top             =   120
            Width           =   16350
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   5490
            TabIndex        =   262
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   10035
            TabIndex        =   261
            Top             =   945
            Width           =   840
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   7875
            TabIndex        =   260
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Net £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   12555
            TabIndex        =   259
            Top             =   6780
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   13800
            TabIndex        =   258
            Top             =   6780
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   7260
            TabIndex        =   257
            Top             =   6780
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   4635
            TabIndex        =   256
            Top             =   6780
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Job No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   6180
            TabIndex        =   255
            Top             =   6780
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   14985
            TabIndex        =   254
            Top             =   6780
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   195
            TabIndex        =   253
            Top             =   6780
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "N/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   3495
            TabIndex        =   252
            Top             =   6780
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   1635
            TabIndex        =   251
            Top             =   6780
            Width           =   1125
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Prop/Unit"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   675
            TabIndex        =   250
            Top             =   6780
            Width           =   690
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   1
            Left            =   120
            Top             =   120
            Width           =   16395
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   240
            Index           =   0
            Left            =   315
            TabIndex        =   272
            Top             =   900
            Width           =   16140
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   330
            Index           =   9
            Left            =   120
            TabIndex        =   273
            Top             =   6705
            Width           =   16335
         End
         Begin MSForms.ComboBox cmbPropertyHistory 
            Height          =   285
            Left            =   12285
            TabIndex        =   265
            Top             =   315
            Visible         =   0   'False
            Width           =   3525
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6218;503"
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
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   8985
            TabIndex        =   263
            Top             =   885
            Visible         =   0   'False
            Width           =   645
         End
      End
      Begin VB.Frame fraTab3 
         Height          =   10410
         Left            =   -75000
         TabIndex        =   221
         Top             =   270
         Width           =   16605
         Begin VB.TextBox txtTotalOSAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   14670
            MaxLength       =   80
            TabIndex        =   404
            Top             =   9630
            Width           =   1200
         End
         Begin VB.TextBox txtRctTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   13320
            MaxLength       =   80
            TabIndex        =   401
            Top             =   9630
            Width           =   1200
         End
         Begin VB.CommandButton cmdPrintallocationhistory 
            Caption         =   "Print Allocation history"
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
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   396
            Top             =   9630
            Width           =   3060
         End
         Begin VB.CommandButton cmdPrintPaymentList 
            Caption         =   "Print Payment List"
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
            Left            =   1350
            Style           =   1  'Graphical
            TabIndex        =   380
            Top             =   9630
            Width           =   1440
         End
         Begin VB.CommandButton cmdSearchPurchPayHistory 
            Caption         =   "Sea&rch"
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
            Left            =   3825
            Style           =   1  'Graphical
            TabIndex        =   372
            Top             =   9630
            Width           =   1170
         End
         Begin VB.TextBox txtDisplayMaxPurchPayHist 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10845
            MaxLength       =   80
            TabIndex        =   370
            Top             =   9630
            Width           =   1065
         End
         Begin VB.CommandButton cmdPurchasePaymentHistory 
            Caption         =   ". ."
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
            Left            =   4440
            TabIndex        =   363
            Top             =   225
            Width           =   315
         End
         Begin VB.CommandButton cmdOpSupSearch 
            Caption         =   ". ."
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
            Left            =   10935
            TabIndex        =   234
            Top             =   225
            Width           =   315
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchPPHistory 
            Height          =   5880
            Left            =   90
            TabIndex        =   222
            Top             =   1215
            Width           =   16320
            _ExtentX        =   28787
            _ExtentY        =   10372
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
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
            _Band(0).Cols   =   12
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchPPHistorySplit 
            Height          =   1800
            Left            =   90
            TabIndex        =   223
            Top             =   7470
            Width           =   16320
            _ExtentX        =   28787
            _ExtentY        =   3175
            _Version        =   393216
            Cols            =   12
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   12
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Allocation View  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   7920
            TabIndex        =   407
            Top             =   7110
            Width           =   1650
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "OS Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   14895
            TabIndex        =   405
            Top             =   810
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total OSAmount:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   51
            Left            =   14670
            TabIndex        =   403
            Top             =   9405
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   13365
            TabIndex        =   402
            Top             =   9405
            Width           =   990
         End
         Begin VB.Label Label2 
            Caption         =   "Display : "
            Height          =   195
            Left            =   9810
            TabIndex        =   371
            Top             =   9675
            Width           =   690
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   585
            TabIndex        =   365
            Top             =   255
            Width           =   465
         End
         Begin MSForms.TextBox txtPurchasePaymentHistory 
            Height          =   300
            Left            =   1155
            TabIndex        =   364
            Top             =   225
            Width           =   3285
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "5794;529"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSupSearchHis 
            Height          =   300
            Left            =   8775
            TabIndex        =   344
            Top             =   225
            Width           =   2160
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3810;529"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Property"
            Height          =   195
            Index           =   41
            Left            =   3135
            TabIndex        =   243
            Top             =   7230
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "N/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   5775
            TabIndex        =   242
            Top             =   7230
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   375
            TabIndex        =   241
            Top             =   7230
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   48
            Left            =   14445
            TabIndex        =   240
            Top             =   7185
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   44
            Left            =   6720
            TabIndex        =   239
            Top             =   5520
            Width           =   45
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   43
            Left            =   6720
            TabIndex        =   238
            Top             =   5520
            Width           =   45
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   45
            Left            =   6855
            TabIndex        =   237
            Top             =   7230
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   47
            Left            =   13485
            TabIndex        =   236
            Top             =   7185
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Net £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   46
            Left            =   12405
            TabIndex        =   235
            Top             =   7185
            Width           =   390
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   7965
            TabIndex        =   233
            Top             =   270
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   57
            Left            =   9720
            TabIndex        =   232
            Top             =   855
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   5280
            TabIndex        =   231
            Top             =   840
            Width           =   1035
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   660
            Index           =   9
            Left            =   120
            Top             =   135
            Width           =   16305
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   7680
            TabIndex        =   230
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   58
            Left            =   13050
            TabIndex        =   229
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   51
            Left            =   360
            TabIndex        =   228
            Top             =   855
            Width           =   240
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier A/C"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   54
            Left            =   4140
            TabIndex        =   227
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   53
            Left            =   3180
            TabIndex        =   226
            Top             =   855
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   52
            Left            =   1215
            TabIndex        =   225
            Top             =   855
            Width           =   345
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Client"
            Height          =   195
            Index           =   40
            Left            =   975
            TabIndex        =   224
            Top             =   7230
            Width           =   390
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   660
            Index           =   10
            Left            =   90
            Top             =   135
            Width           =   16350
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   240
            Index           =   157
            Left            =   120
            TabIndex        =   244
            Top             =   7200
            Width           =   16305
         End
         Begin VB.Label Label20 
            BackColor       =   &H00E0FFFF&
            Height          =   240
            Index           =   158
            Left            =   135
            TabIndex        =   245
            Top             =   765
            Width           =   16275
         End
      End
      Begin VB.Frame fraEditDemand 
         Height          =   10335
         Left            =   -74880
         TabIndex        =   162
         Top             =   480
         Width           =   1215
         Begin VB.CommandButton cmdfixNL 
            Caption         =   "Fix NL"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   409
            Top             =   5940
            Width           =   1080
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "&View"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   323
            Top             =   2520
            Width           =   1080
         End
         Begin MyHoverButton.Button cmdNew 
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   1
            Top             =   360
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
            Picture         =   "frmPurchaseExpense.frx":09C6
            HoverPicture    =   "frmPurchaseExpense.frx":09E2
            DisabledPicture =   "frmPurchaseExpense.frx":09FE
            DownPicture     =   "frmPurchaseExpense.frx":0A1A
            MouseIcon       =   "frmPurchaseExpense.frx":0A36
            Caption         =   "&Add New"
            HoverCaption    =   "Add New"
            DownCaption     =   "Add New"
         End
         Begin VB.CommandButton cmdCopy 
            BackColor       =   &H000000FF&
            Caption         =   "&Copy"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2970
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   9225
            Width           =   1080
         End
         Begin VB.CommandButton cmdPrintPI_List 
            Caption         =   "Print List"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4140
            Width           =   1080
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1410
            Width           =   1080
         End
         Begin VB.CommandButton cmdPostDemands 
            Caption         =   "Post to Hist."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5070
            Width           =   1080
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Sea&rch"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   8505
            Width           =   1080
         End
      End
      Begin TabDlg.SSTab tabPayment 
         Height          =   10305
         Left            =   135
         TabIndex        =   53
         Top             =   360
         Width           =   16665
         _ExtentX        =   29395
         _ExtentY        =   18177
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         ForeColor       =   4194368
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Purchase Payment"
         TabPicture(0)   =   "frmPurchaseExpense.frx":0A52
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label20(50)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(11)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(10)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(9)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(8)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(7)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1(5)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(4)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label1(3)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1(2)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label10(9)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label10(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label3(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label10(7)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label10(1)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label3(5)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "lblAllocating(1)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line2(0)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Line2(1)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "flxSCrPoA"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "flxSPayment"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtSPayment"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Frame8(1)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Frame5(5)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtCrPayment"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Frame5(1)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtAllocatedDiff(1)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "cmeRevereseAllocation"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "cmdPayAllocate"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "cmdAdvanceProgr"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "cmdFix"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Frame1"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).ControlCount=   34
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   90
            TabIndex        =   417
            Top             =   5400
            Width           =   15810
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   426
               Top             =   180
               Width           =   210
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   915
               TabIndex        =   425
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "P/U ID"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   3240
               TabIndex        =   424
               Top             =   180
               Width           =   450
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   4155
               TabIndex        =   423
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ref"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   5250
               TabIndex        =   422
               Top             =   180
               Width           =   225
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   6840
               TabIndex        =   421
               Top             =   180
               Width           =   510
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amount £"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   12195
               TabIndex        =   420
               Top             =   180
               Width           =   675
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "O/S Amt. £"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   13155
               TabIndex        =   419
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Payment £"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   14115
               TabIndex        =   418
               Top             =   180
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdFix 
            Caption         =   "Fix"
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
            Left            =   12825
            TabIndex        =   406
            Top             =   9495
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdAdvanceProgr 
            Caption         =   "Advance"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   10980
            Style           =   1  'Graphical
            TabIndex        =   399
            Top             =   9450
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton cmdPayAllocate 
            Caption         =   "All&ocation Only"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   182
            Top             =   9450
            Width           =   1700
         End
         Begin VB.CommandButton cmeRevereseAllocation 
            Caption         =   "&Reverse Allocation"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   9450
            Width           =   1700
         End
         Begin VB.TextBox txtAllocatedDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   1
            Left            =   12915
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "0.00"
            Top             =   8910
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Allocation:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   745
            Index           =   1
            Left            =   5145
            TabIndex        =   129
            Top             =   8700
            Visible         =   0   'False
            Width           =   3705
            Begin VB.CommandButton cmdPayAllocateSave 
               Caption         =   "Save"
               Enabled         =   0   'False
               Height          =   355
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   178
               Top             =   255
               Width           =   1080
            End
            Begin VB.CommandButton cmdPayAllocationDiscard 
               Caption         =   "Clear"
               Height          =   355
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   179
               Top             =   255
               Width           =   1080
            End
            Begin VB.CommandButton cmdPayAutomatic 
               Caption         =   "Automatic"
               Height          =   355
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   181
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "allocation ref."
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   195
               Index           =   5
               Left            =   1440
               TabIndex        =   130
               Top             =   0
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.TextBox txtCrPayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   13095
            MaxLength       =   13
            TabIndex        =   127
            Top             =   3780
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D5D5D5&
            Caption         =   "Payments:"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   745
            Index           =   5
            Left            =   120
            TabIndex        =   123
            Top             =   9240
            Width           =   5415
            Begin VB.CommandButton cmdEditPayment 
               Caption         =   "Edit"
               Height          =   355
               Left            =   3270
               Style           =   1  'Graphical
               TabIndex        =   114
               Top             =   275
               Width           =   960
            End
            Begin VB.CommandButton cmdSPayAll 
               Caption         =   "Pay &All"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   355
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   113
               Top             =   275
               Width           =   960
            End
            Begin VB.CommandButton cmdSPFull 
               Caption         =   "Pay in &Full"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   355
               Left            =   1170
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   275
               Width           =   960
            End
            Begin VB.CommandButton cmdPaymentDiscard 
               Caption         =   "Clear"
               Height          =   355
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   115
               Top             =   275
               Width           =   960
            End
            Begin VB.CommandButton cmdSPSave 
               Caption         =   "Save"
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
               Height          =   355
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   111
               Top             =   275
               Width           =   960
            End
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00D5D5D5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1440
            Index           =   1
            Left            =   75
            TabIndex        =   86
            Top             =   420
            Width           =   16530
            Begin VB.Frame Frame8 
               BackColor       =   &H00DEDEDE&
               Caption         =   "Bank Balance:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004040&
               Height          =   1275
               Index           =   0
               Left            =   14085
               TabIndex        =   410
               Top             =   45
               Width           =   2400
               Begin VB.TextBox txtAvailableBankBal1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   413
                  Text            =   "0.00"
                  Top             =   855
                  Width           =   1125
               End
               Begin VB.TextBox txtRetentions1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   412
                  Text            =   "0.00"
                  Top             =   525
                  Width           =   1125
               End
               Begin VB.TextBox txtBankBal1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   411
                  Text            =   "0.00"
                  Top             =   195
                  Width           =   1125
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Avail.Bank Bal£"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   416
                  Top             =   855
                  Width           =   1050
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Retentions  £"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
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
                  TabIndex        =   415
                  Top             =   525
                  Width           =   930
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Balance  £"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
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
                  TabIndex        =   414
                  Top             =   195
                  Width           =   1050
               End
            End
            Begin VB.CommandButton cmdSupplierType 
               Caption         =   ". ."
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
               Left            =   3420
               TabIndex        =   97
               Top             =   495
               Width           =   345
            End
            Begin VB.CommandButton cmdOpenClient 
               Caption         =   ". ."
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
               Left            =   3420
               TabIndex        =   95
               Top             =   130
               Width           =   345
            End
            Begin VB.TextBox txtSupAcBal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
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
               Left            =   10170
               Locked          =   -1  'True
               TabIndex        =   167
               Text            =   "0.00"
               Top             =   165
               Width           =   1215
            End
            Begin VB.TextBox txtSPDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   10170
               TabIndex        =   108
               Top             =   900
               Width           =   1215
            End
            Begin VB.TextBox txtSPaymentTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10170
               MaxLength       =   11
               TabIndex        =   107
               Text            =   "0.00"
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtSPReference 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   5235
               MaxLength       =   12
               TabIndex        =   106
               Top             =   915
               Width           =   2370
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00DEDEDE&
               Caption         =   "Payment Analysis:"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004040&
               Height          =   1275
               Index           =   3
               Left            =   11700
               TabIndex        =   87
               Top             =   45
               Width           =   2355
               Begin VB.TextBox txtPaymentTotal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   90
                  Text            =   "0.00"
                  Top             =   195
                  Width           =   1080
               End
               Begin VB.TextBox txtPaymentEntered 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   89
                  Text            =   "0.00"
                  Top             =   525
                  Width           =   1080
               End
               Begin VB.TextBox txtDiffPay 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000014&
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   285
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   88
                  Text            =   "0.00"
                  Top             =   855
                  Width           =   1080
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total                £"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   120
                  TabIndex        =   93
                  Top             =   195
                  Width           =   930
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Entered         £"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   120
                  TabIndex        =   92
                  Top             =   525
                  Width           =   930
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Difference    £"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   5
                  Left            =   120
                  TabIndex        =   91
                  Top             =   855
                  Width           =   960
               End
            End
            Begin MSForms.CommandButton cmdAmountType 
               Height          =   285
               Left            =   8325
               TabIndex        =   104
               Top             =   540
               Width           =   345
               Caption         =   "; ;"
               Size            =   "609;512"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdBankAc 
               Height          =   285
               Left            =   8655
               TabIndex        =   102
               Top             =   180
               Width           =   345
               Caption         =   "; ;"
               Size            =   "609;503"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.TextBox txtBankAc 
               Height          =   285
               Left            =   6165
               TabIndex        =   101
               Top             =   180
               Width           =   2475
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "4366;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtBankCode 
               Height          =   285
               Left            =   5220
               TabIndex        =   100
               Top             =   180
               Width           =   900
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "1587;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtSupplierType 
               Height          =   270
               Left            =   1530
               TabIndex        =   96
               Top             =   495
               Width           =   1890
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "3334;467"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtSPSupplier 
               Height          =   270
               Left            =   855
               TabIndex        =   98
               Top             =   855
               Width           =   2565
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "4524;467"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtClientIDPurPay 
               Height          =   270
               Left            =   855
               TabIndex        =   94
               Top             =   135
               Width           =   2565
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "4524;467"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Account Category:"
               Height          =   255
               Index           =   12
               Left            =   180
               TabIndex        =   314
               Top             =   495
               Width           =   1455
            End
            Begin MSForms.Label lblPayPostingDate 
               Height          =   285
               Left            =   11370
               TabIndex        =   308
               Top             =   900
               Width           =   210
               ForeColor       =   8421504
               BackColor       =   16761024
               Caption         =   " P"
               Size            =   "370;503"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Client"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   180
               TabIndex        =   306
               Top             =   150
               Width           =   435
            End
            Begin MSForms.CommandButton cmdSPSupplier 
               Height          =   285
               Left            =   3435
               TabIndex        =   99
               Top             =   855
               Width           =   315
               Caption         =   "; ;"
               Size            =   "556;503"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/C Balance"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   9120
               TabIndex        =   168
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank A/C"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   4140
               TabIndex        =   134
               Top             =   180
               Width           =   630
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amt  £"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   9120
               TabIndex        =   120
               Top             =   540
               Width           =   825
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Date"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   9120
               TabIndex        =   119
               Top             =   900
               Width           =   975
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   4125
               TabIndex        =   118
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Type"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   29
               Left            =   4125
               TabIndex        =   117
               Top             =   555
               Width           =   975
            End
            Begin MSForms.CommandButton cmdSPAmtType 
               Height          =   285
               Left            =   8685
               TabIndex        =   105
               Top             =   540
               Width           =   315
               Caption         =   "; ;"
               Size            =   "556;503"
               FontName        =   "Myriad Web"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   116
               Top             =   870
               Width           =   600
            End
            Begin MSForms.TextBox txtPayAmtType 
               Height          =   285
               Left            =   5220
               TabIndex        =   103
               Top             =   540
               Width           =   3105
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "5477;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.TextBox txtSPayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   13140
            MaxLength       =   11
            TabIndex        =   110
            Top             =   3375
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
            Height          =   2610
            Left            =   135
            TabIndex        =   109
            Top             =   2355
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   4604
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
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
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCrPoA 
            Height          =   2865
            Left            =   135
            TabIndex        =   128
            Top             =   5895
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   5054
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
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
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   720
            X2              =   16605
            Y1              =   5355
            Y2              =   5355
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   675
            X2              =   16560
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Label lblAllocating 
            BackStyle       =   0  'Transparent
            Caption         =   "Allocating..."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   210
            Index           =   1
            Left            =   14310
            TabIndex        =   133
            Top             =   9225
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total allocation difference:"
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
            Left            =   10665
            TabIndex        =   132
            Top             =   8910
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "credit row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   1
            Left            =   15795
            TabIndex        =   124
            Top             =   5040
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "allocating row no"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   7
            Left            =   15615
            TabIndex        =   122
            Top             =   2250
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "  Allocation View  "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Index           =   1
            Left            =   6930
            TabIndex        =   125
            Top             =   4995
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   126
            Top             =   5250
            Width           =   465
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   121
            Top             =   1905
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   85
            Top             =   2115
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1110
            TabIndex        =   84
            Top             =   2115
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PROP ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2775
            TabIndex        =   83
            Top             =   2115
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3735
            TabIndex        =   82
            Top             =   2115
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   4935
            TabIndex        =   81
            Top             =   2115
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   6495
            TabIndex        =   80
            Top             =   2115
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   11535
            TabIndex        =   79
            Top             =   2115
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O/S Amt. £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   12735
            TabIndex        =   78
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   13860
            TabIndex        =   77
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment £"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   15795
            TabIndex        =   76
            Top             =   1575
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   15750
            TabIndex        =   75
            Top             =   1260
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label20 
            Height          =   240
            Index           =   50
            Left            =   120
            TabIndex        =   177
            Top             =   2115
            Width           =   15300
         End
      End
   End
   Begin VB.Label lblBankOD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   17040
      TabIndex        =   311
      Top             =   810
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblBankOD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   17040
      TabIndex        =   310
      Top             =   450
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBankOD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   17040
      TabIndex        =   309
      Top             =   90
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSForms.TextBox txtRecoverable 
      Height          =   255
      Index           =   1
      Left            =   17820
      TabIndex        =   186
      Top             =   225
      Width           =   2775
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4895;450"
      Value           =   "text box for tab control"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proj.:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   17040
      TabIndex        =   11
      Top             =   2385
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CC:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   17040
      TabIndex        =   10
      Top             =   2025
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CC:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   17040
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proj.:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   17040
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "frmPurchaseExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vatOptionEnabled As Boolean
Dim bChangesMade  As Boolean            'This variable is uesed as flag that user has made any changes
Public lFund      As Long
Public sEditPPR   As Single             'This variable keep track if flxSpayment grid has clicked or flxSCrPoA grid has been clicked

Private nTaxCode  As Double             'Tax Rate for Invoice
Private sVCFound  As Single             'Vat code found either in Supplier (1) or Global Data (2)
Private iSelected    As Integer
Private iPIEdit      As Integer
Private sTextBox     As String
Private sSearchSwitch     As String
Private iCurEditRow  As Integer
Private bEditMode    As Boolean        'Is the PI in Edit mode?
Private bFormLoaded  As Boolean
Private lLastID      As Long, dDrOS As Double, dCrOS As Double
Private BANK_TYPE    As String
Private Const iXflxPI = 18

Dim iCurRow             As Integer, dDrOS_Adj As Double, dCrOS_Adj As Double
Dim Rst1                As New ADODB.Recordset
Dim bAdjustment         As Boolean, sAddChoice As String
Dim iGridIdentity       As Integer, NC As String
Dim szUndoText          As String, dSortIndex() As Integer
Private bTotalPayTyped  As Boolean
Dim cGridSPTotal        As Currency
Dim baChangesMade()     As Boolean, iCrPoARowSel As Integer
Dim cTempReceiptAmt     As Currency, cTempPaymentAmt As Currency
Dim iSupplier           As Integer
'Dim bClicked            As Boolean
Dim szaSupplierBal()    As String      'Supplier   balance
Dim szaLandlordBal()    As String      'Landlord   balance
Dim szaClientBal()      As String      'Client     balance
Dim szaAgentBal()       As String      'Agent      balance
Dim szaSuppBalbyClient() As String     'Supplier   balance by client
Dim iDayTerms           As String
Dim bPayAll             As Boolean
Dim szaProperty()       As String
Dim baBankRecon()          As Boolean
Dim iFlxSPayCol       As Integer
Private cTotalSI As Currency, cTotalAdjI As Currency
Dim szTran2Fix As String
Dim searchResultOn As Boolean 'this variable shall be true when you are  showing some filtered  value in the grid. if false then you are showing full records without any filter
Dim lastVat_ID_fromClient As Integer
Dim lastVat_Code_fromClient As String
Dim lastVat_Rate_fromClient As Double
Dim iSelectedFundCategoryID As Integer
Dim areYouProcessingRentPayable As Boolean
Dim dblclientPaymentAmount As Double
Dim CSID As String
Private Enum ComponentMode
   DefaultMode = 0
   newLine = 1
   EditLine = -1
   GridLostFocus = -2
   GridRowOnSelection = 2
   SavedMode = 3
   RefundMode = -3
   ExpensesMode = 4
End Enum
Private Type TJobAmount
     JobID As String
     amount As Double
End Type
Dim VTJobAmount() As TJobAmount
Dim szNominal() As String
Public editexception As Boolean 'this variable will disable the validation in case you shift+ click
Dim frmLockingDialogisActive As Boolean
Dim UserSessionID As String
Dim colTransactionIDOtherPayGrid As String
Dim colTransactionIDOtherPIGrid As String
Dim bSearchClientNameFocus As Boolean
Dim strSupplierTypeOnSelection As String
Dim strSupplierTypeOnSelectionOnPI As String
'Dim PurchaseInvoice As New PuchaseInvoiceCol
Public Sub RedeclareArray()
   ReDim Preserve baChangesMade(flxSPayment.Rows) As Boolean
End Sub

Private Sub cboBC_Change()
'Comment out by anol issue 571.
'Date 25 Aug 2015 Note 1148

'   If IsNull(cboBC) Then Exit Sub
'
'   Dim iRow As Integer
'
'   For iRow = 1 To flxSPayment.Rows - 1
'      If flxSPayment.TextMatrix(iRow, 23) <> cboBC.Column(2) And _
'            flxSPayment.TextMatrix(iRow, 23) <> "" Then
'         flxSPayment.RowHeight(iRow) = 0
'         If flxSPayment.TextMatrix(iRow, 0) <> "" Then
'            iRow = iRow + 1
'            Do While flxSPayment.TextMatrix(iRow, 0) = "-"
'               flxSPayment.RowHeight(iRow) = 0
'               iRow = iRow + 1
'               If iRow = flxSPayment.Rows Then Exit Do
'            Loop
'            iRow = iRow - 1
'         End If
'      Else
'         flxSPayment.RowHeight(iRow) = 240
'         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
'            iRow = iRow + 1
'            Do While flxSPayment.TextMatrix(iRow, 0) = "-"
'               flxSPayment.RowHeight(iRow) = 0
'               iRow = iRow + 1
'               If iRow = flxSPayment.Rows Then Exit Do
'            Loop
'            iRow = iRow - 1
'         End If
'      End If
'   Next iRow
End Sub

'Private Sub cboBC_LostFocus()
'       'issue 571 Validation
'   'Added by anol 20 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
''   If cboBC.ListIndex = -1 Then
''        ShowMsgInTaskBar "Please select a valid bank account"
''        tabPurExp.Tab = 1
''        'if i put a setfocus it fallin a infinite loop with the next control's lost focus
'''        If frmMMain.rtxtMessageDisplay.text = "" Then
'''            cboBC.SetFocus
'''        End If
''        Exit Sub
''   End If
'   Dim strTemp As String
'   Dim adoConn As New ADODB.Connection
'  strTemp = Replace(cboBC.text, "'", "''")
'   If Trim(cboBC.text) <> "" Then
'        adoConn.Open getConnectionString
''        szSQL = "SELECT BANK_AC_Name " & _
''                "FROM tlbClientBanks " & _
''                "where BANK_AC_Name='" & cboBC.text & "';"
'    szSQL = "SELECT Name " & _
'                "FROM NominalLedger " & _
'                "where Name='" & strTemp & "';"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'            tabPurExp.Tab = 1
'            MsgBox "Please select a valid bank account", vbInformation, "Select a bank account"
'            cboBC.text = ""
'            cboBC.SetFocus
'        End If
'        adoRst.Close
'        adoConn.Close
'        Set adoConn = Nothing
'    End If
'End Sub

'Private Sub cboClient_Change()
''   On Error GoTo ERR
''   If IsNull(cboClient.Value) Then Exit Sub
''
''   Dim iRow As Integer
''
''   If cboClient.Value = "ALL" Then
''      For iRow = 1 To flxSPayment.Rows - 1
''         flxSPayment.RowHeight(iRow) = 240
''         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
''            iRow = iRow + 1
''            While flxSPayment.TextMatrix(iRow, 0) = "-"
''               flxSPayment.RowHeight(iRow) = 0
''               iRow = iRow + 1
''            Wend
''            iRow = iRow - 1
''         End If
''      Next iRow
''
''      Exit Sub
''   End If
''
''   Dim adoConn As New ADODB.Connection
''
''   adoConn.Open getConnectionString
''   cboBC.Clear
''   PrepareCboBC adoConn
''   adoConn.Close
''   Set adoConn = Nothing
''   'added by anol 20 July 2015 issue 571
''   'The details displayed in purchase payment for 1/ Bank Account 2/ Payment Type, 3/ Reference and 4/ Total Amt £ should be cleared
'''every time a user changes client.
''      cmbSPAmtType.ListIndex = -1
''      txtSPReference.text = ""
''      txtSPaymentTotal.text = "0.00"
''   'added by anol 09 July 2015
''   'issue 571
''   '11. Purchase payment debit is not filtering by client. It is not clearing.
''   For iRow = 1 To flxSCrPoA.Rows - 1
''      If flxSCrPoA.TextMatrix(iRow, 18) <> cboClient.Value And _
''            flxSCrPoA.TextMatrix(iRow, 18) <> "" Then
''         flxSCrPoA.RowHeight(iRow) = 0
'''         If flxSPayment.TextMatrix(iRow, 0) <> "" Then
'''            iRow = iRow + 1
'''            While flxSPayment.TextMatrix(iRow, 0) = "-"
'''               flxSPayment.RowHeight(iRow) = 0
'''               iRow = iRow + 1
'''            Wend
'''            iRow = iRow - 1
'''         End If
''      Else
''         flxSCrPoA.RowHeight(iRow) = 240
'''         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
'''            iRow = iRow + 1
'''            While flxSPayment.TextMatrix(iRow, 0) = "-"
'''               flxSPayment.RowHeight(iRow) = 0
'''               iRow = iRow + 1
'''            Wend
'''            iRow = iRow - 1
'''         End If
''      End If
''   Next iRow
'''Error Occurs here 02 Dec 2014
'''Issue No 508
'''outstanding- anol
''   For iRow = 1 To flxSPayment.Rows - 1
''      If flxSPayment.TextMatrix(iRow, 23) <> cboClient.Value And _
''            flxSPayment.TextMatrix(iRow, 23) <> "" Then
''         flxSPayment.RowHeight(iRow) = 0
''         If flxSPayment.TextMatrix(iRow, 0) <> "" Then
''            iRow = iRow + 1
''            While flxSPayment.TextMatrix(iRow, 0) = "-"
''               flxSPayment.RowHeight(iRow) = 0
''               iRow = iRow + 1
''            Wend
''            iRow = iRow - 1
''         End If
''      Else
''         flxSPayment.RowHeight(iRow) = 240
''         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
''            iRow = iRow + 1
''            While flxSPayment.TextMatrix(iRow, 0) = "-"
''               flxSPayment.RowHeight(iRow) = 0
''               iRow = iRow + 1
''            Wend
''            iRow = iRow - 1
''         End If
''      End If
''   Next iRow
''
''   Exit Sub
''ERR:
'End Sub

'Private Sub cboClient_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'    cmdACType.SetFocus
'       ' cmdAccounts.SetFocus
'    End If
'End Sub

'Private Sub cboClient_LostFocus()
'    'issue 571 Validation
'   'Added by anol 21 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim adoConn    As New ADODB.Connection
'   If Trim(cboClient.text) <> "" Then
'        adoConn.Open getConnectionString
'        szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'                "FROM CLIENT " & _
'                "where CLIENTNAME='" & cboClient.text & "';"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'            'It shall give message if client is wrong
'            'and cleared down the property
'            tabPurExp.Tab = 1
'            MsgBox "You must select the client", vbInformation, "Select a client"
'            cboClient.SetFocus
'
'        Else
'             'if client is not wrong it shall load proper property
'            cboClient_Change
'        End If
'        adoConn.Close
'        Set adoConn = Nothing
'    Else
'       ' txtClientID.text = ""
'        ' cmbClient.ListIndex = -1
'
'       ' txtPropID.text = ""
'    End If
'End Sub



Private Sub chkAllPurchaseHistory_Click()
    Dim iRow As Integer, i As Integer

   If chkAllPurchaseHistory.Value Then
      For i = 1 To flxPurchHistory.Rows - 1
         flxPurchHistory.TextMatrix(i, 1) = ""
      Next i
      For i = 1 To flxPurchHistory.Rows - 2
         flxPurchHistory.TextMatrix(i, 1) = "X"
      Next i

      flxPurchHistory.row = flxPurchHistory.Rows - 1

      flxPurchHistory_Click
   Else
      For i = 1 To flxPurchHistory.Rows - 1
         flxPurchHistory.TextMatrix(i, 1) = ""
      Next i

      ConfigFlxSplit flxPurchHistorySplit, 29
   End If
End Sub

Private Sub chkIsMgtFee_Click()
    cmdSavePI.Enabled = True
End Sub

Private Sub chkProperty_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If chkProperty.Value = 0 Then
        txtPropID.text = "ALL"
        cmdOpProperty.Enabled = True
        LoadFlxPurchase adoConn
        
    Else
        txtPropID.text = ""
        cmdOpProperty.Enabled = False
        'SortTheGrid flxPurchase, txtIDClient, txtPropID, txtSupplier
        LoadFlxPurchase adoConn
    End If
    fmeLoading.Visible = False
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub chkPropertyHist_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If chkPropertyHist.Value = 0 Then
        'chkPropertyHist.Caption = "Excl"
        txtPropertyIDHist.text = "ALL"
        cmdOpPropertyHist.Enabled = True
        Call LoadFlxPurchHistory(adoConn, "")
    Else
        'chkPropertyHist.Caption = "Excl"
        txtPropertyIDHist.text = ""
        cmdOpPropertyHist.Enabled = False
         Call LoadFlxPurchHistory(adoConn, "")
    End If
    fmeLoading.Visible = False
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub chkShowBal_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    
    If chkShowBal.Value = 1 Then
            If frmMMain.frmPI_SupplierBalance_isUptoDate = False Then
                SupplierAccountBalance adoConn
                frmMMain.frmPI_SupplierBalance_isUptoDate = True
            End If
            adoConn.Execute "Update ShoppingCentre set PIShowBal=true"
            UpdateBalance
     Else
            adoConn.Execute "Update ShoppingCentre set PIShowBal=false"
            UpdateBalanceWithZero
     End If
     adoConn.Close
     Set adoConn = Nothing
End Sub

Private Sub cmdAdvanceProgr_Click()

    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update PayTransactions T,tlbPayment P set T.ClientID=P.ClientID,PPOrPAOrPC='PP'&P.slnumber,PaymentSAGEAC=P.SageAccountNumber  where FromTran=P.TransactionID and P.Type in(8) "
    adoConn.Execute "Update PayTransactions T,tlbPayment P set T.ClientID=P.ClientID,PPOrPAOrPC='PA'&P.slnumber,PaymentSAGEAC=P.SageAccountNumber where FromTran=P.TransactionID and P.Type in(9) "
    adoConn.Execute "Update PayTransactions T,tlbPayment P set T.ClientID=P.ClientID,PPOrPAOrPC='PPR'&P.slnumber,PaymentSAGEAC=P.SageAccountNumber where FromTran=P.TransactionID and P.Type in(24) "
    adoConn.Execute "Update PayTransactions T,tlbPayment P set T.ClientID=P.ClientID,PIOrPPR='PI'&P.slnumber,PaymentSAGEAC=P.SageAccountNumber where ToTran=P.TransactionID and P.Type in(6) "
    adoConn.Execute "Update PayTransactions T,tlbPayment P set T.ClientID=P.ClientID,PIOrPPR='Pc'&P.slnumber,PaymentSAGEAC=P.SageAccountNumber where ToTran=P.TransactionID and P.Type in(7) "
    adoConn.Close
    MsgBox "ClientID updated in allocation table"


End Sub

Private Sub cmdAmountType_Click()
    On Error GoTo Err
    sTextBox = "3"
    tabPayment.Enabled = False
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    loadPaymentType "RECEIPT AMOUNT TYPE", adoConn
    adoConn.Close
    Set adoConn = Nothing
    
    picSupplierType.Left = txtPayAmtType.Left
    picSupplierType.Top = txtPayAmtType.Top
    picSupplierType.Visible = True
    picSupplierType.ZOrder 0
    FocusControl flxSupplierType
    If flxSupplierType.Rows > 1 And flxSupplierType.row = 0 Then
        flxSupplierType.row = 1
    End If
    tabPayment.Enabled = False
    tabPurExp.Enabled = False
    Exit Sub
Err:
End Sub
Private Sub loadPaymentType(szValue As String, adoConn As ADODB.Connection)
    Dim SQLStr1 As String, i As Integer
    Dim adoRST As New ADODB.Recordset
    Label1(0).Caption = "Payment Type"
    flxSupplierType.Clear
    flxSupplierType.Rows = 2
    flxSupplierType.Cols = 3
    flxSupplierType.RowHeight(0) = 0
    flxSupplierType.ColWidth(0) = 120
    flxSupplierType.ColWidth(1) = 900
    flxSupplierType.ColWidth(2) = 1200
  
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
    flxSupplierType.Rows = adoRST.RecordCount + 1
    i = 1
    While Not adoRST.EOF
        flxSupplierType.TextMatrix(i, 1) = adoRST!c
        flxSupplierType.TextMatrix(i, 2) = adoRST!V
        adoRST.MoveNext
        i = i + 1
    Wend
    adoRST.Close
    Set adoRST = Nothing

End Sub
'Private Sub cmdACType_LostFocus()
'     If cmdACType.ListIndex = -1 Then
'        ShowMsgInTaskBar "Please select a valid supplier type"
'        cmdACType.SetFocus
'     End If
'End Sub





Private Sub cmdBankAc_Click()
    sTextBox = "BankAcPay"
    txtSearchClientID.text = ""
    txtSearchClientName.text = ""
    picClient.Left = 4545
    picClient.Top = 1350
    Call LoadflxBankAC("")
    tabPayment.Enabled = False
    tabPurExp.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Function LoadflxBankAC(Filter As String)
   On Error GoTo Error_Handler
   Dim adoConn As New ADODB.Connection
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   
   'Configure grid start
   lblClientID.Caption = "Bank Code"
   lblClientName.Caption = "Bank Account"
   
   'txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   txtSearchClientID.Width = 900
   
   txtSearchClientName.Visible = True
   'txtSearchClientName.text = ""
   txtSearchClientName.Left = 1200
   txtSearchClientName.Width = 2200
   flxClient.Clear
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 200
   flxClient.ColWidth(1) = 900
   flxClient.ColWidth(2) = 2800
   picClient.Width = 3500
   cmdPicCLose.Left = 3200
  
   Label20(14).Visible = False
   flxClient.Rows = 2
   flxClient.Height = 2845
   flxClient.Width = 3400
   picClient.Height = 3595
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   Label20(12).Caption = "Bank Code"
   Label20(13).Caption = "Bank AC"
   lblClientName.Left = 1200
   Label20(12).Width = 1400
   Label20(12).Left = 250
   Label20(13).Width = 3600
   Label20(13).Left = 1300 'Label20(12).Left + flxClient.ColWidth(0)
   'Configure grid End
   If txtClientIDPurPay.text = "ALL" Then
        txtBankCode.text = ""
        txtBankAc.text = ""
        Exit Function
   End If
   adoConn.Open getConnectionString
   If txtClientIDPurPay.text = "ALL" Then
      szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                  "NominalLedger.Name AS BNN, tlbClientBanks.CurrentBalance AS BAL, AllowOverDraft, OverDraftLimit,DEFAULT_AC  " & _
              "FROM tlbClientBanks, NominalLedger " & _
              "WHERE tlbClientBanks.NominalCode = NominalLedger.Code " & _
              "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance, AllowOverDraft, OverDraftLimit,DEFAULT_AC;"
   Else
      szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                  "NominalLedger.Name AS BNN, tlbClientBanks.CurrentBalance AS BAL, AllowOverDraft, OverDraftLimit,DEFAULT_AC  " & _
              "FROM tlbClientBanks, NominalLedger " & _
              "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
                  "tlbClientBanks.CLIENT_ID = '" & txtClientIDPurPay.text & "' AND " & _
                  "NominalLedger.ClientID = '" & txtClientIDPurPay.text & "' " & _
              "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance, AllowOverDraft, OverDraftLimit,DEFAULT_AC;"
   End If
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If adoRST.EOF Then
      MsgBox "Please setup bank account for the client.", vbInformation, "Warning"
   Else
      flxClient.Rows = adoRST.RecordCount + 1
      iRec = 1
      If Filter <> "" Then
            adoRST.Filter = Filter
      End If
      While Not adoRST.EOF
        flxClient.TextMatrix(iRec, 0) = "" 'For Selection Row
        flxClient.TextMatrix(iRec, 1) = adoRST.Fields.Item("BNC").Value
        flxClient.TextMatrix(iRec, 2) = adoRST.Fields.Item("BNN").Value
        Debug.Print flxClient.RowHeight(iRec)
        'Debug.Print adoRst.Fields.Item("BNN").Value
        iRec = iRec + 1
        adoRST.MoveNext
      Wend
   End If
   adoConn.Close
   Set adoConn = Nothing
   Exit Function

   
Error_Handler:
   Set adoRST = Nothing
   MsgBox Err.description
   
   
End Function

Private Sub cmdClearSearch_Click()
    txtSearchNo.text = ""
    txtSearchRef.text = ""
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If tabPurExp.Tab = 0 Then
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
''            LoadFlxPurchase adoConn
'            fmeLoading.Visible = False
'            cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
'           ' Call LoadFlxPurchaseFilter(adoConn, 3)
''            cmdSearch.Caption = "Clear Sea&rch"
'            fmeLoading.Visible = False
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
''            cmdSearch.Caption = "Clear Sea&rch"
''            Call LoadFlxPurchaseFilter(adoConn, 4)
'            fmeLoading.Visible = False
'        End If
    ElseIf tabPurExp.Tab = 2 Then
         Call LoadFlxPurchHistory(adoConn, "")
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
'             Call LoadFlxPurchHistory(adoconn, "")
'             cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
''             Call LoadFlxPurchHistory(adoConn, "3")
''             cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
''             cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
''             Call LoadFlxPurchHistory(adoConn, "4")
'        End If
    ElseIf tabPurExp.Tab = 3 Then
        Call LoadFlxPurchPPHistory(adoConn, "")
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
'             Call LoadFlxPurchPPHistory(adoconn, "")
'             cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
'
'        End If
    End If
    adoConn.Close
    Set adoConn = Nothing
    
End Sub

Private Sub cmdClosePrintHIst_Click()
    picPurchaseHistory.Visible = False
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdFix_Click()
    Dim adoConn As New ADODB.Connection
    If MsgBox("Do you want to Run the fix?", vbYesNo, "Confirm?") = vbYes Then
        adoConn.Open getConnectionString
        adoConn.Execute "Update tlbPayment P, tblPurinv V set P.slnumber=V.slnumber where P.PI=V.MY_ID and P.Type=V.TransactionTYpe and P.slnumber<>V.slnumber"
        adoConn.Close
        MsgBox "Data has been Updated."
    End If
End Sub

Private Sub cmdfixNL_Click()
    Dim Conn1 As New ADODB.Connection
    If MsgBox("are you sure you want to fix NL Posting for MA Control Accounts?", vbYesNo, "Please confirm") = vbYes Then
        Conn1.Open getConnectionString
        Conn1.Execute "Update NLPosting N,Supplier S,NominalLedger L SET N.NOMINAL_CODE=L.Code where N.ACCOUNT_NUMBER=S.SupplierID " & _
                      "AND L.ClientID=N.ClientID AND NOMINAL_CODE='1300' AND S.Type='Agent' AND L.CAName='Managing Agents control Account (B/S)';"
        Conn1.Execute "Update NLPosting N,Supplier S,NominalLedger L SET N.NOMINAL_CODE=L.Code where N.ACCOUNT_NUMBER=S.SupplierID " & _
                      "AND L.ClientID=N.ClientID AND NOMINAL_CODE='1300' AND (S.Type='Client' OR S.Type='LLORD') AND L.CAName='Client/Landlord Control Account (B/S)';"
        MsgBox "Update done"
        Conn1.Close
    End If
End Sub

Private Sub cmdOkSearch_Click()
        fraSearch.Visible = False
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If tabPurExp.Tab = 0 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                    Call LoadFlxPurchaseFilter(adoConn, 3)
                    searchResultOn = True
                    'cmdSearch.Caption = "Clear Sea&rch"
                    fmeLoading.Visible = False
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                    cmdSearch.Caption = "Clear Sea&rch"
                    Call LoadFlxPurchaseFilter(adoConn, 4)
                    searchResultOn = True
                    fmeLoading.Visible = False
                End If
            ElseIf tabPurExp.Tab = 2 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchHistory(adoConn, "3")
                     searchResultOn = True
                     'cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     'cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchHistory(adoConn, "4")
                     searchResultOn = True
                End If
            ElseIf tabPurExp.Tab = 3 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchPPHistory(adoConn, "3")
                     searchResultOn = True
                     'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchPPHistory(adoConn, "4")
                     searchResultOn = True
                End If
            End If
            adoConn.Close
       FocusControl cmdSearchOK
End Sub

Private Sub cmdOpProperty_Click()
   chkShowBal.Visible = False
   sTextBox = "PROPERTYFILTER"
   LoadPropertyList ""

   tabPurExp.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4855
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 60
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 80
'   fraList.Left = txtProperty.Left + fraLay(0).Left + 100
   fraList.Left = txtPropID.Left - 400
   fraList.Top = 800
   fraList.Visible = True
   fraList.ZOrder 0
   
   
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
End Sub



Private Sub cmdOpPropertyHist_Click()
   sTextBox = "PROPERTYHIST"
   chkShowBal.Visible = False
   LoadPropertyList ""

   tabPurExp.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4815
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
'   fraList.Left = txtProperty.Left + fraLay(0).Left + 100
   fraList.Left = txtPropID.Left - 100
   fraList.Top = 800
   fraList.Visible = True
   fraList.ZOrder 0
   cmdGridUnitLookup(0).ZOrder 0
   cmdGridUnitLookup(0).Left = 4550
  
   
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
End Sub

Private Sub cmdOpSupSearch_Click()
'written by anol 12 Dec 2015
   chkShowBal.Visible = False
   sTextBox = "PAYHIST"
   LoadSupplierAccount ""
   
   
   tabPurExp.Enabled = False
  
   fraList.Left = 5755
   fraList.Top = 740
   fraList.Width = 5350
   fraList.Visible = True
   fraList.ZOrder 0
   txtSearch1.Visible = True
   txtSearch2.Visible = True
   txtSearch1.text = ""
   txtSearch2.text = ""
   txtSearch1.SetFocus
End Sub

Private Sub cmdPrintallocationhistory_Click()
   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   
   
   
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PaymentAllocation.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
 

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintHistCancel_Click()
    picPurchaseHistory.Visible = False
End Sub

Private Sub cmdPrintHistOK_Click()
   On Error GoTo Err
   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   
   If Trim(txtStartDate.text) = "" Then
        MsgBox "Please enter start Date", vbCritical, "Warning"
        FocusControl txtStartDate
        Exit Sub
   End If
    If Trim(txtEndDate.text) = "" Then
        MsgBox "Please enter end date Date", vbCritical, "Warning"
        FocusControl txtEndDate
        Exit Sub
   End If
   If IsDate(txtStartDate.text) Then
        If DateDiff("d", txtStartDate.text, Date) < 0 Then
            MsgBox "Start Date cannot be greater than current date", vbCritical, "Warning"
            FocusControl txtStartDate
            Exit Sub
        End If
   Else
        MsgBox "Please enter a valid start date", vbCritical, "Warning"
        FocusControl txtStartDate
        Exit Sub
   End If
    If IsDate(txtEndDate.text) Then
        If DateDiff("d", txtEndDate.text, Date) < 0 Then
            MsgBox "End Date cannot be greater than current date", vbCritical, "Warning"
            FocusControl txtEndDate
            Exit Sub
        End If
   Else
        MsgBox "Please enter a valid end date", vbCritical, "Warning"
         FocusControl txtEndDate
        Exit Sub
   End If
   cmdPrintHistOK.Enabled = False
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tblPurInv SET Prn = 'N';"

'   szMY_ID = ""
'   For i = 1 To flxPurchHistory.Rows - 2
'      If flxPurchHistory.RowHeight(i) > 0 Then _
'         szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "', "
'   Next i
'   szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "'"
'
'   adoconn.Execute "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ");"

'   Written by anol error was coming query too complex. Need to split query for load balancing .issue 743
 If tabPurExp.Tab = 2 Then
            szMY_ID = ""
            Dim K As Integer
            Dim j As Integer
            j = flxPurchHistory.Rows
            K = CInt(j / 50)
           
           If K = j / 50 Then
                'No no need to do ceiling, this is fully divisible
                K = j / 50
           Else
                K = CInt(j / 50) + 1 'This is ceiling function
           End If
           For K = 0 To K - 1
                szMY_ID = ReturnStringPI_ID(K * 50, (K + 1) * 50 - 1)
                If Trim(szMY_ID) <> "" Then
                    adoConn.Execute "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ") and " & _
                        "tran_date<=#" & Format(txtEndDate.text, "dd mmm yyyy") & "# and tran_date>=#" & Format(txtStartDate.text, "dd mmm yyyy") & "#;"
        '            Debug.Print "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ") and tran_date<=#" & txtEndDate.text & "# and tran_date>=#" & txtStartDate.text & "#;"
                 End If
           Next
          ' adoconn.Execute "UPDATE tblPurInv SET Prn = 'N' where Prn = 'Y' and tran_date>=#" & txtEndDate.text & "# and tran_date<=#" & txtStartDate.text & "#;"
           'Select * from  tblPurInv where tran_date>=#11/10/2018# AND tran_date<=#11/10/2018#
           adoConn.Close
           Set adoConn = Nothing
        
           'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PI_List1.rpt")
           '2019-10-24 New modification come so I decided to implement new report issue 803
           Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PI_transactionHist.rpt")
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
           Report.EnableParameterPrompting = False
           Report.DiscardSavedData
        
           Report.ParameterFields(1).AddCurrentValue "N"
           Report.ParameterFields(2).AddCurrentValue CDate(Format(txtStartDate.text, "dd mmmm yyyy")) '"20181212" 'txtStartDate.text
           Report.ParameterFields(3).AddCurrentValue CDate(Format(txtEndDate.text, "dd mmmm yyyy"))  'txtEndDate.text
           Report.ParameterFields(4).AddCurrentValue IIf(txtPropertyIDHist.Tag = "", "ALL", txtPropertyIDHist.Tag)
           Report.ParameterFields(5).AddCurrentValue IIf(txtClientIdlist.Tag = "", "ALL", txtClientIdlist.Tag)
           Report.ParameterFields(6).AddCurrentValue "Purchase Transaction History Report"
           Report.ParameterFields(7).AddCurrentValue txtSupplierSearc.text
           cmdPrintHistOK.Enabled = True
   End If
   If tabPurExp.Tab = 0 Then
           cmdPrintPI_List.Enabled = False
           cmdPrintPI_List.Enabled = False
           szMY_ID = ""
           For i = 1 To flxPurchase.Rows - 2
                   If flxPurchase.RowHeight(i) > 0 Then _
                         szMY_ID = szMY_ID + "'" + flxPurchase.TextMatrix(i, 0) + "', "
           Next i
           szMY_ID = szMY_ID + "'" + flxPurchase.TextMatrix(i, 0) + "'"
           If Trim(szMY_ID) <> "" Then
                szMY_ID = "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ")  and " & _
                        "tran_date<=#" & Format(txtEndDate.text, "dd mmm yyyy") & "# and tran_date>=#" & Format(txtStartDate.text, "dd mmm yyyy") & "#;"
           End If
           adoConn.Execute szMY_ID
           adoConn.Close
           Set adoConn = Nothing
        
           Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PI_List1.rpt")
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
           Report.EnableParameterPrompting = False
           Report.DiscardSavedData
        
            Report.ParameterFields(1).AddCurrentValue "N"
            Report.ParameterFields(2).AddCurrentValue CDate(txtStartDate.text)
            Report.ParameterFields(3).AddCurrentValue CDate(txtEndDate.text)
            Report.ParameterFields(4).AddCurrentValue IIf(txtPropID.Tag = "", "ALL", txtPropID.Tag)
            Report.ParameterFields(5).AddCurrentValue IIf(txtIDClient.Tag = "", "ALL", txtIDClient.Tag)
            Report.ParameterFields(6).AddCurrentValue "Purchase Invoice Listing Report"
            cmdPrintPI_List.Enabled = True
   End If
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   cmdPrintHistOK.Enabled = True
   Exit Sub
Err:
   MsgBox Err.description
   cmdPrintHistOK.Enabled = True
End Sub

Private Sub cmdPrintPaymentList_Click()
    
   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tlbPayment SET isPrint = false ;"


    szMY_ID = ""
    Dim K As Integer
    Dim j As Integer
    j = flxPurchHistory.Rows
    K = CInt(j / 50)
   
   If K = j / 50 Then
        'No no need to do ceiling, this is fully divisible
        K = j / 50
   Else
        K = CInt(j / 50) + 1 'This is ceiling function
   End If
   For K = 0 To K - 1
        szMY_ID = ReturnStringPP_ID(K * 50, (K + 1) * 50 - 1)
        If szMY_ID <> "" Then
            adoConn.Execute "UPDATE tlbPayment SET isPrint = true WHERE transactionID IN (" & szMY_ID & ");"
        End If
        
   Next
   
   
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Payment_List.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue "N"
   Report.ParameterFields(6).AddCurrentValue "Purchase Payment Report"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPurchasePaymentHistory_Click()
    sTextBox = "PaymentHistory"
    chkShowBal.Visible = False
    tabPurExp.Enabled = False
    tabPayment.Enabled = False
    picClient.Left = 405
    picClient.Top = 655
    picClient.Visible = True
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdSavePIRef_Click()
        Dim adoConn As New ADODB.Connection
        Dim rsPI As New ADODB.Recordset
        Dim szSQL As String
        Dim SlNumber As Integer
        Dim slType As Integer
        adoConn.Open getConnectionString
       
        szSQL = "SELECT * FROM tblPurInv WHERE MY_ID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
        rsPI.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
        If Not rsPI.EOF Then
            rsPI.Fields.Item("INV_NO").Value = txtReference.text
            flxPurchase.TextMatrix(iPIEdit, 7) = txtReference.text
            rsPI.Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
            flxPurchase.TextMatrix(iPIEdit, 16) = lblPostingDate.ToolTipText
            rsPI.Fields.Item("PropertyID").Value = txtProperty.text
        End If
        rsPI.Update
        rsPI.Close
        
          adoConn.Execute "Update tblPurInvSRec set TRANS='" & txtProperty.text & "' where ParentID  = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
          adoConn.Execute "Update tlbPayment set UnitID='" & txtProperty.text & "' where PI  = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
        
        
        szSQL = "SELECT * FROM tlbPayment WHERE PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
        rsPI.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
        If Not rsPI.EOF Then
            rsPI!ref = txtReference.text
            SlNumber = rsPI!SlNumber
            slType = rsPI!Type
            rsPI!postingDate = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
        End If
        rsPI.Update
        rsPI.Close
'        szSQL = "SELECT * FROM NLPOSTING WHERE TRANS_ID = '" & slnumber & "' AND TRANSACTION_TYPE=" & slType & " and DeleteFlag=false;"
'        rsPI.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
'        If Not rsPI.EOF Then
'            rsPI!Reference = txtReference.text
'            rsPI!POSTED_DATE = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
'        End If
'        rsPI.Update
'        rsPI.Close
        adoConn.Execute "Update NLPOSTING Set Reference='" & Replace(txtReference.text, "'", "") & "', POSTED_DATE=#" & _
                Format(lblPostingDate.ToolTipText, "dd mmmm yyyy") & "# ,PROPERTY_ID='" & txtProperty.text & "' WHERE TRANS_ID = '" & SlNumber & "' AND TRANSACTION_TYPE=" & slType & " and DeleteFlag=false"
        adoConn.Close
        MsgBox "PI Reference has been updated.", vbInformation, "Saved"
        cmdSavePIRef.Visible = False
End Sub

Private Sub cmdSearch_Click()
        chkProperty.Value = 0
        fraSearch.Left = 1395
        fraSearch.Top = 6210
        Label6(3).Caption = "Supplier"
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        If cmdSearch.Caption = "Clear Sea&rch" Then
             Call cmdSearchCancel_Click
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             fmeLoading.Visible = False
             cmdSearch.Caption = "Sea&rch"
             fraSearch.Visible = False
        Else
           
            If fraSearch.Visible = False Then
                'cmdSearch.Caption = "Clear &Search"
                fraSearch.Visible = True
                txtSearchNo.SetFocus
            Else
               ' cmdSearch.Caption = "Sea&rch"
                fraSearch.Visible = False
            End If
        End If
        adoConn.Close
    
End Sub

Private Sub cmdSearchCancel_Click()

             Dim adoConn As New ADODB.Connection
             adoConn.Open getConnectionString
             If tabPurExp.Tab = 0 Then
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchaseFilter(adoConn, 1)
                     searchResultOn = True
                 Else
                     Call LoadFlxPurchase(adoConn)
                 End If
                 fmeLoading.Visible = False

             ElseIf tabPurExp.Tab = 2 Then
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchHistory(adoConn, "1")
                     searchResultOn = True
                 Else
                     Call LoadFlxPurchHistory(adoConn, "")
                 End If

            ElseIf tabPurExp.Tab = 3 Then
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchPPHistory(adoConn, "1")
                     searchResultOn = True
                 Else
                        If Trim(txtSearchNo.text) <> "" Then
                            'do nothing
                        ElseIf Trim(txtSearchRef.text) <> "" Then
                            'do nothing
                        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                             Call LoadFlxPurchPPHistory(adoConn, "3")
                             searchResultOn = True
                             'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                             'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                             Call LoadFlxPurchPPHistory(adoConn, "4")
                             searchResultOn = True
                        End If
                 End If
             End If
             
             
             
    fmeLoading.Visible = False
   ' fraSearch.Visible = False
    FocusControl txtSearchNo
    adoConn.Close
    searchResultOn = False
End Sub

Private Sub cmdSearchOK_Click()
    'This is search close button
    fraSearch.Visible = False
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    If tabPurExp.Tab = 0 Then
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
''            LoadFlxPurchase adoConn
'            fmeLoading.Visible = False
'            cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
'           ' Call LoadFlxPurchaseFilter(adoConn, 3)
''            cmdSearch.Caption = "Clear Sea&rch"
'            fmeLoading.Visible = False
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
''            cmdSearch.Caption = "Clear Sea&rch"
''            Call LoadFlxPurchaseFilter(adoConn, 4)
'            fmeLoading.Visible = False
'        End If
'    ElseIf tabPurExp.Tab = 2 Then
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
'             Call LoadFlxPurchHistory(adoconn, "")
'             cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
''             Call LoadFlxPurchHistory(adoConn, "3")
''             cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
''             cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
''             Call LoadFlxPurchHistory(adoConn, "4")
'        End If
'    ElseIf tabPurExp.Tab = 3 Then
'        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
'             Call LoadFlxPurchPPHistory(adoconn, "")
'             cmdSearch.Caption = "Sea&rch"
'        ElseIf Trim(txtSearchNo.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchRef.text) <> "" Then
'            'do nothing
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
''             Call LoadFlxPurchPPHistory(adoConn, "3")
''             cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
''             cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
''             Call LoadFlxPurchPPHistory(adoConn, "4")
'        End If
'    End If
'    adoconn.Close
'    Set adoconn = Nothing
End Sub



Private Sub cmdSearchPurchaseHistory_Click()
    Dim adoConn As New ADODB.Connection
    chkPropertyHist.Value = 0
    adoConn.Open getConnectionString
   
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    Label6(3).Caption = "Ref"
    'in fact tab 0 and 3 condition is not usefull here
    If tabPurExp.Tab = 0 Then
        fraSearch.Left = 1395
        fraSearch.Top = 6210
        If cmdSearch.Caption = "Clear Sea&rch" Then
'             Call LoadFlxPurchase(adoConn)
             txtSearchFromD.text = ""
             txtSearchToD.text = ""
             fmeLoading.Visible = False
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
        
    ElseIf tabPurExp.Tab = 2 Then
        fraSearch.Left = 3420
        fraSearch.Top = 5400
        If cmdSearchPurchaseHistory.Caption = "Clear Sea&rch" Then
              'Call LoadFlxPurchHistory(adoConn, "")
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             cmdSearchPurchaseHistory.Caption = "Sea&rch"
             fraSearch.Visible = False
         Else
                If fraSearch.Visible = False Then
                    fraSearch.Visible = True
                    txtSearchNo.SetFocus
                Else
                    fraSearch.Visible = False
                End If
        End If
 
    ElseIf tabPurExp.Tab = 3 Then
        fraSearch.Left = 3420
        fraSearch.Top = 5400
        If cmdSearchPurchPayHistory.Caption = "Clear Sea&rch" Then
              'LoadFlxPurchPPHistory adoConn, ""
              txtSearchFromD.text = ""
              txtSearchToD.text = ""
'             Call LoadFlxPurchase(adoConn)
'             fmeLoading.Visible = False
             cmdSearchPurchPayHistory.Caption = "Sea&rch"
             fraSearch.Visible = False
        Else
                If fraSearch.Visible = False Then
                    fraSearch.Visible = True
                    txtSearchNo.SetFocus
                Else
                    fraSearch.Visible = False
                End If
        End If
    End If
    
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdSearchPurchPayHistory_Click()
    'procedure written by anol 20171113 issue 504
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
'    txtSearchNo.text = ""
'    txtSearchRef.text = ""
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    Label6(3).Caption = "Ref"
    'in fact tab 0 and 2 condition is not usefull here
    If tabPurExp.Tab = 0 Then
        fraSearch.Left = 1395
        fraSearch.Top = 6210
        fmeLoading.Visible = False
        cmdSearch.Caption = "Sea&rch"
        fraSearch.Visible = True
    ElseIf tabPurExp.Tab = 2 Then
        fraSearch.Left = 3420
        fraSearch.Top = 5400
        txtSearchNo.text = ""
        txtSearchRef.text = ""
        fraSearch.Visible = True
    ElseIf tabPurExp.Tab = 3 Then
        fraSearch.Left = 3420
        fraSearch.Top = 5400
        txtSearchNo.text = ""
        txtSearchRef.text = ""
        fraSearch.Visible = True

    End If
    FocusControl txtSearchNo
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdSupplierType_Click()
    sTextBox = "2"
    tabPurExp.Enabled = False
    tabPayment.Enabled = False
   
    Call LoadflxSupplierType
    picSupplierType.Top = txtSupplierType.Top + 1200
    picSupplierType.Left = txtSupplierType.Left + 300
    picSupplierType.Visible = True
    picSupplierType.ZOrder 0
    FocusControl flxSupplierType
    If flxSupplierType.Rows > 1 And flxSupplierType.row = 0 Then
        flxSupplierType.row = 1
    End If
End Sub
Private Sub LoadflxSupplierType()
    Dim iRow As Integer
    Label1(0).Caption = "Account Category"
    flxSupplierType.Clear
    flxSupplierType.Rows = 6
    flxSupplierType.Cols = 3
    flxSupplierType.ColWidth(0) = 60
    flxSupplierType.ColWidth(1) = 1600
    flxSupplierType.ColWidth(2) = 0
    flxSupplierType.RowHeight(0) = 0
    
    iRow = 1
    flxSupplierType.TextMatrix(iRow, 1) = "All Categories"
    flxSupplierType.TextMatrix(iRow, 2) = "ALL"
    
    iRow = 2
    flxSupplierType.TextMatrix(iRow, 1) = "Supplier"
    flxSupplierType.TextMatrix(iRow, 2) = "SUPPLIER"
    iRow = 3
    flxSupplierType.TextMatrix(iRow, 1) = "Client"
    flxSupplierType.TextMatrix(iRow, 2) = "CLIENT"
    iRow = 4
    flxSupplierType.TextMatrix(iRow, 1) = "Managing Agent"
    flxSupplierType.TextMatrix(iRow, 2) = "AGENT"
    iRow = 5
    flxSupplierType.TextMatrix(iRow, 1) = "Landlord"
    flxSupplierType.TextMatrix(iRow, 2) = "LLORD"
    
End Sub
Private Sub cmdSupplierTypeClose_Click()
    picSupplierType.Visible = False
     tabPurExp.Enabled = True
    tabPayment.Enabled = True
    
End Sub

Private Sub Command1_Click()
    Call SendEmail("", "", "", "", "", "", , , , , "")
End Sub

Private Sub flxPurchase_DblClick()
    Call cmdEdit_Click(1)
End Sub

Private Sub flxPurchase_RowColChange()
    Call InstantLockingCheck
'    Dim adoconn As New ADODB.Connection
'    Dim rsLockDialog As New ADODB.Recordset
'    Dim selcol As Integer
'    Dim strSQL As String
'   If flxPurchase.TextMatrix(flxPurchase.row, 26) <> "" Then ' This procedure is only for unlock the record on each cell browsing written by anol 20190412
'        adoconn.Open getConnectionString
'        strSQL = "Select DateTimeStamp ,UserSessionID " & _
'               "from tlbPayment as Pt  where TransactionID=" & flxPurchase.TextMatrix(flxPurchase.row, 20) & ""
'        rsLockDialog.Open strSQL, adoconn, adOpenStatic, adLockReadOnly
'        If Not rsLockDialog.EOF Then
'                If Len(rsLockDialog("UserSessionID").Value) = 0 Then
'                      selcol = flxPurchase.col
'                      flxPurchase.col = 1
'                      flxPurchase.CellBackColor = vbWhite
'                      flxPurchase.col = selcol
'                      'now you need to lock it for this screen
'                       adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Batch Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
'                       SystemUser & "',MachineName='" & WS_Name & "'," & _
'                       "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID =" & flxPurchase.TextMatrix(flxPurchase.row, 20) & ""
'                End If
'        End If
'        rsLockDialog.Close
'        Set rsLockDialog = Nothing
'        adoconn.Close
'        Set adoconn = Nothing
'   End If
End Sub

Private Sub flxPurchPPHistory_DblClick()
        'issue 520
        'written by anol 20180212
        'this function shall edit PPR and PP and PA AND only zero values
        Dim adocon As New ADODB.Connection
        adocon.Open getConnectionString
        If flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 9) > 0 Then Exit Sub

        'Code for Loading PPR,PP AND PA
        'If Left(flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 2), 3) = "PPR" Then
            Load frmPaymentEdit
            With frmPaymentEdit
            'Debug.Print flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 5)
            .Caption = .Caption & " Refund - " & flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 2) 'Invoice ID
            '.cmbSPSupplier.ListIndex = FindComboIndex(.cmbSPSupplier, flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 5), 0)
            .txtSPSupplier.text = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 5)
            .txtDate.text = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 4) 'for type 24 this is Pdate else Ddate
            .lblPostingDate.ToolTipText = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 12)
            .txtAmount.text = Format(flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 9), "0.00")
            .cboBC.ListIndex = FindComboIndex(.cboBC, flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 15), 0) 'BankCode
            .txtReference.text = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 7) 'ref
            .cmbSPAmtType.ListIndex = FindComboIndex(.cmbSPAmtType, flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 13), 0) 'PayAmtType
             .cmbFund.ListIndex = FindComboIndex(.cmbFund, flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 11), 0)
             .txtDetails.text = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 8)
             .TransactionID = frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 0)
             .InvoiceNO = frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 2)
            'added by anol on 25 Aug 2015 issue 571 note 1148
            If flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 14) <> "" Then
            ' .cboClient.ListIndex = FindComboIndex(.cboClient, flxSCrPoA.TextMatrix(flxSCrPoA.row, 18), 0)
               .txtClient.text = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 14)
            End If
            .LoadFlxPaymentSplit adocon
            adocon.Close
            End With
            frmPaymentEdit.Left = 100
            frmPaymentEdit.Top = 100
            frmPaymentEdit.Show
        'End If
        
        
End Sub

Private Sub flxPurchPPHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If flxPurchPPHistory.MouseCol = 2 Then
            flxPurchPPHistory.ToolTipText = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.MouseRow, 4)
        End If
End Sub

Private Sub flxSCrPoA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 1 And cmdEditPayment.Enabled Then
        sEditPPR = 2
        editexception = True
        MsgBox "Please ensure the total amount is unchanged after editing to ensure the Bank Reconciliation is unaffected and remains the same after editing", vbInformation, "Warning!!"
        cmdEditPayment_Click
    End If
End Sub

Private Sub flxSPayment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 And Shift = 1 And cmdEditPayment.Enabled Then
        sEditPPR = 1
        If Val(flxSPayment.TextMatrix(flxSPayment.row, 8)) - Val(flxSPayment.TextMatrix(flxSPayment.row, 9)) <> 0 And Val(flxSPayment.TextMatrix(flxSPayment.row, 8)) <> 0 Then
             MsgBox "Please un allocate all the payment against this invoice.", vbInformation, "Warning!!"
            Exit Sub
        End If
        editexception = True
        MsgBox "Please ensure the total amount is unchanged after editing to ensure the Bank Reconciliation is unaffected and remains the same after editing", vbInformation, "Warning!!"
        cmdEditPayment_Click
    End If
End Sub

Private Sub flxSPayment_RowColChange() 'this procedure is written by anol 20190412 for instant unlocking all rows
    Dim adoConn As New ADODB.Connection
    Dim rsLockDialog As New ADODB.Recordset
    Dim selcol As Integer
    Dim selRow As Integer
    Dim strSQL As String
    Dim colTransactionIDHere As String
    Dim i As Integer
    selcol = flxSPayment.col
    selRow = flxSPayment.row
    If Len(colTransactionIDOtherPayGrid) > 0 Then ' This procedure is only for unlock the record on each cell browsing written by anol 20190412
      'colTransactionIDOtherPayGrid varibale contains the transaction ID that is locked by other screen
        adoConn.Open getConnectionString
        strSQL = "Select DateTimeStamp ,UserSessionID,transactionID " & _
               "from tlbPayment as Pt  where  (UserSessionID='' or isnull(UserSessionID)) AND TransactionID in (" & colTransactionIDOtherPayGrid & ")" 'UserSessionID='' or isnull(UserSessionID)
        rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
        
        While Not rsLockDialog.EOF
                flxSPayment.col = 0
                For i = 1 To flxSPayment.Rows - 1
                    If flxSPayment.TextMatrix(i, 19) = rsLockDialog("transactionID").Value Then
                          flxSPayment.row = i
                          flxSPayment.CellBackColor = vbWhite
                          'now you need to lock it for this screen
                           colTransactionIDHere = colTransactionIDHere & flxSPayment.TextMatrix(i, 19) & ","
                          
                    End If
                 Next i
              rsLockDialog.MoveNext
        Wend
        flxSPayment.col = selcol
        flxSPayment.row = selRow
       
        If Len(colTransactionIDHere) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
            colTransactionIDHere = Left(colTransactionIDHere, Len(colTransactionIDHere) - 1)
        End If
        If Len(colTransactionIDHere) > 0 Then
            'again locking those records for current screen
            adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & _
                           SystemUser & "',MachineName='" & WS_Name & "'," & _
                           "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in  (" & colTransactionIDHere & ")"
        End If
        rsLockDialog.Close
        Set rsLockDialog = Nothing
        adoConn.Close
        Set adoConn = Nothing
   End If
End Sub

Private Sub flxSPayment_Scroll()
    If txtSPayment.Visible Then      'The grid is in edting mode
      txtSPayment.text = szUndoText
      flxSPayment.Enabled = True
      txtSPayment.Visible = False
   End If
   If txtCrPayment.Visible Then
      txtCrPayment.text = Format(szUndoText, "0.00")
      flxSPayment.Enabled = True
      txtCrPayment.Visible = False
   End If
End Sub



Private Sub flxSupplierType_Click()
     tabPurExp.Enabled = True
     tabPayment.Enabled = True
    If sTextBox = "2" Then
        If flxSupplierType.row > 0 Then
            txtSupplierType.text = flxSupplierType.TextMatrix(flxSupplierType.row, 1)
            txtSupplierType.Tag = flxSupplierType.TextMatrix(flxSupplierType.row, 2)
            'here clearing the supplier name and its ID
            txtSPSupplier.text = ""
            txtSPSupplier.Tag = ""
            'now need to clear everything
            txtSPReference.text = ""
            txtSPaymentTotal.text = "0.00"
            txtPaymentTotal.text = "0.00"
            txtPaymentEntered.text = "0.00"
            txtDiffPay.text = "0.00"
            txtAllocatedDiff(1).text = "0.00"
            txtSupAcBal.text = "0.00"
            txtSupAcBal.ForeColor = vbBlack
            ConfigFlxSPayment
            ConfigFlxSCrPoA
            FocusControl cmdSPSupplier
        End If
     Else
         If flxSupplierType.row > 0 Then
                txtPayAmtType.Tag = flxSupplierType.TextMatrix(flxSupplierType.row, 1)
                txtPayAmtType.text = flxSupplierType.TextMatrix(flxSupplierType.row, 2)
                FocusControl txtSPReference
         End If
     End If
     picSupplierType.Visible = False
End Sub

Private Sub flxSupplierType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplierType_Click
    End If
End Sub

Private Sub fraCmds_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MsgBox Me.ActiveControl.Name
End Sub

Private Sub txtAccountSearch_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If Index = 1 And KeyAscii = 13 Then
        flxSupplier(1).SetFocus
    End If
    If Index = 2 And KeyAscii = 13 Then
        flxSupplier(1).SetFocus
    End If
    
End Sub

Private Sub txtBankAc_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdBankAc.SetFocus
    End If

End Sub

Private Sub txtBankCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtBankAc.SetFocus
    End If
    
End Sub

Private Sub txtClientID_Change()
'Added by anol 13 Aug 2015
    If Left(fraLay(1).Caption, 15) = "Transaction ID:" Then
        flxPI.Tag = "EditedOrAdded"
        cmdSavePI.Enabled = True
    End If
    txtProperty.text = ""
   
   'Resolved by BOSL
   'Issue No: 0000467
   'Added By: Asif. 04 Sep 2014
   txtUnit(0).text = ""
End Sub

'Private Sub cboClientPI_Click()
'   txtProperty.text = ""
'
'   'Resolved by BOSL
'   'Issue No: 0000467
'   'Added By: Asif. 04 Sep 2014
'   txtUnit(0).text = ""
'
'End Sub

'Private Sub cboClientPI_GotFocus()
''Added by anol 21 May 2015
'    SelTxtInCtrl cboClientPI
'   If fraList.Visible Then fraList.Visible = False
'End Sub

'Private Sub cboClientPI_LostFocus()
'  'issue 571 Validation
'   'Added by anol 21 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim adoConn    As New ADODB.Connection
'   If Trim(cboClientPI.text) <> "" Then
'        adoConn.Open getConnectionString
'        szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'                "FROM CLIENT " & _
'                "where CLIENTNAME='" & cboClientPI.text & "';"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'            'It shall give message if client is wrong
'            'and cleared down the property
'            tabPurExp.Tab = 0
'            MsgBox "You must select the client", vbInformation, "Select a client"
'            cboClientPI.SetFocus
'            txtProperty.text = ""
'        Else
'             'if client is not wrong it shall load proper property
'            cboClientPI_Click
'        End If
'        adoConn.Close
'        Set adoConn = Nothing
'    Else
'       ' txtClientID.text = ""
'        ' cmbClient.ListIndex = -1
'        txtProperty.text = ""
'       ' txtPropID.text = ""
'    End If
'End Sub
Private Function isValidClient() As Boolean
  'issue 571 Validation
   'Added by anol 21 May 2015
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim adoConn    As New ADODB.Connection
   If Trim(txtClientID.text) <> "" Then
        adoConn.Open getConnectionString
        szSQL = "SELECT CLIENTID, CLIENTNAME " & _
                "FROM CLIENT " & _
                "where CLIENTID='" & txtClientID.text & "';"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If adoRST.EOF Then
            isValidClient = False
        Else
             'if client is correct it shall load proper property
            isValidClient = True
        End If
        adoConn.Close
        Set adoConn = Nothing
    End If
End Function
Private Function isValidProperty() As Boolean
  'issue 571 Validation
   'Added by anol 21 May 2015
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim adoConn    As New ADODB.Connection
   If Trim(txtProperty.text) = "ZZZZ" Then
        txtProperty.text = ""
        isValidProperty = True
        Exit Function
   End If
   If Trim(txtProperty.text) <> "" Then
        adoConn.Open getConnectionString
        szSQL = "SELECT PropertyID, PropertyName " & _
                "FROM Property " & _
                "where PropertyID='" & txtProperty.text & "';"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If adoRST.EOF Then
            isValidProperty = False
        Else
             'if client is correct it shall load proper property
            isValidProperty = True
        End If
        adoConn.Close
        Set adoConn = Nothing
    End If
End Function
Private Sub chkRecover_Click()
   If txtRecoverable(0).text <> "" Then
      txtRecoverable(0).Enabled = True
      Exit Sub
   End If

   If chkRecover.Value = 1 Then
      txtRecoverable(0).text = "100"
      txtRecoverable(0).Enabled = True
      txtRecoverable(0).SetFocus
   Else
      txtRecoverable(0).Enabled = False
      txtRecoverable(0).text = ""
      cmdJobNo(0).SetFocus
   End If
End Sub

Private Sub chkSelectAllDemands_Click()
   Dim iRow As Integer, i As Integer

   If chkSelectAllDemands.Value Then
      For i = 1 To flxPurchase.Rows - 1
         flxPurchase.TextMatrix(i, 1) = ""
      Next i
      For i = 1 To flxPurchase.Rows - 2
         flxPurchase.TextMatrix(i, 1) = "X"
      Next i

      flxPurchase.row = flxPurchase.Rows - 1

      flxPurchase_Click
   Else
      For i = 1 To flxPurchase.Rows - 1
         flxPurchase.TextMatrix(i, 1) = ""
      Next i

      ConfigFlxPurchaseSplit
   End If
End Sub

'Private Sub cmbClient_Change()
''Resolved by BOSL
''Issue No: 0000467
''Added By: Asif. 04 Sep 2014
'LoadPropertyDropDown cmbProperty
'''''''''''''''''''''''''''''''
'End Sub

Private Sub txtIDClient_change_made()
   If Not bFormLoaded Then Exit Sub
   SortTheGrid flxPurchase, txtIDClient, txtPropID, txtSupplier
   flxPurchaseSplit.Clear
End Sub

Private Sub cmdOpClient_Click()
    sTextBox = "1"
    chkShowBal.Visible = False
    tabPurExp.Enabled = False
    picClient.Visible = True
    picClient.Left = 1300
    picClient.Top = 800
    
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdOpenSupp_Click()
    chkShowBal.Visible = False
   sTextBox = "PIHIST"
   LoadSupplierAccount ""
   
   tabPayment.Enabled = False
   tabPurExp.Enabled = False
  
   fraList.Left = 6975
   fraList.Top = 800
   fraList.Visible = True
   fraList.ZOrder 0
   txtSearch1.Visible = True
   txtSearch2.Visible = True
   txtSearch1.text = ""
   txtSearch2.text = ""
   txtSearch1.SetFocus
End Sub

'Private Sub txtClientIdlist_Change_made()
'   If Not bFormLoaded Then Exit Sub
'   SortTheGrid flxPurchHistory, txtClientIdlist, cmbPropertyHistory, txtSupplierSearc
'   flxPurchHistorySplit.Clear
'End Sub

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

      'If Len(cmbPropCombo.text) > 0 Then
         If cmbPropCombo.text = "ALL" Then
            sFlag = sFlag + 10
         Else
            If flxGrid.TextMatrix(iRow, 11) = cmbPropCombo.text Then sFlag = sFlag + 10
         End If
'      Else
'         sFlag = sFlag + 10
'      End If

      If Len(txtSupplierSearc.text) > 0 Then
         If cmbSuppCombo.text = "ALL" Then
            sFlag = sFlag + 1
         Else
            If flxGrid.TextMatrix(iRow, 5) = cmbSuppCombo.text Then sFlag = sFlag + 1
         End If
      Else
         sFlag = sFlag + 1
      End If

      If sFlag = 111 Then
         flxGrid.RowHeight(iRow) = 240
      Else
         flxGrid.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub





Private Sub txtClientID_GotFocus()
    txtClientID.Locked = True
End Sub

Private Sub txtDisplayMaxPurchaseHist_Change()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update ShoppingCentre set MaxPurChaseHist=" & Val(txtDisplayMaxPurchaseHist.text) & ""
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub txtDisplayMaxPurchaseHist_GotFocus()
    SelTxtInCtrl txtDisplayMaxPurchaseHist
End Sub

Private Sub txtDisplayMaxPurchaseHist_KeyPress(KeyAscii As Integer)
    Dim adoConn As New ADODB.Connection
    If KeyAscii = 13 Then
          adoConn.Open getConnectionString
          LoadFlxPurchHistory adoConn, ""
          adoConn.Close
          Set adoConn = Nothing
    End If
    DigitTextKeyPress txtDisplayMaxPurchaseHist, KeyAscii
End Sub


Private Sub txtDisplayMaxPurchPayHist_Change()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        adoConn.Execute "Update ShoppingCentre set MaxPurPaymentHist=" & Val(txtDisplayMaxPurchPayHist.text) & ""
        adoConn.Close
        Set adoConn = Nothing
    
End Sub

Private Sub txtDisplayMaxPurchPayHist_GotFocus()
    SelTxtInCtrl txtDisplayMaxPurchPayHist
End Sub

Private Sub txtDisplayMaxPurchPayHist_KeyPress(KeyAscii As Integer)
    Dim adoConn As New ADODB.Connection
    If KeyAscii = 13 Then
        adoConn.Open getConnectionString
        LoadFlxPurchPPHistory adoConn, ""
        adoConn.Close
        Set adoConn = Nothing
    End If
    DigitTextKeyPress txtDisplayMaxPurchPayHist, KeyAscii
End Sub

Private Sub txtInv_Change(Index As Integer)
    If Left(fraLay(1).Caption, 15) = "Transaction ID:" And Index = 0 Then
        flxPI.Tag = "EditedOrAdded"
        cmdSavePI.Enabled = True
    End If
End Sub

Private Sub txtInv_LostFocus(Index As Integer)
   
        
End Sub



Private Sub txtPayAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdAmountType
    End If
End Sub

Private Sub txtProperty_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = vbKeyTab Then
'        FocusControl cmdTypeList
'   End If
End Sub

Private Sub txtProperty_LostFocus()
    If cmdaddnewline.Enabled Then
        FocusControl cmdTypeList
    End If
End Sub

Private Sub txtPropertyIDHist_Change()
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    Call LoadFlxPurchHistory(adoconn, "")
'
'    adoconn.Close
'    Set adoconn = Nothing
End Sub

Private Sub txtPropID_Change()
'   If Not bFormLoaded Then Exit Sub
'   SortTheGrid flxPurchase, txtIDClient, txtPropID, txtSupplier
'   flxPurchaseSplit.Clear
End Sub

Private Sub cmbSC_Click()
   If txtSupplierID.text <> "" Then
      ConfigFlxPI
      PIComponents "DefaultMode"
   End If
   If cmbSC.text = "Landlord" Then
            txtClientID.text = ""
            txtClientID.Tag = ""
    End If
End Sub



Private Sub cmbSPAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtSPReference.SetFocus
    End If
End Sub

'Private Sub cmbSPAmtType_LostFocus()
'     'issue 571 Validation
'   'Added by anol 20 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim strTemp As String
''   If cmbSPAmtType.ListIndex = -1 Then
''        tabPurExp.Tab = 1
''        ShowMsgInTaskBar "Please select a valid Amount Type to proceed"
''        FocusControl cmbSPAmtType
''        Exit Sub
''   End If
'   Dim adoConn    As New ADODB.Connection
'   strTemp = Replace(cmbSPAmtType.text, "'", "''")
'   If Trim(cmbSPAmtType.text) <> "" Then
'        adoConn.Open getConnectionString
'        szSQL = "SELECT PRIMARYCODE.VALUE, SECONDARYCODE.CODE, SECONDARYCODE.VALUE, SECONDARYCODE.DESCRIPTION FROM PRIMARYCODE, SECONDARYCODE WHERE Flexible = TRUE AND PRIMARYCODE.CODE = 'RAT' AND PRIMARYCODE.CODE = SECONDARYCODE.PRIMARYCODE AND SECONDARYCODE.VALUE='" & strTemp & "'"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'            tabPurExp.Tab = 1
'            MsgBox "Please select a valid Payment Type to proceed", vbInformation, "Amount Type"
'            FocusControl cmbSPAmtType
'        End If
'        adoConn.Close
'        Set adoConn = Nothing
'    End If
'End Sub

Private Sub cmbSPSupplier_GotFocus()
'   bClicked = False
End Sub

Private Sub cmbSPSupplier_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If KeyAscii = 13 And Not bClicked Then
'      cmbSPSupplier_Click
'      Exit Sub
'   End If
   KeyAscii = 0
End Sub

Private Sub cmbSPSupplier_LostFocus()
'   If Not bClicked Then cmbSPSupplier_Click
End Sub

'Private Sub cmbSPSupplier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   iSupplier = cmbSPSupplier.ListIndex
'End Sub

Private Sub txtIDClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOpClient.SetFocus
    End If
    
End Sub

Private Sub txtPropID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOpProperty.SetFocus
    End If
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDate
    End If
End Sub

Private Sub txtReference_LostFocus()
    Dim adoConn As New ADODB.Connection
    Dim rsRefCheck As New ADODB.Recordset
    Dim temp As String
    Dim strTemp As String
    If Trim(txtReference.text) = "" Then Exit Sub
    temp = Replace(txtSupplierID.text, "'", "''")
    strTemp = Replace(txtReference.text, "'", "''")
    'If cmdEdit(1).Enabled = True Then ' that means form is in new save mode
        adoConn.Open getConnectionString
        rsRefCheck.Open "Select * from tblPurInv where SUPP_AC='" & temp & "' AND INV_NO='" & strTemp & "'", adoConn, adOpenStatic, adLockReadOnly
        If Not rsRefCheck.EOF Then
                MsgBox "An invoice with the same reference already exists for this supplier", vbInformation, "Warning"
        End If
        rsRefCheck.Close
        Set rsRefCheck = Nothing
        adoConn.Close
        Set adoConn = Nothing
'     End If
End Sub

Private Sub txtSearch1_GotFocus()
    sSearchSwitch = "ID"
End Sub

Private Sub txtSearch2_GotFocus()
     sSearchSwitch = "Name"
End Sub

Private Sub txtSearchClientID_GotFocus()
     bSearchClientNameFocus = False
End Sub

Private Sub txtSearchClientName_GotFocus()
    bSearchClientNameFocus = True
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
            fraSearch.Visible = False
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If tabPurExp.Tab = 0 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                    Call LoadFlxPurchaseFilter(adoConn, 3)
                    searchResultOn = True
                    'cmdSearch.Caption = "Clear Sea&rch"
                    fmeLoading.Visible = False
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                    cmdSearch.Caption = "Clear Sea&rch"
                    Call LoadFlxPurchaseFilter(adoConn, 4)
                    searchResultOn = True
                    fmeLoading.Visible = False
                End If
            ElseIf tabPurExp.Tab = 2 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchHistory(adoConn, "3")
                     searchResultOn = True
                     'cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     'cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchHistory(adoConn, "4")
                     searchResultOn = True
                End If
            ElseIf tabPurExp.Tab = 3 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchPPHistory(adoConn, "3")
                     searchResultOn = True
                     'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchPPHistory(adoConn, "4")
                     searchResultOn = True
                End If
            End If
            adoConn.Close
            FocusControl cmdSearchOK
        
    End If
    TextBoxKeyPrsDate txtSearchFromD, KeyAscii
End Sub

Private Sub txtSearchFromD_LostFocus()
    If txtSearchFromD.text <> "" Then
        TextBoxFormatDate txtSearchFromD
        SelTxtInCtrl txtSearchToD
     End If
End Sub

Private Sub txtSearchNo_Change()
'    txtSearchFromD.text = ""
'    txtSearchToD.text = ""
'    txtSearchRef.text = ""
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    If tabPurExp.Tab = 0 Then
'        If Len(txtSearchNo.text) > 0 Then
'            Call LoadFlxPurchaseFilter(adoConn, 1)
'        Else
'            Call LoadFlxPurchase(adoConn)
'        End If
'        fmeLoading.Visible = False
'        If Len(txtSearchNo.text) > 0 Then
'            cmdSearch.Caption = "Clear Sea&rch"
'        Else
'            cmdSearch.Caption = "Sea&rch"
'        End If
'    ElseIf tabPurExp.Tab = 2 Then
'        If Len(txtSearchNo.text) > 0 Then
'            Call LoadFlxPurchHistory(adoConn, "1")
'        Else
'            Call LoadFlxPurchHistory(adoConn, "")
'        End If
'        If Len(txtSearchNo.text) > 0 Then
'            cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'        Else
'            cmdSearchPurchaseHistory.Caption = "Sea&rch"
'        End If
'   ElseIf tabPurExp.Tab = 3 Then
'
'        If Len(txtSearchNo.text) > 0 Then
'            Call LoadFlxPurchPPHistory(adoConn, "1")
'        Else
'            Call LoadFlxPurchPPHistory(adoConn, "")
'        End If
'        If Len(txtSearchNo.text) > 0 Then
'            cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'        Else
'            cmdSearchPurchPayHistory.Caption = "Sea&rch"
'        End If
'    End If
'
'
'    adoConn.Close
'    Set adoConn = Nothing
    
End Sub

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
             txtSearchFromD.text = ""
             txtSearchToD.text = ""
             txtSearchRef.text = ""
             Dim adoConn As New ADODB.Connection
             adoConn.Open getConnectionString
             If tabPurExp.Tab = 0 Then
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchaseFilter(adoConn, 1)
                     searchResultOn = True
                 Else
                     Call LoadFlxPurchase(adoConn)
                 End If
                 fmeLoading.Visible = False
'                 If Len(txtSearchNo.text) > 0 Then
'                     cmdSearch.Caption = "Clear Sea&rch"
'                 Else
'                     cmdSearch.Caption = "Sea&rch"
'                 End If
             ElseIf tabPurExp.Tab = 2 Then
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchHistory(adoConn, "1")
                     searchResultOn = True
                 Else
                     Call LoadFlxPurchHistory(adoConn, "")
                 End If
'                 If Len(txtSearchNo.text) > 0 Then
'                     cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'                 Else
'                     cmdSearchPurchaseHistory.Caption = "Sea&rch"
'                 End If
            ElseIf tabPurExp.Tab = 3 Then
                
                 If Len(txtSearchNo.text) > 0 Then
                     Call LoadFlxPurchPPHistory(adoConn, "1")
                     searchResultOn = True
                 Else
                     Call LoadFlxPurchPPHistory(adoConn, "")
                 End If
'                 If Len(txtSearchNo.text) > 0 Then
'                     cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'                 Else
'                     cmdSearchPurchPayHistory.Caption = "Sea&rch"
'                 End If
             End If
             
             
             adoConn.Close
             Set adoConn = Nothing
    End If
End Sub

'Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
'    'remove unwanted characters from the string
'    Dim strCur As String
'    strCur = "!@#$%^&*()?><~`+=|\/.',{}[];:-%_20"
'
'    For iCount = 0 To Len(strCur)
'        txtSearchNo.text = Replace(txtSearchNo.text, Mid(strCur, iCount + 1, 1), "")
'    Next
'    If KeyAscii = 13 Then
'        txtSearchRef.SetFocus
'    End If
'End Sub

Private Sub txtSearchRef_Change()
'    txtSearchFromD.text = ""
'    txtSearchToD.text = ""
'    txtSearchNo.text = ""
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    If tabPurExp.Tab = 0 Then
'        If Len(txtSearchRef.text) > 0 Then
'            Call LoadFlxPurchaseFilter(adoConn, 2)
'        Else
'            Call LoadFlxPurchase(adoConn)
'        End If
'        fmeLoading.Visible = False
'        If Len(txtSearchRef.text) > 0 Then
'             cmdSearch.Caption = "Clear Sea&rch"
'        Else
'             cmdSearch.Caption = "Sea&rch"
'        End If
'    ElseIf tabPurExp.Tab = 2 Then
'        If Len(txtSearchRef.text) > 0 Then
'            Call LoadFlxPurchHistory(adoConn, "2")
'        Else
'            Call LoadFlxPurchHistory(adoConn, "")
'        End If
'        If Len(txtSearchNo.text) > 0 Then
'            cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'        Else
'            cmdSearchPurchaseHistory.Caption = "Sea&rch"
'        End If
'    ElseIf tabPurExp.Tab = 3 Then
'
'        If Len(txtSearchRef.text) > 0 Then
'            Call LoadFlxPurchPPHistory(adoConn, "2")
'        Else
'            Call LoadFlxPurchPPHistory(adoConn, "")
'        End If
'        If Len(txtSearchRef.text) > 0 Then
'            cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'        Else
'            cmdSearchPurchPayHistory.Caption = "Sea&rch"
'        End If
'    End If
'
'
'
'    adoConn.Close
'    Set adoConn = Nothing
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            txtSearchFromD.text = ""
            txtSearchToD.text = ""
            txtSearchNo.text = ""
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If tabPurExp.Tab = 0 Then
                If Len(txtSearchRef.text) > 0 Then
                    Call LoadFlxPurchaseFilter(adoConn, 2)
                    searchResultOn = True
                Else
                    Call LoadFlxPurchase(adoConn)
                End If
                fmeLoading.Visible = False
'                If Len(txtSearchRef.text) > 0 Then
'                     cmdSearch.Caption = "Clear Sea&rch"
'                     searchResultOn = True
'                Else
'                     cmdSearch.Caption = "Sea&rch"
'                End If
            ElseIf tabPurExp.Tab = 2 Then
                If Len(txtSearchRef.text) > 0 Then
                    Call LoadFlxPurchHistory(adoConn, "2")
                    searchResultOn = True
                Else
                    Call LoadFlxPurchHistory(adoConn, "")
                End If
'                If Len(txtSearchNo.text) > 0 Then
'                    cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'                Else
'                    cmdSearchPurchaseHistory.Caption = "Sea&rch"
'                End If
            ElseIf tabPurExp.Tab = 3 Then
               
                If Len(txtSearchRef.text) > 0 Then
                    Call LoadFlxPurchPPHistory(adoConn, "2")
                    searchResultOn = True
                Else
                    Call LoadFlxPurchPPHistory(adoConn, "")
                End If
'                If Len(txtSearchRef.text) > 0 Then
'                    cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'                Else
'                    cmdSearchPurchPayHistory.Caption = "Sea&rch"
'                End If
            End If
            adoConn.Close
            Set adoConn = Nothing
            txtSearchFromD.SetFocus
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
            fraSearch.Visible = False
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If tabPurExp.Tab = 0 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                    Call LoadFlxPurchaseFilter(adoConn, 3)
                    searchResultOn = True
                    'cmdSearch.Caption = "Clear Sea&rch"
                    fmeLoading.Visible = False
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                    'cmdSearch.Caption = "Clear Sea&rch"
                    Call LoadFlxPurchaseFilter(adoConn, 4)
                    searchResultOn = True
                    fmeLoading.Visible = False
                End If
            ElseIf tabPurExp.Tab = 2 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchHistory(adoConn, "3")
                     searchResultOn = True
                    ' cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     'cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchHistory(adoConn, "4")
                     searchResultOn = True
                End If
            ElseIf tabPurExp.Tab = 3 Then
                If Trim(txtSearchNo.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchRef.text) <> "" Then
                    'do nothing
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
                     Call LoadFlxPurchPPHistory(adoConn, "3")
                     searchResultOn = True
                     'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
                     'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
                     Call LoadFlxPurchPPHistory(adoConn, "4")
                     searchResultOn = True
                End If
            End If
            adoConn.Close
       FocusControl cmdSearchCancel
    End If
    TextBoxKeyPrsDate txtSearchToD, KeyAscii
End Sub

Private Sub txtSearchToD_LostFocus()
    If txtSearchToD.text <> "" Then TextBoxFormatDate txtSearchToD
End Sub

Private Sub txtSPSupplier_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdSPSupplier
    End If
End Sub

Private Sub txtSupplierSearc_Change()
'   If Not bFormLoaded Then Exit Sub
'   SortTheGrid flxPurchHistory, txtClientIdlist, cmbPropertyHistory, txtSupplierSearc
'   flxPurchHistorySplit.Clear
'   flxPurchHistorySplit.Rows = 2
End Sub

Private Sub txtSupplierType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdSupplierType
    End If

End Sub

Private Sub txtSupSearchHis_Change()
'   Dim iRow As Integer
'
'   If txtSupSearchHis.text = "ALL" Then
'      For iRow = 1 To flxPurchPPHistory.Rows - 1
'         flxPurchPPHistory.RowHeight(iRow) = 240
'      Next iRow
'
'      Exit Sub
'   End If
'
'   For iRow = 1 To flxPurchPPHistory.Rows - 1
'      If flxPurchPPHistory.TextMatrix(iRow, 5) = txtSupSearchHis.text Then
'         flxPurchPPHistory.RowHeight(iRow) = 240
'      Else
'         flxPurchPPHistory.RowHeight(iRow) = 0
'      End If
'   Next iRow
End Sub

Private Sub cmdSPSupplier_Click() 'loading picAccounts for supplier search
'    On Error GoTo ERR
    'mark 5
    chkShowBal.Visible = False
    sTextBox = "SPay"
    Dim strMarkonError As String
    strMarkonError = "1"
    txtAccountSearch(1).text = ""
    txtAccountSearch(2).text = ""
    If txtClientIDPurPay.text = "" Then
        ShowMsgInTaskBar "Please select a client", "Y", "N"
        FocusControl txtClientIDPurPay
        Exit Sub
    End If
    If txtSupplierType.text = "" Then
        MsgBox "Please select Supplier type", vbInformation, "Please Select Type"
        FocusControl txtSupplierType
        Exit Sub
    End If

    txtAccountSearch(1).text = ""
    txtAccountSearch(1).Width = lblSearch0(3).Left - lblSearch0(2).Left
    txtAccountSearch(1).Left = lblSearch0(2).Left
    strMarkonError = "2"
    txtAccountSearch(2).Width = lblSearch0(4).Left - lblSearch0(3).Left
    txtAccountSearch(2).Left = lblSearch0(3).Left
   'added by anol 30 Aug 2015
   'issue 571 Note 1156
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    'Debug.Print time
    ' I am going to rem this part 20181112 Because balance is already loaded when the form is starting
    'may be this part will afect after making a transacttion and update the balance Need to check that
    
'    If cmdACType.ListIndex = 1 Then
'            SupplierAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 2 Then
'            ClientAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 3 Then
'            AgentAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 4 Then
'            LandlordAccountBalance adoConn
'    End If
''    If frmMMain.frmPI_SupplierBalance_isUptoDate = False Then
''          SupplierAccountBalance adoconn
''          frmMMain.frmPI_SupplierBalance_isUptoDate = True
''    End If
''    If frmMMain.frmPI_SupplierBalanceByCL_isUptoDate = False Then
''            SupplierAcBalByClient2 adoconn  'load the client balance column
''            frmMMain.frmPI_SupplierBalanceByCL_isUptoDate = True
''    End If
'   Debug.Print time
    strMarkonError = "3"
    
    'added by anol 22 Aug 2016
    strMarkonError = "4"
    'SupplierAcBalByClient adoConn  'load the client balance column
    '20180823
    'Debug.Print time
    
'    Debug.Print time
     'end of addition
    strMarkonError = "5"
    Call LoadSupplierOnPayment(adoConn, "")  'Load all supplier to the pop up grid of supplier on payment screen
    strMarkonError = "6"
'    Debug.Print time
 '   Something non reachable code rem by anol 20181114
''    If lblSearch0(5).Caption = "NotLoaded" Then
''       strMarkonError = "7"
''       LoadflxSupplier adoConn
''       lblSearch0(5).Caption = "Loaded"
''    End If
    adoConn.Close
    Set adoConn = Nothing
    strMarkonError = "8"
    picAccounts.Left = tabPurExp.Left + tabPayment.Left + Frame8(1).Left + txtSPSupplier.Left
    picAccounts.Top = tabPurExp.Top + tabPayment.Top + Frame8(1).Top + txtSPSupplier.Top + 300
    picAccounts.Visible = True
    tabPurExp.Enabled = False
    picAccounts.ZOrder 0
   'anol 06 July 2015
   'txtAccountSearch(0).SetFocus
    strMarkonError = "9"
    FocusControl txtAccountSearch(1)
    strMarkonError = "10"
    ' flxSupplier(1).ColWidth(4) = 2400
     Exit Sub
Err:
    MsgBox "Please report this problem to PCM consulting." & vbCrLf & Err.description, vbInformation, "cmdSPSupplier_click: " & strMarkonError
End Sub

Private Sub cmdAccSel_Click()
    Dim adoConn As New ADODB.Connection
    Dim iRow As Integer
     chkShowBal.Visible = False
    ' If lblSearch0(5).Caption = "NotLoaded" Then
    'Set the ADO Connections to the dataset
    If adoConn.State = 0 Then
        adoConn.Open getConnectionString
    End If
    sTextBox = "PILIST"
    txtAccountSearch(0).text = ""
    txtAccountSearch(1).text = ""
    txtAccountSearch(2).text = ""
    txtAccountSearch(6).text = ""
    txtAccountSearch(7).text = ""
    LoadflxSupplier adoConn
    lblSearch0(5).Caption = "Loaded"
    adoConn.Close
    'End If
    cmdGridUnitLookup(2).Left = 5640
    
    Set adoConn = Nothing
    Set adoConn = Nothing
    'Resolved by BOSL
    'Modified By anol 20 Aug 2014
    picAccList.Left = 6500
    picAccList.Top = 800
    picAccList.Visible = True
    tabPurExp.Enabled = False
    picAccList.ZOrder 0
    txtAccountSearch(4).SetFocus
End Sub

Private Sub cmdACList_Click(Index As Integer) 'we don't need here indexed because there is only one control with this name
    Dim adoConn As New ADODB.Connection
    Dim rsShowBal As New ADODB.Recordset
    
    adoConn.Open getConnectionString
    sTextBox = "A/C"
    chkShowBal.Visible = True
    fraList.Width = 5490
    Call LoadSupplierAccount("") 'loading the supplier list in the grid
    'For show bal button load on first click of the supplier button
    rsShowBal.Open "Select PIShowBal From ShoppingCentre", adoConn, adOpenStatic, adLockReadOnly
    If Not rsShowBal.EOF Then
         If CBool(rsShowBal!PIShowBal) Then
            chkShowBal.Value = 1
         Else
            chkShowBal.Value = 0
         End If
    Else
         chkShowBal.Value = 0
    End If
    rsShowBal.Close
    Set rsShowBal = Nothing
    adoConn.Close
    Set adoConn = Nothing
    tabPayment.Enabled = False
    txtSearch1.Visible = True
    txtSearch2.Visible = True
    txtSearch1.text = ""
    txtSearch2.text = ""
    fraList.Width = 5200
    cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
    Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
    flxSupplier(0).Width = fraList.Width - 50
    flxSupplier(0).ColWidth(0) = 1600
    lblSearch1.Left = flxSupplier(0).Left + 1600
    txtSearch2.Left = flxSupplier(0).Left + 1600
    lblSearch2.Left = txtSearch2.Left + 2400
    txtSearch1.Width = 1550
    fraList.Left = txtSupplierID.Left + 100
    fraList.Top = txtSupplierID.Top + 370
    fraList.Visible = True
    fraList.ZOrder 0
    FocusControl txtSearch1
    tabPurExp.Enabled = False
End Sub
Private Sub focustxtSearch1()
On Error GoTo Err
     txtSearch1.SetFocus
     Exit Sub
Err:
End Sub
Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplier(0).Rows - 1
      For j = 0 To UBound(szaSupplierBal, 2) - 1
         If flxSupplier(0).TextMatrix(i, 0) = szaSupplierBal(0, j) Then
            flxSupplier(0).TextMatrix(i, 4) = Format(szaSupplierBal(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBal, 2) Then flxSupplier(0).TextMatrix(i, 4) = ""
   Next i
End Sub
Private Sub UpdateBalanceWithZero()
   Dim i As Integer

   For i = 1 To flxSupplier(0).Rows - 1
       flxSupplier(0).TextMatrix(i, 4) = ""
   Next i
End Sub
'Private Sub ArrayUpperVal()
'On Error GoTo ERR
'    Exit Sub
'ERR:
'End Sub
Private Sub LoadSupplierAccount(Filter As String)
   Dim adoConn As New ADODB.Connection
   Dim rstRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer
   Debug.Print 3

'ConfigFlxSupplier - Configuring flxSupplier grid
   With flxSupplier(0)
       fraList.Width = 5295
       flxSupplier(0).Width = 5175
       cmdGridUnitLookup(0).Left = 4945
      .Clear
      .Cols = 7
      .ColWidth(0) = 1400 'does not need actually
      .ColWidth(1) = 2400
      .ColAlignment(0) = vbLeftJustify
      .ColAlignment(1) = vbLeftJustify
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 900
      '.ColAlignment(4) = vbRightJustify
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      
      '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
      lblSearch0(0).Width = 700
      lblSearch0(0).Left = 60
      lblSearch0(0).Caption = "A/C ID"
      
      lblSearch1.Width = 2600
      lblSearch1.Left = lblSearch0(0).Left + .ColWidth(0)
      lblSearch1.Caption = "Name"
      
      lblSearch2.Width = 750
      lblSearch2.Left = 4020
      lblSearch2.Visible = True
      lblSearch2.Caption = "A/C Bal"

      txtSearch1.Width = 1400
      txtSearch1.Left = 70

      txtSearch2.Width = 2200
      txtSearch2.Left = txtSearch1.Left + .ColWidth(0)
      
     
      
      

      ' Error Handler
      'On Error GoTo ErrorHandler

      'Set the RDO Connections to the dataset
      adoConn.Open getConnectionString

      '~~~Added By Senthuran~~~ Code to configuer Label Caption
      
      
      iRow = 1
      If sTextBox = "A/C" Then
            If cmbSC.text = "Supplier" Then
                    szSQL = "SELECT SupplierID, SupplierName, NominalCode,VATCode, tlbVatCode.VAT_CODE, PaymentTerms, VAT_RATE " & _
                       "FROM Supplier LEFT JOIN tlbVatCode " & _
                           "ON Supplier.VATCode = cstr(tlbVatCode.VAT_ID) " & _
                       "WHERE Supplier.TYPE = 'SUPPLIER' " & _
                       "ORDER BY SupplierName;"
            ElseIf cmbSC.text = "Client" Then
'                    szSQL = "SELECT ClientID, ClientName, spare2 " & _
'                          "FROM Client " & _
'                          "ORDER BY ClientName;"
                      szSQL = "SELECT SupplierID as ClientID, SupplierName as ClientName , NominalCode as spare2, VATCode,  tlbVatCode.VAT_CODE, PaymentTerms, VAT_RATE " & _
                       "FROM Supplier LEFT JOIN tlbVatCode " & _
                           "ON Supplier.VATCode = cstr(tlbVatCode.VAT_ID) " & _
                       "WHERE Supplier.TYPE = 'Client' " & _
                       "ORDER BY SupplierName;"
            ElseIf cmbSC.text = "Landlord" Then
                    szSQL = "SELECT SupplierID, SupplierName, NominalCode, VATCode,  tlbVatCode.VAT_CODE, PaymentTerms, VAT_RATE " & _
                       "FROM Supplier LEFT JOIN tlbVatCode " & _
                           "ON Supplier.VATCode = cstr(tlbVatCode.VAT_ID) " & _
                       "WHERE Supplier.TYPE = 'LLORD' " & _
                       "ORDER BY SupplierName;"
                       
            ElseIf cmbSC.text = "Managing Agent" Then
                      szSQL = "SELECT SupplierID, SupplierName, NominalCode,VATCode, tlbVatCode.VAT_CODE, PaymentTerms, VAT_RATE " & _
                       "FROM Supplier LEFT JOIN tlbVatCode " & _
                           "ON Supplier.VATCode = cstr(tlbVatCode.VAT_ID) " & _
                       "WHERE Supplier.TYPE = 'AGENT' " & _
                       "ORDER BY SupplierName;"
            End If
      ElseIf sTextBox = "PILIST" Or sTextBox = "PIHIST" Or sTextBox = "PAYHIST" Then
             szSQL = "SELECT SupplierID, SupplierName, NominalCode, VATCode,tlbVatCode.VAT_CODE, PaymentTerms, VAT_RATE " & _
                       "FROM Supplier LEFT JOIN tlbVatCode " & _
                           "ON Supplier.VATCode = cstr(tlbVatCode.VAT_ID) " & _
                       "WHERE Supplier.TYPE = 'SUPPLIER' Or Supplier.TYPE = 'LLORD' or Supplier.TYPE = 'AGENT' " & _
                       "ORDER BY SupplierName;"
                       
                    
                    .TextMatrix(iRow, 0) = "ALL"
                    .TextMatrix(iRow, 1) = "All Supplier"
                    '.AddItem ""
                     iRow = 2
      
      End If
           
            
      rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      If Filter <> "" Then
            rstRst.Filter = Filter
      End If
      '.Clear
      .Rows = rstRst.RecordCount + iRow + 1
      
      If sTextBox = "A/C" And cmbSC.text = "Client" Then
           
            While Not rstRst.EOF
               .TextMatrix(iRow, 0) = rstRst!ClientID
               .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!ClientName), "", rstRst!ClientName)
               .TextMatrix(iRow, 2) = IIf(IsNull(rstRst!spare2), "", rstRst!spare2) 'what is this spare 2??? That is nominal code for client
               .TextMatrix(iRow, 3) = IIf(IsNull(rstRst!VAT_CODE) Or rstRst!VAT_CODE = "", "", rstRst!VAT_CODE & "##" & rstRst!VAT_RATE)
               .TextMatrix(iRow, 6) = IIf(IsNull(rstRst!VatCode), "", rstRst!VatCode)
               '.RowHeight(iRow) = 240
               rstRst.MoveNext
               'If Not rstRst.EOF Then .AddItem ""
               iRow = iRow + 1
            Wend
      Else
            While Not rstRst.EOF
               .TextMatrix(iRow, 0) = rstRst!SupplierID
               .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!SupplierName), "", rstRst!SupplierName)
               .TextMatrix(iRow, 2) = IIf(IsNull(rstRst!nominalCode), "", rstRst!nominalCode)
               .TextMatrix(iRow, 3) = IIf(IsNull(rstRst!VAT_CODE) Or rstRst!VAT_CODE = "", "", rstRst!VAT_CODE & "##" & rstRst!VAT_RATE)
                'Column 4 is for balance
               .TextMatrix(iRow, 5) = IIf(IsNull(rstRst!PaymentTerms), "", rstRst!PaymentTerms)
               .TextMatrix(iRow, 6) = IIf(IsNull(rstRst!VatCode), "", rstRst!VatCode)
               ' .RowHeight(iRow) = 240
               rstRst.MoveNext
               iRow = iRow + 1
            Wend
      End If
   End With
   

   rstRst.Close
   
   adoConn.Close
   If flxSupplier(0).Rows > 1 Then
        flxSupplier(0).row = 1
   End If
   If chkShowBal.Value = 1 Then
       UpdateBalance
   End If

   Set rstRst = Nothing
   Set adoConn = Nothing
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
   
   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
End Sub

'Private Sub LoadCboClientPI(adoConn As ADODB.Connection)
'   Dim szSQL   As String
'   Dim adoRst  As New ADODB.Recordset
'
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CT " & _
'           "FROM   CLIENT " & _
'           "ORDER BY CLIENTNAME;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientPI.Column() = Data()
'   cboClientPI.ListIndex = 0
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub LoadBankAccount()
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 800
   flxSupplier(0).ColWidth(1) = 2500
   
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   lblSearch0(0).Width = 700
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + 1000

   txtSearch1.Width = 700
   txtSearch1.Left = 40
   
   txtSearch2.Width = 2400
   txtSearch2.Left = txtSearch1.Left + 1000
   
   lblSearch0(0).Caption = "Bank Code"
   lblSearch1.Caption = "Bank Name"
   lblSearch2.Visible = False
   
   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   ' Error Handler
   On Error GoTo Error_Handler

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      ShowMsgInTaskBar "Please setup bank account for the client.", , "N"
   Else
      rRow = 1
      While Not adoRST.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = adoRST.Fields.Item("BNC").Value
         flxSupplier(0).TextMatrix(rRow, 1) = adoRST.Fields.Item("BNN").Value
         flxSupplier(0).AddItem ""
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
   End If

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ShowMsgInTaskBar "Prestige Database Error: ", , "N"
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdACType_Click()
'added by anol 30 Aug 2015 issue 571 note 1156
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    cmbSPSupplier.text = ""
'    txtSupAcBal.text = "0.00"
'    If cmdACType.ListIndex = 1 Then
'            SupplierAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 2 Then
'            ClientAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 3 Then
'            AgentAccountBalance adoConn
'    ElseIf cmdACType.ListIndex = 4 Then
'            LandlordAccountBalance adoConn
'    End If
'    adoConn.Close
End Sub

Private Sub cmdACType_KeyPress(KeyAscii As MSForms.ReturnInteger)
      If KeyAscii = 13 Then
    'cmdACType.SetFocus
      cmdSPSupplier.SetFocus
    End If
End Sub


'Private Function isVatOptionEnabledProperty(Conn As ADODB.Connection) As Boolean
'    Dim rsGlobalData As New ADODB.Recordset
'    isVatOptionEnabledProperty = False
'    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtProperty.text & "'", Conn, adOpenStatic, adLockReadOnly
'    If Not rsGlobalData.EOF Then
'            If IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value) = 0 Then
'                    isVatOptionEnabledProperty = False
'             Else
'                    isVatOptionEnabledProperty = True
'             End If
'    End If
'    rsGlobalData.Close
'    Set rsGlobalData = Nothing
'End Function
Private Function LoadVatOption(Conn As ADODB.Connection) As Integer
    Dim rsGlobalData As New ADODB.Recordset
    LoadVatOption = 0
    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtProperty.text & "' AND vatOptionEnabled=true ", Conn, adOpenStatic, adLockReadOnly
    If Not rsGlobalData.EOF Then
            LoadVatOption = IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value)
    End If
    rsGlobalData.Close
    Set rsGlobalData = Nothing
End Function
Private Sub cmdaddnewline_Click()
     Dim adoConn As New ADODB.Connection
     If txtSupplierID.text = "" Then
         ShowMsgInTaskBar "You must select Account Code from the list.", "Y", "N"
         FocusControl cmdACList(0)
         Exit Sub
      End If
      If txtDate.text = "" Then
         ShowMsgInTaskBar "You must enter the date from the list.", "Y", "N"
         FocusControl txtDate
         Exit Sub
      End If
      If txtDueDate.text = "" Then
         ShowMsgInTaskBar "You must enter the due date from the list.", "Y", "N"
         FocusControl txtDueDate
         Exit Sub
      End If
      If txtClientID.text = "" Then
         ShowMsgInTaskBar "You must select the client.", "Y", "N"
        FocusControl cmdClientSerc
         Exit Sub
      End If
        
      If IsDate(lblPostingDate.ToolTipText) = True Then
              Dim szSQL As String
              If IsNull(txtClientID.text) Then
                  ShowMsgInTaskBar "Please select a client", "Y", "N"
                  cmdClientSerc.SetFocus
                  Exit Sub
              End If
              adoConn.Open getConnectionString
              If IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 0 Then
                  ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
                  adoConn.Close
                  Exit Sub
              ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 9 Then
                  ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
                  adoConn.Close
                  Exit Sub
              End If
      End If
      'issue 571 note 1148
      'Added by anol 25 Aug 2015
      If Trim(txtProperty.text) = "" Then
            If MsgBox("You have not selected a property. Do you wish to add a property?", vbYesNo, "Select a Property") = vbYes Then
                cmdTypeList.SetFocus
                Exit Sub
            End If
      End If
      If Trim(txtProperty.text) = "" Then
          cmdUnitList.Enabled = False
      Else
          cmdUnitList.Enabled = True
      End If
      cmdNCList.Enabled = True
      cmdaddnewline.Enabled = False
      
      'added by anol 09 July 2015
      'Split line Enable mode
      txtUnit(0).Enabled = True
   
      txtJobNo.Enabled = True
      cmdJobNo(0).Enabled = True
      txtNet_(0).Enabled = True
      txtNC(0).Enabled = True
      cmdNCList.Enabled = True
      cmdUpdate(1).Enabled = True
      cmdTaxList(0).Enabled = True
      cmdDeptList.Enabled = True
      txtDetails_(0).Enabled = True
      txtVat_(0).Enabled = True
      cmdSchedules(0).Enabled = True
      cmdUpdate(2).Enabled = True
      txtUnit(0).Locked = True
'      fraLay(1).Enabled = False
      'Do all tax code visibility code here
      If adoConn.State = 1 Then
        adoConn.Close
      End If
      adoConn.Open getConnectionString
       If LoadVatOption(adoConn) = 0 Then
            vatOptionEnabled = False
       Else
            vatOptionEnabled = True
       End If
       If vatOptionEnabled = True Then
            cmdTaxList(0).Enabled = True
            txtVat_(0).Enabled = True
      Else
            cmdTaxList(0).Enabled = False
            txtVat_(0).Enabled = False
            txtVat_(0).text = ""
            lblVatCode(0).Caption = ""
            lblVatCode(0).Tag = -1
       End If
      If cmbSC.text = "Supplier" And txtProperty.text = "" Then
        'check client if opt enabled then take from supplier
            Dim rstRec As New ADODB.Recordset
            Dim rstVat As New ADODB.Recordset
            Dim rstSupplier As New ADODB.Recordset
            Dim bOptedtoTax As Boolean
           ' adoConn.Open getConnectionString
    
            szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtClientID.text & "'"
            rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rstRec.EOF Then
                    bOptedtoTax = True
            End If
            rstRec.Close
            Set rstRec = Nothing
            If bOptedtoTax = True Then 'if check client if opt enabled then take the tax code from supplier 2) 2.  Where supplier type = "Supplier" then IF while entering a supplier transaction and the client "Opted to tax" TRUE AND NO property is selected then the VAT Code will come from the client.(this must be from supplier)
                  szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S  " & _
                 "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true AND SupplierID='" & txtSupplierID.text & "' ORDER BY SupplierID;"
                  rstVat.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                  'if vat_ID in supplier  is "-1" that means it does not have a vat code
                  If Not rstVat.EOF Then
                          lblVatCode(0).Tag = IIf(IsNull(rstVat.Fields("VAT_ID").Value), "-1", rstVat.Fields("VAT_ID").Value)
                          nTaxCode = IIf(IsNull(rstVat.Fields("VAT_RATE").Value), "0.00", rstVat.Fields("VAT_RATE").Value)
                          lblVatCode(0).Caption = IIf(IsNull(rstVat.Fields("VAT_CODE").Value), "", rstVat.Fields("VAT_CODE").Value)
                  Else
                          lblVatCode(0).Tag = -1
                          nTaxCode = 0
                          lblVatCode(0).Caption = ""
                  End If
                  rstVat.Close
                  Set rstVat = Nothing
            Else        'if check client if opt not enabled  then clear vats'1.  1)Where supplier type = "Supplier" then IF while entering a supplier transaction AND the CLIENT "Opted to tax" FALSE AND NO property is selected then the VAT Code will come from the CLIENT and will show as empty.
                          lblVatCode(0).Tag = -1
                          nTaxCode = 0
                          lblVatCode(0).Caption = ""
            End If
     ElseIf cmbSC.text = "Supplier" And txtProperty.text <> "" Then
           'bOptedtoTax =true false in both cases you are taking va code from supplier
'                 szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S  " & _
'                 "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where SupplierID='" & txtAc(0).text & "' ORDER BY SupplierID;"

                 szSQL = "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    txtProperty.text & "' AND vatOptionEnabled=true"
                 
                  rstVat.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                  'if vat_ID in supplier  is "-1" that means it does not have a vat code
                  If Not rstVat.EOF Then
                          lblVatCode(0).Tag = IIf(IsNull(rstVat.Fields("VAT_ID").Value), "-1", rstVat.Fields("VAT_ID").Value)
                          nTaxCode = IIf(IsNull(rstVat.Fields("VAT_RATE").Value), "0.00", rstVat.Fields("VAT_RATE").Value)
                          lblVatCode(0).Caption = IIf(IsNull(rstVat.Fields("VAT_CODE").Value), "", rstVat.Fields("VAT_CODE").Value)
                  Else
                          lblVatCode(0).Tag = -1
                          nTaxCode = 0
                          lblVatCode(0).Caption = ""
                  End If
                  rstVat.Close
                  Set rstVat = Nothing
            
      
     'else if now  look at  If txtProperty.text = "" And cmbSC.text = "client" or manaign agent then and also when u fill up property and click add new
      ElseIf cmbSC.text = "Client" And txtProperty.text <> "" Then
'            szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtClientID.text & "'"
'            rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'            If Not rstRec.EOF Then
'                    bOptedtoTax = True
'            End If
'            rstRec.Close
'            Set rstRec = Nothing
'            If bOptedtoTax = True Then
            'what ever be the client value take it from the property I have checked both conditions
                 Dim rsGlobalData As New ADODB.Recordset
                 rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    txtProperty.text & "' AND vatOptionEnabled=true", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsGlobalData.EOF Then
                          lblVatCode(0).Tag = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                          nTaxCode = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                          lblVatCode(0).Caption = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                 Else
                          lblVatCode(0).Tag = -1
                          nTaxCode = 0
                          lblVatCode(0).Caption = ""
                 End If
                 rsGlobalData.Close
                 Set rsGlobalData = Nothing
'            End If
        ElseIf cmbSC.text = "Client" And txtProperty.text = "" Then
            szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtClientID.text & "'"
            rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rstRec.EOF Then
                          lblVatCode(0).Tag = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "-1", rstRec.Fields("VAT_ID").Value)
                          nTaxCode = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "0.00", rstRec.Fields("VAT_RATE").Value)
                          lblVatCode(0).Caption = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
             Else
                          lblVatCode(0).Tag = -1
                          nTaxCode = 0
                          lblVatCode(0).Caption = ""
            End If
            rstRec.Close
            Set rstRec = Nothing
      ElseIf (cmbSC.text = "Managing Agent" Or cmbSC.text = "Landlord") And txtProperty.text <> "" Then
'            szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtClientID.text & "'"
'            rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'            If Not rstRec.EOF Then
'                          bOptedtoTax = True
'            Else
'                          bOptedtoTax = False
'            End If
'
'            rstRec.Close
'            Set rstRec = Nothing
'            If bOptedtoTax = True Then
                'what ever at client vat code will come from property
                   
                    rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID where PropertyID='" & _
                                       txtProperty.text & "' AND vatOptionEnabled=true", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsGlobalData.EOF Then
                             lblVatCode(0).Tag = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                             nTaxCode = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                             lblVatCode(0).Caption = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                    Else
                             lblVatCode(0).Tag = -1
                             nTaxCode = 0
                             lblVatCode(0).Caption = ""
                    End If
                    rsGlobalData.Close
                    Set rsGlobalData = Nothing
'            Else
'            End If
        ElseIf (cmbSC.text = "Managing Agent" Or cmbSC.text = "Landlord") And txtProperty.text = "" Then
            szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtClientID.text & "'"
            rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rstRec.EOF Then
                          bOptedtoTax = True
            Else
                          bOptedtoTax = False
            End If

            rstRec.Close
            Set rstRec = Nothing
            If bOptedtoTax = True Then
                    szSQL = "SELECT  V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM (Supplier S LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where optedtoTax=true  AND SupplierID='" & txtSupplierID.text & "'"
                    rsGlobalData.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                    If Not rsGlobalData.EOF Then
                             lblVatCode(0).Tag = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                             nTaxCode = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                             lblVatCode(0).Caption = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                    Else
                             lblVatCode(0).Tag = -1
                             nTaxCode = 0
                             lblVatCode(0).Caption = ""
                    End If
                    rsGlobalData.Close
                    Set rsGlobalData = Nothing
            End If
      End If
      adoConn.Close
      Set adoConn = Nothing
      DoEvents
      FocusControl cmdDeptList
End Sub

Private Sub cmdCancel_Click(Index As Integer)
   If bEditMode Then
      If MsgBox("Do you want to cancel Edit?", vbQuestion + vbYesNo, "Edit Record") = vbNo Then Exit Sub
      bEditMode = False
      fraEditDemand.Enabled = True
      fraTab0.Enabled = True
      iSelected = 0
      ConfigFlxPI
   Else
      If MsgBox("Do you want to cancel?" & Chr(13) & "If you wish to save the data you already entered click No", vbQuestion + vbYesNo, "Add Record") = vbNo Then Exit Sub

      ConfigFlxPI
   End If
   PIComponents "DefaultMode"

   HandleCommandButton "Cancel"
'   cmdUnitList.Enabled = False
   flxPI.Enabled = True
   flxPI.col = 0
   flxPI.CellBackColor = vbWhite

'   fraLay(1).Enabled = True
End Sub

Private Sub cmdChqRemittNo_Click()
    'Bellow line commented by anol 29 Apr 2015
   'txtChqNo.Locked = True
   Frame4(0).Visible = False
   cmdChqRemittNo.Enabled = False
   Call SavePaymentTransactions
   cmdChqRemittNo.Enabled = True
   If areYouProcessingRentPayable = True Then
        Dim adoConn As New ADODB.Connection
        Dim dblClientACBalance As Double
        Dim dblBankBalance As Double
        'dblBankBalance = BankAccBalance(adoConn, txtBankCode.text, txtClientIDPurPay.text)
        Dim iRows As Integer
        adoConn.Open getConnectionString
        dblClientACBalance = GetClientACBalance
        dblBankBalance = BankAccBalance(adoConn, txtBankCode.text, txtClientIDPurPay.text)
        adoConn.Execute "Update RentSummaryStatement set  ClientACBalance=" & dblClientACBalance & ",BankACBalance=" & dblBankBalance & " where StatementID=" & CSID & ""
         If IsLoadedAndVisible("frmRentPayable") Then
            For iRows = 1 To frmRentPayable.flxPayFees.Rows - 1
                 If frmRentPayable.flxPayFees.TextMatrix(iRows, 2) = "CS" & CSID Then
                             frmRentPayable.flxPayFees.TextMatrix(iRows, 15) = dblBankBalance
                             frmRentPayable.flxPayFees.TextMatrix(iRows, 16) = -Format(dblclientPaymentAmount, "0.00") 'this is client payment
                             frmRentPayable.flxPayFees.TextMatrix(iRows, 20) = Format(dblClientACBalance, "0.00") 'this is client Ac balance
                     Exit For
                 End If
            Next iRows
         End If
         adoConn.Close
         areYouProcessingRentPayable = False
    End If
End Sub

Private Sub cmdChqRemittNo_LostFocus()
   On Error Resume Next

   optRemittanceOnly.SetFocus
End Sub

Private Sub cmdChqRemittYes_Click()
    'Below line commented by anol 29 Apr 2015
  ' txtChqNo.Locked = True
'  On Error GoTo ERR
   Frame4(0).Visible = False

   Dim lPayID As Long
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
   cmdChqRemittYes.Enabled = False
   lPayID = SavePaymentTransactions
   cmdChqRemittYes.Enabled = True
   If areYouProcessingRentPayable = True Then
        Dim adoConn As New ADODB.Connection
        Dim dblClientACBalance As Double
        Dim dblBankBalance As Double
        Dim iRows As Integer
        
        adoConn.Open getConnectionString
        dblClientACBalance = GetClientACBalance
        dblBankBalance = BankAccBalance(adoConn, txtBankCode.text, txtClientIDPurPay.text)
        adoConn.Execute "Update RentSummaryStatement set  ClientACBalance=" & dblClientACBalance & ",BankACBalance=" & dblBankBalance & " where StatementID=" & CSID & ""
         If IsLoadedAndVisible("frmRentPayable") Then
            For iRows = 1 To frmRentPayable.flxPayFees.Rows - 1
                 If frmRentPayable.flxPayFees.TextMatrix(iRows, 2) = "CS" & CSID Then
                            frmRentPayable.flxPayFees.TextMatrix(iRows, 15) = dblBankBalance
                             frmRentPayable.flxPayFees.TextMatrix(iRows, 16) = -Format(dblclientPaymentAmount, "0.00") 'this is client payment
                             frmRentPayable.flxPayFees.TextMatrix(iRows, 20) = Format(dblClientACBalance, "0.00") 'this is client Ac balance
                     Exit For
                 End If
            Next iRows
         End If
         adoConn.Close
         areYouProcessingRentPayable = False
    End If
    
   If optRemittanceOnly.Value Then
   
   '********************These are the Remittance Only options*****************************
      If UCase(strSupplierTypeOnSelection) = UCase("Supplier") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurSupplierRemittance.rpt") 'Remittance only report supplier
      End If
      If UCase(strSupplierTypeOnSelection) = UCase("Client") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurClientRemittance.rpt") 'Remittance only report client
      End If
      If UCase(strSupplierTypeOnSelection) = UCase("Agent") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurMAgentRemittance.rpt") 'Remittance only report agent
      End If
      
      If UCase(strSupplierTypeOnSelection) = "Landlord" Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurLandLordRemittance.rpt") 'Remittance only report Landlord
      End If
      
      
   Else '********************These are the Cheque options*****************************
      'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurPaymentCheque.rpt") 'Remittance with Cheque
      If UCase(strSupplierTypeOnSelection) = UCase("Supplier") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurSupplierCheque.rpt") 'Cheque Supplier
      End If
      If UCase(strSupplierTypeOnSelection) = UCase("Client") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurClientCheque.rpt") 'Cheque Client
      End If
      If UCase(strSupplierTypeOnSelection) = UCase("Agent") Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurAgentCheque.rpt") 'Cheque Agent
      End If
      
      If UCase(strSupplierTypeOnSelection) = "Landlord" Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurLandlordCheque.rpt") 'Cheque Landlord
      End If
   End If

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Call Sleep(100)
   Report.ParameterFields(1).AddCurrentValue (lPayID)

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   Exit Sub
Err:
    MsgBox Err.description
End Sub

Private Sub LoadDeptBk()
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 800
   flxSupplier(0).ColWidth(1) = 2500
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.

   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2400
   lblSearch1.Left = lblSearch0(0).Left + 800
   
   txtSearch1.Width = 700
   txtSearch1.Left = 40
   
   txtSearch2.Width = 2400
   txtSearch2.Left = txtSearch1.Left + 800
   
         '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Fund"
   lblSearch1.Caption = "Name"
   lblSearch2.Visible = False
   
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      flxSupplier(0).Clear
      flxSupplier(0).TextMatrix(0, 0) = "Dept. ID"
      flxSupplier(0).TextMatrix(0, 1) = "Department Name"

      rRow = 1
      While Not adoRST.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = adoRST.Fields.Item("FundID").Value
         flxSupplier(0).TextMatrix(rRow, 1) = adoRST.Fields.Item("FundName").Value
         rRow = rRow + 1
         adoRST.MoveNext
         If Not adoRST.EOF Then flxSupplier(0).AddItem ""
      Wend
   End If

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

Private Sub cmdClientSerc_Click()
    sTextBox = "PICLIENTID"
    chkShowBal.Visible = False
    tabPurExp.Enabled = False
    txtSearchClientID.text = ""
    txtSearchClientName.text = ""
    picClient.Left = 7205
    picClient.Top = 710
    picClient.Visible = True
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient(Filter As String)
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 6
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   
   txtSearchClientID.Width = 1530
   picClient.Width = 5295
   flxClient.Width = 5175
   cmdPicCLose.Left = 5010
   txtSearchClientName.Left = 1620
   lblClientName.Left = 1875
   txtSearchClientID.Left = 45
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3420
   picClient.Height = 4095
   flxClient.Height = 3345
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT, V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM ((CLIENT C INNER JOIN Supplier S ON C.ClientID=S.SupplierID) " & _
           "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) ORDER BY CLIENTID;"
     
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
            rstRec.Filter = Filter
   End If
   flxClient.Rows = rstRec.RecordCount + 2
   If tabPurExp.Tab = 0 Then
        If sTextBox = "1" Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 240
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
                flxClient.row = 1
                flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
                flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
                flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("CT").Value), "", rstRec.Fields("CT").Value)
                flxClient.TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
                flxClient.TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "", rstRec.Fields("VAT_ID").Value)
                flxClient.TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "", rstRec.Fields("VAT_RATE").Value)
                rstRec.MoveNext
               rRow = rRow + 1
           Wend
        Else
           rRow = 1
            While Not rstRec.EOF
                flxClient.row = 1
                flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
                flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
                flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("CT").Value), "", rstRec.Fields("CT").Value)
                flxClient.TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
                flxClient.TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "", rstRec.Fields("VAT_ID").Value)
                flxClient.TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "", rstRec.Fields("VAT_RATE").Value)
               rstRec.MoveNext
               rRow = rRow + 1
            Wend
        End If
   End If
   If tabPurExp.Tab = 1 Then
            rRow = 1
            While Not rstRec.EOF
               flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               rstRec.MoveNext
               rRow = rRow + 1
            Wend
   End If
   If tabPurExp.Tab = 2 Or tabPurExp.Tab = 3 Then
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 240
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               rstRec.MoveNext
               rRow = rRow + 1
           Wend
   End If

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cmdClinetAddAtch_Click()
    'file attachment
    '01 Dec 2015 added by anol
     If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
     If Not cmdEdit(1).Enabled Then 'means edit mode
        AddNewAttachmentInCombo cmbFiles, "PI", flxPurchase.TextMatrix(iPIEdit, 0)
     Else 'means in addnew mode
        AddNewAttachmentInCombo cmbFiles, "PI", cmdSavePI.Tag
     End If
     ShowMsgInTaskBar "The file has been saved successfully."
     Frame17.Visible = False
End Sub

Private Sub cmdCopy_Click()
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + cmdCopy.Top + Me.Top + fraEditDemand.Top + tabPurExp.Top + 1150
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + fraEditDemand.Left + tabPurExp.Left + cmdCopy.Left + 80
   frmPopUpMenu.CallingFrom "Purchase"
   frmPopUpMenu.Show
End Sub

Private Sub cmdCopyReceipt_Click()
   If flxPurchPPHistory.row < 1 Then
      MsgBox "Please select a transaction from the grid.", vbInformation + vbOKOnly, "Selection"
      flxPurchPPHistory.SetFocus
      Exit Sub
   End If

'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + cmdCopyReceipt.Top + _
'                      Me.Top + tabPurExp.Top + 1160
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + _
'                       tabPurExp.Left + _
'                       cmdCopyReceipt.Left + 80
   If Left(flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 2), 3) = "PPR" Then
      frmPopUpMenu.CallingFrom "PURCHASE_PAYMENT_PPR"
   Else
      If Left(flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 2), 2) = "PP" Then
         frmPopUpMenu.CallingFrom "PURCHASE_PAYMENT"
      End If
      If Left(flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 2), 2) = "PA" Then
         frmPopUpMenu.CallingFrom "PURCHASE_PAYMENT_ACCOUNT"
      End If
   End If
   frmPopUpMenu.Show
'Debug.Print frmPopUpMenu.Top & " " & frmPopUpMenu.Left
End Sub

Private Sub cmdCross_Click()
    Frame17.Visible = False
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then
        MsgBox "Please select file name from combo", vbInformation, "You need to select a file name"
        Exit Sub
   End If
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   If Not cmdEdit(1).Enabled Then 'means edit mode
         'fixed by anol 20160915
        DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), flxPurchase.TextMatrix(iPIEdit, 0), "PI"
   Else 'means in addnew mode
        DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), cmdSavePI.Tag, "PI" 'fixed by anol 20160915
   End If
   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
   Frame17.Visible = False
End Sub

Private Sub cmdDeptList_Click()
   'MousePointer = vbHourglass
    fraList.Height = 4325
   chkShowBal.Visible = False
   sTextBox = "Fund"
'   flxSupplier(0).Height = fraList.Height - 750
   LoadFund ""              'main function loading fund grid
   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4865
   fraList.Height = 4045
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 60
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtDept(0).Left + 100
   fraList.Top = txtDept(0).Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   
   'MousePointer = vbDefault
  'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
End Sub

Private Sub LoadFund(Filter As String)
   flxSupplier(0).Rows = 2
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   flxSupplier(0).ColAlignment(0) = vbLeftJustify
   flxSupplier(0).ColAlignment(1) = vbLeftJustify
   Dim iSel As Integer

         '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   txtSearch1.Width = 1460
   txtSearch1.Left = 40
   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   ' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim rsFundMatrix As New ADODB.Recordset
   
   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoConn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
   Else
        iSel = 1
        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
                txtProperty.text & "' and ClientID='" & txtClientID & "' and isDeleted=false"
   End If
   rsFundMatrix.Close
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        adoRST.Filter = Filter
   End If
   If adoRST.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxSupplier(0).Clear
      flxSupplier(0).Rows = 2
   Else
      flxSupplier(0).Clear
      flxSupplier(0).Rows = adoRST.RecordCount + 2
                 '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(tabPurExp.Tab).Caption = "Fund Code"
      lblSearch1.Caption = "Fund Name"
      lblSearch2.Visible = False
      
      flxSupplier(0).RowHeight(0) = 0
      'flxSupplier(0).Rows = 2

      rRow = 1
      While Not adoRST.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = adoRST.Fields.Item("FundCode").Value
         flxSupplier(0).TextMatrix(rRow, 1) = adoRST.Fields.Item("FundName").Value
         flxSupplier(0).TextMatrix(rRow, 2) = adoRST.Fields.Item("FundID").Value
         flxSupplier(0).TextMatrix(rRow, 3) = adoRST.Fields.Item("CategoryCode").Value
         rRow = rRow + 1
         adoRST.MoveNext
         'If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      Wend
   End If

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

Private Sub cmdEdit_Click(Index As Integer)
       Dim iCol As Integer
      'added by anol 09 July 2015
      'Split line Enable mode
      fraSearch.Visible = False
      cmdaddnewline.Enabled = True
      cmdEdit(0).Enabled = True 'Enabling 'edit' in the main grid 'procedure edit is calling from main form

txtDate.Enabled = True
chkIsMgtFee.Enabled = True
cmdClientSerc.Enabled = True
cmdTypeList.Enabled = True
cmdACList(0).Enabled = True
cmbSC.Enabled = True
txtDueDate.Enabled = True
cmdSavePI.Visible = True


      flxPI.Enabled = True
   If Index = 0 And flxPI.row > 0 Then
      cmdDelete.Enabled = False
      cmdNCList.Enabled = True
      cmdNCList.SetFocus
      txtUnit(0).Enabled = True
      cmdUnitList.Enabled = True
      txtJobNo.Enabled = True
      cmdJobNo(0).Enabled = True
      txtNet_(0).Enabled = True
      txtNC(0).Enabled = True
      cmdNCList.Enabled = True
      cmdUpdate(1).Enabled = True
      cmdTaxList(0).Enabled = True
      cmdDeptList.Enabled = True
      txtDetails_(0).Enabled = True
      txtVat_(0).Enabled = True
      cmdSchedules(0).Enabled = True
      cmdUpdate(2).Enabled = True
      flxPI.Enabled = True
      txtRecoverable(0).Enabled = True
      chkRecover.Enabled = True
      'End of addition
   End If
   If Index = 0 Then                                     'Edit the line
''         'added by anol 02 Jan 2015
''         'issue 469
''         txtInv(0).Locked = False
''         txtUnit(0).Locked = False
''         txtNC(0).Locked = False
''         txtDept(1).Locked = False
''         txtDept(0).Locked = False
''         txtPFName.Locked = False
''         txtJobNo.Locked = False
''         txtDetails_(0).Locked = False
''         txtNet_(0).Locked = False
''         txtVat_(0).Locked = False
''         txtSchedules.Locked = False
''         txtRecoverable(0).Locked = False
''         txtTotal.Locked = False
''         'End of modification
      If flxPI.RowHeight(flxPI.row) = 0 Then
            'added by anol 17 Aug 2015
            MsgBox "Select a record in the grid to edit.", vbInformation, "Warning!!"
            Exit Sub
      End If
      If flxPI.TextMatrix(flxPI.row, 1) = "" Then Exit Sub

      If iSelected = 0 Then
         'ShowMsgInTaskBar "Select a record in the grid to edit.", "Y", "N"
          MsgBox "Select a record in the grid to edit.", vbInformation, "Warning!!"
         Exit Sub
      End If

      PIComponents "EditLine"
      cmdEdit(0).Enabled = False 'added by anol issue 571 note 1125 , 09 july 2015 'disabling the button for better understanding the mode
      With flxPI
         txtUnit(0).text = .TextMatrix(.row, 21)
         txtNC(0).text = .TextMatrix(.row, 6) 'Nominal Code
         txtDept(0).text = .TextMatrix(.row, 7) 'Fund Code
         txtDept(1).text = .TextMatrix(.row, 8)
         txtPFName.text = .TextMatrix(.row, 25)
         txtJobNo.text = .TextMatrix(.row, 9)
         txtDetails_(0).text = .TextMatrix(.row, 11)
         txtNet_(0).text = .TextMatrix(.row, 12)
         lblVatCode(0).Caption = Trim(.TextMatrix(.row, 13))
         txtVat_(0).text = .TextMatrix(.row, 14)
         txtSchedules.text = .TextMatrix(.row, 20)
         txtRecoverable(0).text = .TextMatrix(.row, 22)
         chkRecover.Value = IIf(Val(txtRecoverable(0).text) > 0, 1, 0)
         txtTotal.text = .TextMatrix(.row, 15)
        
   
         sAddChoice = IIf(.TextMatrix(.row, 4) = "Invoice", "IN", "CN")
         bEditMode = True
         .TextMatrix(.row, 19) = "1"
        
   
         HandleCommandButton "Edit"
         .Enabled = False
   'Below line comment out by anol 16 Aug 2015
         '.row = 0
      End With
   End If

   If Index = 1 Then                                  'Edit the PI/PC
        'reverese control activation of view button
        fraLay(1).Enabled = True
        cmdDelete.Enabled = True
        cmdEdit(0).Enabled = True
        cmdOpenFileView.Visible = False
        cmdviewMenu.Visible = True
        
        Dim rCount As Integer
        Dim iIncDec As Integer
        Dim iRow As Integer
        For rCount = 1 To flxPurchase.Rows - 1
            If flxPurchase.TextMatrix(rCount, 1) = "X" Then
                iIncDec = iIncDec + 1
                iPIEdit = rCount
            End If
        Next
        If iIncDec <> 1 Then
            MsgBox "Please select one PI/PC only.", vbInformation + vbOKOnly, "PI/PC Selection"
            chkSelectAllDemands.Value = 0
            For rCount = 1 To flxPurchase.Rows - 1
                If flxPurchase.TextMatrix(rCount, 1) = "X" Then
                   flxPurchase.TextMatrix(rCount, 1) = ""
                   'rem by anol 20181118 I think this will be a slow process when dat will increase
'                   flxPurchase.row = rCount
'                     For iRow = 1 To flxPurchase.Cols - 1
'                       flxPurchase.col = iRow
'                       flxPurchase.CellBackColor = RGB(255, 255, 255)
'                    Next iRow
                    
                End If
            Next
        
            Exit Sub
        End If
      'end
'       'added by anol 01 Sep 2016
'        flxPurchase.TextMatrix(flxPurchase.row, 0) = ""
'        flxPurchase.CellBackColor = vbCyan 'RGB(179, 233, 174)
'            If flxPurchase.row > 0 Then
'            If flxPurchase.TextMatrix(flxPurchase.row, 0) = "" Then
'                 'If flxDemands.CellBackColor = RGB(174, 179, 233) Then
'                     For iCol = 1 To flxPurchase.Cols - 1
'                        flxPurchase.col = iCol
'                        flxPurchase.CellBackColor = RGB(174, 199, 200)
'                     Next iCol
'                ' End If
'            End If
'         End If
         
      If iPIEdit = 0 Then Exit Sub
    
      Dim X As Byte

      If Not IsPossible2Edit Then
         'added by anol 20160912
        flxPurchase.TextMatrix(iPIEdit, 1) = ""
        'flxPurchase.CellBackColor = vbCyan 'RGB(179, 233, 174)
            If flxPurchase.row > 0 Then
            If flxPurchase.TextMatrix(iPIEdit, 1) = "" Then
                 'If flxDemands.CellBackColor = RGB(174, 179, 233) Then
                     For iCol = 2 To flxPurchase.Cols - 1
                        flxPurchase.col = iCol
                        flxPurchase.CellBackColor = RGB(174, 205, 200)
                     Next iCol
                ' End If
            End If
         End If
         'ShowMsgInTaskBar "The transaction is fully/partially paid.", "Y", "N"
         Exit Sub
      End If
      'added by  anol 01 Dec 2015
      LoadAttachmentFiles cmbFiles, flxPurchase.TextMatrix(iPIEdit, 0), "PI"
      'End of addition
      'added by  anol 01 Dec 2015
      cmdOpenFileView.Visible = False
      cmdviewMenu.Visible = True
      'End of addition
      
      
    'added by anol issue 571 date 13 Aug 2015
      cmdEdit(1).Enabled = False                 'At the saving time system will know it PI is in EDIT mode
      cmdNCList.Enabled = False
      'cmdNCList.SetFocus
      txtUnit(0).Enabled = False
      txtUnit(0).text = ""
      cmdUnitList.Enabled = False
      txtJobNo.Enabled = False
      txtJobNo.text = ""
      cmdJobNo(0).Enabled = False
      txtNet_(0).Enabled = False
      txtNet_(0).text = ""
      txtNC(0).Enabled = False
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtTotal.text = ""
      cmdNCList.Enabled = False
      cmdUpdate(1).Enabled = False
      cmdTaxList(0).Enabled = False
      cmdDeptList.Enabled = False
      txtDetails_(0).Enabled = False
      txtDetails_(0).text = ""
      txtVat_(0).Enabled = False
      txtVat_(0).text = ""
      cmdSchedules(0).Enabled = False
      txtRecoverable(0).Enabled = False
      chkRecover.Enabled = False
      cmdUpdate(2).Enabled = False
      cmdUpdate(1).Enabled = False
      FocusControl cmdACList(0)
      'End of addition
      
     'added by anol issue 571 date 23 July 2015
'     flxPI.ColWidth(11) = Label7(8).Left - Label7(7).Left - 200 '"Details"
      With flxPurchase
            If .TextMatrix(iPIEdit, 20) = "SUPPLIER" Then
                cmbSC.ListIndex = 0
            ElseIf .TextMatrix(iPIEdit, 20) = "CLIENT" Then
                cmbSC.ListIndex = 1
            ElseIf .TextMatrix(iPIEdit, 20) = "LLORD" Then
                cmbSC.ListIndex = 3
            ElseIf .TextMatrix(iPIEdit, 20) = "AGENT" Then
                cmbSC.ListIndex = 2
            End If
            txtTransType.text = .TextMatrix(iPIEdit, 3)
            txtSupplierID.text = .TextMatrix(iPIEdit, 5)
            cmdACList(0).Enabled = IIf(.TextMatrix(iPIEdit, 9) <> .TextMatrix(iPIEdit, 12), False, True)
            txtDate.text = .TextMatrix(iPIEdit, 4)
            txtDueDate.text = Format(.TextMatrix(iPIEdit, 13), "dd/mm/yyyy")
            txtSupplierName.text = .TextMatrix(iPIEdit, 6)
            txtReference.text = .TextMatrix(iPIEdit, 7)
            txtClientID.text = .TextMatrix(iPIEdit, 14)
            txtProperty.text = .TextMatrix(iPIEdit, 11)
            lblPostingDate.ToolTipText = .TextMatrix(iPIEdit, 16)
            'chkIsMgtFee.Value = IIf(.TextMatrix(iPIEdit, 26), 1, 0)
             
            
            LoadSplit4Edit flxPI, 0 'here loading the Detail lines in the grid for editing  Main Function
            
            fraLay(0).Left = 80
            fraLay(0).Top = 360
            cmdNew(0).Enabled = False
            fraLay(1).Caption = "Transaction ID: " & .TextMatrix(iPIEdit, 2)
            fraLay(1).Tag = .TextMatrix(iPIEdit, 19) 'putting the serial no /slnumber
            'MsgBox fraLay(1).Tag
            ' cmdUpdate(1).Enabled = True
            
            cmdSavePI.Enabled = False  'changes to false by anol 13 July 2015
         
         If .TextMatrix(iPIEdit, 17) = "" Then
            cmdPO.Enabled = False
         Else
            cmdPO.Enabled = True
         End If
      End With

   End If
         'added by anol 20160912
       flxPurchase.TextMatrix(iPIEdit, 1) = ""
       'flxPurchase.CellBackColor = vbCyan 'RGB(179, 233, 174)
           If flxPurchase.row > 0 Then
               If flxPurchase.TextMatrix(iPIEdit, 1) = "" Then
                    'If flxDemands.CellBackColor = RGB(174, 179, 233) Then
                    'coloring is tardy when too many rows by anol 20181118
    ''                    For iCol = 1 To flxPurchase.Cols - 1
    ''                       flxPurchase.col = iCol
    ''                       flxPurchase.CellBackColor = RGB(174, 205, 200)
    ''                    Next iCol
                   ' End If
               End If
        End If
End Sub

Private Function IsPossible2Edit() As Boolean
   Dim adoPay As New ADODB.Recordset
   Dim adoRST As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer

   adoConn.Open getConnectionString
   szSQL = "SELECT (Amount - OSAmount) AS Paid, UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
           "FROM tlbPayment " & _
           "WHERE tlbPayment.PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"

   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    
    If Not adoPay.EOF Then
        If Val(adoPay.Fields.Item("Paid").Value) <> 0 Then
           IsPossible2Edit = False
           MsgBox "The transaction is fully/partially paid.", vbInformation, "Paid"
        Else
           IsPossible2Edit = True
        End If
    Else
        IsPossible2Edit = True
    End If
   
    If IsPossible2Edit = True Then
        If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 And szSQL <> UserSessionID Then 'szSQL <> UserSessionID shall be always true bcoz PI is generating only one seeion through out this module.
                flxPurchase.col = 1
                flxPurchase.row = iPIEdit
                flxPurchase.CellBackColor = vbRed
                'MsgBox "Selected invoice is locked by another user. Please wait untill other user release this record.", vbInformation, "Warning"
                MsgBox "The selected invoice is currently locked by '" & IIf(IsNull(adoPay("WindowsUserName").Value), "", adoPay("WindowsUserName").Value) & _
                "' on '" & IIf(IsNull(adoPay("MachineName").Value), "", adoPay("MachineName").Value) & "' in the '" & IIf(IsNull(adoPay("Module").Value), "", adoPay("Module").Value) & "'" & vbCrLf & "" & _
                        "screen for the Client '" & IIf(IsNull(adoPay("ClientID").Value), "", adoPay("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
                IsPossible2Edit = False
            Else 'lock row for this user in database
                flxPurchase.col = 1
                flxPurchase.row = iPIEdit
                flxPurchase.CellBackColor = vbWhite
'                adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
'                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "'"
        
            End If
        End If
     End If
     adoPay.Close
     Set adoPay = Nothing
   'Check if there is any saved transaction before edit this invoice
       If IsPossible2Edit = True Then
            szSQL = "SELECT B.*, T.TransactionID, T.PayAmt, T.PayDt,T.PostingDate,T.ref " & _
                    "FROM tblBatchPayment AS B, tblBatchTransaction AS T " & _
                    "WHERE B.BP = T.BP AND B.Generated = FALSE AND TransactionID = " & Val(flxPurchase.TextMatrix(iPIEdit, 21)) & " AND T.PayAmt > 0;"
            
            adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not adoRST.EOF Then
                   MsgBox "The selected invoice cannot be edited. This transaction has been saved in Batch Payments for future payment.", vbInformation, "Warning"
                  IsPossible2Edit = False
            End If
            adoRST.Close
            Set adoRST = Nothing
        End If
        If IsPossible2Edit = True Then 'the reason I need to lock it here because it has passed all the tests here to open PI and now I can lock
             adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "'"
        End If
        adoConn.Close
        Set adoConn = Nothing
End Function
Private Function InstantLockingCheck() As Boolean 'unlocking for all row
   Dim adoPay As New ADODB.Recordset
   Dim rsLockDialog As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim selRow As Integer
   Dim selcol As Integer
   Dim i As Integer
   Dim j As Integer
   Dim strSQL As String
   selRow = flxPurchase.row
   selcol = flxPurchase.col
   
   
   adoConn.Open getConnectionString
   ' I am doing some tests here
   ' on loading time full table vs selected row
   
''   szSQL = "SELECT TransactionID,(Amount - OSAmount) AS Paid, UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
''           "FROM tlbPayment "
''
''   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''   'adoPay.Filter = "(UserSessionID)='' or UserSessionID=null" 'black ones
''   'adoPay.Filter = "(UserSessionID)<>''" 'locked ones ' SO filter works fine Need to implement that so one query shall solve purpose
''   'Debug.Print adoPay.RecordCount
''   Dim LockedTransaction() As String
''   ReDim LockedTransaction(adoPay.RecordCount) As String
''   Dim UnLockedTransaction() As String
''   ReDim UnLockedTransaction(adoPay.RecordCount) As String
''   Debug.Print time
''   i = 0
''   While Not adoPay.EOF
''        If Len(IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)) > 0 Then
''            LockedTransaction(i) = adoPay("TransactionID").Value
''        End If
''        i = i + 1
''   adoPay.MoveNext
''   Wend
''   flxPurchase.col = 1
''   For j = 1 To flxPurchase.Rows - 1
''
''            If IsInArray(flxPurchase.TextMatrix(j, 21), LockedTransaction) = True Then
''                'record is locked mark it red
''            Else
''                'rest of them unlock unlock
''            End If
''           'adoPay.Filter = "(TransactionID)='" & flxPurchase.TextMatrix(j, 21) & "'" 'adoPay.Filter is a very slow porecess we should  use instr
'''           If Not adoPay.EOF Then    'szSQL <> UserSessionID shall be always true bcoz PI is generating only one session thrgh out this module.
'''                   If Len(IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)) > 0 Then
'''                        'flxPurchase.col = 1
'''                        flxPurchase.row = j
'''                        flxPurchase.CellBackColor = vbRed
'''                    Else 'lock for this user
'''                        'flxPurchase.col = 1
'''                        flxPurchase.row = j
'''                        flxPurchase.CellBackColor = vbWhite
'''                        flxPurchase.TextMatrix(j, 22) = ""
'''                        flxPurchase.TextMatrix(j, 23) = ""
'''                        flxPurchase.TextMatrix(j, 24) = ""
'''                        flxPurchase.TextMatrix(j, 25) = ""
'''                    End If
'''            End If
''           ' Debug.Print j
'''             flxPurchase.col = 1 '20% slow meth0d
'''                        flxPurchase.row = j
''   Next j
''   Debug.Print time
''   adoPay.Close
''   adoConn.Close
''    flxPurchase.col = selcol
''        flxPurchase.row = selRow
 '  Exit Function
   'first part is instant lock
   szSQL = "SELECT TransactionID,(Amount - OSAmount) AS Paid, UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
           "FROM tlbPayment " & _
           "WHERE tlbPayment.PI = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "';"
'    szSQL = "SELECT TransactionID,SlNumber, TransactionType,(Amount - OSAmount) AS Paid, UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
'           "FROM tlbPayment where (type=6 or type =7) AND (Amount - OSAmount)=0 AND (isnull(UserSessionID) OR UserSessionID='') order by 2,3 Desc;"
   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  'locking status show for current row
   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then   'szSQL <> UserSessionID shall be always true bcoz PI is generating only one session thrgh out this module.
                flxPurchase.col = 1
                'flxPurchase.row = flxPurchase.row
                flxPurchase.CellBackColor = vbRed
                InstantLockingCheck = False
                colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoPay("TransactionID").Value), "", adoPay("TransactionID").Value) & ","
            Else 'lock for this user
                        flxPurchase.col = 1
                        i = flxPurchase.row
                        flxPurchase.CellBackColor = vbWhite
    '                    adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
    '                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "'"
                        'Need to clear the locking flag
                        flxPurchase.TextMatrix(i, 22) = ""
                        flxPurchase.TextMatrix(i, 23) = ""
                        flxPurchase.TextMatrix(i, 24) = ""
                        flxPurchase.TextMatrix(i, 25) = ""
            End If
           
   End If
   'second part isnstant unlock
   
'   strSQL = "Select DateTimeStamp ,UserSessionID,transactionID " & _
'               "from tlbPayment as Pt  where  UserSessionID='' AND TransactionID in (" & colTransactionIDOther & ")"
       
      
      'MsgBox colTransactionIDOtherPIGrid
      
      If Len(colTransactionIDOtherPIGrid) > 0 Then
            szSQL = "SELECT TransactionID,SlNumber, Type,(Amount - OSAmount) AS Paid, UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
                 "FROM tlbPayment where (type=6 or type =7) AND (Amount - OSAmount)=0 AND (isnull(UserSessionID) OR UserSessionID='') " & _
                 " AND TransactionID in (" & colTransactionIDOtherPIGrid & ") order by 2,3 Desc;"
            rsLockDialog.Open szSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
             While Not rsLockDialog.EOF
                      flxPurchase.col = 1
                      For j = 1 To flxPurchase.Rows - 1
                           ' If "PI9339" = flxPurchase.TextMatrix(j, 2) Then
                                'Debug.Print flxPurchase.TextMatrix(j, 2)
                            'End If
                            'If j = 3 And rsLockDialog("SlNumber").Value = 9339 Then
                               ' Debug.Print flxPurchase.TextMatrix(j, 2)
                            'End If
'                           Debug.Print flxPurchase.TextMatrix(j, 2)
                           'Debug.Print rsLockDialog("SlNumber").Value
                          If flxPurchase.TextMatrix(j, 21) = rsLockDialog("transactionID").Value And i <> j Then 'no need to update row of first part check
                                
                                flxPurchase.row = j
                                'flxPurchase.col = 1
                                'Debug.Print j
                                flxPurchase.CellBackColor = vbWhite
                          End If
                       Next j
                    rsLockDialog.MoveNext
              Wend
        End If
        'second part ends here
        
        flxPurchase.col = selcol
        flxPurchase.row = selRow
        
        
        
        
   adoPay.Close
   adoConn.Close
   flxPurchase.row = selRow
   flxPurchase.col = selcol
   Set adoPay = Nothing
   Set adoConn = Nothing
End Function
Private Sub LoadSplit4Edit(flxGrid As MSHFlexGrid, iID As Integer)
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim NumOfSpaces As Integer
   Dim adoPay As New ADODB.Recordset

   adoConn.Open getConnectionString
  szSQL = "SELECT isManagementFee " & _
           "FROM tblPurInv " & _
           "WHERE MY_ID = '" & flxPurchase.TextMatrix(iPIEdit, iID) & "';"

   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoPay.EOF Then
        chkIsMgtFee.Value = IIf(adoPay("isManagementFee").Value = True, 1, 0)
   End If
   adoPay.Close
   
   
   szSQL = "SELECT DISTINCT tblPurInvSRec.*, Fund.FundName, Fund.FundCode, tblPurInv.PO " & _
           "FROM tblPurInvSRec, Fund, tblPurInv " & _
           "WHERE tblPurInvSRec.ParentID = '" & flxPurchase.TextMatrix(iPIEdit, iID) & "' AND " & _
                 "tblPurInvSRec.DEPT_ID = Fund.FundID AND " & _
                 "tblPurInvSRec.ParentID = tblPurInv.My_ID " & _
           "ORDER BY TRAN_ID;"
   adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim i As Integer
   
   flxGrid.Rows = 2
   ReDim VTJobAmount(adoInvSp.RecordCount)
   With flxGrid
      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 1) = txtSupplierID.text
         .TextMatrix(iRow, 2) = txtDate.text
         'FIXED BY ANOL 12 Dec 2015
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("TRANS").Value), "", adoInvSp.Fields.Item("TRANS").Value) 'property ID from tblPurInvSRec table
         .TextMatrix(iRow, 4) = IIf(txtTransType.text = "Invoice", "Invoice", "Credit")
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 6) = adoInvSp.Fields.Item("NOMINAL_CODE").Value 'this is Reference and that textbox is set on click flxpurchase. flxPurchaseis the main list of invoices.
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value) 'this is fund ID
         .TextMatrix(iRow, 9) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 12) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
          If Len(adoInvSp.Fields.Item("TAX_CODE").Value) < 8 Then  'reason of this small programming is alignment is left and we need  right alignment
               NumOfSpaces = 14 - Len(adoInvSp.Fields.Item("TAX_CODE").Value)
          End If
         .TextMatrix(iRow, 13) = Space$(NumOfSpaces) & adoInvSp.Fields.Item("TAX_CODE").Value
         '.ColAlignment(13) = vbRightJustify
         .TextMatrix(iRow, 14) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 15) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 20) = IIf(IsNull(adoInvSp.Fields.Item("ScheduleID").Value), "", adoInvSp.Fields.Item("ScheduleID").Value)
         .TextMatrix(iRow, 21) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 22) = adoInvSp.Fields.Item("RecoverablePt").Value
         .TextMatrix(iRow, 23) = adoInvSp.Fields.Item("MY_ID").Value 'MY_ID from tblPurInvSRec table
         .TextMatrix(iRow, 24) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 25) = adoInvSp.Fields.Item("FundName").Value
         'Modifed by anol 15 Sep 2014
         'Edit was not possible
         .TextMatrix(iRow, 26) = IIf(IsNull(adoInvSp.Fields.Item("PO").Value), "", adoInvSp.Fields.Item("PO").Value)
          'adoInvSp.Fields.Item("PO").Value
          'anol 09 MAR 2015
          If .TextMatrix(iRow, 9) <> "" Then
               i = i + 1
               VTJobAmount(i).JobID = .TextMatrix(iRow, 9)
               VTJobAmount(i).amount = Val(.TextMatrix(iRow, 15))
          End If
          
         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      .row = 0
   End With

   adoInvSp.Close
   Set adoInvSp = Nothing

   UpdateTotalPICN
End Sub

Private Sub HandleCommandButton(szButton As String)
On Error GoTo Err
   Select Case szButton
      Case "Save"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = True
         cmdCancel(0).Enabled = False

         flxPI.Enabled = True
         flxPI.col = 0
         flxPI.CellBackColor = vbWhite

         ConfigFlxPI

      Case "Add Invoice"
         'cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = False

      Case "Edit"
         cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = False

      Case "Cancel"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = True
         cmdCancel(0).Enabled = False

      Case "Update Record"
         cmdSavePI.Enabled = True
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = True
         flxPI.Enabled = True
         flxPI.row = 0
   End Select
   Exit Sub
Err:
End Sub

Private Sub cmdEdit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraInvCrChoice.Visible Then cmdManualDmdCancel_Click
End Sub

Private Sub cmdEditPayment_Click()
   'If sEditPPR = 2 Then GoTo EditPayment

'   If flxSPayment.TextMatrix(flxSPayment.row, 20) <> "" Then
'      ShowMsgInTaskBar "The refund has been reconciled and cannot be edited", "Y", "N"
'      Exit Sub
'   End If
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If sEditPPR = 1 Then 'Edit PPR load values form the first grid
            If Not editexception Then 'this variable will disable the validation in case you shift+ click
               If Val(flxSPayment.TextMatrix(flxSPayment.row, 8)) <> Val(flxSPayment.TextMatrix(flxSPayment.row, 9)) Then
                   frmPaymentEdit.BoolReconciled = True
'                  ShowMsgInTaskBar "You cannot edit a part allocated refund. Please unallocate fully before edit", "Y", "N"
'                  Exit Sub
               Else
                    frmPaymentEdit.BoolReconciled = False
               End If
            End If
            '**** Code for edit PPR starts here
           LoadForm frmPaymentEdit
           
           With frmPaymentEdit
               .TransactionID = frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 19)
              .LoadFlxPaymentSplit adoConn
              .Caption = .Caption & " Refund - " & flxSPayment.TextMatrix(flxSPayment.row, 1) 'transaction ID
              '.cmbSPSupplier.ListIndex = cmbSPSupplier.ListIndex
              .txtSPSupplier.text = txtSPSupplier.Tag
              .txtDate.text = flxSPayment.TextMatrix(flxSPayment.row, 5) 'for type 24 this is Pdate else Ddate
              .cboBC.ListIndex = FindComboIndex(.cboBC, flxSPayment.TextMatrix(flxSPayment.row, 21), 0) 'BankCode
              
              .txtReference.text = flxSPayment.TextMatrix(flxSPayment.row, 6) 'ref
              .cmbSPAmtType.ListIndex = FindComboIndex(.cmbSPAmtType, flxSPayment.TextMatrix(flxSPayment.row, 22), 0) 'PayAmtType
               .cmbFund.ListIndex = FindComboIndex(.cmbFund, .flxPaymentSplit.TextMatrix(1, 1), 0)
               .txtProperty = .flxPaymentSplit.TextMatrix(1, 9)
                .lblPostingDate.ToolTipText = .flxPaymentSplit.TextMatrix(1, 8)
             ' .cmbFund.ListIndex = FindComboIndex(.cmbFund, flxSPayment.TextMatrix(flxSPayment.row, 13), 0)
              .txtDetails.text = flxSPayment.TextMatrix(flxSPayment.row, 7)
              .txtAmount.text = Format(flxSPayment.TextMatrix(flxSPayment.row, 8), "0.00")
              'added by anol on 25 Aug 2015 issue 571 note 1148
              If flxSPayment.TextMatrix(flxSPayment.row, 23) <> "" Then
                ' .cboClient.ListIndex = FindComboIndex(.cboClient, flxSCrPoA.TextMatrix(flxSCrPoA.row, 18), 0)
                   .txtClient.text = flxSPayment.TextMatrix(flxSPayment.row, 23)
              End If
             ' .TransactionID = frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 19)
              .InvoiceNO = frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 1)
              '.LoadFlxPaymentSplit adoconn
           End With
        
           Me.Enabled = False
'           frmPaymentEdit.Left = 100
'           frmPaymentEdit.Top = 100
'           frmPaymentEdit.Show
    '**** Code for edit PPR Ends here
  End If
   
   
   
   
    '**** Code for edit payment Ends here
'EditPayment:
    If sEditPPR = 2 Then 'Edit PA,PP. load some values from the second grid
            'validation on locking
            flxSCrPoA.col = 0
            Dim rsLockCheck As New ADODB.Recordset
            rsLockCheck.Open "Select UserSessionID,WindowsUserName,MachineName,Module,ClientID from tlbPayment where transactionID=" & flxSCrPoA.TextMatrix(flxSCrPoA.row, 10) & "", adoConn, adOpenStatic, adLockReadOnly
            If Not rsLockCheck.EOF Then
                If rsLockCheck("UserSessionID").Value <> "" And rsLockCheck("UserSessionID").Value <> UserSessionID Then
                    flxSCrPoA.CellBackColor = vbRed
                    MsgBox "The selected " & flxSCrPoA.TextMatrix(flxSCrPoA.row, 1) & " is currently locked by '" & IIf(IsNull(rsLockCheck("WindowsUserName").Value), "", rsLockCheck("WindowsUserName").Value) & _
                            "' on '" & IIf(IsNull(rsLockCheck("MachineName").Value), "", rsLockCheck("MachineName").Value) & "' in the '" & IIf(IsNull(rsLockCheck("Module").Value), "", rsLockCheck("Module").Value) & "'" & vbCrLf & "" & _
                            "screen for the Client '" & IIf(IsNull(rsLockCheck("ClientID").Value), "", rsLockCheck("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
                    Exit Sub
                Else
                    flxSCrPoA.CellBackColor = vbWhite
                    'you need to now lock it for your screen because other person has released it
                     adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                   "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID=" & flxSCrPoA.TextMatrix(flxSCrPoA.row, 10) & ""
                End If
           End If
           rsLockCheck.Close
           Set rsLockCheck = Nothing
            'Resolved by BOSL
            'Issue 470
            'Modified by Anol 08 Sep 2014
            If Not editexception Then 'this variable will disable the validation in case you shift+ click
                If Not baBankRecon(flxSCrPoA.row) Then
                     'MsgBox "This Payment is Bank reconciled and cannot be edited.", vbInformation + vbOKOnly, "Edit Payment"
                   
                   ' Exit Sub
                   frmPaymentEdit.BoolReconciled = True
                Else
                   frmPaymentEdit.BoolReconciled = False
                End If
            End If
            'End of modification
           If InStr(flxSCrPoA.TextMatrix(flxSCrPoA.row, 1), "Purchase Payment") = 0 Then
              MsgBox "Please select a payment to edit.", vbInformation, "Warning"
              Exit Sub
           End If
        
           If Val(flxSCrPoA.TextMatrix(flxSCrPoA.row, 7)) <> Val(flxSCrPoA.TextMatrix(flxSCrPoA.row, 8)) Then
              MsgBox "You can't edit a part allocated payment. Please unallocate fully before edit.", vbInformation, "Warning"
              Exit Sub
           End If
        
           Load frmPaymentEdit
           With frmPaymentEdit
              .Caption = .Caption & " - " & flxSCrPoA.TextMatrix(flxSCrPoA.row, 0)
              .txtSPSupplier.text = txtSPSupplier.Tag
              .txtDate.text = flxSCrPoA.TextMatrix(flxSCrPoA.row, 4)
              .cboBC.ListIndex = FindComboIndex(.cboBC, flxSCrPoA.TextMatrix(flxSCrPoA.row, 15), 0)
              .txtReference.text = flxSCrPoA.TextMatrix(flxSCrPoA.row, 16)
              .cmbSPAmtType.ListIndex = FindComboIndex(.cmbSPAmtType, flxSCrPoA.TextMatrix(flxSCrPoA.row, 17), 0)
              .cmbFund.ListIndex = FindComboIndex(.cmbFund, flxSCrPoA.TextMatrix(flxSCrPoA.row, 14), 0)
              .txtDetails.text = flxSCrPoA.TextMatrix(flxSCrPoA.row, 6)
              .txtAmount.text = Format(flxSCrPoA.TextMatrix(flxSCrPoA.row, 7), "0.00")
               .TransactionID = frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 10)
              .InvoiceNO = frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 0)
              'added by anol on 25 Aug 2015 issue 571 note 1148
              If flxSCrPoA.TextMatrix(flxSCrPoA.row, 18) <> "" Then
                '.cboClient.ListIndex = FindComboIndex(.cboClient, flxSCrPoA.TextMatrix(flxSCrPoA.row, 18), 0)
                .txtClient.text = flxSCrPoA.TextMatrix(flxSCrPoA.row, 18)
              End If
              .LoadFlxPaymentSplit adoConn
           End With
           
           frmPaymentEdit.Left = 100
           frmPaymentEdit.Top = 100
           Me.Enabled = False
           
           frmPaymentEdit.Show
    End If
    adoConn.Close
End Sub



Private Sub cmdGridUnitLookup_Click(Index As Integer)
   fraEditDemand.Enabled = True
   fraTab0.Enabled = True

   If Index = 3 Then
        Call cmdClose_Click(0)
   End If
   If Index = 2 Then
      tabPurExp.Enabled = True
      picAccList.Visible = False
      Exit Sub
   End If
   If Index = 1 Then
      tabPurExp.Enabled = True
      picAccounts.Visible = False
      Exit Sub
   End If
   
   fraList.Visible = False

   tabPurExp.Enabled = True
   tabPayment.Enabled = True
   If Index <> 0 Then
        fraList.Height = 2565
   End If
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Added by Anol 25 Mar 2015
    If Index = 0 Then
    
        cmdACList(0).Enabled = True
        If cmdACList(0).Enabled = True Then
            cmdACList(0).SetFocus
        End If
    End If
End Sub

Private Sub cmdJobNo_Click(Index As Integer)
   'Resolved by BOSL
   'Issue 474 Note 9
   'Modified by Anol 28 Sep 2014
    chkShowBal.Visible = False
   If txtProperty.text = "" Then
      ShowMsgInTaskBar "Please select a property to view the list of jobs", "Y", "Y"
      Exit Sub
   End If
   'End of modification
   MousePointer = vbHourglass
   sTextBox = "job"
   LoadJobSheet ""

   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtJobNo.Left + 100
   fraList.Top = txtJobNo.Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   
   MousePointer = vbDefault
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
   'resolved by BOSL
   'Issue 453 Note 6
   'Modified by Anol 24 Sep 2014
   txtJobNo.Locked = False
   'End If
End Sub

Private Sub LoadJobSheet(Filter As String)
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPurExp.Tab).Width = 1400
   lblSearch0(tabPurExp.Tab).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(tabPurExp.Tab).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1460
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   Dim rRow As Integer, szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Job No."
   lblSearch1.Caption = "Job Name"
   lblSearch2.Visible = False

   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'   szSQL = "SELECT ID, Job_DiaryName " & _
'           "FROM   PropertyMaintHistory " & _
'           "WHERE  AssignedTo = '" & txtAc(0).text & "' " & _
'           "ORDER BY ID;"

'10/12/2013
'Salia: any invoice will be connected with any job. any job (supplier/internal) can be attached.
  'Resolved by BOSL
'issue 474 note 9
'Modified by anol 28 Sep 2014
   szSQL = "SELECT ID, Job_DiaryName " & _
           "FROM   PropertyMaintHistory " & _
           "WHERE  RecordType = 'J' AND PropertyMaintHistory.PropertyID='" & txtProperty.text & "'" & _
           "ORDER BY ID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        rstRec.Filter = Filter
   End If
   flxSupplier(0).Rows = rstRec.RecordCount + 2
   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0
        
         rRow = 1
         flxSupplier(0).TextMatrix(rRow, 0) = ""
         flxSupplier(0).TextMatrix(rRow, 1) = ""
         'flxSupplier(0).AddItem ""
         flxSupplier(0).RowHeight(rRow) = 240
       rRow = 2
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!Id
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!Job_DiaryName), "", rstRec!Job_DiaryName)
         'flxSupplier(0).RowHeight(rRow) = 240
         rstRec.MoveNext
         'If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdManualDmdCancel_Click()
   fraInvCrChoice.Visible = False
   cmdNew(tabPurExp.Tab).Enabled = True
   tabPurExp.Enabled = True
End Sub

Private Sub cmdManualDmdOk_Click()
    txtDate.Enabled = True
    chkIsMgtFee.Enabled = True
    cmdClientSerc.Enabled = True
    cmdTypeList.Enabled = True
    cmdACList(0).Enabled = True
    cmbSC.Enabled = True
    txtDueDate.Enabled = True
    cmdSavePI.Visible = True

   fraLay(0).Left = 60
   fraLay(0).Top = 360
   
   fraEditDemand.Enabled = False
   fraTab0.Enabled = False
   
    fraSearch.Visible = False
   'added by anol 30 Aug 2015
    fraLay(1).Caption = ""
   fraInvCrChoice.Visible = False
   bAdjustment = False

   bEditMode = False                      'Adding new PI

   tabPurExp.Enabled = True

   If Len(txtDate.text) < 10 Then
      txtDate.text = Format(Date, "dd/mm/yyyy")
      lblPostingDate.ToolTipText = txtDate.text
   End If
   '28 July 2015
    cmdACList(0).Enabled = True
    cmdaddnewline.Enabled = True
    txtReference.text = ""
    txtNC(0).text = ""
    txtUnit(0).text = ""
    txtNCName.text = ""
    txtDept(0).text = ""
    txtPFName.text = ""
    txtJobNo.text = ""
    txtSchedules.text = ""
    txtDetails_(0).text = ""
    txtNet_(0).text = ""
    txtVat_(0).text = ""
    txtTotal.text = ""
    'added by anol 09 July 2015
      'Split line Enable mode
      cmdUpdate(1).Enabled = False
      txtUnit(0).Enabled = False
      cmdUnitList.Enabled = False
      txtJobNo.Enabled = False
      cmdJobNo(0).Enabled = False
      txtNet_(0).Enabled = False
      txtNC(0).Enabled = False
      
      cmdNCList.Enabled = False
      cmdUpdate(1).Enabled = False
      cmdTaxList(0).Enabled = False
      cmdDeptList.Enabled = False
      txtDetails_(0).Enabled = False
      txtVat_(0).Enabled = False
      cmdSchedules(0).Enabled = False
      cmdUpdate(2).Enabled = False
      'End of addition
   If optManualInv.Value Then
      sAddChoice = "IN"
      txtTransType.text = "Invoice"
   End If
   If optManualCrNote.Value Then
      sAddChoice = "CN"
      txtTransType.text = "Credit Note"
   End If
   If optManualAdjInv.Value Then
      sAddChoice = "AI"
      txtTransType.text = "Adjustment Invoice"
   End If
   If optManualAdjCrNote.Value Then
      sAddChoice = "AC"
      txtTransType.text = "Adjustment Credit Note"
   End If

   If sAddChoice = "AI" Or sAddChoice = "AC" Then
      bAdjustment = True
   Else
      bAdjustment = False
   End If

   HandleCommandButton "Add Invoice"
   'cmdACList(0).SetFocus added by anol 20161014
    cmdACList0Focus
    Dim strTemp As String
    txtClientID.ForeColor = vbBlack
    strTemp = isControlAccountSet(txtClientID.text)
    If Len(strTemp) > 0 Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        strTemp = ""
        'picClient.Visible = False
        txtClientID.ForeColor = vbRed
        Exit Sub
    End If
End Sub
Private Sub cmdACList0Focus()
    On Error GoTo Err
    cmdACList(0).SetFocus
    Exit Sub
Err:
   
End Sub
Private Sub LoadNominalCodeBk()
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 800
   flxSupplier(0).ColWidth(1) = 2300
   flxSupplier(0).ColAlignment = vbLeftJustify

   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + 800
   
   txtSearch1.Width = 700
   txtSearch1.Left = 40
   
   txtSearch2.Width = 2400
   txtSearch2.Left = txtSearch1.Left + 800
   
   lblSearch0(0).Caption = "Code"
   lblSearch1.Caption = "Name"
   lblSearch2.Visible = False

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As New ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT NominalLedger.* FROM NominalLedger ORDER BY CODE;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim i As Integer
   i = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(i, 0) = adoRST.Fields.Item("Code").Value
      flxSupplier(0).TextMatrix(i, 1) = adoRST.Fields.Item("Name").Value
      flxSupplier(0).AddItem ""
      i = i + 1
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

Private Sub cmdNCList_Click()
   chkShowBal.Visible = False
   fraList.Height = 4325
   sTextBox = "NC"
   LoadNominalCode "" 'Main function to load the values
'   flxSupplier(0).Height = fraList.Height - 750
   tabPurExp.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 60
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtNC(0).Left + 100
   fraList.Top = txtNC(0).Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
  
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
 
End Sub

Private Sub LoadNominalCode(Filter As String)
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPurExp.Tab).Width = 1400
   lblSearch0(tabPurExp.Tab).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(tabPurExp.Tab).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1460
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(tabPurExp.Tab).Caption = "N/C"
   lblSearch1.Caption = "Name"
   lblSearch2.Visible = False

' Error Handler
'   On Error GoTo Error_Handler

   Dim adoConn As New ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim rsShoppingCentre As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString
   
   If frmMMain.IsRibbonVersion Then
'      szSQL = "SELECT N.* " & _
'              "FROM NominalLedger AS N " & _
'              "WHERE N.ClientID = '" & cboClientPI.Value & "' AND " & _
'                    "Posting AND (ISNULL(CAType) OR CAType='') " & _
'              "ORDER BY N.Code;"
'Fixed by anol 27 july 2015
'Here implement Restricted to Budget issue 889 anol 2020-11-19
        rsShoppingCentre.Open "Select isRestrictedtoBudget from shoppingCentre", adoConn, adOpenStatic, adLockReadOnly
        If rsShoppingCentre("isRestrictedtoBudget").Value = False Then
                szSQL = "SELECT N.* " & _
                 "FROM NominalLedger AS N " & _
                 "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
                 "Posting AND CAFixed=0 AND CODE NOT IN " & _
                 "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
                 " ORDER BY N.Code;"
         Else   'Here go to the service charge budget  and get only budgeted nominal code
         'previous SQL  before 2020-11-19 when it was not loadind fund category wise (AND restricted to budget was not implemented)
'                szSQL = "SELECT N.* " & _
'                 "FROM NominalLedger AS N " & _
'                 "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
'                 "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
'                 "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
'                 " ORDER BY N.Code;"
               If iSelectedFundCategoryID = 2 Then 'When you select a fund and that is only  'Service Charge' Category then load fundcode from the budgeted funds
'                    szSQL = "SELECT N.* From GlobalSC G,GlobalSCDtls D,Fund F,NominalLedger N,Property P where P.CBY=G.FinancialYear AND N.ClientID=P.ClientID AND G.BudgetID=D.BudgetID AND D.NC=N.Code " & _
'                           "AND F.FundID=G.fund AND N.ClientID='" & txtClientID.text & "' AND G.PropertyID='" & txtProperty.text & "' AND F.CategoryCode=" & iSelectedFundCategoryID & " AND " & _
'                           "N.Posting AND (ISNULL(N.CAType) OR N.CAType='') AND N.CODE NOT IN " & _
'                           "(SELECT NominalCode FROM tlbClientBanks  where ClientID = '" & txtClientID.text & "')" & _
'                           " ORDER BY N.Code;"
'
                          szSQL = " Select * from ( SELECT N.*,G.FinancialYear,G.PropertyID   From GlobalSC G,GlobalSCDtls" & _
                        " D,Fund F,NominalLedger N where  G.BudgetID=D.BudgetID AND D.NC=N.Code AND F.FundID=G.fund AND" & _
                        " N.ClientID='" & txtClientID.text & "' AND G.PropertyID='" & txtProperty.text & "' AND F.CategoryCode=2 AND N.Posting AND CAFixed=0 " & _
                        " AND N.CODE NOT IN (SELECT NominalCode FROM tlbClientBanks  where ClientID" & _
                        " = '" & txtClientID.text & "') ORDER BY N.Code) K  INNER JOIN Property P" & _
                        " ON K.FinancialYear=P.CBY AND P.ClientID= K.clientID AND P.PropertyID=K.PropertyID; "


               Else     'Load normal nominal codes  CAFixed=0 means (in option control account) set allow posting= Yes then show here as NC
                    szSQL = "SELECT N.* " & _
                            "FROM NominalLedger AS N " & _
                            "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
                            "Posting AND CAFixed=0 AND CODE NOT IN " & _
                            "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
                            " ORDER BY N.Code;"
               End If
         End If
         rsShoppingCentre.Close
         Set rsShoppingCentre = Nothing
   Else
    
      szSQL = "SELECT N.* " & _
              "FROM NominalLedger AS N " & _
              "ORDER BY N.Code;"
   End If

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        adoRST.Filter = Filter
   End If
   Dim iRows As Integer
   flxSupplier(0).Clear
   flxSupplier(0).Rows = adoRST.RecordCount + 2
   flxSupplier(0).Cols = 3
   iRows = 1
   While Not adoRST.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRST.Fields.Item("Code").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRST.Fields.Item("Name").Value
      'If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRST.MoveNext
   Wend
  
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   flxSupplier(0).RowHeight(0) = 0

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdOClientList_Click()
    sTextBox = "PIHistory"
    tabPurExp.Enabled = False
    tabPayment.Enabled = False
    chkShowBal.Visible = False
    picClient.Left = 405
    picClient.Top = 800
    picClient.Visible = True
    LoadflxClient ""
    txtSearchClientID.SetFocus
End Sub



Private Sub cmbPropertyHistory_Click()
   If Not bFormLoaded Then Exit Sub
   SortTheGrid flxPurchHistory, txtClientIdlist, cmbPropertyHistory, txtSupplierSearc
   flxPurchHistorySplit.Clear
End Sub

Private Sub AllocDiscard()
   Dim iRow As Integer

   For iRow = 1 To flxSPayment.Rows - 1
'      If flxSPayment.TextMatrix(iRow, 15) = "A" Then
         flxSPayment.TextMatrix(iRow, 10) = "0.00"
         flxSPayment.TextMatrix(iRow, 15) = ""
         flxSPayment.TextMatrix(iRow, 16) = ""
'      End If
   Next iRow

   For iRow = 1 To flxSCrPoA.Rows - 1
      flxSCrPoA.TextMatrix(iRow, 9) = "0.00"
   Next iRow

   flxSPayment.Enabled = True
   flxSCrPoA.Enabled = True
   cmdPayAllocateSave.Enabled = False
   lblAllocating(1).Visible = False
   Frame5(5).Enabled = True                     'Receipt - Saving
   Frame5(1).Enabled = True                    'Allocation - Saving
   txtAllocatedDiff(1).text = "0.00"
   cmdPayAutomatic.Enabled = True
End Sub

Private Sub cmdOpenClient_Click()
    sTextBox = "Payment"
    chkShowBal.Visible = False
    tabPurExp.Enabled = False
    tabPayment.Enabled = False
    picClient.Left = 405
    picClient.Top = 1255
    
    txtSearchClientName.text = ""
    txtSearchClientID.text = ""
    picClient.Visible = True
    fmeLoading.Visible = True
    fmeLoading.Refresh
    LoadflxClient ""
    fmeLoading.Visible = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then
        MsgBox "Please select file name from combo", vbInformation, "You need to select a file name"
        Exit Sub
   End If
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
   Frame17.Visible = False
End Sub

Private Sub cmdOpenFileView_Click()
    If cmbFiles.text = "" Then
        MsgBox "Please select file name from combo", vbInformation, "You need to select a file name"
        Exit Sub
   End If
   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
   Frame17.Visible = False
   'cmdOpenFileView.Visible = False
End Sub

Private Sub cmdPayAllocationDiscard_Click()
   Dim iRow As Integer

   If MsgBox("Are you sure to discard allocation?", vbQuestion + vbYesNo, "Clear Allocation") = vbNo Then Exit Sub

   AllocDiscard
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    tabPurExp.Enabled = True
    tabPayment.Enabled = True
End Sub

Private Sub cmdPO_Click()
   If flxPurchase.TextMatrix(iPIEdit, 17) = "" Then Exit Sub
   If IsLoadedAndVisible("frmPO_Amend") Then
      ShowMsgInTaskBar "Purchase Order form is already open", "Y", "N"
      Exit Sub
   End If

   Load frmPO_Amend

   With frmPO_Amend
       Dim adoConn As New ADODB.Connection
       adoConn.Open getConnectionString
       Dim rsPO As New ADODB.Recordset
       rsPO.Open "Select slNumber from tblPurInv where MY_ID='" & flxPurchase.TextMatrix(iPIEdit, 17) & "'", adoConn, adOpenStatic, adLockReadOnly
       If rsPO.EOF Then
            rsPO.Close
            Set rsPO = Nothing
            adoConn.Close
            Set adoConn = Nothing
            Exit Sub
       End If
       
        
       
      
     ' .Caption = "View Purchase Order - PO" & flxPurchase.TextMatrix(iPIEdit, 17) 'Slnumber which has transaction type 25 Can be derived from PO column of current record
      .Caption = "View Purchase Order - PO" & rsPO("slnumber").Value  'Slnumber which has transaction type 25 Can be derived from PO column of current record
       rsPO.Close
       Set rsPO = Nothing
       adoConn.Close
       Set adoConn = Nothing
      .szPO = flxPurchase.TextMatrix(iPIEdit, 0) ' I do net see any use of this variable and current row selection is wrong
      
'      .LoadDate
      .bEditMode = True
      .szCallerForm = "I"
      .sPI = flxPurchase.TextMatrix(iPIEdit, 17) 'the one has transaction Type 6

      .Show
      .txtInv(0).SetFocus
   End With
   'If iPIEdit = 0 Then Exit Sub
   With flxPurchase
'      frmPO_Amend.txtTransType.text = .TextMatrix(iPIEdit, 3)
      frmPO_Amend.txtAc(0).text = .TextMatrix(iPIEdit, 5)
      frmPO_Amend.cmdACList(0).Enabled = IIf(.TextMatrix(iPIEdit, 9) <> .TextMatrix(iPIEdit, 12), False, True)
      frmPO_Amend.txtDate.text = .TextMatrix(iPIEdit, 4)
      frmPO_Amend.txtDueDate.text = Format(.TextMatrix(iPIEdit, 13), "dd/mm/yyyy")
      frmPO_Amend.txtSupplierName.text = .TextMatrix(iPIEdit, 6)
      frmPO_Amend.txtInv(0).text = .TextMatrix(iPIEdit, 7)
      frmPO_Amend.txtClientID.text = .TextMatrix(iPIEdit, 14)
      frmPO_Amend.txtProperty.text = .TextMatrix(iPIEdit, 11)

      LoadSplit4Edit frmPO_Amend.flxPI, 17

      frmPO_Amend.cmdNew(0).Enabled = False
      frmPO_Amend.cmdUpdate(1).Enabled = True
      frmPO_Amend.cmdSavePI.Visible = False
      frmPO_Amend.cmdEdit.Visible = False
      frmPO_Amend.cmdDelete.Visible = False
      frmPO_Amend.cmdCancel(0).Visible = False
   End With
   frmPO_Amend.UpdateTotalPICN
   Me.Enabled = False
End Sub

Private Sub cmdPostDemands_Click()
   Dim szTemp As String

'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + Me.Top + tabPurExp.Top + fraEditDemand.Top + cmdPostDemands.Top + 1140
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + fraEditDemand.Left + tabPurExp.Left + cmdPostDemands.Left + 80
   frmPopUpMenu.CallingFrom "PostPurchase"

   szTemp = SelectedPurInvID()

   If szTemp <> "" Then
      frmPopUpMenu.optSelPI.Value = True
   Else
      frmPopUpMenu.optSelPI.Value = False
      frmPopUpMenu.optSelPI.Enabled = False
   End If
   chkProperty.Value = 0
   frmPopUpMenu.Show 1
End Sub

Private Sub cmdPrintListHistory_Click()
   sTextBox = "PurchaseTransactionHistoryReport"
   
   picPurchaseHistory.Top = 6480
   picPurchaseHistory.Left = 1485
   Label6(10).Caption = "Purchase History Print"
   picPurchaseHistory.Visible = True
   FocusControl txtStartDate
   txtStartDate.text = "01/01/2000"
   txtEndDate.text = Format(Date, "dd/mm/yyyy")
End Sub
Private Function ReturnStringPP_ID(i As Long, j As Long) As String
    Dim ReturnString As String
    On Error GoTo Err
    For j = i To j
        If flxPurchPPHistory.TextMatrix(j, 0) = "" Then
                GoTo NextI
        End If
        If flxPurchPPHistory.RowHeight(j) > 0 Then
            ReturnString = ReturnString & flxPurchPPHistory.TextMatrix(j, 0)
            ReturnString = ReturnString & ","
        End If
NextI:
    Next j
Err:
    If Len(ReturnString) > 0 Then
        ReturnStringPP_ID = Left(ReturnString, Len(ReturnString) - 1)
    End If
    
End Function
Private Function ReturnStringPI_ID(i As Long, j As Long) As String
    On Error GoTo Err
    For j = i To j
        If flxPurchHistory.TextMatrix(j, 0) = "" Then
                GoTo NextI
        End If
        If flxPurchHistory.RowHeight(j) > 0 Then
            ReturnStringPI_ID = ReturnStringPI_ID & "'" & flxPurchHistory.TextMatrix(j, 0) & "'"
            ReturnStringPI_ID = ReturnStringPI_ID & ","
        End If
NextI:
    Next j
Err:
    If Len(ReturnStringPI_ID) > 0 Then
        ReturnStringPI_ID = Left(ReturnStringPI_ID, Len(ReturnStringPI_ID) - 1)
    End If
    
End Function
Private Sub cmdPrintPI_List_Click()
   sTextBox = "PurchaseTransactionReport"
   picPurchaseHistory.Top = 6480
   picPurchaseHistory.Left = 1485
   Label6(10).Caption = "Purchase Transaction Report"
   picPurchaseHistory.Visible = True
   FocusControl txtStartDate
   txtStartDate.text = "01/01/2000"
   txtEndDate.text = Format(Date, "dd/mm/yyyy")
   

End Sub

Private Sub cmdRevHistory_Click()
'Added by anol 10 Jun 2015
'issue 0000572: Purchases and expenses - Reverse postings to history not present
   Dim szSQL As String
   Dim iRow As Integer, szPurID As String
   Dim j As Long
   Dim K As Integer
   Dim SelPurHisID() As String
   Dim adoConn As New ADODB.Connection
   'in a grid if there is hundreds of records it cannot process with one SQL Need to seperate them or I can use an array to buld up Update query one Batch can hold upto 50
   'invoices
   'On Error GoTo Catch_Error
   Call SelPurHistory(SelPurHisID())
   j = UBound(SelPurHisID())
   If j = 0 Then
        MsgBox "Please select a purchase invoice to reverese.", vbCritical + vbOKOnly, "Purchase invoice"
        Exit Sub
   End If
   If MsgBox("Are you sure you wish to reverse the selected transactions from history?", vbQuestion + vbYesNo, "Purchase Invoice History") = vbNo Then Exit Sub
  

   adoConn.Open getConnectionString
   K = CInt(j / 50)
   
   If K = j / 50 Then
        'No no need to do ceiling, this is fully divisible
        K = j / 50
   Else
        K = CInt(j / 50) + 1 'This is ceiling function
   End If
   For K = 0 To K - 1
           
           szPurID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurHisID())
           If szPurID = "" Then
                Exit For
           End If
       If Trim(szPurID) <> "" Then
           szSQL = "UPDATE tblPurInv " & _
           "SET    History = FALSE " & _
           "WHERE  My_ID IN (" & szPurID & ") ; "
            adoConn.Execute szSQL
       End If
   Next

'   szSQL = "UPDATE tblPurInv " & _
'           "SET    History = FALSE " & _
'           "WHERE  SlNumber IN (" & SelPurHistory & ") ; "
'    szSQL = "UPDATE tblPurInv " & _
'           "SET    History = FALSE " & _
'           "WHERE  My_ID IN (" & szPurID & ") ; "
           
'           AND " & _
'                  "  DemandID NOT IN (" & _
'                  "     SELECT DemandRef " & _
'                  "     From tlbReceipt " & _
'                  "     WHERE Type = 1 AND Amount > OSAmount);"
'Debug.Print szSQL
'   adoConn.Execute szSQL

   LoadFlxPurchHistory adoConn, ""
   LoadFlxPurchase adoConn
   fmeLoading.Visible = False
   adoConn.Close
   Set adoConn = Nothing
   chkAllPurchaseHistory.Value = 0
   ShowMsgInTaskBar "System has reversed " & j & " invoices.", "Y", "P"
   Exit Sub

Catch_Error:
   MsgBox "Select a purchase invoice to reverse.", vbCritical + vbOKOnly, "Purchase invoice"
End Sub

 Private Function ReturnString(i As Long, j As Long, ByRef SelPurHisID() As String) As String
    On Error GoTo Err
    For j = i To j
        If SelPurHisID(j) = "" Then
                Exit For
        End If
        ReturnString = ReturnString & SelPurHisID(j)
        ReturnString = ReturnString & ","
    Next j
Err:
    If Len(ReturnString) > 0 Then
        ReturnString = Left(ReturnString, Len(ReturnString) - 1)
    End If
 End Function
Private Sub SelPurHistory(ByRef SelPurHisID() As String)
   Dim i As Integer
   Dim j As Integer
   For i = 1 To flxPurchHistory.Rows - 1
      If flxPurchHistory.TextMatrix(i, 1) = "X" Then
'         SelPurHistory = SelPurHistory & "'" & CStr(flxPurchHistory.TextMatrix(i, 0)) & "'"
'         SelPurHistory = SelPurHistory & ","
         j = j + 1
      End If
   Next i
   ReDim SelPurHisID(j)
   j = 0
   For i = 1 To flxPurchHistory.Rows - 1
      If flxPurchHistory.TextMatrix(i, 1) = "X" Then
         SelPurHisID(j) = "'" & CStr(flxPurchHistory.TextMatrix(i, 0)) & "'"
         j = j + 1
      End If
   Next i
'   If Len(SelPurHistory) > 0 Then
'        SelPurHistory = Left(SelPurHistory, Len(SelPurHistory) - 1)
'   End If
End Sub
Private Sub cmdSavePI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If txtTotal.text <> "" Then
'      ShowMsgInTaskBar "Please update the invoice line.", "Y", "N"
'      cmdUpdate(1).SetFocus
'   End If
End Sub

Private Sub cmdSchedules_Click(Index As Integer)
   chkShowBal.Visible = False
   sTextBox = "Schedules"
   LoadSchedules ""

   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtSchedules.Left + 100
   fraList.Top = txtSchedules.Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
  
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
End Sub

Private Sub LoadSchedules(Filter As String)
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify
   
   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPurExp.Tab).Width = 1400
   lblSearch0(tabPurExp.Tab).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(tabPurExp.Tab).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1.Width = 1460
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)
   
   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset
   
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(tabPurExp.Tab).Caption = "Schedule ID"
      lblSearch1.Caption = "Schedule Name"
      lblSearch2.Visible = False
      
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   szSQL = "SELECT ScheduleID, ScheduleName " & _
           "FROM Schedule " & _
           "ORDER BY ScheduleID;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
        rstRec.Filter = Filter
   End If
   flxSupplier(0).Rows = rstRec.RecordCount + 2
   If Not rstRec.EOF Then


      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0
   
        rRow = 1
        flxSupplier(0).TextMatrix(rRow, 0) = ""
        flxSupplier(0).TextMatrix(rRow, 1) = ""
        'flxSupplier(0).AddItem ""
        flxSupplier(0).RowHeight(rRow) = 240
        rRow = 2
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!ScheduleID
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!ScheduleName), "", rstRec!ScheduleName)
         flxSupplier(0).RowHeight(rRow) = 240
         rstRec.MoveNext
        ' If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close
End Sub

Private Function DateValidation2() As Boolean
        Dim iRow As Integer
        For iRow = 1 To flxSPayment.Rows - 1
            If (flxSPayment.TextMatrix(iRow, 0) = "+" Or _
                    flxSPayment.TextMatrix(iRow, 0) = ">") And _
                    Val(flxSPayment.TextMatrix(iRow, 10)) <> 0 Then
                        If (flxSPayment.TextMatrix(iRow, 30) = True Or flxSPayment.TextMatrix(iRow, 31) = True) Then
                            If DateDiff("d", flxSPayment.TextMatrix(iRow, 5), txtSPDate.text) <> 0 Then
                                Exit Function
                            End If
                        End If
            End If
        Next iRow
        DateValidation2 = True
End Function
Private Sub cmdSPSave_Click()
   Dim tSageDept As Byte, lDemandID As Long
   Dim iRow As Integer
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rsBankCheck As New ADODB.Recordset
   If txtClientIDPurPay.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientIDPurPay.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
   If txtClientIDPurPay.text = "" Then
      MsgBox "Please select a client", vbInformation, "Warning"
      FocusControl txtClientIDPurPay
      Exit Sub
   End If
   If txtSPSupplier.text = "" Then
      MsgBox "Please select a Supplier", vbInformation, "Warning"
      FocusControl cmdSPSupplier
      Exit Sub
   End If
    
   If txtBankAc.text = "" Then
      MsgBox "Please select a bank account.", vbInformation, "Warning"
      FocusControl cmdBankAc
      Exit Sub
   End If

   If Not bChangesMade Then
      MsgBox "There is no transaction to save.", vbInformation, "Warning"
      Exit Sub
   End If
    If Not DateValidation2 Then
        If MsgBox("The payment date entered is different from the due date of the selected invoice(s). Do you wish to continue? ", vbYesNo + vbDefaultButton2, "Please confirm") = vbYes Then
        Else
            txtSPDate.text = flxSPayment.TextMatrix(1, 5)
            FocusControl txtSPDate
            Exit Sub
        End If
   End If
   If Val(txtSPaymentTotal.text) < 0 Then
      If MsgBox("Would you like to book a payment refund?", vbQuestion + vbYesNo, "Refund") = vbNo Then
         txtSPaymentTotal.text = "0.00"
         Exit Sub
      Else
         Load frmFund4RptPay
         frmFund4RptPay.szCallerForm = "PI"
         frmFund4RptPay.Show 1
         If lFund = -1 Then
            ShowMsgInTaskBar "You can not book a payment refund without a fund.", , "N"
            SelTxtInCtrl txtSPaymentTotal
            txtSPaymentTotal.SetFocus
            Exit Sub
         End If
      End If
   End If

   If txtPayAmtType.text = "" Then
      MsgBox "Please select an Amount Type.", vbInformation, "Warning"
      cmdAmountType.SetFocus
      Exit Sub
   End If

   If txtSPDate.text = "" Then
      MsgBox "Please enter the Date.", vbInformation, "Warning"
      txtSPDate.SetFocus
      Exit Sub
   Else
      If lblPayPostingDate.ToolTipText = "" Then lblPayPostingDate.ToolTipText = txtSPDate.text
   End If

   If txtSPReference.text = "" Then
      If MsgBox("Do you want to save transaction without reference?", vbQuestion + vbYesNo, "Save") = vbNo Then
         txtSPReference.SetFocus
         Exit Sub
      End If
   End If

    If txtBankAc.text = "" Then
        MsgBox "Please select the bank account", vbInformation, "select the bank account"
        cmdBankAc.Enabled = True
        cmdBankAc.SetFocus
        Exit Sub
    End If
    If txtClientIDPurPay = "" Then
        MsgBox "Please select the client", vbInformation, "select the client"
        txtClientIDPurPay.SetFocus
        Exit Sub
    End If
    'issue 519
    'it should not be possible to change the posting date on a transaction
     'if the posting date on that transaction falls within a closed financial period
    adoConn.Open getConnectionString
    'SELECT * FROM tblpurinv AS T
    'INNER JOIN rentsummarystatement AS R
    'ON (T.inv_no = 'SS' & R.StatementID)
    
   'Here is the logic if you are selecting wrong bank account in which differes from Rentpayable bank account.
   'We should check from the client statement , the Bank code selected for this Rent Payable an d compare with selected bank account for paying current PI
   'written by anol 20230728
    For iRow = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(iRow, 30) = "" Then 'isrentpayable flag is  empty then do not check anything
        Else
             If cmdPayAllocate.Caption = "All&ocation Only" Then
                      If flxSPayment.TextMatrix(iRow, 30) = True And flxSPayment.TextMatrix(iRow, 14) = "6" Then
                          szSQL = "SELECT R.BankCode FROM tblpurinv AS T INNER JOIN rentsummarystatement AS R ON " & _
                          "(T.inv_no = 'SS'& R.StatementID) WHERE MY_ID='" & flxSPayment.TextMatrix(iRow, 12) & "'"
                          rsBankCheck.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                          If RecordCount(rsBankCheck) > 0 Then
                              If rsBankCheck("BankCode").Value <> txtBankCode.text Then
                                  rsBankCheck.Close
                                  adoConn.Close
                                   MsgBox "Please select a bank account that matches the bank account on the client statement used to generate this rent payable invoice", vbInformation, "Warning"
                                  Exit Sub
                              End If
                          End If
                     End If
              End If
        End If
    Next
    

    If IsPeriodStatus(lblPayPostingDate.ToolTipText, txtClientIDPurPay.text, adoConn) = 0 Then
        MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
        adoConn.Close
        Exit Sub
    ElseIf IsPeriodStatus(lblPayPostingDate.ToolTipText, txtClientIDPurPay.text, adoConn) = 9 Then
        MsgBox "The Posting date does not fall in any existing financial period", vbInformation, "Warning"
        adoConn.Close
        Exit Sub
    End If
    adoConn.Close
'Resolved by BOSL
'issue 525
'Overdraft implementation on purchase payment, date: 09 Apr 2015, by anol
   If Val(txtSPaymentTotal.text) > 0 Then
      GetBankAccountBalance
       'Dim dblBankBanlance As Double
       
       adoConn.Open getConnectionString
       'following line is the old bank balance determination now this has been modified with retention value added plus
       'lblBankOD(0).Caption = BankAccBalance(adoconn, txtBankCode.text, txtClientIDPurPay.text)
       '# this checking added by anol on 20160912 Mon
       If IsDate(lblPayPostingDate.ToolTipText) = True Then
              'I am using txtClientIDPurPay.text ,because this contain the client ID 20160927
              If IsNull(txtClientIDPurPay.text) Or txtClientIDPurPay.text = "" Then
                  MsgBox "Please select a client", "Y", "N"
                  cmdOpenClient.SetFocus
                  Exit Sub
              End If
              
              If IsPeriodStatus(lblPayPostingDate.ToolTipText, txtClientIDPurPay.text, adoConn) = 0 Then
                  MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
                  adoConn.Close
                  Exit Sub
              ElseIf IsPeriodStatus(lblPayPostingDate.ToolTipText, txtClientIDPurPay.text, adoConn) = 9 Then
                  MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
                  adoConn.Close
                  Exit Sub
              End If
              If DateDiff("d", lblPayPostingDate.ToolTipText, txtSPDate.text) > 0 Then
                    MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
                    Exit Sub
              End If
      End If
      
       '#
       adoConn.Close
      If Val(lblBankOD(0).Caption) - Val(txtSPaymentTotal.text) < 0 Then 'Account balance-Current Amount
         If lblBankOD(1).Caption = "True" Then 'OverDraftAllowed?
            If Val(lblBankOD(2).Caption) > 0 Then 'OverDraftAmount
               If (Val(lblBankOD(0).Caption) - Val(txtSPaymentTotal.text)) * (-1) > Val(lblBankOD(2).Caption) Then 'BankBalance - currentamount * (-1) > OverDraftAmount
               '-22+25=3 > 25
                  If MsgBox("This Bank Account is over its overdraft limit. Do you wish to continue?", vbQuestion + vbYesNo, "Bank Overdrawn") = vbNo Then Exit Sub
               End If
            End If
         Else
            MsgBox "This Bank Account cannot go overdrawn.", vbInformation + vbOKOnly, "Bank Overdraft"
            Exit Sub
         End If
      End If
   End If

   If Val(txtSPaymentTotal.text) > cGridSPTotal Then
      If MsgBox("There is an unallocated payment of £" & _
                 Format(Val(txtSPaymentTotal.text) - cGridSPTotal, "0.00") & "." & Chr(10) & Chr(10) & _
                "Do you wish to Post this as a Payment on Account?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
         txtSPaymentTotal.SetFocus
         SelTxtInCtrl txtSPaymentTotal
         Exit Sub
      Else
         Load frmFund4RptPay
         frmFund4RptPay.szCallerForm = "PI"
         frmFund4RptPay.Show 1
         If lFund = -1 Then
            MsgBox "You can not save a Payment without a fund.", vbInformation, "Warning"
            SelTxtInCtrl txtSPaymentTotal
            txtSPaymentTotal.SetFocus
            Exit Sub
         End If
      End If
   Else
      If Val(txtSPaymentTotal.text) > 0 Then _
         If MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save") = vbNo Then Exit Sub
   End If

   If Val(txtSPaymentTotal.text) > 0 Then
      Frame4(0).Top = 5760 'Activating option frame for remittence
      Frame4(0).Left = 840
      Frame4(0).Visible = True
'      cmdChqRemittYes.SetFocus
      FocusControl cmdChqRemittYes
   End If
   'Modified by anol 1 july 2015
   cmdEditPayment.Enabled = False
    'End of modification
   UpdateSuppBalInSearch
   If Val(txtSPaymentTotal.text) < 0 Then
        cmdPayAllocateSave.Enabled = False
        Call SavePaymentTransactions 'here is the main procedure
        Call ClearLockFromPIPayment     'Written by anol 2019 05 14
        cmdPayAllocateSave.Enabled = True
   End If
   If IsLoadedAndVisible("frmCashbook") = True Then
            frmCashbook.cboBC_Click
   End If
End Sub

Private Sub GetBankAccountBalance()
   Dim adoConn    As New ADODB.Connection
   Dim rstSet     As New ADODB.Recordset

   adoConn.Open getConnectionString

   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   szSQL = "SELECT CB.CurrentBalance AS BAL, AllowOverDraft, OverDraftLimit " & _
           "FROM tlbClientBanks AS CB " & _
           "WHERE CB.NominalCode = '" & txtBankCode.text & "' AND " & _
               "CB.CLIENT_ID = '" & txtClientIDPurPay.text & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup bank account for the client."
   Else
      lblBankOD(0).Caption = adoRST.Fields.Item("BAL").Value
      lblBankOD(1).Caption = adoRST.Fields.Item("AllowOverDraft").Value
      lblBankOD(2).Caption = IIf(IsNull(adoRST.Fields.Item("OverDraftLimit").Value), "0", adoRST.Fields.Item("OverDraftLimit").Value)
   End If
   adoRST.Close
   Dim Balance As Double
   ' find current Balance for the selected bank account and selected client ID by anol 2023-05-24
   szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) as AMTT from (" & _
            "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & txtBankCode & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & txtClientIDPurPay & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & txtBankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & txtClientIDPurPay & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & txtBankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & txtClientIDPurPay & "'   group by Type )"
                       
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
      Balance = IIf(IsNull(adoRST.Fields.Item("AMTT").Value), 0, adoRST.Fields.Item("AMTT").Value)
   End If
   adoRST.Close
                       
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where  isDeleted=false and  BankCode='" & txtBankCode & "' and ClientID='" & txtClientIDPurPay & "' "
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        Balance = Balance - IIf(IsNull(adoRST.Fields.Item("DAmt").Value), 0, adoRST.Fields.Item("DAmt").Value) 'adoRst.Fields.Item("DAmt").Value
    End If
    adoRST.Close
    lblBankOD(0).Caption = Balance

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

Private Sub UpdateSuppBalInSearch()
   Dim i As Integer

'   For i = 0 To cmbSPSupplier.ListCount - 1
'      If szaSupplierBal(0, i) = txtSPSupplier.Tag Then
'         szaSupplierBal(1, i) = Val(szaSupplierBal(1, i)) - Val(txtSPaymentTotal.text)
'         Exit For
'      End If
'   Next i
   'cmbSPSupplier.Column(2) = Format(Val(cmbSPSupplier.Column(2)) - Val(txtSPaymentTotal.text), "0.00")
   For i = 1 To flxSupplier(1).Rows - 1
        If flxSupplier(1).TextMatrix(i, 1) = txtSPSupplier.Tag Then
            flxSupplier(1).TextMatrix(i, 3) = Format(Val(flxSupplier(1).TextMatrix(i, 3)) - Val(txtSPaymentTotal.text), "0.00")
        End If
   Next
End Sub

'Save Purchase Payment
'Private Function ReturnPINo() As String
'    'This function is written by  anol 29 Apr 2015
'    'When printing payment ref no is showing PP:Date
'    'But we need to show the PI reference
'    Dim i As Integer
'    Dim strPI As String
'    For i = 1 To flxSPayment.Rows
'        If flxSPayment.Rows = 2 Then
'            strPI = strPI + "PI" + flxSPayment.TextMatrix(i, 11)
'        ElseIf flxSPayment.Rows > 2 Then
'            strPI = strPI + ", PI" + flxSPayment.TextMatrix(i, 11)
'        End If
'    Next i
'    ReturnPINo = strPI
'End Function
Private Function GetClientACBalance() As Double   'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT" & _
            " from tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND P.ClientID ='" & txtClientIDPurPay.text & "'  and SP.TYPE in ('CLIENT') AND P.PDate <=#" & Format(txtSPDate.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalance = -IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function SavePaymentTransactions() As Long
   Dim adoConn    As New ADODB.Connection
   Dim rstSet     As New ADODB.Recordset
   Dim rsSSR As New ADODB.Recordset
   Dim lSlNumber  As Long
   Dim cPoA       As Currency            'Payment on Account
   Dim szPurchaseTranID As String
   Dim rsInvoice As New ADODB.Recordset
   Dim cSumSplits As Double
   cPoA = -1

   If Val(txtSPaymentTotal.text) > cGridSPTotal Then
      cPoA = Val(txtSPaymentTotal.text) - cGridSPTotal
   End If

   adoConn.Open getConnectionString
   adoConn.BeginTrans
'  Generate the next Transaction Id from the tlbPayment
   Dim lSp_ID As Long, lSPTran_ID As Long, szSQL As String
   Dim lSPTran_ID_split As Long

   szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbPayment;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSp_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
   szSQL = "SELECT MAX(TransactionID) AS TID FROM PayTransactions;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
   
   szSQL = "SELECT MAX(TransactionID) AS TID FROM PayTransactionsSplit;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID_split = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close

   If cGridSPTotal > 0 Then              'if any payment in the grid have been made
      
      Dim iRow As Integer

      bTotalPayTyped = False

'*******************************************************  PAYMENT ******************************************
      szSQL = "SELECT * FROM tlbPayment;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      lSlNumber = SlNumber("PP", "tlbPayment", adoConn)

      rstSet.AddNew
      lSp_ID = lSp_ID + 1

      rstSet!CreatedBy = User
      rstSet!CreatedDate = Now
      rstSet!TransactionID = lSp_ID
      
      rstSet!szTransactionID = rstSet!TransactionID
      SavePaymentTransactions = lSp_ID

      rstSet!Type = 8 'Purchase Payment
      rstSet!SageAccountNumber = txtSPSupplier.Tag

      rstSet!PDate = Format(txtSPDate.text, "dd mmmm yyyy")
      rstSet!ref = "PP" & Format(Now, "yymmddhhmmss")
      rstSet!Details = "PURCHASE PAYMENT"
      rstSet!amount = CCur(cGridSPTotal)
      'Below line has been added by anol 20181030. Amount of payment has been fully allocated of type 8. so OSamount is zero'it was not showing in the reverse pay allocation
      rstSet!OSAmount = 0
      
      rstSet!IsSageUpdate = False
      rstSet!UpdateSage = False
      rstSet!PaymentView = False
      rstSet!ExtRef = txtSPReference.text
      rstSet!PayAmtType = txtPayAmtType.Tag
      rstSet!BankCode = txtBankCode.text
      rstSet!nominalCode = rstSet!BankCode
      rstSet!SlNumber = lSlNumber
      rstSet!ChqNo = txtChqNo.text
      '########I am finding the fund From The PI SPlits by anol 20181116
      For iRow = 1 To flxSPayment.Rows - 1
         If baChangesMade(iRow) And _
               CCur(IIf(flxSPayment.TextMatrix(iRow, 10) = "", 0, flxSPayment.TextMatrix(iRow, 10))) > 0 And _
               (flxSPayment.TextMatrix(iRow, 0) = "" Or flxSPayment.TextMatrix(iRow, 0) = "-") Then
                rstSet!fundID = flxSPayment.TextMatrix(iRow, 13)
        End If
      Next iRow
               
     
      rstSet!ClientID = txtClientIDPurPay.text
      'Resolved by anol 02 Dec 2014
      'issue 468
      rstSet!postingDate = Format(lblPayPostingDate.ToolTipText, "dd mmmm yyyy")
      rstSet.Update
      rstSet.Close

      cGridSPTotal = 0

'  ################################################# Create PP splits  ###############################################
      szSQL = "SELECT * FROM tlbPaymentSplit;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      Dim rsPaytransaction As New ADODB.Recordset
      Dim splitID As Integer
      For iRow = 1 To flxSPayment.Rows - 1
         If baChangesMade(iRow) And _
               CCur(IIf(flxSPayment.TextMatrix(iRow, 10) = "", 0, flxSPayment.TextMatrix(iRow, 10))) > 0 And _
               (flxSPayment.TextMatrix(iRow, 0) = "" Or flxSPayment.TextMatrix(iRow, 0) = "-") Then
               
                lSPTran_ID_split = lSPTran_ID_split + 1
            With rstSet
               .AddNew
               .Fields.Item("TransactionID").Value = UniqueID()
               .Fields.Item("PayHeader").Value = lSp_ID
               .Fields.Item("FundID").Value = flxSPayment.TextMatrix(iRow, 13)
               .Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 10)
'               .Fields.Item("OSAmount").Value = flxSPayment.TextMatrix(iRow, 10)
               If flxSPayment.TextMatrix(iRow, 0) = "" Then
                  .Fields.Item("SplitID").Value = 1
               Else
                  .Fields.Item("SplitID").Value = flxSPayment.TextMatrix(iRow, 1)
               End If
                splitID = .Fields.Item("SplitID").Value
               .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
               .Fields.Item("Description").Value = flxSPayment.TextMatrix(iRow, 7)
               .Fields.Item("AllocTranID").Value = flxSPayment.TextMatrix(iRow, 3)
               'Below line adde by anol 29 Apr 2015
               'When printing payment ref no is showing PP:Date
               'But we need to show the PI reference
               .Fields.Item("PIRef").Value = flxSPayment.TextMatrix(iRow, 12) '
               .Fields.Item("PayTransactionIDSplit").Value = lSPTran_ID_split
               .Fields.Item("propertyID").Value = flxSPayment.TextMatrix(iRow, 4)
               'if PI is rent payable and then I shall find empty string  and I shall not write the CSIDclientstatement or insert null if it is empty
               If flxSPayment.TextMatrix(iRow, 32) <> "" Then
                    .Fields.Item("ClientStatementID").Value = flxSPayment.TextMatrix(iRow, 32)
               End If
               .Update
            End With
            With rsPaytransaction
                szSQL = "SELECT * FROM PayTransactionsSplit;"
               rsPaytransaction.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                .AddNew
                !TranType = "AL"
               
                !TransactionID = lSPTran_ID_split
                !Alloc_Unalloc = 1
                !FromTran = lSp_ID                   'Payment transaction ID
                'added by anol 2021-10-16 . Here the fix was for payment was not writing unitID that means propertyID in tlbpayment which was causing trouble in find
                'supplier OS in Rent summary statement
                'you cannot touch this AllocTranID  field you are using this field while unallocation
                'adoconn.Execute "Update tlbpaymentSplit SET UNIT_ID='" & flxSPayment.TextMatrix(iRow, 4) & "',PayTransactionID=" & lSPTran_ID & " where transactionID='" & lSp_ID & "' and splitID=" & rsInvoice!SplitId & " "
                !ToTran = CLng(flxSPayment.TextMatrix(iRow, 19)) 'PI transaction ID
               ' !AllocDate = Format(Date, "DD MMMM YYYY") 'Issue 437 Allocation Date modification

                !AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
                !PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 10))

                !BankCode = txtBankCode.text
                !nominalCode = !BankCode
                '!SlNumber = lSlNumber
                 'Need to write fund ID 2021-09-13
                !fundID = CLng(flxSPayment.TextMatrix(iRow, 13))

                !VATAMOUNT = CalculateVatAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 19)), splitID, !PaymentAmount)
                !NetAmount = CalculateNetAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 19)), splitID, !PaymentAmount)

                !VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
                !SplitIDofPI = splitID
                !deleteFlag = False
                .Update
            End With
                rsPaytransaction.Close
         End If
      Next iRow
      rstSet.Close

'  ################################################# Update the OS balance of PI #####################################
      For iRow = 1 To flxSPayment.Rows - 1
         If baChangesMade(iRow) And _
            CCur(IIf(flxSPayment.TextMatrix(iRow, 10) = "", 0, flxSPayment.TextMatrix(iRow, 10))) > 0 And _
                  flxSPayment.TextMatrix(iRow, 0) <> "-" Then
'           If there is only one split of this inv then 1st column is "".
'              in this case system will update one line each in tlbPayment and tlbPaymentSplit

'           update the OS balance of the invoice in the payment table
            szSQL = "SELECT P.* " & _
                    "FROM tlbPayment AS P " & _
                    "WHERE P.TransactionID = " & flxSPayment.TextMatrix(iRow, 19) & ";"
            rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

            rstSet!OSAmount = Round(CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                              CCur(flxSPayment.TextMatrix(iRow, 10)) - _
                              CCur(IIf(flxSPayment.TextMatrix(iRow, 11) = "", 0, flxSPayment.TextMatrix(iRow, 11))), 2)
            rstSet!OSAmount = Round(rstSet!OSAmount, 2)
            rstSet!PaymentView = IIf(rstSet!OSAmount > 0, True, False)
            If rstSet!PaymentView = False Then
                rstSet!DateTimeStamp = ""
                rstSet!Module = ""
                rstSet!UserSessionID = ""
                rstSet!WindowsUserName = ""
                rstSet!MachineName = ""
                rstSet!PrestigeUserName = ""
                rstSet!ServerIPaddress = ""
            End If
            rstSet.Update
            rstSet.Close
          '20220802
          'Here I am writing new section for updating rent summary statement/Client statement update client Payment column
          'if the invoice is rent payable then update the amount-osamount to client Payment column
            'Dim szSQL As String
           
            Dim iRows As Integer
            Dim dblBalance As Double
           
            Dim adoRST As New ADODB.Recordset
            szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & txtClientIDPurPay.text & "'"
            adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                    CSID = IIf(IsNull(adoRST("IDbyCL").Value), 0, adoRST("IDbyCL").Value)
            adoRST.Close
            Set adoRST = Nothing
            szSQL = "SELECT P.* " & _
                    "FROM tlbPayment AS P,tblPurinv V " & _
                    "WHERE V.MY_ID=P.PI AND P.TransactionID = " & flxSPayment.TextMatrix(iRow, 19) & " AND V.isRentPayable=true;"
            rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rstSet.EOF Then
                    areYouProcessingRentPayable = True
'                    dblBalance = GetClientACBalance
                    dblclientPaymentAmount = rstSet("Amount").Value - rstSet("OSAmount").Value
                    adoConn.Execute "Update RentSummaryStatement set ClientPayments=" & _
                    -dblclientPaymentAmount & " where StatementID=" & CSID & ""
'                    If IsLoadedAndVisible("frmRentPayable") Then
'                       'frmAutoBankReconciliation.LoadDataExternally adoConn
'                       For iRows = 1 To frmRentPayable.flxPayFees.Rows - 1
'                            If frmRentPayable.flxPayFees.TextMatrix(iRows, 2) = "CS" & CSID Then
'                                        frmRentPayable.flxPayFees.TextMatrix(iRows, 12) = dblclientPaymentAmount 'this is client payment
'                                        frmRentPayable.flxPayFees.TextMatrix(iRows, 16) = dblBalance  'this is client Ac balance
'                                Exit For
'                            End If
'                       Next iRows
'                    End If
            End If
            rstSet.Close
            
          'END
           '     Create the relationship btn the PI and PP header. Relationship btn PI split and PP split is recorded in
'        paymnet split table.
            szSQL = "SELECT * " & _
                    "FROM PayTransactions;"
            With rstSet
               .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
               .AddNew
               !TranType = "AL"
               lSPTran_ID = lSPTran_ID + 1
               !TransactionID = lSPTran_ID
               !Alloc_Unalloc = 1
               !FromTran = lSp_ID                   'Payment transaction ID
               !ToTran = CLng(flxSPayment.TextMatrix(iRow, 19)) 'PI transaction ID
              ' !AllocDate = Format(Date, "DD MMMM YYYY") 'Issue 437 Allocation Date modification
               !fundID = CLng(flxSPayment.TextMatrix(iRow + 1, 13))  'there is fundID in the next  row. because next row contains the split
               !AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
               !PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 10))
               !BankCode = txtBankCode.text
               !nominalCode = !BankCode
               !SlNumber = lSlNumber
               !deleteFlag = False
               .Update
               .Close
            End With

         End If
         If baChangesMade(iRow) And _
                  CCur(IIf(flxSPayment.TextMatrix(iRow, 10) = "", 0, flxSPayment.TextMatrix(iRow, 10))) > 0 And _
                  (flxSPayment.TextMatrix(iRow, 0) = "-" Or flxSPayment.TextMatrix(iRow, 0) = "") Then
            If flxSPayment.TextMatrix(iRow, 0) = "" Then
               adoConn.Execute "UPDATE tlbPaymentSplit " & _
                               "SET OSAmount = " & Round(CCur(flxSPayment.TextMatrix(iRow, 9)), 2) - _
                                 CCur(flxSPayment.TextMatrix(iRow, 10)) - _
                                 CCur(IIf(flxSPayment.TextMatrix(iRow, 11) = "", 0, _
                                 flxSPayment.TextMatrix(iRow, 11))) & " " & _
                               "WHERE TransactionID = '" & flxSPayment.TextMatrix(iRow, 3) & "' AND " & _
                                 "SplitID = 1;"
            End If
            If flxSPayment.TextMatrix(iRow, 0) = "-" Then
               adoConn.Execute "UPDATE tlbPaymentSplit " & _
                               "SET OSAmount = " & Round(CCur(flxSPayment.TextMatrix(iRow, 9)), 2) - _
                                 CCur(flxSPayment.TextMatrix(iRow, 10)) - _
                                 CCur(IIf(flxSPayment.TextMatrix(iRow, 11) = "", 0, _
                                 flxSPayment.TextMatrix(iRow, 11))) & " " & _
                               "WHERE TransactionID = '" & flxSPayment.TextMatrix(iRow, 3) & "' AND " & _
                                 "SplitID = " & flxSPayment.TextMatrix(iRow, 1) & ";"
            End If
         End If
      Next iRow
   End If
   
'**************************************************************************************
'  Saving Payment on Account            ************* PoA *****************************
'**************************************************************************************
   If cPoA > 0 Then
      lSp_ID = lSp_ID + 1

      szSQL = "SELECT * FROM tlbPayment;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      lSlNumber = SlNumber("PA", "tlbPayment", adoConn)

      rstSet.AddNew
      rstSet!TransactionID = lSp_ID
      rstSet!szTransactionID = rstSet!TransactionID
      SavePaymentTransactions = lSp_ID

      rstSet!Type = 9 'CByte(sdoPA)                   'tlbTransactionType.TYPE_ID (4) = Purchase Payment on Account
      rstSet!SageAccountNumber = txtSPSupplier.Tag

      rstSet!PDate = Format(txtSPDate.text, "dd mmmm yyyy")
      rstSet!ref = "PA" & Format(Now, "yymmddhhmmss")
      rstSet!Details = "PAYMENT ON ACCOUNT"
      rstSet!amount = cPoA
      rstSet!OSAmount = rstSet!amount           'amount to be allocated
      rstSet!PaymentView = True
      rstSet!BankCode = txtBankCode.text
      rstSet!nominalCode = rstSet!BankCode
      rstSet!ExtRef = txtSPReference.text
      rstSet!PayAmtType = txtPayAmtType.Tag
      rstSet!SlNumber = lSlNumber
      rstSet!fundID = lFund
      rstSet!ClientID = txtClientIDPurPay.text
      rstSet!postingDate = Format(lblPayPostingDate.ToolTipText, "dd mmmm yyyy")
      rstSet.Update
      rstSet.Close

'     Saving the split(s) of the header
      szSQL = "SELECT * FROM tlbPaymentSplit;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

      With rstSet
         .AddNew
         .Fields.Item("TransactionID").Value = UniqueID()
         .Fields.Item("PayHeader").Value = lSp_ID
         .Fields.Item("FundID").Value = lFund
         .Fields.Item("Amount").Value = cPoA
         .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
         .Fields.Item("SplitID").Value = 1
         .Fields.Item("DueDate").Value = Format(txtSPDate.text, "dd mmmm yyyy")
         .Fields.Item("Description").Value = "Payment on Account"
         .Update
      End With
      rstSet.Close
   End If

'**************************************************************************************
'  Saving Refund            ************* PAYMENT  REFUND *****************************
'**************************************************************************************
   If Val(txtSPaymentTotal.text) < 0 Then
      lSp_ID = lSp_ID + 1

      szSQL = "SELECT * FROM tlbPayment;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      lSlNumber = SlNumber("PPR", "tlbPayment", adoConn)

      rstSet.AddNew
      rstSet!TransactionID = lSp_ID
      rstSet!szTransactionID = rstSet!TransactionID
      SavePaymentTransactions = lSp_ID

      rstSet!Type = 24
      rstSet!SageAccountNumber = txtSPSupplier.Tag
      rstSet!PDate = Format(txtSPDate.text, "dd mmmm yyyy")
      rstSet!ref = "PPR" & Format(Now, "yymmddhhmmss")
      rstSet!Details = "Purchase Payment Refund"
      rstSet!amount = txtSPaymentTotal.text * (-1)
      rstSet!OSAmount = rstSet!amount
      rstSet!PaymentView = True
      rstSet!BankCode = txtBankCode.text
      rstSet!nominalCode = rstSet!BankCode
      rstSet!ExtRef = txtSPReference.text
      rstSet!PayAmtType = txtPayAmtType.Tag
      rstSet!SlNumber = lSlNumber
      rstSet!fundID = lFund
      rstSet!ClientID = txtClientIDPurPay.text
      rstSet!postingDate = Format(lblPayPostingDate.ToolTipText, "dd mmmm yyyy")
      rstSet.Update
      rstSet.Close

'     Saving the split(s) of the header
      szSQL = "SELECT * FROM tlbPaymentSplit;"
      rstSet.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

      With rstSet
         .AddNew
         .Fields.Item("TransactionID").Value = UniqueID()
         .Fields.Item("PayHeader").Value = lSp_ID
         .Fields.Item("FundID").Value = lFund
         .Fields.Item("Amount").Value = txtSPaymentTotal.text * (-1)
         .Fields.Item("OSAmount").Value = txtSPaymentTotal.text * (-1)
         .Fields.Item("SplitID").Value = 1
         .Fields.Item("DueDate").Value = Format(txtSPDate.text, "dd mmmm yyyy")
         .Fields.Item("Description").Value = "Purchase Payment Refund"
         .Update
      End With
      rstSet.Close
   End If

   MousePointer = vbDefault

'
''--------------------------------------------------------------------------------------------
''  Export Transactions to Nominal Ledger (NLPosting table)
'   Export_PPnPPR_2_NL adoConn
'--------------------------------------------------------------------------------------

   bChangesMade = False
   cmdBankAc.Enabled = True

   Set rstSet = Nothing
   
'issue 523
'Modified by anol 20 Jan 2014
   UpdateBankAcBal_Minus adoConn, Val(txtSPaymentTotal.text), txtBankCode.text, txtClientIDPurPay.text

   

'  ****************                 UPDATING OPENED ### CASHBOOK ### GRIDS
   UpdatingCB adoConn
'  ***********************************************************************

   If IsLoadedAndVisible("frmAutoBankReconciliation") Then
      frmAutoBankReconciliation.LoadDataExternally adoConn
   End If
   szTran2Fix = "" 'Need to clear this before it stores any value ot this will cause a problem
   'If PI_Check(adoconn, szTran2Fix) = False Or PayAllocation_check(adoconn, szTran2Fix) = False Then
   If PI_Check(adoConn, szTran2Fix) = False Then 'true means data is in good state
        adoConn.RollbackTrans
        adoConn.Close
        MsgBox "An error occurred while payment, transaction has been rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked"
   ElseIf PayAllocation_check(adoConn, szTran2Fix) = False Then 'true means data is in good state
        adoConn.RollbackTrans
        adoConn.Close
        MsgBox "An error occurred while payment, transaction has been rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked"
   Else
        adoConn.CommitTrans
        frmMMain.frmSupplier_SupplierList_isUptoDate = False
        frmMMain.frmPI_SupplierBalance_isUptoDate = False
        frmMMain.frmPI_SupplierBalanceByCL_isUptoDate = False
        adoConn.Close
        adoConn.Open getConnectionString
        '--------------------------------------------------------------------------------------------
        '  Export Transactions to Nominal Ledger (NLPosting table)
'        If strSupplierTypeOnSelection = "AGENT" Then
'            Export_PPnPPR_2_NL_MAgent adoConn
'        Else
            Export_PPnPPR_2_NL adoConn
'        End If
         'added by anol 20170923
        cmdSPSupplier_Click
        MsgBox "The Transactions have been saved successfully.", vbInformation, "Transaction Saved"
        adoConn.Close
        Set adoConn = Nothing
   End If
   
   adoConn.Open getConnectionString
   Call ConfigFlxSPayment
   Call LoadFlxSPayment(adoConn)
   Call ConfigFlxSCrPoA
   Call LoadFlxSCrPoA(adoConn)
   Call UpdtSupplierAccountBalance(adoConn)
    adoConn.Close
'   txtSPSupplier.text = ""
'   txtSPSupplier.Tag = ""
'   txtSPReference.text = ""
'   txtBankCode.text = ""
'   txtBankAc.text = ""
   txtSPDate.text = Format(Date, "dd/MM/yyyy")
   txtSPaymentTotal.text = "0.00"
   txtPaymentEntered.text = "0.00"
   
'   txtSupAcBal.text = "0.00"
   flxSCrPoA.Enabled = True
   ReDim baChangesMade(flxSPayment.Rows) As Boolean

   
End Function
Private Function CalculateVatAmountFormPIsplit(adoConn As ADODB.Connection, szPITranID As String, ByVal szSpitID As Integer, dblPayment As Double) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select VAT+NET_AMOUNT as amt,VAT from tblPurInvSRec S ,tlbPayment P where P.PI =S.ParentID AND P.TransactionID=" & _
            szPITranID & " and TRAN_ID='" & szSpitID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
             CalculateVatAmountFormPIsplit = dblPayment / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
             CalculateVatAmountFormPIsplit = CalculateVatAmountFormPIsplit * IIf(IsNull(rsDemandSplitRecords("VAT").Value), 0, rsDemandSplitRecords("VAT").Value)
    End If
    rsDemandSplitRecords.Close
    
End Function
Private Function CalculateNetAmountFormPIsplit(adoConn As ADODB.Connection, szDemandID As String, ByVal szSpitID As Integer, dblPayment As Double) As Double
    Dim rsDemandSplitRecords As New ADODB.Recordset
    rsDemandSplitRecords.Open "Select VAT+NET_AMOUNT as amt,NET_AMOUNT,VAT from tblPurInvSRec S ,tlbPayment P  where  P.PI =S.ParentID AND P.TransactionID=" & _
                szDemandID & "  and TRAN_ID='" & szSpitID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsDemandSplitRecords.EOF Then
             If IIf(IsNull(rsDemandSplitRecords("VAT").Value), 0, rsDemandSplitRecords("VAT").Value) = 0 Then
                    CalculateNetAmountFormPIsplit = dblPayment
             Else
                    CalculateNetAmountFormPIsplit = dblPayment / IIf(IsNull(rsDemandSplitRecords("amt").Value), 0, rsDemandSplitRecords("amt").Value)
                    CalculateNetAmountFormPIsplit = CalculateNetAmountFormPIsplit * IIf(IsNull(rsDemandSplitRecords("NET_AMOUNT").Value), 0, rsDemandSplitRecords("NET_AMOUNT").Value)
             End If
    End If
    rsDemandSplitRecords.Close

End Function
Private Sub cmdSPayAll_Click()
'   If Val(txtSPaymentTotal.text) < 0 Then Exit Sub
'
'   Dim iRow As Integer, cDiff As Currency
'
'   If bChangesMade Then
'      For iRow = 1 To flxSPayment.Rows - 1
'         If Val(flxSPayment.TextMatrix(iRow, 10)) > 0 Then
'            txtSPaymentTotal.text = Val(txtSPaymentTotal.text) - Val(flxSPayment.TextMatrix(iRow, 10))
'            flxSPayment.TextMatrix(iRow, 10) = "0.00"
'         End If
'      Next iRow
'   End If
'
'   If bTotalPayTyped Then cDiff = Val(txtSPaymentTotal.text)
'
'   For iRow = 1 To flxSPayment.Rows - 1
'      If flxSPayment.TextMatrix(iRow, 2) <> "ADJI" And flxSPayment.TextMatrix(iRow, 9) <> "" Then
'         If bTotalPayTyped Then
'            If cDiff > Val(flxSPayment.TextMatrix(iRow, 9)) Then
'               flxSPayment.TextMatrix(iRow, 10) = flxSPayment.TextMatrix(iRow, 9)
'               cDiff = cDiff - Val(flxSPayment.TextMatrix(iRow, 9))
'            Else
'               flxSPayment.TextMatrix(iRow, 10) = Format(cDiff, "0.00")
'               cDiff = 0
'               baChangesMade(iRow) = IIf(Val(flxSPayment.TextMatrix(iRow, 10)) > 0, True, False)
'               Exit For
'            End If
'         Else
'            flxSPayment.TextMatrix(iRow, 10) = flxSPayment.TextMatrix(iRow, 9)
'            'Marked incorrect by anol 22 July 2015
'            txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + Val(flxSPayment.TextMatrix(iRow, 9)), "0.00")
'         End If
'         baChangesMade(iRow) = IIf(Val(flxSPayment.TextMatrix(iRow, 10)) > 0, True, False)
'      End If
'   Next iRow
'
'   cGridSPTotal = TotalPaymentEntered
'   txtPaymentEntered.text = Format(cGridSPTotal, "0.00")
'   bPayAll = True
'copied from demand3 by anol 22 July 2015
'issue 571
  Dim iPayTran As Integer, i As Integer
 'addded by anol on 20160526 Pay all was not filling last row
  ReDim baChangesMade(flxSPayment.Rows) As Boolean
   For i = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(i, 0) = "+" Or _
            flxSPayment.TextMatrix(i, 0) = "" Or _
            flxSPayment.TextMatrix(i, 0) = ">" Then _
      iPayTran = iPayTran + 1
      
   Next i

   flxSPayment.row = 1

   For i = 1 To iPayTran
      cmdSPFull_Click
   Next i

   'cmeRevereseAllocation.Enabled = False
   'cmdAllocate.Enabled = False
 
'   For i = 1 To flxSPayment.Rows - 1
'      'addded by anol on 20160526 Pay all was not working on several press on the button
'      flxSPayment.TextMatrix(flxSPayment.row, 10) = 0
'   Next i
   bTotalPayTyped = False
   'Exit Sub
End Sub

Private Sub cmdSPClose_Click()
    
   Unload Me
End Sub

Private Sub cmdSPFull_Click()
   On Error GoTo ErrorHandler

   
   If flxSPayment.row > 0 And flxSPayment.row <= flxSPayment.Rows - 1 Then
      flxSPayment.col = 0
      If flxSPayment.CellBackColor = vbRed Then
                MsgBox "Selected invoice is locked by another user. Please wait untill other user release this record.", vbInformation, "Warning"
                GoTo XX
      End If
       
      flxSPayment.col = 9
      If Val(flxSPayment.TextMatrix(flxSPayment.row, 10)) > 0 And Not bTotalPayTyped Then
         txtSPaymentTotal.text = Val(txtSPaymentTotal.text) - Val(flxSPayment.TextMatrix(flxSPayment.row, 10))
      End If

      If bTotalPayTyped Then                'Payment amount has put in the "Total Payment Amt"
         If Val(txtDiffPay.text) > Val(flxSPayment.TextMatrix(flxSPayment.row, 9)) Then
            flxSPayment.TextMatrix(flxSPayment.row, 10) = flxSPayment.TextMatrix(flxSPayment.row, 9)
         Else
            flxSPayment.TextMatrix(flxSPayment.row, 10) = IIf(baChangesMade(flxSPayment.row), flxSPayment.TextMatrix(flxSPayment.row, 10), txtDiffPay.text)
         End If
      Else
         flxSPayment.TextMatrix(flxSPayment.row, 10) = flxSPayment.TextMatrix(flxSPayment.row, 9)
         txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + CCur(flxSPayment.TextMatrix(flxSPayment.row, 10)), "0.00")
      End If

      
      baChangesMade(flxSPayment.row) = IIf(Val(flxSPayment.TextMatrix(flxSPayment.row, 10)) > 0, True, False)

      txtPaymentEntered.text = Format(cGridSPTotal, "0.00")
      'By anol 22 July 2015 issue 571
        Dim iInc As Integer
        iCurRow = flxSPayment.row
        iInc = 1
        If flxSPayment.TextMatrix(flxSPayment.row, 0) = "+" Or flxSPayment.TextMatrix(flxSPayment.row, 0) = ">" Then
           iInc = SpreadHeaderInSplits
        End If
        If flxSPayment.TextMatrix(flxSPayment.row, 0) = "-" Then
           SumUpHeaderBySplits
        End If
        cGridSPTotal = TotalPaymentEntered
        txtPaymentEntered.text = Format(cGridSPTotal, "0.00")
        'End of modification
     ' flxSPayment.row = flxSPayment.row + 1
      'flxSPayment_Click
      If flxSPayment.row < flxSPayment.Rows - 1 Then
           flxSPayment.row = flxSPayment.row + iInc
           HighLightRowFlxGrid flxSPayment, flxSPayment.row
      End If
XX:
   End If
   Exit Sub

ErrorHandler:
   Debug.Print "Reached the end of the records"
End Sub

Private Sub cmdTaxList_Click(Index As Integer)
   'Resolved by BOSL
    'Issue 453 Incorrect filtering solve of Note 3
    'Modified by Anol 14 Aug 2014
    
'   On Error GoTo Err
   chkShowBal.Visible = False
   sTextBox = "VAT"
   Call LoadVAT

   tabPayment.Enabled = False

   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   txtSearch2.Width = 1000
   fraList.Width = 2400
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtVat_(0).Left - 400
   fraList.Top = txtVat_(0).Top + txtVat_(0).Height
   fraList.Visible = True
   fraList.ZOrder 0
 
  'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
Err:
End Sub

Private Sub LoadVAT()
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 1000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "RATE"

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPurExp.Tab).Width = 900
   lblSearch0(tabPurExp.Tab).Left = 50
   lblSearch1.Width = 1900
   lblSearch1.Left = lblSearch0(tabPurExp.Tab).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 960
   txtSearch1.Left = 40

   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(tabPurExp.Tab).Caption = "CODE"
   lblSearch1.Caption = "RATE"
   lblSearch2.Visible = False

   flxSupplier(0).RowHeight(0) = 0

   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString

   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode where IN_USE;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

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
   Conn2.Close
   
   Set rstRec = Nothing
   Set Conn2 = Nothing
End Sub

Private Sub LoadVATBk()
   flxSupplier(0).Clear
   flxSupplier(0).Cols = 4
   flxSupplier(0).ColWidth(0) = 800
   flxSupplier(0).ColWidth(1) = 800

   flxSupplier(0).TextMatrix(0, 1) = "CODE"
   flxSupplier(0).TextMatrix(0, 2) = "RATE"
    
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   lblSearch0(0).Width = 600
   lblSearch0(0).Left = 50
   lblSearch1.Width = 600
   lblSearch1.Left = lblSearch0(0).Left + 800
   
   txtSearch1.Width = 600
   txtSearch1.Left = 40
   
   txtSearch2.Width = 600
   txtSearch2.Left = txtSearch1.Left + 800
   
   lblSearch0(0).Caption = "Code"
   lblSearch1.Caption = "Rate"
   lblSearch2.Visible = False

   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString

   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

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
   Conn2.Close

   Set rstRec = Nothing
   Set Conn2 = Nothing
End Sub

Private Sub cmdTypeList_Click()
    chkShowBal.Visible = False
    sTextBox = "PROPERTY"
    If txtSupplierID.text = "" Then
       cmdACList(0).SetFocus
       ShowMsgInTaskBar "Please select the " & cmbSC.text & ".", "Y", "N"
       Exit Sub
    End If
    tabPurExp.Enabled = False
    LoadPropertyList "" 'loading the properties
    
    tabPayment.Enabled = False
    txtSearch1.Visible = True
    txtSearch2.Visible = True
    
    txtSearch1.text = ""
    txtSearch2.text = ""
    
    fraList.Width = 4815
    cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
    Shape4(0).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
    flxSupplier(0).Width = fraList.Width - 50
    fraList.Left = txtProperty.Left + txtProperty.Width - fraList.Width + fraLay(0).Left + tabPurExp.Left
    fraList.Top = txtProperty.Top + fraLay(0).Top + tabPurExp.Top '+ 380
    fraList.Visible = True
    fraList.ZOrder 0
    
    
    'Resolved by BOSL
    'Issue 553 PRESTIGE GUI IMPROVEMENT
    'Modified by Anol 25 Mar 2015
    'flxSupplier(0).SetFocus
    txtSearch1.SetFocus
    'End of modification
    'Resolved by BOSL
    'Issue No: 0000467
    'Added By: Asif. 04 Sep 2014
    txtUnit(0).text = ""
   
End Sub

Private Sub cmdUnitList_Click()
'   If txtProperty.text = "" Then
'      cmdTypeList.SetFocus
'      Exit Sub
'   End If
    chkShowBal.Visible = False
   fraList.Height = 2925

   Dim Conn2 As New ADODB.Connection
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString

   LoadUnitList "", Conn2
   Conn2.Close
   Set Conn2 = Nothing

   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(tabPurExp.Tab).Left = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width
   Shape4(tabPurExp.Tab).Width = fraList.Width - cmdGridUnitLookup(tabPurExp.Tab).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   flxSupplier(0).Refresh
   fraList.Left = txtUnit(tabPurExp.Tab).Left + 100
   fraList.Top = txtUnit(tabPurExp.Tab).Top + 380
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "UNIT"
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Modified by Anol 25 Mar 2015
   'flxSupplier(0).SetFocus
   txtSearch1.SetFocus
   'Modified by Anol 30 Apr 2015
   txtUnit(0).Locked = False
   'End of addition
End Sub

Private Sub LoadPropertyList(Filter As String)
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxSupplier(0).RowHeight(0) = 0
   flxSupplier(0).Cols = 6
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
   flxSupplier(0).ColWidth(3) = 0
   flxSupplier(0).ColWidth(4) = 0
   flxSupplier(0).ColWidth(5) = 0

   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColAlignment(0) = vbLeftJustify
   flxSupplier(0).ColAlignment(1) = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1.Width = 1490
   txtSearch1.Left = 40
   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   lblSearch0(0).Caption = "Property ID"
   lblSearch1.Caption = "Property Name"
   lblSearch2.Visible = False
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'   On Error Resume Next
    

   rRow = 1
   
    If sTextBox = "PROPERTY" Then
        rRow = 1
        If txtClientID.text <> "ALL" Then
            'Modification in SQL written by anol 2020-10-07
            szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as VAT_CODE1,G.VATRate  as  Rate,G.vatOptionEnabled " & _
                  "FROM ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN tlbVatCode V ON G.VATRate=V.VAT_ID) " & _
                  "WHERE ClientID = '" & txtClientID.text & "' " & _
                  "ORDER BY P.PropertyID;"
        Else
            szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as as VAT_CODE1,G.VATRate  as  Rate,vatOptionEnabled " & _
                  "FROM ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN  tlbVatCode V  ON G.VATRate=V.VAT_ID) " & _
                  "ORDER BY P.PropertyID;"
        End If
        
        rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Filter <> "" Then
            rstRec.Filter = Filter
        End If
        flxSupplier(0).Rows = rstRec.RecordCount + 2
        While Not rstRec.EOF
            flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
            If rRow = 1 Then
                Debug.Print rstRec.Fields.Item(0).Value
            End If
            flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
            If (IIf(IsNull(rstRec.Fields("vatOptionEnabled").Value), "", rstRec.Fields("vatOptionEnabled").Value)) = 1 Then
                flxSupplier(0).TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields("Rate").Value), "", rstRec.Fields("Rate").Value) 'VAT_ID
                flxSupplier(0).TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("RateValue").Value), "0.00", rstRec.Fields("RateValue").Value)
                flxSupplier(0).TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value) 'like T9
                flxSupplier(0).TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("vatOptionEnabled").Value), "", rstRec.Fields("vatOptionEnabled").Value) 'like T9
            Else
                flxSupplier(0).TextMatrix(rRow, 2) = "" 'VAT_ID
                flxSupplier(0).TextMatrix(rRow, 3) = ""
                flxSupplier(0).TextMatrix(rRow, 4) = "" 'like T9
                flxSupplier(0).TextMatrix(rRow, 5) = "0" '0 means VAT disabled for the property
            End If
            rstRec.MoveNext
            rRow = rRow + 1
        Wend
         rstRec.Close
         Set rstRec = Nothing
 ElseIf sTextBox = "PROPERTYFILTER" Then 'Properties loading into the filter
        rRow = 1
        If txtIDClient.text <> "ALL" Then
            szSQL = "SELECT PropertyID, PropertyName " & _
                  "FROM Property " & _
                  "WHERE ClientID = '" & txtIDClient.text & "' " & _
                  "ORDER BY PropertyID;"
        Else
            szSQL = "SELECT PropertyID, PropertyName " & _
                  "FROM Property " & _
                  "ORDER BY PropertyID;"
        End If
        
        rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         If Filter <> "" Then
            rstRec.Filter = Filter
        End If
        flxSupplier(0).Rows = rstRec.RecordCount + 3
        flxSupplier(0).TextMatrix(1, 0) = "ALL"
        flxSupplier(0).TextMatrix(1, 1) = "ALL Properties"
        'flxSupplier(0).AddItem ""
        rRow = 2
        While Not rstRec.EOF
            flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
            flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
            'flxSupplier(0).RowHeight(rRow) = 240
            rstRec.MoveNext
            'If Not rstRec.EOF Then flxSupplier(0).AddItem ""
            rRow = rRow + 1
        Wend
        rstRec.Close
  ElseIf sTextBox = "PROPERTYHIST" Then  'issue 629 creating new filter
            rRow = 1
            If txtClientIdlist.text <> "ALL" Then
               szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "WHERE ClientID = '" & txtClientIdlist.text & "' " & _
                       "ORDER BY PropertyID;"
            Else
                 szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "ORDER BY PropertyID;"
            End If
            rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Filter <> "" Then
                rstRec.Filter = Filter
            End If
            flxSupplier(0).Rows = rstRec.RecordCount + 3
            flxSupplier(0).TextMatrix(1, 0) = "ALL"
            flxSupplier(0).TextMatrix(1, 1) = "ALL Properties"
            'flxSupplier(0).AddItem ""

            rRow = 2
            While Not rstRec.EOF
               flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               'flxSupplier(0).RowHeight(rRow) = 240
               rstRec.MoveNext
               'If Not rstRec.EOF Then flxSupplier(0).AddItem ""
               rRow = rRow + 1
            Wend
              rstRec.Close
  End If
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub LoadUnitList(szSQL As String, Conn2 As ADODB.Connection)
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPurExp.Tab).Width = 1400
   lblSearch0(tabPurExp.Tab).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(tabPurExp.Tab).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1460
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   Dim rRow As Integer
   Dim rstRec As New ADODB.Recordset

   If szSQL = "" Then
      szSQL = "SELECT UnitNumber, UnitName " & _
              "FROM Units, Property AS P " & _
              "WHERE Units.PropertyID = P.PropertyID AND " & _
                  "P.PropertyID = '" & txtProperty.text & "' AND " & _
                  "P.ClientID = '" & txtClientID.text & "' " & _
              "ORDER BY UnitNumber"
              'Issue 469 note 2
              'modifed by anol 28 Dec 2014
       If txtProperty.text = "" Then
             szSQL = "SELECT UnitNumber, UnitName " & _
              "FROM Units, Property AS P " & _
              "WHERE Units.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & txtClientID.text & "' " & _
              "ORDER BY UnitNumber"
      End If
   End If
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0
      '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(tabPurExp.Tab).Caption = "Unit ID"
      lblSearch1.Caption = "Unit Name"
      lblSearch2.Visible = False

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!UnitNumber
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!UnitName), "", rstRec!UnitName)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
End Sub



'Private Sub LoadUnit()
'   flxSupplier(0).Clear
'   flxSupplier(0).Rows = 2
'   flxSupplier(0).Cols = 2
'
'   flxSupplier(0).ColWidth(0) = 1000
'   flxSupplier(0).ColWidth(1) = 2500
'   flxSupplier(0).ColAlignment(0) = vbRightJustify
'   flxSupplier(0).ColAlignment(1) = vbLeftJustify
'
'   flxSupplier(0).ColWidth(2) = 0
'   flxSupplier(0).ColWidth(3) = 0
'   lblSearch0(0).Width = 700
'   lblSearch0(0).Left = 50
'   lblSearch1.Width = 1500
'   lblSearch1.Left = lblSearch1.Left + 500
'
'   txtSearch1.Width = 700
'   txtSearch1.Left = 40
'
'   txtSearch2.Width = 2000
'   txtSearch2.Left = txtSearch1.Left + 1200
'
'   lblSearch0(0).Caption = "Unit ID"
'   lblSearch1.Caption = "Unit Name"
'   lblSearch2.Visible = False
'
'   Dim rRow As Integer
'   Dim Conn2 As New ADODB.Connection
'
'   Dim szSQL As String
'   Dim rstRec As New ADODB.Recordset
'
'   'Reset screen to show all the units in cboUnits.
'   'Set the RDO Connections to the dataset
'   Conn2.Open getConnectionString
'
'   szSQL = "SELECT Units.UnitNumber, Units.UnitName, Units.UnitPostCode, Units.TotalArea, Units.RentalPrice " & _
'               "FROM Client, Property, Units WHERE Property.PropertyID = Units.PropertyID AND Client.ClientID = Property.ClientID"
''Debug.Print szSQL
'   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
'
'   If rstRec.EOF = False Then
'      rRow = 1
'      While Not rstRec.EOF
'         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!UnitNumber
'         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!UnitName), "", rstRec!UnitName)
'         rstRec.MoveNext
'         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
'         rRow = rRow + 1
'      Wend
'   End If
'
'   rstRec.Close
'   Conn2.Close
'End Sub

Private Sub cmdUpdate__Click(Index As Integer)

End Sub

Private Sub cmdUpdate_Click(Index As Integer)
''   'added by anol 02 Jan 2015
''   'issue 469
''         txtInv(0).Locked = False
''         txtUnit(0).Locked = False
''         txtNC(0).Locked = False
''         txtDept(1).Locked = False
''         txtDept(0).Locked = False
''         txtPFName.Locked = False
''         txtJobNo.Locked = False
''         txtDetails_(0).Locked = False
''         txtNet_(0).Locked = False
''         txtVat_(0).Locked = False
''         txtSchedules.Locked = False
''         txtRecoverable(0).Locked = False
''         txtTotal.Locked = False
''         'End of modification

   If Index = 0 Or Index = 1 Then                               'Add New Line
      cmdDelete.Enabled = True
      If txtSupplierID.text = "" Then
         ShowMsgInTaskBar "You must select Account Code from the list.", "Y", "N"
         cmdACList(0).SetFocus
         Exit Sub
      End If
      If txtDate.text = "" Then
         ShowMsgInTaskBar "You must enter the date from the list.", "Y", "N"
         txtDate.SetFocus
         Exit Sub
      End If
      If txtDueDate.text = "" Then
         ShowMsgInTaskBar "You must enter the due date from the list.", "Y", "N"
         txtDueDate.SetFocus
         Exit Sub
      End If
      If txtClientID.text = "" Then
         ShowMsgInTaskBar "You must select the client.", "Y", "N"
         cmdClientSerc.SetFocus
         Exit Sub
      End If
   End If

   If Index = 1 Then                                  'This means when OK button is clicked
        If cmbSC.text <> "Managing Agent" And chkIsMgtFee.Value = 1 Then
             MsgBox "Please select supplier type Managing Agent if you are creating a Management Fee", vbInformation, "Management Fee"
             Exit Sub
        End If
      If txtNC(0).text = "" Then
         ShowMsgInTaskBar "You must select Nominal Code from the list.", "Y", "N"
         If cmdNCList.Enabled = True Then
            cmdNCList.SetFocus
         End If
         Exit Sub
      End If
      If txtDept(0).text = "" Then
         ShowMsgInTaskBar "You must select a fund from the list.", "Y", "N"
         cmdDeptList().SetFocus
         Exit Sub
      End If
      If Val(txtNet_(0).text) <= 0 Then
         ShowMsgInTaskBar "You must enter the amount.", "Y", "N"
         If txtNet_(0).Enabled Then txtNet_(0).SetFocus
         Exit Sub
      End If
      If chkRecover.Value = 1 And Val(txtRecoverable(0).text) = 0 Then
         ShowMsgInTaskBar "You must enter the amount.", "Y", "N"
         txtNet_(0).SetFocus
         Exit Sub
      End If
      
      With flxPI
          If Not bEditMode Then                                 ' ****************  ADD NEW PI  ************************
             If Not (.Rows = 2 And .TextMatrix(1, 1) = "") Then
                .AddItem ""
             End If
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = txtSupplierID.text
            .TextMatrix(.Rows - 1, 2) = txtDate.text
            .TextMatrix(.Rows - 1, 3) = txtProperty.text
            .TextMatrix(.Rows - 1, 4) = IIf(sAddChoice = "IN" Or sAddChoice = "AI", "Invoice", "Credit")
            .TextMatrix(.Rows - 1, 5) = txtUnit(tabPurExp.Tab).text
            .TextMatrix(.Rows - 1, 6) = txtNC(0).text   'Nominal code
            .TextMatrix(.Rows - 1, 7) = txtDept(0).text 'Fund Code
            .TextMatrix(.Rows - 1, 8) = txtDept(1).text 'Fund ID
            .TextMatrix(.Rows - 1, 9) = txtJobNo.text
            .TextMatrix(.Rows - 1, 11) = txtDetails_(tabPurExp.Tab).text
            .TextMatrix(.Rows - 1, 12) = txtNet_(0).text
            .TextMatrix(.Rows - 1, 13) = Space(10) & lblVatCode(tabPurExp.Tab).Caption
            .TextMatrix(.Rows - 1, 14) = txtVat_(0).text
            .TextMatrix(.Rows - 1, 20) = txtSchedules.text
            .TextMatrix(.Rows - 1, 21) = txtUnit(tabPurExp.Tab).text
            .TextMatrix(.Rows - 1, 22) = IIf(txtRecoverable(0).text = "", 0, txtRecoverable(0).text)
            .TextMatrix(.Rows - 1, 15) = Format(Val(txtTotal.text), "0.00")
            .TextMatrix(.Rows - 1, 23) = UniqueID()
            .TextMatrix(.Rows - 1, 24) = txtDept(0).text  'Fund Code
            .TextMatrix(.Rows - 1, 25) = txtPFName.text
         Else                                                  ' ****************  Update PI  ************************
            .TextMatrix(iCurEditRow, 5) = txtUnit(tabPurExp.Tab).text
            .TextMatrix(iCurEditRow, 6) = txtNC(0).text
            .TextMatrix(iCurEditRow, 7) = txtDept(0).text  'Fund Code
            .TextMatrix(iCurEditRow, 8) = txtDept(1).text  'Fund Name
            .TextMatrix(iCurEditRow, 11) = txtDetails_(tabPurExp.Tab).text
            .TextMatrix(iCurEditRow, 12) = txtNet_(0).text
            .TextMatrix(iCurEditRow, 13) = Space(10) & lblVatCode(tabPurExp.Tab).Caption
            .TextMatrix(iCurEditRow, 14) = txtVat_(0).text
            .TextMatrix(iCurEditRow, 19) = ""
            .TextMatrix(iCurEditRow, 20) = txtSchedules.text
            .TextMatrix(iCurEditRow, 21) = txtUnit(tabPurExp.Tab).text
            .TextMatrix(iCurEditRow, 22) = IIf(txtRecoverable(0).text = "", 0, txtRecoverable(0).text)
            .TextMatrix(iCurEditRow, 15) = Format(Val(txtTotal.text), "0.00")
            .TextMatrix(iCurEditRow, 23) = UniqueID()
            .TextMatrix(iCurEditRow, 24) = txtDept(0).text
            .TextMatrix(iCurEditRow, 25) = txtPFName.text
            
           
            HandleCommandButton "Update Record"
         End If
           'anol 09 July 2015 issue 571 note 1125
            flxPI.Tag = "EditedOrAdded"
            'anol 09 July 2015 issue 571 note 1125
            bEditMode = False
            cmdSavePI.Enabled = True
           'end of modification
            cmdEdit(0).Enabled = True
            PIComponents "NewLine"
      End With

      UpdateTotalPICN

'      If txtProperty.text = "" Then
    If cmdNCList.Enabled = True Then
            cmdNCList.SetFocus
         End If
'      Else
'         cmdUnitList.SetFocus
'      End If
    'anol 22 Jun 2015 issue 571 note 1097
   
   End If
    If Index = 1 Then
        cmdCancel(0).Enabled = True
        cmdSavePI.Enabled = True
        'added by anol 20 July 2015
        'Split line Enable mode
        cmdNCList.Enabled = False
        cmdaddnewline.Enabled = True
        txtUnit(0).Enabled = False
        cmdUnitList.Enabled = False
        txtJobNo.Enabled = False
        cmdJobNo(0).Enabled = False
        txtNet_(0).Enabled = False
        txtNC(0).Enabled = False
        cmdNCList.Enabled = False
        cmdUpdate(1).Enabled = False
        cmdTaxList(0).Enabled = False
        cmdDeptList.Enabled = False
        txtDetails_(0).Enabled = False
        txtVat_(0).Enabled = False
        cmdSchedules(0).Enabled = False
        cmdUpdate(2).Enabled = False
        chkRecover.Enabled = False
        txtRecoverable(0).Enabled = False
        cmdaddnewline.SetFocus
        bEditMode = False
        'flxPI.Enabled = True
        'End of addition
    
    
    End If
   If Index = 2 Then
      If flxPI.TextMatrix(1, 1) = "" Then
            fraLay(1).Enabled = True
            lblVatCode(0).Caption = ""
      End If
      cmdDelete.Enabled = True
      flxPI.Enabled = True
      PIComponents "EditLine"
     'added by anol issue 571 date 13 Aug 2015
      Dim iCol       As Integer
      'flxPI.row = iSelected
      For iCol = 0 To flxPI.Cols - 1
            flxPI.col = iCol
            flxPI.CellBackColor = vbWhite
      Next iCol
      flxPI.TextMatrix(flxPI.row, iXflxPI) = ""
      flxPI.row = 0
      'cmdEdit(1).Enabled = False                 'At the saving time system will know it PI is in EDIT mode
      cmdNCList.Enabled = False
      'cmdNCList.SetFocus
      txtUnit(0).Enabled = False
      txtUnit(0).text = ""
      cmdUnitList.Enabled = False
      txtJobNo.Enabled = False
      txtJobNo.text = ""
      cmdJobNo(0).Enabled = False
      txtNet_(0).Enabled = False
      txtNet_(0).text = ""
      txtNC(0).Enabled = False
      cmdNCList.Enabled = False
      cmdUpdate(1).Enabled = False
      cmdTaxList(0).Enabled = False
      cmdDeptList.Enabled = False
      txtDetails_(0).Enabled = False
      txtDetails_(0).text = ""
      txtVat_(0).Enabled = False
      txtVat_(0).text = ""
      cmdSchedules(0).Enabled = False
      cmdUpdate(2).Enabled = False
      cmdUpdate(1).Enabled = False
      cmdaddnewline.Enabled = True
      cmdaddnewline.SetFocus
      cmdEdit(0).Enabled = True
      chkRecover.Enabled = False
      txtRecoverable(0).Enabled = False
       'anol 17
      bEditMode = False
      'End of addition
   End If
    cmdUpdate(1).Enabled = False
End Sub

Private Sub UpdateTotalPICN()
   Dim i As Integer

   txtPICNNet.text = "0"
   txtPICNVat.text = "0"
   txtPICNTotal.text = "0"

   For i = 1 To flxPI.Rows - 1
      txtPICNNet.text = Val(txtPICNNet.text) + Val(flxPI.TextMatrix(i, 12))
      txtPICNVat.text = Val(txtPICNVat.text) + Val(flxPI.TextMatrix(i, 14))
      txtPICNTotal.text = Val(txtPICNTotal.text) + Val(flxPI.TextMatrix(i, 15))
   Next i

   txtPICNNet.text = Format(txtPICNNet.text, "0.00")
   txtPICNVat.text = Format(txtPICNVat.text, "0.00")
   txtPICNTotal.text = Format(txtPICNTotal.text, "0.00")
End Sub

Private Sub cmdView_Click()
    'View the PI/PC
    'added by  anol 02 Dec 2015
    'IPEdit variable should re evalute it is not getting the correct result issu 652
        Dim rCount As Integer
        Dim iIncDec As Integer
        Dim iRow As Integer
        For rCount = 1 To flxPurchase.Rows - 1
            If flxPurchase.TextMatrix(rCount, 1) = "X" Then
                iIncDec = iIncDec + 1
                iPIEdit = rCount
            End If
        Next
        If iIncDec <> 1 Then
            MsgBox "Please select one PI/PC only.", vbInformation + vbOKOnly, "PI/PC Selection"
            chkSelectAllDemands.Value = 0
            For rCount = 1 To flxPurchase.Rows - 1
                If flxPurchase.TextMatrix(rCount, 1) = "X" Then
                   flxPurchase.TextMatrix(rCount, 1) = ""
                   flxPurchase.row = rCount
                     For iRow = 1 To flxPurchase.Cols - 1
                       flxPurchase.col = iRow
                       flxPurchase.CellBackColor = RGB(255, 255, 255)
                    Next iRow
                End If
            Next
        
            Exit Sub
        End If
        
        
      If iPIEdit = 0 Then Exit Sub
      Dim X As Byte
      LoadAttachmentFiles cmbFiles, flxPurchase.TextMatrix(iPIEdit, 0), "PI"
      'fraLay(1).Enabled = False
      cmdDelete.Enabled = False
      cmdEdit(0).Enabled = False
      cmdOpenFileView.Visible = True
      cmdviewMenu.Visible = False
      


      
      cmdEdit(1).Enabled = False
      cmdNCList.Enabled = False
      txtUnit(0).Enabled = False
      txtUnit(0).text = ""
      cmdUnitList.Enabled = False
      txtJobNo.Enabled = False
      txtJobNo.text = ""
      cmdJobNo(0).Enabled = False
      txtNet_(0).Enabled = False
      txtNet_(0).text = ""
      txtNC(0).Enabled = False
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtTotal.text = ""
      cmdNCList.Enabled = False
      cmdUpdate(1).Enabled = False
      cmdTaxList(0).Enabled = False
      cmdDeptList.Enabled = False
      txtDetails_(0).Enabled = False
      txtDetails_(0).text = ""
      txtVat_(0).Enabled = False
      txtVat_(0).text = ""
      cmdSchedules(0).Enabled = False
      txtRecoverable(0).Enabled = False
      chkRecover.Enabled = False
      cmdUpdate(2).Enabled = False
      cmdUpdate(1).Enabled = False
      cmdSavePIRef.Visible = True
'      flxPI.ColWidth(11) = Label7(8).Left - Label7(7).Left - 200 '"Details"
      With flxPurchase
            If .TextMatrix(iPIEdit, 20) = "SUPPLIER" Then
               cmbSC.ListIndex = 0
            ElseIf .TextMatrix(iPIEdit, 20) = "CLIENT" Then
               cmbSC.ListIndex = 1
            ElseIf .TextMatrix(iPIEdit, 20) = "LLORD" Then
               cmbSC.ListIndex = 3
            ElseIf .TextMatrix(iPIEdit, 20) = "AGENT" Then
               cmbSC.ListIndex = 2
            End If
            txtTransType.text = .TextMatrix(iPIEdit, 3)
            txtSupplierID.text = .TextMatrix(iPIEdit, 5)
            cmdACList(0).Enabled = IIf(.TextMatrix(iPIEdit, 9) <> .TextMatrix(iPIEdit, 12), False, True)
            txtDate.text = .TextMatrix(iPIEdit, 4)
            txtDueDate.text = Format(.TextMatrix(iPIEdit, 13), "dd/mm/yyyy")
            txtSupplierName.text = .TextMatrix(iPIEdit, 6)
            txtReference.text = .TextMatrix(iPIEdit, 7)
            txtClientID.text = .TextMatrix(iPIEdit, 14)
            txtProperty.text = .TextMatrix(iPIEdit, 11)
            lblPostingDate.ToolTipText = .TextMatrix(iPIEdit, 16)
            
            LoadSplit4Edit flxPI, 0
            fraLay(0).Left = 120
            fraLay(0).Top = 360
            cmdNew(0).Enabled = False
            fraLay(1).Caption = "Transaction ID: " & .TextMatrix(iPIEdit, 2)
            cmdSavePI.Enabled = False  'changes to false by anol 13 July 2015
            If .TextMatrix(iPIEdit, 17) = "" Then
               cmdPO.Enabled = False
            Else
               cmdPO.Enabled = True
            End If
      End With
    'Now code for locking the controls
    cmdaddnewline.Enabled = False
    cmdUpdate(1).Enabled = False
    
          txtDate.Enabled = False
        chkIsMgtFee.Enabled = False
        cmdClientSerc.Enabled = False
        cmdTypeList.Enabled = True
        cmdACList(0).Enabled = False
        cmbSC.Enabled = False
        txtDueDate.Enabled = False
        cmdSavePI.Visible = False
        
   
End Sub

Private Sub cmdviewMenu_Click()
    Frame17.Visible = True
End Sub

Private Sub cmeRevereseAllocation_Click()
   If txtSPSupplier.text = "" Then
      ShowMsgInTaskBar "Please select a supplier.", , "N"
      Exit Sub
   End If

'   Me.Hide
    LoadForm frmRevPayment
    frmRevPayment.LoadFlxCrPoA
'   frmRevPayment.Show
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub



Private Sub flxClient_Click()
' sTextBox = "PICLIENTID"
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rslandlord As New ADODB.Recordset
    tabPurExp.Enabled = True
    tabPayment.Enabled = True
    If sTextBox = "PICLIENTID" Then
            If cmbSC.text = "Landlord" Then
                rslandlord.Open "Select * from PropertyLandlord L ,Property P where L.PropertyID=P.PropertyID and clientID ='" & _
                            flxClient.TextMatrix(flxClient.row, 0) & "'", adoConn, adOpenStatic, adLockReadOnly
                If rslandlord.EOF Then
                    MsgBox "Landlord is not linked to this client (via the property)"
                    txtClientID.text = ""
                    txtClientID.Tag = ""
                    picClient.Visible = False
                    FocusControl cmdClientSerc
                    Exit Sub
                End If
                rslandlord.Close
            End If
            txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
'            'Removing all previous preset vat/tax Code
'            lblVatCode(0).Tag = ""
'            nTaxCode = 0
'            lblVatCode(0).Caption = ""
            
            Dim strTemp As String
            txtClientID.ForeColor = vbBlack
            If Len(Trim(txtClientID.text)) > 0 Then
                strTemp = isControlAccountSet(txtClientID.text)
                If Len(strTemp) > 0 Then
                    MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & _
                    vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
                    strTemp = ""
                    picClient.Visible = False
                    txtClientID.ForeColor = vbRed
                    Exit Sub
                End If
            End If
'this is the loading structure
'             flxClient.TextMatrix(rRow, 3) = IIf(IsNull(rstRec.Fields("VAT_CODE").Value), "", rstRec.Fields("VAT_CODE").Value)
'             flxClient.TextMatrix(rRow, 4) = IIf(IsNull(rstRec.Fields("VAT_ID").Value), "", rstRec.Fields("VAT_ID").Value)
'             flxClient.TextMatrix(rRow, 5) = IIf(IsNull(rstRec.Fields("VAT_RATE").Value), "", rstRec.Fields("VAT_RATE").Value)
            If flxClient.TextMatrix(flxClient.row, 4) <> "" Then    'flxClient.TextMatrix(flxClient.row, 4) when this is empty you cannot load next numerical value resulting in an error
                lastVat_Code_fromClient = flxClient.TextMatrix(flxClient.row, 3)
                lastVat_ID_fromClient = flxClient.TextMatrix(flxClient.row, 4)
                lastVat_Rate_fromClient = flxClient.TextMatrix(flxClient.row, 5)
'                lblVatCode(0).Caption = lastVat_Code_fromClient
'                lblVatCode(0).Tag = lastVat_ID_fromClient
'                nTaxCode = lastVat_Rate_fromClient
            Else
                lastVat_Code_fromClient = ""
                lastVat_ID_fromClient = -1
                lastVat_Rate_fromClient = 0
            End If
            FocusControl cmdTypeList
    ElseIf sTextBox = "1" Then 'filter on client PI list
            txtIDClient.text = flxClient.TextMatrix(flxClient.row, 0)
            txtIDClient.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtPropID.text = "ALL"
            chkProperty.Value = 0
            Call LoadFlxPurchase(adoConn)
            fmeLoading.Visible = False
            cmdOpClient.SetFocus
        
    ElseIf sTextBox = "Payment" Then  'Purchase payment tab
        cGridSPTotal = 0 'clearing global form variable  for changing new client
        txtSPaymentTotal.text = "0.00"
        txtPaymentEntered.text = "0.00"
        txtPaymentTotal.text = "0.00"
        txtDiffPay.text = "0.00"
        tabPayment.Enabled = True
        txtClientIDPurPay.text = flxClient.TextMatrix(flxClient.row, 0)
        txtClientIDPurPay.ForeColor = vbBlack
        If Trim(txtClientIDPurPay.text) <> "" Then
            strTemp = isControlAccountSet(txtClientIDPurPay.text)
            If Len(strTemp) > 0 Then
                MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & _
                vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
                strTemp = ""
                picClient.Visible = False
                txtClientIDPurPay.ForeColor = vbRed
                Exit Sub
            End If
            
            frmMMain.frmPI_SupplierBalanceByCL_isUptoDate = False
            If SupplierIDFromComBo <> "" Then
                'Unlock Previous locked item
                fmeLoading.Visible = True
                fmeLoading.Refresh
                adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                       "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
                LoadFlxSPayment adoConn 'issue 638
                LoadFlxSCrPoA adoConn
                fmeLoading.Visible = False
                DisplaylockScreen adoConn
                
            End If
            Call txtClientIDPurPay_Change_made
            FocusControl cmdSupplierType
             Call LoadDefaultBankAC(adoConn) 'when you select a client it is loading a deafult Bank AC
             Call updateBankBalance
        Else 'when client is empty
            txtBankCode.text = ""
            txtBankAc.text = ""
        End If
       
     ElseIf sTextBox = "PIHistory" Then 'PIHistory
        txtClientIdlist.text = flxClient.TextMatrix(flxClient.row, 0)
        txtClientIdlist.Tag = flxClient.TextMatrix(flxClient.row, 1)
        txtPropertyIDHist.text = "ALL"
        chkPropertyHist.Value = 0
        LoadFlxPurchHistory adoConn, ""
        cmdOClientList.SetFocus
    ElseIf sTextBox = "PaymentHistory" Then
        txtPurchasePaymentHistory.text = flxClient.TextMatrix(flxClient.row, 0)
        '
        If Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxPurchPPHistory(adoConn, "3")
             searchResultOn = True
             'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
             'cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
             Call LoadFlxPurchPPHistory(adoConn, "4")
             searchResultOn = True
        Else
            LoadFlxPurchPPHistory adoConn, ""
        End If
                        
                        
        cmdPurchasePaymentHistory.SetFocus
    ElseIf sTextBox = "BankAcPay" Then
        txtBankCode.text = flxClient.TextMatrix(flxClient.row, 1)
        txtBankAc.text = flxClient.TextMatrix(flxClient.row, 2)
        Call updateBankBalance
        FocusControl cmdAmountType
    End If
    adoConn.Close
    Set adoConn = Nothing
    picClient.Visible = False
End Sub

Public Sub updateBankBalance()
        Dim adoConn As New ADODB.Connection
        Dim adoRST As New ADODB.Recordset
        adoConn.Open getConnectionString
        Dim Balance As Double
        Dim szSQL As String
   ' find current Balance for the selected bank account and selected client ID by anol 2023-05-24
   szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) as AMTT from (" & _
            "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & txtBankCode & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & txtClientIDPurPay.text & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & txtBankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & txtClientIDPurPay & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & txtBankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & txtClientIDPurPay & "'   group by Type )"
                       
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
      txtBankBal1.text = IIf(IsNull(adoRST.Fields.Item("AMTT").Value), 0, adoRST.Fields.Item("AMTT").Value)
      txtBankBal1.text = Format(txtBankBal1.text, "0.00")
   End If
   adoRST.Close
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where isDeleted=false and BankCode='" & txtBankCode & "' and ClientID='" & txtClientIDPurPay & "'"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        txtRetentions1.text = IIf(IsNull(adoRST.Fields.Item("DAmt").Value), 0, adoRST.Fields.Item("DAmt").Value)   'adoRst.Fields.Item("DAmt").Value
        txtRetentions1.text = Format(txtRetentions1.text, "0.00")
    End If
    adoRST.Close
    
    
'   szSQL = "SELECT * from tlbClientBanks where NominalCode and client_ID='" & txtClientList & "'"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   If Not adoRst.EOF Then
'   End If
   txtAvailableBankBal1.text = Val(txtBankBal1.text) - Val(txtRetentions1.text)
   txtAvailableBankBal1.text = Format(txtAvailableBankBal1.text, "0.00")
'   txtAvailableBankBal1
'   txtRetentions1
   adoConn.Close
End Sub
Private Function SupplierIDFromComBo() As String
On Error GoTo Err
    SupplierIDFromComBo = ""
    SupplierIDFromComBo = txtSPSupplier.Tag

Exit Function
Err:

End Function
Private Sub flxClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If flxClient.row = 1 And KeyCode = vbKeyUp Then
        txtSearchClientID.SetFocus
    End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    If KeyAscii = 13 Then
'         tabPurExp.Enabled = True
'        If tabPurExp.Tab = 0 Then
'            If fraLay(0).Top = 360 Then
'                txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
'                cmdTypeList.SetFocus
'             Else
'                txtIDClient.text = flxClient.TextMatrix(flxClient.row, 0)
'                Call txtIDClient_change_made
'                txtPropID.text = "ALL"
'                cmdOpClient.SetFocus
'             End If
'        End If
'        If tabPurExp.Tab = 1 Then
'            tabPayment.Enabled = True
'            txtClientIDPurPay = flxClient.TextMatrix(flxClient.row, 0)
'            Call txtClientIDPurPay_Change_made
'            cmdACType.SetFocus
'        End If
'        If tabPurExp.Tab = 2 Then
'            txtClientIdlist.text = flxClient.TextMatrix(flxClient.row, 0)
'            LoadFlxPurchHistory adoconn, ""
'            cmdOClientList.SetFocus
'        End If
'        picClient.Visible = False
'    End If
'    adoconn.Close
        If KeyAscii = 13 Then
            flxClient_Click
        End If
End Sub



Private Sub flxPI_Click()
    SelectOnly1RowFlxGrid flxPI, flxPI.row, iXflxPI
End Sub

Private Sub flxPI_EnterCell()
'added by anol 17 Aug 2015
    'Select1RowFlxGrid flxPI, flxPI.row, iXflxPI
End Sub

Private Sub flxPI_SelChange()
     SelectOnly1RowFlxGrid flxPI, flxPI.row, iXflxPI
End Sub

Private Function SelectFlxGridRowNocolor(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
        ' conFlxGrid.CellBackColor = RGB(255, 255, 255)
      'Next iRow
      SelectFlxGridRowNocolor = -1
   Else
      conFlxGrid.TextMatrix(iSelRow, iColID) = "X"

      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
         'conFlxGrid.CellBackColor = RGB(174, 179, 233)
      'Next iRow
      SelectFlxGridRowNocolor = 1
   End If
End Function
Private Sub flxPurchase_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection

   If flxPurchase.TextMatrix(flxPurchase.row, 0) = "" Then Exit Sub
   If flxPurchase.RowHeight(flxPurchase.row) = 0 Then
      iPIEdit = 0
      Exit Sub
   End If

   adoConn.Open getConnectionString

'   HighLightRowFlxGrid flxPurchase, flxPurchase.row
   SelectFlxGridRowNocolor 1, flxPurchase, flxPurchase.row

   iPIEdit = flxPurchase.row

   ConfigFlxPurchaseSplit

   With flxPurchaseSplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyID, U.UnitNumber, " & _
                  "P.PropertyName, U.UnitName " & _
              "FROM (tblPurInvSRec AS S " & _
                  "LEFT JOIN  Property AS P ON S.TRANS = P.PropertyID) " & _
                  "LEFT JOIN Units AS U ON S.UNIT_ID = U.UnitNumber " & _
              "WHERE S.ParentID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' " & _
              "ORDER BY TRAN_ID;"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = flxPurchase.TextMatrix(flxPurchase.row, 14)
                                 
         .TextMatrix(iRow, 3) = flxPurchase.TextMatrix(flxPurchase.row, 11)
         .TextMatrix(iRow, 4) = IIf(IsNull(adoInvSp.Fields.Item("UnitName").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyName").Value), "", _
                                 adoInvSp.Fields.Item("PropertyName").Value), adoInvSp.Fields.Item("UnitName").Value)
         .TextMatrix(iRow, 5) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 7) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 8) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 9) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 10) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 11) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 12) = adoInvSp.Fields.Item("RecoverablePt").Value & "%"

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoConn.Close
   Set adoInvSp = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadpurchaseSplit(adoConn As ADODB.Connection)
    Dim szSQL As String, iRow As Integer
    Dim adoInvSp As New ADODB.Recordset

    ConfigFlxPurchaseSplit

   With flxPurchaseSplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyID, U.UnitNumber, " & _
                  "P.PropertyName, U.UnitName " & _
              "FROM (tblPurInvSRec AS S " & _
                  "LEFT JOIN  Property AS P ON S.TRANS = P.PropertyID) " & _
                  "LEFT JOIN Units AS U ON S.UNIT_ID = U.UnitNumber " & _
              "WHERE S.ParentID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' " & _
              "ORDER BY TRAN_ID;"

      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = IIf(IsNull(adoInvSp.Fields.Item("UnitNumber").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyID").Value), "", _
                                 adoInvSp.Fields.Item("PropertyID").Value), adoInvSp.Fields.Item("UnitNumber").Value)
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("UnitName").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyName").Value), "", _
                                 adoInvSp.Fields.Item("PropertyName").Value), adoInvSp.Fields.Item("UnitName").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 9) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 10) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("RecoverablePt").Value & "%"

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

'   adoConn.Close
   Set adoInvSp = Nothing
'   Set adoConn = Nothing
End Sub
Private Sub flxPurchase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraInvCrChoice.Visible Then cmdManualDmdCancel_Click
End Sub

Private Sub flxPurchaseSplit_DblClick()
   'Resolved by BOSL
    'issue 453-1:High The recoverable amount window which appears if you double click the
    'recoverable field in the lower grid causes the program to crash if it is canceled
    'Modified by anol 13 Aug 2014
    On Error GoTo Cancel
    Dim r As Single

   If flxPurchaseSplit.col = 11 Then
      r = InputBox("Please enter the value of recoverable amount:", "Recoverable Amount")
      
      If r > 100 Then
         ShowMsgInTaskBar "Recoverable amount cannot be more than 100%.", "Y", "N"
         Exit Sub
      End If
      If r < 0 Then
         ShowMsgInTaskBar "Recoverable amount cannot be less than 0.", "Y", "N"
         Exit Sub
      End If

      flxPurchaseSplit.TextMatrix(flxPurchaseSplit.row, 11) = r

      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString

      adoConn.Execute "UPDATE tblPurInvSRec " & _
                      "SET    RecoverablePt = " & r & " " & _
                      "WHERE  MY_ID = '" & flxPurchaseSplit.TextMatrix(flxPurchaseSplit.row, 0) & "';"
      adoConn.Close
      Set adoConn = Nothing
   End If
Cancel:
End Sub

Private Sub flxPurchHistory_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection

   If flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) = "" Then Exit Sub
   If flxPurchHistory.RowHeight(flxPurchHistory.row) = 0 Then Exit Sub
   'below line has been added by anol
   'Date 10 Jun 2015
   'issue 0000572: Purchases and expenses - Reverse postings to history not present
   Call SelectFlxGridRow(1, flxPurchHistory, flxPurchHistory.RowSel)
   adoConn.Open getConnectionString

'   HighLightRowFlxGrid flxPurchHistory, flxPurchHistory.row

   ConfigFlxSplit flxPurchHistorySplit, 29

   With flxPurchHistorySplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyName AS XX " & _
              "FROM tblPurInvSRec AS S LEFT JOIN Property AS P ON S.TRANS = P.PropertyID " & _
              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' " & _
              "ORDER BY TRAN_ID;"
'
'      szSQL = szSQL + " UNION "
'
'      szSQL = szSQL + _
'              "SELECT DISTINCT S.*, U.UnitName AS XX " & _
'              "FROM tblPurInvSRec AS S, Units AS U " & _
'              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' AND " & _
'                  "S.TRANS = 'Unit' AND S.UNIT_ID = U.UnitNumber " & _
'              "ORDER BY TRAN_ID"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("XX").Value), "", adoInvSp.Fields.Item("XX").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = adoInvSp.Fields.Item("DEPT_ID").Value
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = adoInvSp.Fields.Item("NET_AMOUNT").Value
         .TextMatrix(iRow, 9) = adoInvSp.Fields.Item("VAT").Value
         .TextMatrix(iRow, 10) = adoInvSp.Fields.Item("TOTAL_AMOUNT").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoConn.Close
   Set adoInvSp = Nothing
   Set adoConn = Nothing
End Sub

Private Sub flxPurchPPHistory_Click()
   Dim adoConn    As New ADODB.Connection
   Dim adoInvSp   As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer

   If flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 0) = "" Then Exit Sub
   If flxPurchPPHistory.RowHeight(flxPurchPPHistory.row) = 0 Then Exit Sub
   
   ConfigFlxSplit flxPurchPPHistorySplit, 39

   adoConn.Open getConnectionString

   With flxPurchPPHistorySplit
'         Adding the split of the header
              szSQL = "SELECT S.*, I.CL_ID, I.PropertyID,P.NominalCode as NC " & _
              "FROM (PayTransactions AS S INNER JOIN " & _
                    "tlbPayment AS P ON S.ToTran = P.TransactionID) LEFT JOIN " & _
                    "tblPurInv AS I ON P.PI = I.MY_ID " & _
              "WHERE S.Deleteflag=False and S.FromTran = " & flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 0) & " " & _
              "ORDER BY S.TransactionID"
'Debug.Print szSQL
'      szSQL = szSQL + " UNION "
'
'      szSQL = szSQL + _
'              "SELECT DISTINCT S.* " & _
'              "FROM PayTransactions AS S, Units AS U " & _
'              "WHERE S.FromTran = " & flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 0) & " " & _
'              "ORDER BY TransactionID"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("TransactionID").Value
         .TextMatrix(iRow, 1) = iRow 'adoInvSp.Fields.Item("SlNumber").Value
         .TextMatrix(iRow, 2) = IIf(IsNull(adoInvSp.Fields.Item("CL_ID").Value), "", adoInvSp.Fields.Item("CL_ID").Value)
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("PropertyID").Value), "", adoInvSp.Fields.Item("PropertyID").Value) 'adoInvSp.Fields.Item("PropertyID").Value
         .TextMatrix(iRow, 4) = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 15) 'adoInvSp.Fields.Item("NominalCode").Value
         .TextMatrix(iRow, 7) = flxPurchPPHistory.TextMatrix(flxPurchPPHistory.row, 8)
         .TextMatrix(iRow, 8) = Format(adoInvSp.Fields.Item("PaymentAmount").Value, "0.00")
         .TextMatrix(iRow, 9) = "0.00" 'adoInvSp.Fields.Item("VAT").Value
         .TextMatrix(iRow, 10) = Format(adoInvSp.Fields.Item("PaymentAmount").Value, "0.00")

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoConn.Close
   Set adoInvSp = Nothing
   Set adoConn = Nothing
End Sub

Private Sub flxSPayment_Click()
   If flxSPayment.row = 0 Then Exit Sub

   Dim i             As Integer
   Dim iFlxSPayCol   As Integer
   Dim iCurRowHeight As Integer

   sEditPPR = 1
   If Left(flxSPayment.TextMatrix(flxSPayment.row, 1), 3) = "PPR" Then
      cmdEditPayment.Enabled = True
   Else
      cmdEditPayment.Enabled = False
   End If
   If flxSPayment.col = 0 And flxSPayment.TextMatrix(flxSPayment.row, 0) = "+" Then          'Expanding the grid
      flxSPayment.TextMatrix(flxSPayment.row, 0) = ">"
      iCurRowHeight = flxSPayment.RowHeight(flxSPayment.row)
      i = 1

      While flxSPayment.TextMatrix(flxSPayment.row + i, 0) = "-"
         flxSPayment.RowHeight(flxSPayment.row + i) = iCurRowHeight
         i = i + 1
         If (flxSPayment.row + i) = flxSPayment.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   If flxSPayment.col = 0 And flxSPayment.TextMatrix(flxSPayment.row, 0) = ">" Then          'Squeezing the grid
      flxSPayment.TextMatrix(flxSPayment.row, 0) = "+"
      i = 1
      While flxSPayment.TextMatrix(flxSPayment.row + i, 0) = "-"
         flxSPayment.RowHeight(flxSPayment.row + i) = 0
         i = i + 1
         If (flxSPayment.row + i) = flxSPayment.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   
   
End Sub

Private Sub flxSPayment_dblClick()
    If flxSPayment.RowHeight(flxSPayment.row) = 0 Then Exit Sub
    Dim i As Integer
     'added by anol for locking issue 749 will not be editable on double click
   flxSPayment.col = 0
   If flxSPayment.CellBackColor = vbRed Then
        'MsgBox "Selected invoice is locked by another user. Please wait untill other user release this record.", vbInformation, "Warning"
        MsgBox "The selected invoice is currently locked by  '" & flxSPayment.TextMatrix(flxSPayment.row, 25) & "' on '" & _
                    flxSPayment.TextMatrix(flxSPayment.row, 26) & "' in the '" & flxSPayment.TextMatrix(flxSPayment.row, 27) & "'" & vbCrLf & "" & _
                        "screen for the Client : '" & flxSPayment.TextMatrix(flxSPayment.row, 23) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
        
'        MsgBox "The selected invoice is currently locked by '" & IIf(IsNull(adoPay("WindowsUserName").Value), "", adoPay("WindowsUserName").Value) & _
'                "' on '" & IIf(IsNull(adoPay("MachineName").Value), "", adoPay("MachineName").Value) & "' in the '" & IIf(IsNull(adoPay("Module").Value), "", adoPay("Module").Value) & "'" & vbCrLf & "" & _
'                        "screen for the Client '" & IIf(IsNull(adoPay("ClientID").Value), "", adoPay("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
        Exit Sub
   End If
   
   
   iFlxSPayCol = 10
   flxSPayment.col = iFlxSPayCol

   szUndoText = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
   'added by anol 20170215
        If szUndoText = "" Then Exit Sub
        If Not lblAllocating(1).Visible And flxSPayment.TextMatrix(flxSPayment.row, 2) <> "ADJI" And _
                 Val(txtSPaymentTotal.text) >= 0 And Val(txtAllocatedDiff(1).text) = 0 And cmdPayAllocate.Caption <> "All&ocation Only" Then
                 MsgBox " You must first select a credit amount from the grid below before you can enter any amounts in this grid.", vbInformation + vbOKOnly, "Allocation"
                 Exit Sub
        End If
        
   If cmdPayAllocate.Caption <> "&Payment Only" And flxSPayment.TextMatrix(flxSPayment.row, 2) <> "ADJI" And _
      Val(txtSPaymentTotal.text) >= 0 Then
         txtSPayment.Top = flxSPayment.CellTop + flxSPayment.Top
         txtSPayment.Left = flxSPayment.CellLeft + flxSPayment.Left
         txtSPayment.Width = flxSPayment.ColWidth(iFlxSPayCol)
         'Modified by anol 05 July 2015
         If flxSPayment.RowHeight(flxSPayment.row) = 0 Then
            txtSPayment.Visible = False
         Else
            txtSPayment.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
         End If
         txtSPayment.text = flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)
   
         If flxSPayment.TextMatrix(flxSPayment.row, 0) = "-" And CCur(txtSPayment.text) > 0 Then
            For i = flxSPayment.row To 1 Step -1
               If flxSPayment.TextMatrix(i, 0) = ">" Then Exit For
            Next i
            flxSPayment.TextMatrix(i, 10) = Format(CCur(flxSPayment.TextMatrix(i, 10)) - CCur(txtSPayment.text), "0.00")
         End If
         'Modified by anol 05 July 2015
         If flxSPayment.RowHeight(flxSPayment.row) <> 0 Then
            txtSPayment.Visible = True
            txtSPayment.SetFocus
            SelTxtInCtrl txtSPayment
         End If
   End If

'  ALLOCATION - Place the txtCrPayment text box in the grid to allocate against invoice
   If lblAllocating(1).Visible And Val(flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)) = 0 And Val(txtAllocatedDiff(1).text) > 0 Then
      If (InStr(lblAllocating(1).Caption, "ADJ") > 0 And InStr(flxSPayment.TextMatrix(flxSPayment.row, 2), "ADJ") > 0) Or _
         (InStr(lblAllocating(1).Caption, "ADJ") = 0 And InStr(flxSPayment.TextMatrix(flxSPayment.row, 2), "ADJ") = 0) Then
         txtCrPayment.Top = flxSPayment.CellTop + flxSPayment.Top
         txtCrPayment.Left = flxSPayment.CellLeft + flxSPayment.Left
         txtCrPayment.Width = flxSPayment.ColWidth(iFlxSPayCol)
         txtCrPayment.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
         txtCrPayment.text = Format(flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol), "0.00")
         txtCrPayment.Visible = True
         txtCrPayment.SetFocus
         SelTxtInCtrl txtCrPayment
         txtCrPayment.BackColor = RGB(233, 232, 155)
         Label10(7).Caption = flxSPayment.row
         txtAllocatedDiff(1).text = Format(Val(txtAllocatedDiff(1).text) + Val(flxSPayment.TextMatrix(flxSPayment.row, iFlxSPayCol)), "0.00")
      Else
         If InStr(lblAllocating(1).Caption, "ADJ") > 0 Then
            MsgBox "               Please select an Adjustment Invoice (ADJI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         Else
            MsgBox "                    Please select a Purchase Invoice (SI) to allocate against." & Chr(13) & _
                   "You can only allocate an Adjustment Credit (ADJC) against an Adjustment Invoice (ADJI).", vbCritical + vbOKOnly, "Allocation"
         End If
      End If
   End If
     '10/02/2017  Modification of allocation process
'
'1\ The allocation process needs to be modified so that when the user does an allocation, and saves it, the allocation screen remains open and the user can make the next allocation.'
'2\ When the user double clicks in the amount field of a line to allocate in the debit grid, the system should populate the full amount in the amount field in overtype mode to allow user to modify. On lost focus entry should be accepted as currently.
'Conditions
'For one credit line and one debit line respectively.'
'1/ If credit amount outstanding is less than or equal to the debit amount it will populate the whole of the credit amount'
'2/  If credit amount is greater than the debit amount outstanding  it will populate the whole of the debit amount'
'U:\Support Issues\Austin Chambers\Dyson Properties\Modification of allocation process'
    If cmdPayAllocate.Caption = "&Payment Only" Then
        txtCrPayment.text = MinAmtAllocate
    End If
End Sub
Private Function MinAmtAllocate() As String
   'Written by anol 20170216
   ' find the amount
  ' 1/ If credit amount outstanding is less than or equal to the debit amount it will populate the whole of the credit amount
  '2/  If credit amount is greater than the debit amount outstanding  it will populate the whole of the debit amount

   Dim iRow As Integer, dDr As Double, dCr As Double

   MinAmtAllocate = 0
    'the amount that is currently allocating
   dDr = CDbl(flxSPayment.TextMatrix(flxSPayment.row, 9))
   If CDbl(flxSPayment.TextMatrix(flxSPayment.row, 10)) > 0 Then
        MinAmtAllocate = Format(RoundingNumber(CDbl(flxSPayment.TextMatrix(flxSPayment.row, 10)), 2), "0.00")
        Exit Function
   End If
   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 0) <> "-" Then
      'the receipt which has been already distributed needs to be minus from the credit
         dCr = dCr - CDbl(flxSPayment.TextMatrix(iRow, 10))
      End If
   Next iRow
   For iRow = 1 To flxSCrPoA.Rows - 1
      If flxSCrPoA.TextMatrix(iRow, 0) <> "-" Then
         dCr = dCr + CDbl(flxSCrPoA.TextMatrix(iRow, 9))
      End If
   Next iRow
  
   If dCr <= dDr Then MinAmtAllocate = dCr
   If dCr > dDr Then MinAmtAllocate = dDr
   MinAmtAllocate = Format(RoundingNumber(MinAmtAllocate, 2), "0.00")
End Function
Private Function CalculateDiff() As Double
  'written by anol 20170215 to find the rest amount to be allocated
   Dim iRow As Integer, dDr As Double, dCr As Double

   CalculateDiff = 0
   
   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 0) <> "-" Then
      'the receipt which has been already distributed needs to be minus from the credit
         dDr = dDr + CDbl(flxSPayment.TextMatrix(iRow, 10))
      End If
   Next iRow
   For iRow = 1 To flxSCrPoA.Rows - 1
      If flxSCrPoA.TextMatrix(iRow, 0) <> "-" Then
         dCr = dCr + CDbl(flxSCrPoA.TextMatrix(iRow, 9))
      End If
   Next iRow
   CalculateDiff = dCr - dDr
   
End Function
Private Sub flxSPayment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSPayment_dblClick
    End If
End Sub

Private Sub flxSupplier_Click(Index As Integer)
'    On Error GoTo ERR

    Dim adoConn As New ADODB.Connection
    Dim rstVat As New ADODB.Recordset
    Dim szSQL As String
    If sTextBox = "PIHIST" Then
        txtSupplierSearc.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        txtSupplierSearc.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        fraList.Visible = False
        tabPurExp.Enabled = True
        tabPayment.Enabled = True
        FocusControl cmdOpSupSearch
        adoConn.Open getConnectionString
        Call LoadFlxPurchHistory(adoConn, "")
        adoConn.Close
        Set adoConn = Nothing
        fmeLoading.Visible = False
        Exit Sub
    End If
    If sTextBox = "PROPERTYFILTER" Then
        tabPurExp.Enabled = True
        txtPropID.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        txtPropID.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        
        fraList.Visible = False
        adoConn.Open getConnectionString
        Call LoadFlxPurchaseFilter(adoConn, "")
        adoConn.Close
        Set adoConn = Nothing
        fmeLoading.Visible = False
        Exit Sub
    End If
    
    If sTextBox = "PAYHIST" Then
        txtSupSearchHis.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        fraList.Visible = False
        tabPurExp.Enabled = True
        tabPayment.Enabled = True
        FocusControl cmdOpSupSearch
        adoConn.Open getConnectionString
        Call LoadFlxPurchPPHistory(adoConn, "")
        adoConn.Close
        Set adoConn = Nothing
        fmeLoading.Visible = False
        Exit Sub
    End If

   If sTextBox = "PILIST" Then
      txtSupplier.text = flxSupplier(2).TextMatrix(flxSupplier(2).row, 1)
      ' SortTheGrid function is not a good function when you have a big amount of data in the list. It will slow you down. Need to replace it with query.
      SortTheGrid flxPurchase, txtIDClient, txtPropID, txtSupplier
      flxPurchaseSplit.Clear
      cmdGridUnitLookup_Click (2)
      Exit Sub
   End If

   If Index = 1 Then 'this is flxgrid index
        'added by anol 20 July 2015 issue 571
        'The details displayed in purchase payment for 1/ Payment Type 2/ Reference
        'and 3/ Total Amt £ should be cleared every time a user changes supplier
        txtPayAmtType.text = ""
        txtPayAmtType.Tag = ""
        txtSPReference.text = ""
        txtSPaymentTotal.text = "0.00"
        
        'end of modification
        'cmbSPSupplier.Value = flxSupplier(1).TextMatrix(flxSupplier(1).row, 1)
        txtSPSupplier.text = flxSupplier(1).TextMatrix(flxSupplier(1).row, 2) 'SupplierName
        txtSPSupplier.Tag = flxSupplier(1).TextMatrix(flxSupplier(1).row, 1) 'SupplierID
        strSupplierTypeOnSelection = flxSupplier(1).TextMatrix(flxSupplier(1).row, 0)  'Supplier  Type
        
        cmdGridUnitLookup_Click (1)
        If cmdBankAc.Enabled = True Then
            FocusControl cmdBankAc
        End If
       
        'added by anol 17 Jan 2016
        txtSupAcBal.text = Format(flxSupplier(1).TextMatrix(flxSupplier(1).row, 3), "0.00") ' Format(AccountBalance(adoConn), "0.00")
        If Val(txtSupAcBal.text) >= 0 Then
            txtSupAcBal.ForeColor = vbBlack
        Else
            txtSupAcBal.ForeColor = vbRed
        End If
       
        'End of modification
        'added by anol 20160524
        Call ChangeSupplier
        'End of modification
        Exit Sub
   End If
 
   tabPurExp.Enabled = True

   If sTextBox = "A/C" Then
      bTotalPayTyped = False
      txtSupplierID.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      If cmbSC.text = "Client" Then
         txtClientID.text = txtSupplierID.text
         txtClientID.Locked = True
      Else
         txtClientID.Locked = False
      End If

      txtSupplierName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
'      If txtNC(0).text = "" Then _
'         txtNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
         'below line is added by anol 03 jan 2015
      If txtNC(0).text = "0" Then
         txtNC(0).text = ""
         txtNCName.text = ""
      End If
      FocusControl txtReference
      txtSupplierID.SelStart = Len(txtSupplierID.text)
      iDayTerms = Val(flxSupplier(0).TextMatrix(flxSupplier(0).row, 5))
      txtDueDate.text = DateAdd("d", iDayTerms, Date)
 
      Dim szaTemp() As String
'Rem by anol 2020-10-12 this icose is not needed because we dont want to see the tax values when we click client
'      If InStr(flxSupplier(0).TextMatrix(flxSupplier(0).row, 3), "##") > 0 Then
'         szaTemp = Split(flxSupplier(0).TextMatrix(flxSupplier(0).row, 3), "##") 'column number 3 loading structure IIf(IsNull(rstRst!VAT_CODE) Or rstRst!VAT_CODE = "", "", rstRst!VAT_CODE & "##" & rstRst!VAT_RATE)
'         lblVatCode(0).Caption = szaTemp(0) 'VAT_CODE
'         lblVatCode(0).Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 6) ' Contains vat ID (Numeric)'VAT_CODE
'         If szaTemp(1) = "" Then
'            'nTaxCode = -1 'issue 659 by anol 20181109
'            nTaxCode = 0
'         Else
'            nTaxCode = CDbl(szaTemp(1)) 'VAT_RATE
'         End If
'         sVCFound = 1
'      Else
'         lblVatCode(0).Caption = ""
'         'commented by anol 13 July issue 571 note 1125
'         'txtProperty.text = ""
'         sVCFound = 2
'      End If
       
   End If
   If sTextBox = "PROPERTY" Then

           If cmbSC.text = "Landlord" Then
                adoConn.Open getConnectionString
                Dim rslandlord As New ADODB.Recordset
                rslandlord.Open "Select * from PropertyLandlord L ,Property P where L.PropertyID=P.PropertyID and P.PropertyID ='" & _
                            flxSupplier(0).TextMatrix(flxSupplier(0).row, 0) & "'", adoConn, adOpenStatic, adLockReadOnly
                If rslandlord.EOF Then
                    MsgBox "Landlord is not linked to this property"
                    txtProperty.text = ""
                    txtProperty.Tag = ""
                    fraList.Visible = False
                    FocusControl cmdTypeList
                    Exit Sub
                End If
                rslandlord.Close
                adoConn.Close
                Set adoConn = Nothing
            End If

        txtProperty.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)

            
            If cmdaddnewline.Enabled = True Then
                  cmdaddnewline.SetFocus
            End If

   End If
   If sTextBox = "PROPERTYHIST" Then
        txtPropertyIDHist.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        txtPropertyIDHist.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        adoConn.Open getConnectionString
        Call LoadFlxPurchHistory(adoConn, "")
        
        adoConn.Close
        Set adoConn = Nothing
        fmeLoading.Visible = False
        FocusControl cmdOpenSupp
   End If

   If sTextBox = "UNIT" Then
      txtUnit(tabPurExp.Tab).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      FocusControl cmdDeptList
      txtUnit(tabPurExp.Tab).SelStart = Len(txtUnit(tabPurExp.Tab).text)
      cmdJobNo(0).Enabled = True
   End If
   If sTextBox = "NC" Then
      txtNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      FocusControl cmdJobNo(0)
   End If

   If sTextBox = "Fund" Then
      txtDept(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0) 'FundCode
      txtPFName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1) 'FundName
      txtDept(1).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2) 'FundID
      If flxSupplier(0).TextMatrix(flxSupplier(0).row, 3) = "" Then
        fraList.Visible = False
        Exit Sub
      End If
      iSelectedFundCategoryID = flxSupplier(0).TextMatrix(flxSupplier(0).row, 3) 'fund CategoryCode
      FocusControl cmdNCList
   End If
   
    
   If sTextBox = "VAT" Then
      lblVatCode(tabPurExp.Tab).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      nTaxCode = CSng(flxSupplier(0).TextMatrix(flxSupplier(0).row, 1))
      txtNet__LostFocus (tabPurExp.Tab)
      'Resolved by BOSL
      'issue 453 note 2
      'modified by anol 14 Aug 2014
      If cmdUpdate(1).Enabled = True Then
            cmdUpdate(1).SetFocus
      End If
   End If
   
   If sTextBox = "Schedules" Then
      txtSchedules.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      FocusControl txtDetails_(tabPurExp.Tab)
      txtSchedules.SelStart = Len(txtSchedules.text)
   End If
    
   If sTextBox = "job" Then
      txtJobNo.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      FocusControl cmdSchedules(tabPurExp.Tab)
      txtJobNo.SelStart = Len(txtSchedules.text)
   End If
   If sTextBox = "Bank" Then nTaxCode = TaxRate(1)
   If sTextBox = "VATBank" Then nTaxCode = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)

   tabPayment.Enabled = True
   fraList.Visible = False
  
   Exit Sub
Err:
    MsgBox Err.description
   'MsgBox "Please send a screenshot to PCM of this problem" & vbCrLf & Err.description, vbInformation, "flxSupplier_click: " & strMarkonError
End Sub

Private Sub flxSupplier_KeyPress(Index As Integer, KeyAscii As Integer)
    If sTextBox = "PROPERTYFILTER" And KeyAscii = 13 Then
            tabPurExp.Enabled = True
            txtPropID.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
            fraList.Visible = False
            Exit Sub
    End If

  
'added by anol 12 Dec 2015
    If Index = 0 And tabPurExp.Tab = 2 And sTextBox = "A/C" And KeyAscii = 13 Then
        txtSupplierSearc.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
        fraList.Visible = False
        tabPurExp.Enabled = True
        tabPayment.Enabled = True
        cmdOpenSupp.SetFocus
        Exit Sub
   End If
'    If Index = 0 And tabPurExp.Tab = 0 And KeyAscii = 13 Then
'       tabPurExp.Enabled = True
'       txtPropID.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
'       fraList.Visible = False
'       Exit Sub
'    End If
    'End of addition
   If KeyAscii = 27 Then
      flxSupplier(0).Clear
      flxSupplier(0).Clear

      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2
      fraList.Visible = False
      tabPurExp.Enabled = True
      If sTextBox = "A/C" Then cmdACList(tabPurExp.Tab).SetFocus
      If sTextBox = "UNIT" Or sTextBox = "PROP" Then cmdNCList.SetFocus
      If sTextBox = "NC" Then txtNC(0).SetFocus
      If sTextBox = "Fund" Then txtDept(0).SetFocus
      If sTextBox = "VAT" Then cmdTaxList(tabPurExp.Tab).SetFocus
      Exit Sub
   End If
   If KeyAscii = 13 Then
      flxSupplier_Click (Index)
        If Index = 1 Then
               If sTextBox = "A/C" Then
                    FocusControl cmdACList(tabPurExp.Tab)
               End If
               If sTextBox = "UNIT" Or sTextBox = "PROP" Then
                    FocusControl cmdNCList
               End If
               If sTextBox = "NC" Then
                    FocusControl txtNC(0)
               End If
               If sTextBox = "Fund" Then
                    FocusControl txtDept(0)
                End If
               If sTextBox = "VAT" Then
                    FocusControl cmdTaxList(tabPurExp.Tab)
               End If
      End If
   End If
End Sub

Private Sub DisplaylockScreen(adoConn As ADODB.Connection)
    Dim strSQL As String
   ' Dim adoConn As New ADODB.Connection
    Dim rsLockDialog As New ADODB.Recordset
    Dim szHeader As String
    Dim iRow As Integer
    If frmLockingDialogisActive = False Then
        frmLockingDialogisActive = True
    Else
        Exit Sub
    End If
    'adoConn.Open getConnectionString
    'Select other computers session
    ' Equal to this scrren clinetID and SupplierID
    strSQL = "Select DateTimeStamp ,Module ,ClientID,MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber AS INV,UserSessionID,WindowsUserName,MachineName,PrestigeUserName,ServerIPaddress " & _
            "from tlbPayment INNER JOIN  tlbTransactionTypes AS TT ON tlbPayment.Type = TT.TYPE_ID where tlbPayment.ClientID='" & txtClientIDPurPay.text & _
            "' AND tlbPayment.sageaccountNumber='" & txtSPSupplier.text & "' AND UserSessionID<>'" & UserSessionID & "' AND DateTimeStamp<>'' AND Paymentview=true " & _
             "group by ClientID,DateTimeStamp ,Module ,SlNumber,UserSessionID,WindowsUserName,MachineName,PrestigeUserName," & _
             "ServerIPaddress,MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber order by MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)& SlNumber"
             'rsLockDialog.Close
    rsLockDialog.Open strSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsLockDialog.EOF Then
        With frmLockingDialog
        .Show
        .flxLockedModule.Clear
        .flxLockedModule.Cols = 9
         szHeader$ = "|<DateTimeStamp|<Module|<Client|<Invoice|<UserSessionID|<WindowsUserName|<MachineName" & _
                  "|<PrestigeUserName|<ServerIPaddress"
           .flxLockedModule.FormatString = szHeader$
           
                  
        '.flxLockedModule.RowHeight(0) = 0
'        .flxLockedModule.ColWidth(0) = 130
'        .flxLockedModule.ColWidth(1) = 2000 'DateTimeStamp
'        .flxLockedModule.ColWidth(2) = 1600 'Module
'        .flxLockedModule.ColWidth(3) = 1500 'Invoice
'        .flxLockedModule.ColWidth(4) = 0 '1800'UserSessionID
'        .flxLockedModule.ColWidth(5) = 1200 'WindowsUserName
'        .flxLockedModule.ColWidth(6) = 1200
'        .flxLockedModule.ColWidth(7) = 1000
'        .flxLockedModule.ColWidth(8) = 1000
        .flxLockedModule.ColWidth(0) = 130 'selection grid
        .flxLockedModule.ColWidth(1) = 1600 'DateTimeStamp
        .flxLockedModule.ColWidth(2) = 1350 'Module
        .flxLockedModule.ColWidth(3) = 1000 'Client
        .flxLockedModule.ColWidth(4) = 900 'Invoice
        .flxLockedModule.ColWidth(5) = 0 '1800'UserSessionID
        .flxLockedModule.ColWidth(6) = 1200
        .flxLockedModule.ColWidth(7) = 1100
        .flxLockedModule.ColWidth(8) = 1000
        .flxLockedModule.ColWidth(9) = 1000
       
       
        iRow = 1
        .flxLockedModule.Rows = rsLockDialog.RecordCount + 1
        While Not rsLockDialog.EOF
               .flxLockedModule.TextMatrix(iRow, 1) = IIf(IsNull(rsLockDialog("DateTimeStamp").Value), "", rsLockDialog("DateTimeStamp").Value)
               .flxLockedModule.TextMatrix(iRow, 2) = IIf(IsNull(rsLockDialog("Module").Value), "", rsLockDialog("Module").Value)
               .flxLockedModule.TextMatrix(iRow, 3) = IIf(IsNull(rsLockDialog("ClientID").Value), "", rsLockDialog("ClientID").Value)
               .flxLockedModule.TextMatrix(iRow, 4) = IIf(IsNull(rsLockDialog("inv").Value), "", rsLockDialog("inv").Value)
               .flxLockedModule.TextMatrix(iRow, 5) = IIf(IsNull(rsLockDialog("UserSessionID").Value), "", rsLockDialog("UserSessionID").Value)
               .flxLockedModule.TextMatrix(iRow, 6) = IIf(IsNull(rsLockDialog("WindowsUserName").Value), "", rsLockDialog("WindowsUserName").Value)
               .flxLockedModule.TextMatrix(iRow, 7) = IIf(IsNull(rsLockDialog("MachineName").Value), "", rsLockDialog("MachineName").Value)
               .flxLockedModule.TextMatrix(iRow, 8) = IIf(IsNull(rsLockDialog("PrestigeUserName").Value), "", rsLockDialog("PrestigeUserName").Value)
               .flxLockedModule.TextMatrix(iRow, 9) = IIf(IsNull(rsLockDialog("ServerIPaddress").Value), "", rsLockDialog("ServerIPaddress").Value)
               iRow = iRow + 1
        rsLockDialog.MoveNext
        
        Wend
        End With
    Else
        'frmLockingDialog.Visible = False
    End If
    rsLockDialog.Close
    Set rsLockDialog = Nothing
'    adoConn.Close
'    Set adoConn = Nothing

End Sub
Private Sub Form_Activate()
   Dim rsShowBal As New ADODB.Recordset
   Dim i As Integer
'   cmbSC.ListIndex = 0
   '##########
   Dim Control As Control
   For Each Control In Controls
            i = i + 1
   Next
   If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
        Label50(10).Caption = "totalcontrols count :" & i
        Label50(10).Visible = True
   Else
        Label50(10).Visible = False
   End If
   '##########
    
    
'    Dim Control As Control
'    For Each Control In frmPurchaseExpense.Controls
'
'                Select Case TypeName(Control)
'                    Case "CommandButton"
'                      Debug.Print Control.Caption & "+" & Control.Name
'                      'Debug.Print TypeName(Control)
'                    End Select
'      Next Control
'     Exit Sub
      

   bFormLoaded = True
   Dim adoConn As New ADODB.Connection
   If lblSearch0(5).Caption = "NotLoaded" Then
        adoConn.Open getConnectionString
        'Below line rem by anol 20180220
        'LoadflxSupplier adoConn
        'issue 316 loading is taking too much time
        Debug.Print time
        lblLoading.Caption = "Please wait while loading..."
        fmeLoading.Refresh
        'Debug.Print ReturnLastTenPIPC(adoConn, "13526", 6)
        LoadFlxPurchase adoConn
       
       
        Debug.Print time
'        LandlordAccountBalance adoConn
'        ClientAccountBalance adoConn
'        AgentAccountBalance adoConn
        'Debug.Print time
        lblSearch0(5).Caption = "Loaded"
        adoConn.Close
        Set adoConn = Nothing
        fmeLoading.Visible = False
        lblLoading.Caption = "Please wait while loading..."
   End If
End Sub

Private Sub DeleteInconsistentDatatlbPurinv(adoConn As ADODB.Connection)
    'I have found inconsistent data in the tlbPurinv table but not in the detail table. As welll as not in the tlbpayment table and detail and nlposting
    'This happened despite of using using begin trans and rollback trans in the main save procedure of PI
    'So I have decided to Delete inconsitent data from tlbPurinv table.
    Dim rsPI As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    On Error GoTo Err:
   
     'Select PI.SlNumber, PI.TransactionType FROM tblPurInv AS PI INNER JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI where Pt.PI=''
     'i have found the the data where NLposted is true but data does not exists in other table
     rsPI.Open "SELECT P.SlNumber, P.TRAN_DATE, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID, tblPurInvSRec.TRAN_ID, " & _
     "tblPurInvSRec.Nominal_code , tblPurInvSRec.description FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID " & _
     "WHERE ((P.TransactionType)=6 Or (P.TransactionType)=7)  AND  tran_ID is null;", adoConn, adOpenKeyset, adLockReadOnly
     While Not rsPI.EOF
        rsPayment.Open "Select PI.SlNumber, PI.TransactionType FROM tblPurInv AS PI INNER JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI where Pt.PI='" & _
            rsPI("MY_ID").Value & "'", adoConn, adOpenKeyset, adLockReadOnly
             'Before delete check it is not in the payment table
             If rsPayment.EOF Then
                adoConn.Execute "Insert into SpareTable5(ClientID,Code,CC) values('PI LOAD','" & Date & "' ,'InconsistenttlbPurinv,MY_ID: " & rsPI("MY_ID").Value & ",amount:" & rsPI("TOTAL_AMOUNT").Value & ",INV_NO:" & rsPI("INV_NO").Value & "')"
                adoConn.Execute "Delete from tblPurinv P where P.MY_ID='" & rsPI("MY_ID").Value & "'"
             End If
        rsPI.MoveNext
     Wend
     rsPI.Close
     Set rsPI = Nothing
     Exit Sub
Err:
End Sub

Private Sub Form_Load()
        On Error GoTo Err:
        bFormLoaded = False
        '  'cascade form function created by anol 2019 -06-17
        '    frmMMain.Arrange vbCascade
'        Me.ZOrder 0
        If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
             cmdAdvanceProgr.Visible = True
        Else
             cmdAdvanceProgr.Visible = False
        End If
        UserSessionID = GetTimeStamp
        
        Me.Width = 16920
        Me.Height = 11475
        
        cmdSavePIRef.Left = 4950
        cmdSavePIRef.Top = 270
        
        Me.BackColor = MODULEBACKCOLOR
        chkIsMgtFee.BackColor = MODULEBACKCOLOR
        tabPurExp.BackColor = MODULEBACKCOLOR
        fraLay(0).BackColor = MODULEBACKCOLOR
        fraLay(1).BackColor = MODULEBACKCOLOR
        Frame17.BackColor = MODULEBACKCOLOR
        fraCmds(0).BackColor = MODULEBACKCOLOR
        fraEditDemand.BackColor = MODULEBACKCOLOR
        fraTab0.BackColor = MODULEBACKCOLOR
        fraTab2.BackColor = MODULEBACKCOLOR
        fraTab3.BackColor = MODULEBACKCOLOR
        Frame8(1).BackColor = MODULEBACKCOLOR
        Frame5(5).BackColor = MODULEBACKCOLOR
        Frame5(1).BackColor = MODULEBACKCOLOR
        Frame8(3).BackColor = MODULEBACKCOLOR
        Frame8(0).BackColor = MODULEBACKCOLOR
        Label2.BackColor = MODULEBACKCOLOR
        chkPropertyHist.BackColor = MODULEBACKCOLOR
        
        chkProperty.BackColor = MODULEBACKCOLOR
        
'        tabPurExp.Width = 12780
'        tabPurExp.Height = 8175
        txtSupplier.text = "ALL"
        fraLay(0).Top = Me.Height + 200
        bTotalPayTyped = False
        tabPurExp.Tab = 0
        tabPayment.Tab = 0
        iSelected = 0
        iPIEdit = 0
        cGridSPTotal = 0
        bChangesMade = False
        flxPI.Enabled = True
        nTaxCode = 0
        bEditMode = False
        bPayAll = False
'        MsgBox "1"
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        'filling the client ID
        Dim szSQL As String
        Dim szSQLtemp As String
        Dim adoRstTemp As New ADODB.Recordset
        Dim adoRST As New ADODB.Recordset
        szSQL = "SELECT DISTINCT CLIENTID FROM  Client ORDER BY CLIENTID;"
        
        'When writing management fee I was wriring wrong Slnumber in the tlbpayment
        'now want to promt a messege if wrong data is there in the system
        szSQLtemp = "Select P.* from  tlbPayment P, tblPurinv V where P.PI=V.MY_ID and P.Type=V.TransactionTYpe and P.slnumber<>V.slnumber"
        adoRstTemp.Open szSQLtemp, adoConn, adOpenStatic, adLockReadOnly
        If Not adoRstTemp.EOF Then
            MsgBox "Wrong  Slnumber found in tlbPayment (not same as tlbPurinv) "
            cmdfix.Visible = True
        End If
        adoRstTemp.Close
        
        Call DeleteInconsistentDatatlbPurinv(adoConn)
        
'        MsgBox "2"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        MsgBox "2.1"
        If adoRST.RecordCount > 0 Then
            txtClientID.text = adoRST.Fields("CLIENTID").Value
            txtClientIdlist.text = "ALL" 'adoRst.Fields("CLIENTID").Value
            txtPurchasePaymentHistory.text = "ALL" ' adoRst.Fields("CLIENTID").Value
        End If
        adoRST.Close
'        MsgBox "2.2"
        adoConn.Close
        Set adoConn = Nothing
'        MsgBox "3"
        ConfigFlxSupplier
        
        'Define Purchase Invoice Flex Grid's column
        ConfigFlxPI
        'Load Purchase Invoices which are not posted to SAGE yet
        HandleCommandButton "Cancel"
        
        '   Supplier Payment - All supplier in the combo box
'        MsgBox "4"
        SupplierComboBox
'        MsgBox "5"
        
        cmbSC.AddItem "Supplier", 0
        cmbSC.AddItem "Client", 1
        cmbSC.AddItem "Managing Agent", 2
        cmbSC.AddItem "Landlord", 3
        cmbSC.ListIndex = 0
        'added by anol 30 Aug 2015
        txtSupplierType.text = "All Categories"
        txtSupplierType.Tag = "ALL"
        'end of addition
        UpdateTotalPICN
'        MsgBox "6"
'        If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
'            Call WheelHook(Me.hWnd)
'        End If
        flxSupplier(0).RowHeight(0) = 0
        'added by anol 11 Dec 2015
        txtSupSearchHis.text = "ALL"
        txtIDClient.text = "ALL"
        txtPropID.text = "ALL"
        txtSupplierSearc.text = "ALL"
        FocusControl cmdNew(1)
        Exit Sub
Err:
        MsgBox Err.description
End Sub

Public Sub PostInvoice()
   Dim szPI_ID As String
   Dim szSQL As String
   Dim iPosted As Integer               'Finally posted
   Dim iIP     As Integer               'To be posted
   Dim adoConn As New ADODB.Connection
   Dim rsPI As New ADODB.Recordset
   Dim SelPurID() As String
   Dim j As Long
   Dim K As Integer

   adoConn.Open getConnectionString

   If frmPopUpMenu.optSelPI.Value Then
            szPI_ID = SelectedPurInvIDArr(SelPurID())
            j = UBound(SelPurID())
            If j = 0 Then
                    MsgBox "Please select a purchase invoice to post in history.", vbCritical + vbOKOnly, "Purchase invoice"
                    Exit Sub
            End If
            K = CInt(j / 50)
            If K = j / 50 Then
                 'No no need to do ceiling, this is fully divisible
                 K = j / 50
            Else
                 K = CInt(j / 50) + 1 'This is ceiling function
            End If
            For K = 0 To K - 1
                szPI_ID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurID())
                If szPI_ID = "" Then
                     Exit For
                End If
                If Trim(szPI_ID) <> "" Then
                    szSQL = "UPDATE tblPurInv " & _
                        "SET History = TRUE " & _
                        "WHERE MY_ID IN (" & szPI_ID & ");"
                     adoConn.Execute szSQL
                End If
            Next
            szPI_ID = ""
            GoTo XX
   End If
   If frmPopUpMenu.optPIDtRange.Value Then
            szPI_ID = DateRangePurInvID(CDate(frmPopUpMenu.txtDtRangeFrom.text), CDate(frmPopUpMenu.txtDtRangeTo.text), SelPurID())
            j = UBound(SelPurID())
            If j = 0 Then
                    MsgBox "Please select a purchase invoice to post in history.", vbCritical + vbOKOnly, "Purchase invoice"
                    Exit Sub
            End If
            K = CInt(j / 50)
            If K = j / 50 Then
                 'No no need to do ceiling, this is fully divisible
                 K = j / 50
            Else
                 K = CInt(j / 50) + 1 'This is ceiling function
            End If
            For K = 0 To K - 1
                szPI_ID = ReturnString(K * 50, (K + 1) * 50 - 1, SelPurID())
                If szPI_ID = "" Then
                     Exit For
                End If
                If Trim(szPI_ID) <> "" Then
                  szSQL = "UPDATE tblPurInv " & _
                    "SET History = TRUE " & _
                    "WHERE MY_ID IN (" & szPI_ID & ");"
            
                    adoConn.Execute szSQL
                End If
            Next

            szPI_ID = ""
            GoTo XX
        
   End If
   If frmPopUpMenu.optFP_PI.Value Then
        'szPI_ID = FullyPaidPurInvID(adoConn)
        szSQL = "Select P.* From tlbPayment AS P INNER JOIN tblPurInv AS I ON P.PI = I.MY_ID WHERE P.OSAmount = 0 AND (I.TransactionType = 6 Or I.TransactionType = 7) AND I.History = FALSE"
        rsPI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        j = rsPI.RecordCount
        rsPI.Close
        Set rsPI = Nothing
        adoConn.Execute "Update tlbPayment AS P INNER JOIN tblPurInv AS I ON P.PI = I.MY_ID SET I.History = TRUE WHERE P.OSAmount = 0 AND (I.TransactionType = 6 Or I.TransactionType = 7)  AND I.History = FALSE"
   End If
   If frmPopUpMenu.optSlNoRange.Value Then
        
        szSQL = "Select P.* From tblPurInv AS P  WHERE slNumber>=" & StrDigitVal(frmPopUpMenu.txtPlRangeFrom.text) & " and  slNumber<=" & StrDigitVal(frmPopUpMenu.txtPlRangeTo.text) & "" & _
        " AND TransactionType=" & IIf(UCase(Left(frmPopUpMenu.txtPlRangeTo.text, 2)) = "PI", 6, 7) & " "
        rsPI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        j = rsPI.RecordCount
        rsPI.Close
        Set rsPI = Nothing
        adoConn.Execute "Update tblPurInv Set History = TRUE WHERE slNumber>=" & StrDigitVal(frmPopUpMenu.txtPlRangeFrom.text) & " and  slNumber<=" & StrDigitVal(frmPopUpMenu.txtPlRangeTo.text) & "" & _
        " AND TransactionType=" & IIf(UCase(Left(frmPopUpMenu.txtPlRangeTo.text, 2)) = "PI", 6, 7) & " "
       ' szPI_ID = SlNoRangePurInv(frmPopUpMenu.txtPlRangeFrom.text, frmPopUpMenu.txtPlRangeTo.text)
            
   End If

'   If Len(szPI_ID) = 0 Then
'      ShowMsgInTaskBar "No invoice to post to history.", "Y", "N"
'
'      adoConn.Close
'      Set adoConn = Nothing
'      Exit Sub
'   End If
'
'   szSQL = "UPDATE tblPurInv " & _
'           "SET History = TRUE " & _
'           "WHERE MY_ID IN (" & szPI_ID & ");"
'
'   adoConn.Execute szSQL
XX:
   LoadFlxPurchase adoConn
   fmeLoading.Visible = False
   MousePointer = vbDefault

   adoConn.Close
   Set adoConn = Nothing
   ShowMsgInTaskBar "System has posted " & j & " invoice to history.", "Y", "P"
End Sub

Private Function SlNoRangePurInv(szNoFrom As String, szNoTo As String) As String
   Dim i    As Integer
   Dim szIC As String
   Dim lF   As Long
   Dim lT   As Long

   szIC = UCase(Left(szNoFrom, 2))
   lF = StrDigitVal(szNoFrom)
   lT = StrDigitVal(szNoTo)

   SlNoRangePurInv = ""
   For i = 1 To flxPurchase.Rows - 1
      If Left(flxPurchase.TextMatrix(i, 2), 2) = szIC Then
         If StrDigitVal(flxPurchase.TextMatrix(i, 2)) >= lF And StrDigitVal(flxPurchase.TextMatrix(i, 2)) <= lT Then
            SlNoRangePurInv = SlNoRangePurInv & "'" & flxPurchase.TextMatrix(i, 0) & "'"
            SlNoRangePurInv = SlNoRangePurInv & ","
         End If
      End If
   Next i

   If SlNoRangePurInv <> "" Then
      SlNoRangePurInv = Left(SlNoRangePurInv, Len(SlNoRangePurInv) - 1)
   Else
      SlNoRangePurInv = ""
   End If
End Function

Private Function FullyPaidPurInvID(adoConn As ADODB.Connection) As String
   Dim i          As Integer
   Dim szSQL      As String
   Dim adoRST     As New ADODB.Recordset

   szSQL = "SELECT I.MY_ID " & _
           "FROM tlbPayment AS P INNER JOIN tblPurInv AS I ON P.PI = I.MY_ID " & _
           "WHERE P.OSAmount = 0 AND I.TransactionType = 6 AND I.History = FALSE;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   FullyPaidPurInvID = ""
   If adoRST.RecordCount = 0 Then GoTo NoRecord

   While Not adoRST.EOF
      FullyPaidPurInvID = FullyPaidPurInvID & "'" & adoRST.Fields.Item(0).Value & "'"
      FullyPaidPurInvID = FullyPaidPurInvID & ","
      adoRST.MoveNext
   Wend
   adoRST.Close
   
   If FullyPaidPurInvID <> "" Then
      FullyPaidPurInvID = Left(FullyPaidPurInvID, Len(FullyPaidPurInvID) - 1)
   Else
      FullyPaidPurInvID = ""
   End If

NoRecord:
   Set adoRST = Nothing
End Function

Private Function SelectedPurInvID() As String

   Dim i As Integer

  
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelectedPurInvID = SelectedPurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         SelectedPurInvID = SelectedPurInvID & ","
      End If
   Next i
   
  
   
   If Len(SelectedPurInvID) > 0 Then
      SelectedPurInvID = Left(SelectedPurInvID, Len(SelectedPurInvID) - 1)
   End If
End Function
Private Function SelectedPurInvIDArr(ByRef SelPurID() As String) As String
'This function is for storing Purchase Invoice ID into an array written by anol 20181113
   Dim i As Integer
   Dim j As Long

  
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelectedPurInvIDArr = SelectedPurInvIDArr & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         SelectedPurInvIDArr = SelectedPurInvIDArr & ","
         j = j + 1
      End If
   Next i
   
   ReDim SelPurID(j)
   j = 0
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelPurID(j) = "'" & CStr(flxPurchase.TextMatrix(i, 0)) & "'"
         j = j + 1
      End If
   Next i
   
   If Len(SelectedPurInvIDArr) > 0 Then
      SelectedPurInvIDArr = Left(SelectedPurInvIDArr, Len(SelectedPurInvIDArr) - 1)
   End If
End Function

Private Function DateRangePurInvID(dtFrom As Date, dtTo As Date, SelPurID() As String) As String
   Dim i As Integer
   Dim j As Long

   DateRangePurInvID = ""
   For i = 1 To flxPurchase.Rows - 1
      If CDate(flxPurchase.TextMatrix(i, 4)) >= dtFrom And CDate(flxPurchase.TextMatrix(i, 4)) <= dtTo Then
         DateRangePurInvID = DateRangePurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         DateRangePurInvID = DateRangePurInvID & ","
          j = j + 1
      End If
   Next i
   ReDim SelPurID(j)
   j = 0
   For i = 1 To flxPurchase.Rows - 1
      If CDate(flxPurchase.TextMatrix(i, 4)) >= dtFrom And CDate(flxPurchase.TextMatrix(i, 4)) <= dtTo Then
          SelPurID(j) = "'" & CStr(flxPurchase.TextMatrix(i, 0)) & "'"
          j = j + 1
      End If
   Next i
   If Len(DateRangePurInvID) > 0 Then
      DateRangePurInvID = Left(DateRangePurInvID, Len(DateRangePurInvID) - 1)
   End If
End Function

Private Sub FlxSumUp(conFlxGrid As Control, iNet As Integer, iVat As Integer, conTextBoxNet As Control, conTextBoxVat As Control)
   Dim iRow As Integer, dNet As Double, dVat As Double

   dNet = 0
   dVat = 0

   For iRow = 1 To conFlxGrid.Rows - 1
      dNet = dNet + Val(conFlxGrid.TextMatrix(iRow, iNet))
      dVat = dVat + Val(conFlxGrid.TextMatrix(iRow, iVat))
   Next iRow

   conTextBoxNet.text = CStr(Format(dNet, "0.00"))
   conTextBoxVat.text = CStr(Format(dVat, "0.00"))
End Sub

Private Sub cmdClose_Click(Index As Integer)
'when flxPI.Tag = "Edited" and cmdEdit(1).Enabled=false msgbox do you want to save?
     cmdSavePIRef.Visible = False
     fraEditDemand.Enabled = True
   fraTab0.Enabled = True

   Dim adoConn As New ADODB.Connection
   If flxPI.Tag = "EditedOrAdded" And cmdSavePI.Enabled = True Then
        If MsgBox("Do you wish to save your changes?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
           'If cmdSavePI.Enabled Then cmdSavePI.SetFocus
           'Resolved by BOSL
           'added by anol Date 21 Apr 2015
           'issue 453 Note 5
           cmdSavePI_Click
           'End of modification
           'Exit Sub
        End If
   Else ' you have done no changes
         If cmdEdit(1).Enabled = False Then 'this is the left pane edit button. If this button is disabled that means invoice in edit mode
'            MsgBox "Returning from Editmode  non changes"
                    adoConn.Open getConnectionString
                    adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "' AND Module='Purchase Invoice'"
                   flxPurchase.row = iPIEdit
                    flxPurchase.col = 0
                    flxPurchase.CellBackColor = vbWhite
                   adoConn.Close
                   Set adoConn = Nothing
                   
         End If
   End If
    flxPI.Tag = ""
'   If Not cmdEdit(Index).Enabled Then
'      If MsgBox("Do you want to add another transaction?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
'         If cmdSavePI.Enabled Then cmdSavePI.SetFocus
'         'Resolved by BOSL
'         'added by anol Date 21 Apr 2015
'         'issue 453 Note 5
'
'
'         'cmdSavePI_Click
'         'End of modification
'         Exit Sub
'      End If
'   End If
   fraLay(0).Top = Me.Height + 1300
   cmdEdit(1).Enabled = True
   
   txtSupplierName.text = ""
   txtSupplierID.text = ""
   txtProperty.text = ""
 'comment out by anol 20160912
''   Dim adoConn As New ADODB.Connection
'''   connect to database
''   adoConn.Open getConnectionString
''
''   LoadFlxPurchase adoConn
''
''   adoConn.Close
''   Set adoConn = Nothing
   ConfigFlxPI
   iPIEdit = 0
   FocusControl flxPurchase
End Sub

Private Sub cmdNew_Click(Index As Integer)
     'added by anol on 01 Dec 2015
      'This means form is opening with an add mode
      cmdSavePI.Tag = UniqueID()
      'End of addition
      'added by  anol 01 Dec 2015
      cmdOpenFileView.Visible = False
      cmdviewMenu.Visible = True
      
      'reverese control activation of view button
      fraLay(1).Enabled = True
      cmdDelete.Enabled = True
      cmdEdit(0).Enabled = True
      cmdOpenFileView.Visible = False
      cmdviewMenu.Visible = True
      txtPICNNet.text = "0.00"
      txtPICNVat.text = "0.00"
      txtPICNTotal.text = "0.00"
      'end
      'End of addition
      
   If Index = 1 Then
      fraInvCrChoice.Left = fraEditDemand.Left + fraEditDemand.Width + 100
      fraInvCrChoice.Top = fraEditDemand.Top + tabPurExp.Top + 90 '+ fraCmds(tabPurExp.Tab).Top + cmdNew(tabPurExp.Tab).Top - fraInvCrChoice.Height
   Else
      cmdManualDmdOk_Click
      Exit Sub
      fraInvCrChoice.Left = tabPurExp.Left + fraLay(0).Left + fraCmds(tabPurExp.Tab).Left
      fraInvCrChoice.Top = tabPurExp.Top + fraLay(0).Top + fraCmds(tabPurExp.Tab).Top + cmdNew(tabPurExp.Tab).Top - fraInvCrChoice.Height
   End If
   fraInvCrChoice.Visible = True
   cmdManualDmdOk.SetFocus
   tabPurExp.Enabled = False
'   flxPI.Enabled = False
End Sub

Private Sub PIComponents(ComponentMode As String)
   Select Case ComponentMode

   Case "DefaultMode"
      cmbSC.Enabled = True
      txtSupplierID.text = ""
      txtSupplierName.text = ""
      txtReference.text = ""
      txtDueDate.text = ""
      txtUnit(tabPurExp.Tab).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(tabPurExp.Tab).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False

      txtPICNNet.text = "0.00"
      txtPICNVat.text = "0.00"
      txtPICNTotal.text = "0.00"

   Case "NewLine"
      txtUnit(tabPurExp.Tab).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(tabPurExp.Tab).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False

      txtPICNNet.text = "0.00"
      txtPICNVat.text = "0.00"
      txtPICNTotal.text = "0.00"
      lblVatCode(0).Caption = ""
      lblVatCode(0).Tag = ""

   Case "EditLine"
      txtUnit(tabPurExp.Tab).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(tabPurExp.Tab).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False
   End Select
End Sub

Private Sub cmdSavePI_Click()
'is valid client implement
    Dim iCount As Integer
    Dim rsCheck As New ADODB.Recordset
    If txtClientID.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientID.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
    If Not isValidClient Then
        MsgBox "Please select a valid client", vbInformation, "Select a Client"
        cmdClientSerc.SetFocus
        Exit Sub
    End If
'    If Not isValidProperty Then
'        MsgBox "Please select a valid Property", vbInformation, "Select a Property"
'        cmdTypeList.SetFocus
'        Exit Sub
'    End If
   If cmbSC.text = "Managing Agent" And chkIsMgtFee.Value = 0 Then
        If MsgBox("Do you wish to save this transaction as a Managing Agents Fee " & txtTransType.text & "", vbYesNo, "Please confirm") = vbYes Then
            chkIsMgtFee.Value = 1
            Exit Sub
        End If
   End If
   If Len(txtDueDate.text) <> 10 Then
        MsgBox "Please Enter due date", vbInformation, "Enter due date"
        txtDueDate.SetFocus
        Exit Sub
   End If
   If cmbSC.text <> "Managing Agent" And chkIsMgtFee.Value = 1 Then
        MsgBox "Please select supplier type Managing Agent if you are creating a Management Fee", vbInformation, "Management Fee"
        Exit Sub
   End If
   fraEditDemand.Enabled = True
   fraTab0.Enabled = True
    'validation issue 571
    'added by anol 22 May 2015
    'Comment out by anol 13 Aug 2015
'   If flxPI.Rows = 2 Then
'        If flxPI.TextMatrix(1, 1) = "" Then
'        'Below line modified by anol 13 August
'             MsgBox "You must have atleast one line of transaction in order to save", vbInformation, "Not Saved!"
'             Exit Sub
'        End If
'   End If
         'Resolved by BOSL
         'Issue 468
         'Modified by Anol 25 Nov 2015
    
    Dim adoConn As New ADODB.Connection
    Dim szSQL As String
    If IsDate(lblPostingDate.ToolTipText) = True Then
        adoConn.Open getConnectionString
'        adoConn.BeginTrans
        If IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 0 Then
            MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Posting Date"
            adoConn.Close
            Exit Sub
        ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 9 Then
            MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Posting Date"
            adoConn.Close
            Exit Sub
        End If
        If DateDiff("d", lblPostingDate.ToolTipText, txtDate.text) > 0 Then
              MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
              Exit Sub
        End If
        adoConn.Close
    End If
           
   If frmMMain.rtxtMessageDisplay.text = "Please update the invoice line." Then Exit Sub
   cmdaddnewline.Enabled = True
   FocusControl cmdClose(0)
   If Val(txtNet_(0).text) > 0 Then cmdUpdate_Click 1

   'below line added by anol 05 08 2016
   cmdSavePI.Enabled = False
   
   Dim adoPIHeader As New ADODB.Recordset, adoPISplit As New ADODB.Recordset
   Dim iRow As Integer, uID As String, lT_ID As Long
   Dim lSlNumber As Long

   adoConn.Open getConnectionString
   adoConn.BeginTrans
'  ***************************************************************************************************
'           SAVING HEADER PART OF THE PURCHASE INVOICE                                               '
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInv"
   frmMMain.frmSupplier_SupplierList_isUptoDate = False
   If Not cmdEdit(1).Enabled Then szSQL = szSQL + " WHERE MY_ID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"

   With adoPIHeader
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      If cmdEdit(1).Enabled Then                'Add New Mode
         .AddNew
         'Modified by anol 01 Dec 2015
         If cmdSavePI.Tag = "" Then
                MsgBox "An error has occurred. Please contact PCM LTD.", vbInformation, "Error_ cmdSavePI.Tag_empty"
                Exit Sub
         End If
         uID = cmdSavePI.Tag 'UniqueID()
         .Fields.Item("CreatedBy").Value = User
         .Fields.Item("CreatedDate").Value = Now
         'end of modification
         .Fields.Item("MY_ID").Value = uID
         If sAddChoice = "IN" Or sAddChoice = "AI" Then
            lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
            .Fields.Item("SlNumber").Value = lSlNumber
         End If
         If sAddChoice = "CN" Or sAddChoice = "AC" Then
            lSlNumber = SlNumber("PC", "tblPurInv", adoConn)
            .Fields.Item("SlNumber").Value = lSlNumber
         End If
      Else
         'lSlNumber = StrDigitVal(flxPurchase.TextMatrix(flxPurchase.row, 2))
         lSlNumber = StrDigitVal(flxPurchase.TextMatrix(iPIEdit, 2))
         
      End If
      .Fields.Item("SUPP_AC").Value = txtSupplierID.text
      .Fields.Item("TRAN_DATE").Value = Format(txtDate.text, "DD/MMMM/YYYY")
      .Fields.Item("TransactionType").Value = IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7)
      .Fields.Item("INV_NO").Value = txtReference.text
      .Fields.Item("TOTAL_AMOUNT").Value = CCur(txtPICNTotal.text)
      .Fields.Item("TTP").Value = CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoConn))
      .Fields.Item("History").Value = False
      .Fields.Item("TrfPayment").Value = True
      .Fields.Item("PropertyID").Value = txtProperty.text
      .Fields.Item("CL_ID").Value = txtClientID.text
      .Fields.Item("isManagementFee").Value = CBool(chkIsMgtFee.Value)
      
      '  Resolved By BOSL. Modified by Asif. Issue: 0000503. Date: 22 Jan 2015
      .Fields.Item("NLPost").Value = False
      ''
     
       .Fields.Item("DueDate").Value = Format(txtDueDate.text, "dd mmmm yyyy")
      .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")

'   UPDATE_SAGE --> Adjustment Transaction. If its TRUE then its a Adj Tran                        '
      If bAdjustment Then .Fields.Item("UPDATE_SAGE").Value = True
      .Fields.Item("LastModifiedBy").Value = User
      .Fields.Item("LastModifiedDate").Value = Now
      .Update
      .Close
   End With

'  ***************************************************************************************************
'           B4 SAVING SPLITS, THE PI IS EXPORTED TO PAYMENT TABLE                                    '
'  ***************************************************************************************************
   If cmdEdit(1).Enabled Then                                        'Add New Mode
      szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
      adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
      adoPIHeader.Close
   End If

   szSQL = "SELECT * FROM tlbPayment"                                         'Add New Mode
   frmMMain.frmSupplier_SupplierList_isUptoDate = False
   If Not cmdEdit(1).Enabled Then                                             'Edit    Mode
      szSQL = szSQL + " WHERE PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
   End If

   With adoPIHeader
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
      If cmdEdit(1).Enabled Then
      
         .AddNew
         !TransactionID = lT_ID
            !CreatedBy = User
            !CreatedDate = Now
         !szTransactionID = !TransactionID
         !Pi = uID
      Else
         lT_ID = !TransactionID
      End If
      !Type = IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7)    'PP - Purchase Invoice, look in the tlbTransactionType
      !SageAccountNumber = txtSupplierID.text
      !PDate = Format(txtDate.text, "DD MMMM YYYY")
      !dDate = Format(txtDueDate.text, "DD MMMM YYYY")
      !ref = txtReference.text
      !ExtRef = !ref
      !amount = CCur(txtPICNTotal.text)
      !OSAmount = !amount
      !PaymentView = True
      !Details = flxPI.TextMatrix(1, 11)
      !unitid = txtProperty.text
      !SlNumber = lSlNumber
       For iRow = 1 To flxPI.Rows - 1
            If flxPI.TextMatrix(iRow, 0) <> "" Then
              !fundID = flxPI.TextMatrix(iRow, 8)
            End If
      Next iRow
      !AdjTag = IIf(bAdjustment, "Y", "N")
      !Recoverable = AccuRecoPer()
      !postingDate = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
    'tlbPayment is now writing clinet ID because this is needed for aged creditors report by anol 20170226 issue 308
      !ClientID = txtClientID.text
      .Update
      .Close
   End With

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE in the PAYMENT SPLIT TABLE
'  ***************************************************************************************************
   If Not cmdEdit(1).Enabled Then       'Edit Mode (this is the edit button on the left corner if this is diisabled that means invoice is in edit mode
      adoConn.Execute "DELETE S.* " & _
                      "FROM tlbPayment AS P, tlbPaymentSplit AS S " & _
                      "WHERE PI = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "' AND " & _
                            "P.TransactionID = S.PayHeader;"
   End If
   szSQL = "SELECT * FROM tlbPaymentSplit;"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'Add New Records. At least there is one split line.
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(iRow, 0) <> "" Then
         With adoPISplit
            .AddNew
            .Fields.Item("TransactionID").Value = UniqueID()
            .Fields.Item("PayHeader").Value = lT_ID
            .Fields.Item("SplitID").Value = flxPI.TextMatrix(iRow, 0)
            .Fields.Item("TRANS").Value = flxPI.TextMatrix(iRow, 3)
            .Fields.Item("NOMINAL_CODE").Value = flxPI.TextMatrix(iRow, 6)      'Nominal Code
            .Fields.Item("FundID").Value = flxPI.TextMatrix(iRow, 8)            'Fund ID
            .Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
            .Fields.Item("Description").Value = flxPI.TextMatrix(iRow, 11)
            .Fields.Item("Amount").Value = CCur(flxPI.TextMatrix(iRow, 14)) + CCur(flxPI.TextMatrix(iRow, 12)) 'adding vat and amount in tlbPaymentsplit amount
            .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
            .Fields.Item("DueDate").Value = Format(txtDueDate.text, "DD MMMM YYYY")
            
            .Fields.Item("ScheduleID").Value = IIf(flxPI.TextMatrix(iRow, 20) = "", Null, _
                                                   flxPI.TextMatrix(iRow, 20))
            .Fields.Item("UNIT_ID").Value = flxPI.TextMatrix(iRow, 21)
            .Fields.Item("RecoverablePt").Value = flxPI.TextMatrix(iRow, 22)
            .Fields.Item("AllocTranID").Value = flxPI.TextMatrix(iRow, 23)

            .Update
         End With
      End If
   Next iRow

'  User has deleted all split line.
   If flxPI.TextMatrix(1, 0) = "" Then
      With adoPISplit
         .AddNew
         .Fields.Item("TransactionID").Value = UniqueID()
         .Fields.Item("PayHeader").Value = lT_ID
         .Fields.Item("FundID").Value = 0
         .Fields.Item("Amount").Value = 0
         .Fields.Item("OSAmount").Value = 0
         .Fields.Item("SplitID").Value = 1
         .Fields.Item("DueDate").Value = Format(txtDueDate.text, "DD MMMM YYYY")
         .Fields.Item("Description").Value = "ALL SPLIT DELETED"
         .Fields.Item("NOMINAL_CODE").Value = "0000"
         .Fields.Item("ScheduleID").Value = 0
         .Update
      End With
   End If
   
   adoPISplit.Close

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE
'  ***************************************************************************************************
   If Not cmdEdit(1).Enabled Then                                             'Edit Mode
   '  Resolved By BOSL. Modified by Asif. Issue: 0000503. Date: 22 Jan 2015
'      adoconn.Execute "UPDATE NLPosting AS N, tblPurInvSRec AS S " & _
'                      "SET    N.DeleteFlag = TRUE " & _
'                      "WHERE  N.PARENT_RECORD = S.MY_ID AND " & _
'                             "S.ParentID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
'How many entries can be deleted here? 2?
'ans: all in an invoice

'Re - Fixed by anol 2017-08-10
       adoConn.Execute "UPDATE NLPosting AS N " & _
                      "SET    N.DeleteFlag = TRUE " & _
                      "WHERE  N.TRANSACTION_TYPE = " & IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7) & " AND " & _
                             "N.TRANS_ID = '" & lSlNumber & "';"
'Here you also Need to enter entry with zero values ( anol 2018/02/01 )

'      Debug.Print "UPDATE NLPosting AS N, tblPurInvSRec AS S " & _
'                      "SET    N.DeleteFlag = TRUE " & _
'                      "WHERE  N.PARENT_RECORD = S.MY_ID AND " & _
'                             "S.ParentID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
   '''''''''''''''''''
      adoConn.Execute "DELETE * FROM tblPurInvSRec WHERE ParentID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "';"
   End If
   szSQL = "SELECT * FROM tblPurInvSRec"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

'Add New Records. At least there is only one split line
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(iRow, 0) <> "" Then
         With adoPISplit
            .AddNew
            
            If cmdEdit(1).Enabled Then 'this is the left pane button if this is enable means new invoice
               .Fields.Item("ParentID").Value = uID
            Else
               .Fields.Item("ParentID").Value = flxPurchase.TextMatrix(iPIEdit, 0)
            End If

            .Fields.Item("TRAN_ID").Value = flxPI.TextMatrix(iRow, 0)
            'Modified by anol 18 Aug 2015 Prop ID was not saving in split
            '.Fields.Item("TRANS").Value = flxPI.TextMatrix(iRow, 3)
            .Fields.Item("TRANS").Value = txtProperty.text
            
            .Fields.Item("NOMINAL_CODE").Value = flxPI.TextMatrix(iRow, 6)
            .Fields.Item("DEPT_ID").Value = flxPI.TextMatrix(iRow, 8)
            .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
             If flxPI.TextMatrix(iRow, 9) <> "" Then
                     UpdateJobActualCost .Fields.Item("JOB_ID").Value, .Fields.Item("NET_AMOUNT").Value, adoConn
             End If
            .Fields.Item("COST_CODE").Value = flxPI.TextMatrix(iRow, 10)
            .Fields.Item("description").Value = flxPI.TextMatrix(iRow, 11)
            .Fields.Item("NET_AMOUNT").Value = CCur(flxPI.TextMatrix(iRow, 12))
            .Fields.Item("TAX_CODE").Value = Trim(flxPI.TextMatrix(iRow, 13))
            .Fields.Item("VAT").Value = CCur(flxPI.TextMatrix(iRow, 14))
            .Fields.Item("TOTAL_AMOUNT").Value = CCur(flxPI.TextMatrix(iRow, 12)) + CCur(flxPI.TextMatrix(iRow, 14))
            .Fields.Item("ScheduleID").Value = IIf(flxPI.TextMatrix(iRow, 20) = "", Null, flxPI.TextMatrix(iRow, 20))
            .Fields.Item("UNIT_ID").Value = flxPI.TextMatrix(iRow, 21)
            .Fields.Item("RecoverablePt").Value = flxPI.TextMatrix(iRow, 22)
            .Fields.Item("MY_ID").Value = flxPI.TextMatrix(iRow, 23)
            .Update
         End With
         'Here I am sure there  is at least one row int he split table
         If chkIsMgtFee.Value = 1 Then
             adoConn.Execute "Update tlbAgreement T,ClientProAgr P set T.ReportClientID=P.ClientID,T.ReportPropertyID=P.PropertyID where T.CPA_ID=P.CPA_ID"
    
            szSQL = "SELECT CL_ID,max(Cdate(duedate)) as DDate from tblPurInv where isManagementFee=true AND CL_ID='" & txtClientID.text & "' group by CL_ID"
            rsCheck.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            While Not rsCheck.EOF
                 szSQL = "Update tlbAgreement A,tblPurInv P Set A.LastChargeDate='" & rsCheck("DDate").Value & "'  where isManagementFee=true AND A.ReportClientID='" & rsCheck("CL_ID").Value & "'"
                 adoConn.Execute szSQL
                 rsCheck.MoveNext
            Wend
            rsCheck.Close
         End If
      End If
   Next iRow
'When user edits actual cost against an purchase invoice for a job, the modified amount is being added to the original amount of the invoice
'issue 474
'Resolved by BOSL 09 Mar 2015
      Dim K As Integer
      If cmdEdit(1).Enabled = False Then
         For K = 0 To UBound(VTJobAmount)
           UpdateJobActualCost VTJobAmount(K).JobID, -VTJobAmount(K).amount, adoConn
         Next K
      End If
'end of modification

'  User has deleted all split line.
   If flxPI.TextMatrix(1, 0) = "" Then
      With adoPISplit
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ParentID").Value = uID
         .Fields.Item("TRAN_ID").Value = 1
         .Fields.Item("TRANS").Value = txtProperty.text
         .Fields.Item("NOMINAL_CODE").Value = "0000"
         .Fields.Item("description").Value = "DELETED ALL SPLITS"
         .Fields.Item("NET_AMOUNT").Value = 0
         .Fields.Item("TAX_CODE").Value = "T9"
         .Fields.Item("VAT").Value = 0
         .Fields.Item("TOTAL_AMOUNT").Value = 0
         .Fields.Item("RecoverablePt").Value = 0
         .Update
      End With
   End If
   adoPISplit.Close
    
'   LoadFlxPurchase adoConn
''*****  System will check the header amount with the split total  ***************************
'   Call SiPi_Check(adoConn, "PI", "25876")
''--------------------------------------------------------------------------------------------
''  Export Transactions to Nominal Ledger (NLPosting table)
'   Export_PInPC_2_NL adoConn
'--------------------------------------------------------------------------------------
    szTran2Fix = "" 'Need to clear this before it stores any value ot this will cause a problem
    If PI_Check(adoConn, szTran2Fix) = False Then
         adoConn.RollbackTrans
         adoConn.Close
         MsgBox "An error occurred while saving, transaction rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked"
    Else
        Dim postResult As Boolean
'        If cmbSC.text = "Supplier" Then
'            postResult = Export_PInPC_2_NL(adoConn)
'        ElseIf cmbSC.text = "Client" Or cmbSC.text = "Landlord" Then
'            postResult = Export_PInPC_2_NL_ForClientLandLord(adoConn)
'        ElseIf cmbSC.text = "Managing Agent" Then
'            postResult = Export_PInPC_2_NL_ForAGENT(adoConn)
'        End If
        If cmbSC.text = "Managing Agent" Then
            postResult = Export_PInPC_2_NL_ForAGENT(adoConn)
        ElseIf cmbSC.text = "Client" Or cmbSC.text = "Landlord" Then
            postResult = Export_PInPC_2_NL_ForClientLandLord(adoConn)
        Else
            postResult = Export_PInPC_2_NL(adoConn)
        End If

        
         If postResult = False Then 'this part modified by anol 2021-01-13
'        If Export_PInPC_2_NL(adoconn) = False Then
                'If checkNLPostingPLCAWithPurINV(adoConn) = False Then
                    adoConn.RollbackTrans
                    adoConn.Close
                    MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
                 'End If
        Else
            adoConn.CommitTrans
            
            szSQL = "UPDATE tblPurInv AS P, tblPurInvSRec AS S " & _
                "SET P.NLPost = TRUE " & _
                "WHERE  P.MY_ID = S.ParentID AND NOT P.NLPost AND " & _
                    "(P.TransactionType = 6 OR P.TransactionType = 7);"
            adoConn.Execute szSQL
            frmMMain.frmSupplier_SupplierList_isUptoDate = False
            frmMMain.frmPI_SupplierBalance_isUptoDate = False
            frmMMain.frmPI_SupplierBalanceByCL_isUptoDate = False
            MsgBox "The Transactions have been saved successfully.", vbInformation, "Transactions have been saved"
            If Not cmdEdit(1).Enabled Then ' invoice in edit mode
                'Here I am only loading one row of grid without loading the whole grid by anol 20160912
                LoadFlxGridOneRow adoConn
                LoadpurchaseSplit adoConn
            Else
                LoadFlxPurchase adoConn
                fmeLoading.Visible = False
            End If
            
    '        adoConn.Close
    '        adoConn.Open getConnectionString
            '--------------------------------------------------------------------------------------------
            '  Export Transactions to Nominal Ledger (NLPosting table)
           
            If cmdEdit(1).Enabled = False Then 'this is the left pane edit button. If this button is disabled that means invoice in edit mode
                     adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "' AND Module='Purchase Invoice'"
                   
                 'Call puchaseInvoice.Add(fraLay(1).Tag, flxPI.TextMatrix(iCurEditRow, 7), iCurEditRow, flxPI.TextMatrix(iRow, 23), _
                         flxPI.TextMatrix(iRow, 21), flxPI.TextMatrix(iRow, 8), flxPI.TextMatrix(iRow, 11), , flxPI.TextMatrix(iRow, 0))
    '             If IsObject(PurchaseInvoice) Then
    '                If PurchaseInvoice.Count > 0 Then
    '                   For iCount = 1 To PurchaseInvoice.Count
    '                       Call InsertZeroValues(adoConn, PurchaseInvoice.Item(iCount).SlNumber, PurchaseInvoice.Item(iCount).Nominal_code, _
    '                             PurchaseInvoice.Item(iCount).iCurEditRow, PurchaseInvoice.Item(iCount).PARENT_RECORD, _
    '                            PurchaseInvoice.Item(iCount).UNIT_ID, PurchaseInvoice.Item(iCount).TRANSACTION_DESCRIPTION, PurchaseInvoice.Item(iCount).FUND_ID)     'Insert zero values row in the NLPOSTING table
    '                   Next iCount
    '                End If
    '             End If
    '             Set PurchaseInvoice = Nothing
            End If
    '        ShowMsgInTaskBar "The Transactions have been saved successfully."
            adoConn.Close
            Set adoConn = Nothing
        End If
    End If
   'adoConn.Close

   Set adoPISplit = Nothing
   Set adoPIHeader = Nothing
   'Set adoConn = Nothing
 'below line added by anol 05 08 2016
   PIComponents "DefaultMode"
   If cmdEdit(1).Enabled = False Then
      fraLay(0).Top = Me.Height + 300
      
   Else
      HandleCommandButton "Save"
      'cmdNew(0).SetFocus
     
   End If
 
  ' ShowMsgInTaskBar "Data has been saved successfully."
   'added by anol 10 Jun 2015
   If cmdEdit(1).Enabled = False Then
        cmdEdit(1).Enabled = True
    Else
        cmdNew_Click (0)
        'added by anol 20161014
'        cmdACList(0).SetFocus
        cmdACList0Focus
    End If
    'resolved by BOSL
    'Issue number 453 note2
    'Modified by anol 14 Aug 2014
   adoConn.Open getConnectionString
   SupplierAccountBalance adoConn
   'added by anol 22 Aug 2016 Updating supp bal clientWise
   
   
   'SupplierAcBalByClient adoConn
   SupplierAcBalByClient2 adoConn
   Call LoadSupplierOnPayment(adoConn, "")
   'End modification
   'added by anol 17 Jan 2016
   If txtSPSupplier.text <> "" Then
        UpdtSupplierAccountBalance adoConn
   End If
   'End of modification
   flxPI.Clear
   flxPI.Rows = 2
   lblSearch0(5).Caption = "NotLoaded"
    'Modified by anol 14 Aug 2014
   If adoConn.State = adStateOpen Then
        adoConn.Close
   End If
   Set adoConn = Nothing
   cmdSavePI.Enabled = False
   iPIEdit = 0
   'added by anol 02 Dec 2015
   cmbFiles.Clear
   chkIsMgtFee.Value = 0
End Sub
'Private Function ReturnLastTenPIPC(adoConn As ADODB.Connection, strCurInv As String, transactionType As Integer) As String
'    Dim szSQL As String
'    Dim rsInv As New ADODB.Recordset
'    szSQL = "select slnumber from tblPurinv  where slnumber  IN (SELECT Top 10 slnumber from " & _
'          "tblPurInv where slnumber<=" & strCurInv & " AND  transactionType=" & transactionType & " order by slnumber DESC) and transactionType=" & transactionType & ""
'     rsInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'     If Not rsInv.EOF Then
'        ReturnLastTenPIPC = SQL2String(rsInv, 0)
'     End If
'
'End Function
Private Sub LoadFlxGridOneRow(adoConn As ADODB.Connection)

   
'Written by anol 20160911
' This shall refresh only one row.no need to reload full grid
   
    Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
    Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
 
   

   szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, " & _
               "Pt.OSAmount, QQ.PO, QQ.PO_ID,Supplier.Type,Pt.TransactionID,Pt.UserSessionID  " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "LEFT JOIN (" & _
                  "SELECT Q2.MY_ID, Q1.SLNumber AS PO, Q2.PO AS PO_ID " & _
                  "FROM ( " & _
                  "SELECT MY_ID, SLNumber " & _
                  "From tblPurInv " & _
                  "WHERE TransactionType = 25) AS Q1 INNER JOIN " & _
                  "(SELECT MY_ID, PO " & _
                  "From tblPurInv " & _
                  "WHERE PO <> '') AS Q2 ON Q1.MY_ID = Q2.PO " & _
               ") AS QQ ON PI.MY_ID = QQ.MY_ID " & _
           "Where PI.History = False AND PI.MY_ID='" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' AND (PI.TransactionType = 6 OR " & _
               "PI.TransactionType = 7) " & _
           "ORDER BY 3, 2;"
'Debug.Print szSQL
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iKount = flxPurchase.row
   With flxPurchase
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 13) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 16) = adoInv.Fields.Item("PostingDate").Value
         .TextMatrix(iKount, 17) = IIf(IsNull(adoInv.Fields.Item("PO").Value), "", adoInv.Fields.Item("PO").Value)
         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PO_ID").Value), "", adoInv.Fields.Item("PO_ID").Value)
         
         .TextMatrix(iKount, 19) = IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 20) = IIf(IsNull(adoInv.Fields.Item("Type").Value), "", adoInv.Fields.Item("Type").Value)
         .TextMatrix(iKount, 21) = IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value)
         .TextMatrix(iKount, 22) = IIf(IsNull(adoInv.Fields.Item("UserSessionID").Value), "", adoInv.Fields.Item("UserSessionID").Value)
'######################################################################################################################
'         Adding description of the header from the first split
         szSQL = "SELECT DISTINCT * " & _
                 "FROM tblPurInvSRec " & _
                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
                 "ORDER BY TRAN_ID;"
'Debug.Print szSQL
         adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         bFirstSp = True
         If Not adoInvSp.EOF Then _
            .TextMatrix(iKount, 8) = IIf(IsNull(adoInvSp.Fields.Item("DESCRIPTION").Value), "", adoInvSp.Fields.Item("DESCRIPTION").Value)

         adoInvSp.Close

         adoInv.MoveNext
       
        
      Wend
   End With

   adoInv.Close
   Set adoInv = Nothing

  
End Sub
Private Sub UpdtSupplierAccountBalance(adoConn As ADODB.Connection)
  'written by anol 17 Jan 2016
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Dr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type = 6 Or P.Type = 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.SupplierID = '" & txtSPSupplier.Tag & "' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  
   iIndex = 0
   If Not adoPayDr.EOF Then
      txtSupAcBal.text = Val(adoPayDr.Fields.Item("Dr").Value)
   End If

   adoPayDr.Close

   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Cr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                  "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
                  "FROM tlbPayment AS P " & _
                  "Where P.Type <> 6 And P.Type <> 24 " & _
                  "GROUP BY P.SageAccountNumber) AS P ON P.SageAccountNumber = S.SupplierID " & _
               "WHERE  S.SupplierID = '" & txtSPSupplier.Tag & "' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoPayCr.EOF Then
         txtSupAcBal.text = Val(txtSupAcBal.text) - Val(adoPayCr.Fields.Item("Cr").Value)
   End If

   adoPayCr.Close
   txtSupAcBal.text = Format(txtSupAcBal.text, "0.00")
   End Sub
Private Sub UpdateJobActualCost(szJobId As String, cCost As Currency, adoConn As ADODB.Connection)
   Dim szSQL As String

   If txtTransType.text = "Invoice" Then
      szSQL = "SET ActualCost = ActualCost + " & cCost & " "
   Else
      szSQL = "SET ActualCost = ActualCost - " & cCost & " "
   End If

   szSQL = "UPDATE PropertyMaintHistory " & szSQL & _
           "WHERE ID = '" & szJobId & "';"

   adoConn.Execute szSQL
End Sub

Private Function AccuRecoPer() As Single
   Dim i             As Integer
   Dim cOriginal     As Currency
   Dim cReco         As Currency

   For i = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(1, 1) = "" Then Exit For
      cOriginal = cOriginal + CCur(Val(flxPI.TextMatrix(i, 15)))
      cReco = cReco + CCur(Val(flxPI.TextMatrix(i, 15))) * Val(flxPI.TextMatrix(i, 22)) / 100
   Next i
   If cOriginal = 0 Then
      AccuRecoPer = 0
      Exit Function
   End If
   AccuRecoPer = cReco / cOriginal * 100
End Function

Private Sub flxPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   iSelected = 1
   SelectOnly1RowFlxGrid flxPI, flxPI.row, iXflxPI
End Sub

Private Sub flxPI_RowColChange()
'   cmdEdit(0).Enabled = True
'   cmdDelete.Enabled = True
   iCurEditRow = flxPI.row
   SelectOnly1RowFlxGrid flxPI, flxPI.row, iXflxPI
   
''   'added by anol 02 Jan 2015
''   'issue 469
''   If flxPI.RowHeight(flxPI.row) = 0 Then Exit Sub
''      If flxPI.TextMatrix(flxPI.row, 1) = "" Then Exit Sub
''    With flxPI
''         txtUnit(0).text = .TextMatrix(.row, 21)
''         txtNC(0).text = .TextMatrix(.row, 7)
''         txtDept(1).text = .TextMatrix(.row, 8)
''         txtDept(0).text = .TextMatrix(.row, 24)
''         txtPFName.text = .TextMatrix(.row, 25)
''         txtJobNo.text = .TextMatrix(.row, 9)
''         txtDetails_(0).text = .TextMatrix(.row, 11)
''         txtNet_(0).text = .TextMatrix(.row, 12)
''         lblVatCode(0).Caption = .TextMatrix(.row, 13)
''         txtVat_(0).text = .TextMatrix(.row, 14)
''         txtSchedules.text = .TextMatrix(.row, 20)
''         txtRecoverable(0).text = .TextMatrix(.row, 22)
''         chkRecover.Value = IIf(Val(txtRecoverable(0).text) > 0, 1, 0)
''         txtTotal.text = .TextMatrix(.row, 15)
''
''         txtInv(0).Locked = False
''         txtUnit(0).Locked = True
''         txtNC(0).Locked = True
''         txtDept(1).Locked = True
''         txtDept(0).Locked = True
''         txtPFName.Locked = True
''         txtJobNo.Locked = True
''         txtDetails_(0).Locked = True
''         txtNet_(0).Locked = True
''         'lblVatCode(0).Locked = True
''         txtVat_(0).Locked = True
''         txtSchedules.Locked = True
''         txtRecoverable(0).Locked = True
''         'chkRecover.Value = IIf(Val(txtRecoverable(0).text) > 0, 1, 0)
''         txtTotal.Locked = True
''      End With
End Sub

Private Sub SupplierComboBox()
   ' Error Handler
   On Error GoTo Error_Handler

   Dim conClient As New ADODB.Connection
   Dim rstClient As New ADODB.Recordset
   Dim szSQL As String, rRow As Integer

   ' Error Handler
   On Error GoTo Error_Handler

   conClient.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   If rstClient.EOF Then GoTo NoRes

   rRow = 1

NoRes:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

 Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hwnd)
    UnLoadForm Me
    Call ClearLockFromPIPayment
    frmLockingDialogisActive = False
    UserSessionID = ""
    'frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub
Private Sub ClearLockFromPIPayment()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub fraEditDemand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow

   If fraInvCrChoice.Visible Then cmdManualDmdCancel_Click
End Sub

Private Sub fralay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraLay(Index).MousePointer = vbArrow
   Me.MousePointer = vbArrow
End Sub

Private Sub fraTab0_DblClick()
   'MsgBox fraTab0.Left
End Sub

Private Sub lblPayPostingDate_Click()
   'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClientIDPurPay.text = "" Then
       ShowMsgInTaskBar "Please select a client", "Y"
       txtClientIDPurPay.SetFocus
       Exit Sub
    End If
    DispayCalendar Me, lblPayPostingDate.ToolTipText, txtSPDate.text, txtClientIDPurPay.text
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   If Not isValidClient Then
       ShowMsgInTaskBar "Please select a client", "Y"
       txtClientIDPurPay.SetFocus
       Exit Sub
    End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, txtClientID.text
End Sub

Private Sub optChqRemitt_Click()
   txtChqNo.Locked = False
   SelTxtInCtrl txtChqNo
   txtChqNo.SetFocus
End Sub

Private Sub tabPayment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tabPayment.MousePointer = vbArrow
End Sub

Private Function LoadDefaultBankAC(adoConn As ADODB.Connection)
   On Error GoTo Error_Handler
    
    Dim iRec As Integer
    Dim adoRST As New ADODB.Recordset
    Dim szSQL As String
    Dim iDefaultBankAC As Integer
    iDefaultBankAC = -1
    Dim defaultBankCode As String
    Dim defaultBankAC As String
    Dim bKeepOldValues As Boolean
    If txtClientIDPurPay.text = "ALL" Or txtClientIDPurPay.text = "" Then
         txtBankCode.text = ""
         txtBankAc.text = ""
         Exit Function
    End If
    If txtClientIDPurPay.text = "ALL" Then
        szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                   "NominalLedger.Name AS BNN, tlbClientBanks.CurrentBalance AS BAL, AllowOverDraft, OverDraftLimit,DEFAULT_AC  " & _
               "FROM tlbClientBanks, NominalLedger " & _
               "WHERE tlbClientBanks.NominalCode = NominalLedger.Code " & _
               "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance, AllowOverDraft, OverDraftLimit,DEFAULT_AC;"
    Else
        szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                   "NominalLedger.Name AS BNN, tlbClientBanks.CurrentBalance AS BAL, AllowOverDraft, OverDraftLimit,DEFAULT_AC  " & _
               "FROM tlbClientBanks, NominalLedger " & _
               "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
                   "tlbClientBanks.CLIENT_ID = '" & txtClientIDPurPay.text & "' AND " & _
                   "NominalLedger.ClientID = '" & txtClientIDPurPay.text & "' " & _
               "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance, AllowOverDraft, OverDraftLimit,DEFAULT_AC;"
    End If
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If adoRST.EOF Then
          MsgBox "Please setup bank account for the client."
    Else
        
          While Not adoRST.EOF

                If CBool(adoRST.Fields.Item("DEFAULT_AC").Value) = True Then 'loading default Bank AC
                      iDefaultBankAC = iRec
                      defaultBankCode = adoRST.Fields.Item("BNC").Value
                      defaultBankAC = adoRST.Fields.Item("BNN").Value
                End If
                If txtBankCode.text = adoRST.Fields.Item("BNC").Value Then
                     'if previosly bankcode of for this client has already set to textbox which belongs to this client then do not set deafult bank account keep current bank code
                      bKeepOldValues = True
                End If
                iRec = iRec + 1
                adoRST.MoveNext
          Wend

   End If
      Set adoRST = Nothing
      If iDefaultBankAC < 0 Then
            MsgBox "Please set a default Client Bank Account for: " & txtClientIDPurPay.text & "", vbInformation, "Warning"
      End If
        If txtBankCode.text = "" And defaultBankCode <> "" Then
             txtBankCode.text = defaultBankCode
             txtBankAc.text = defaultBankAC
             Exit Function
        End If
        If bKeepOldValues = True Then
             'do nothing ,it is keeping previously selected bank code
        Else
             'selected bank code is not a part of this client
             txtBankCode.text = ""
             txtBankAc.text = ""
        End If
      
   Exit Function

   ' Error Handling Code
Error_Handler:
   Set adoRST = Nothing
   ShowMsgInTaskBar Err.description, "Y", "P"
End Function
Private Sub tabPurExp_Click(PreviousTab As Integer)
   picPurchaseHistory.Visible = False
   If fraSearch.Visible Then
        fraSearch.Visible = False
   End If
   Dim adoConn As New ADODB.Connection
   Dim iRow As Integer

   'Set the ADO Connections to the dataset
   adoConn.Open getConnectionString
   'added by anol 22 July 2015
   'issue 571
   If tabPurExp.Tab = 1 Then
        cmdSPClose.Visible = True
   Else
         cmdSPClose.Visible = False
   End If
    If tabPurExp.Tab = 0 Then
           cmdSearch.Caption = "Sea&rch"
           cmdSPSave.Enabled = False
           chkProperty.Value = 0
    End If
   If tabPurExp.Tab <> 1 Then 'clearing the  flag when moving between the TABS
         adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "' AND Module='Purchase Payment'"
   End If
   If tabPurExp.Tab = 1 Then
        'cmdSPSave.Enabled = True
        Dim szSQL As String
        Dim adoRST As New ADODB.Recordset
         szSQL = "SELECT DISTINCT CLIENTID " & _
               "FROM  Client " & _
               "ORDER BY CLIENTID;"
       If txtClientIDPurPay.text = "" Then
            adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If adoRST.RecordCount > 0 Then
                txtClientIDPurPay.text = adoRST.Fields("CLIENTID").Value
                'PrepareCboBC adoConn
                Call LoadDefaultBankAC(adoConn)  ''when you select a client it is loading a deafult Bank AC
                Call updateBankBalance
            End If
            adoRST.Close
      End If
  
      If IsLoadedAndVisible("frmCashbook") Then
         For iRow = 1 To frmCashbook.flxStatementReconcile.Rows - 1
            If frmCashbook.flxStatementReconcile.TextMatrix(iRow, 8) <> "" And _
                  frmCashbook.flxStatementReconcile.TextMatrix(iRow, 9) = "" Then
               frmCashbook.SavedPreBankRecTrans adoConn
               UpdatingCB adoConn
               Exit For
            End If
         Next iRow
      End If
     
      ConfigFlxSPayment
      ConfigFlxSCrPoA
      'cmdOpenClient.SetFocus
      'issue 382 Purchase payment screen is freezing after coming back from creating an invoice
      If cmdPayAllocateSave.Enabled Then
'         If MsgBox("Do like to cancel the allocation?", vbYesNo + vbQuestion, "Allocation") = vbNo Then
'            cmdPayAllocateSave.SetFocus
'            Exit Sub
'         End If

         Frame4(1).Visible = False
         tabPurExp.Enabled = True
         tabPayment.Enabled = True

         cmdPayAllocateSave.Enabled = False
         cmdPayAutomatic.Enabled = True
      End If
      Frame8(1).Enabled = True
      Frame5(5).Visible = True
      cmeRevereseAllocation.Visible = Frame5(5).Visible
      Frame5(5).Enabled = True
      Frame5(1).Visible = False
      cmdPayAllocate.Caption = "All&ocation Only"
      cmdSPSupplier.Enabled = True
      Label3(5).Visible = False
      txtAllocatedDiff(1).Visible = False
      Label3(1).Visible = False
      lblAllocating(1).Visible = False

      AllocDiscard
      
        Dim strTemp As String
        txtClientIDPurPay.ForeColor = vbBlack
        strTemp = isControlAccountSet(txtClientIDPurPay.text)
        If Len(strTemp) > 0 Then
            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientIDPurPay.text & _
            vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
            strTemp = ""
            txtClientIDPurPay.ForeColor = vbRed
            Exit Sub
        End If
      
   End If

   If tabPurExp.Tab = 2 Then
        cmdPaymentDiscard.Caption = "Clear"
        chkPropertyHist.Value = 0
        Rst1.Open "Select MaxPurChaseHist From shoppingCentre", adoConn, adOpenKeyset, adLockReadOnly
        If Not Rst1.EOF Then
                txtDisplayMaxPurchaseHist.text = Rst1!MaxPurChaseHist
        End If
        LoadFlxPurchHistory adoConn, ""
        
   End If

   If tabPurExp.Tab = 3 Then
         Rst1.Open "Select MaxPurPaymentHist From shoppingCentre", adoConn, adOpenKeyset, adLockReadOnly
        If Not Rst1.EOF Then
                txtDisplayMaxPurchPayHist.text = Rst1!MaxPurPaymentHist
        End If
        LoadFlxPurchPPHistory adoConn, ""
   End If

   If tabPurExp.Tab = 4 Then
      fraEditDemand.Top = 480
      fraEditDemand.Left = 120
      
'      fraLay(0).Left = 120
'      fraLay(0).Top = 360
   End If

''   If cmbSPSupplier.ListCount = 0 Then
''    'added by anol 22 Aug 2016
''   'SupplierAcBalByClient adoConn
''    'end of addition
''   'modified by anol 22 Aug 2016
''       ' LoadAllSupplierFlxGrd adoConn
''   End If
   If tabPurExp.Tab = 1 Then
'            cmbSPSupplier.ListIndex = -1
        txtSPSupplier.text = ""
        txtSPSupplier.Tag = ""
   End If
   adoConn.Close
   Set adoConn = Nothing
   If tabPurExp.Tab = 1 Then
   'added by anol 19 Oct 2015
        txtClientIDPurPay.Enabled = True
        cmdOpenClient.Enabled = True
        focuscmdOpenClient
   End If
End Sub
Private Sub focuscmdOpenClient()
    On Error GoTo Err
        cmdOpenClient.SetFocus
    Exit Sub
Err:
End Sub
Private Sub tabPurExp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tabPurExp.MousePointer = vbArrow

   If tabPurExp.Tab = 0 Then
      If fraInvCrChoice.Visible Then cmdManualDmdCancel_Click
   End If
End Sub

Private Sub tabPurExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If tabPurExp.Tab = 0 Then
      fraEditDemand.Left = 120
      fraTab0.Left = fraEditDemand.Left + fraEditDemand.Width + 100
   End If
   If tabPurExp.Tab = 1 Then tabPayment.Left = 120
   If tabPurExp.Tab = 2 Then fraTab2.Left = 120
   If tabPurExp.Tab = 3 Then fraTab3.Left = 0
End Sub

Private Sub txtAc_Change(Index As Integer)
'anol 28 July 201?
'    If Len(txtAc(0).text) > 0 Then
'        'cmdSavePI.Enabled = True
'    End If
'Added by anol 05 Jan 2016
    If Left(fraLay(1).Caption, 15) = "Transaction ID:" And Index = 0 Then
        flxPI.Tag = "EditedOrAdded"
        cmdSavePI.Enabled = True
    End If
End Sub

Private Sub txtAc_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdACList_Click (tabPurExp.Tab)
   End If
End Sub

Private Sub txtAccountSearch_Change(Index As Integer)
   Dim i As Integer
   Dim X As Integer
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim tempstr As String
   tempstr = txtAccountSearch(Index).text
   tempstr = Replace(tempstr, "'", "''")
   
 'this part is for searching in in the supplier grid in payment section
    If Index = 1 Then
         If Trim(txtAccountSearch(Index).text) = "" Then
             Call LoadSupplierOnPayment(adoConn, "")
             Exit Sub
         End If
        Call LoadSupplierOnPayment(adoConn, "SupplierID Like '%" & tempstr & "%'")
    ElseIf Index = 2 Then
         If Trim(txtAccountSearch(Index).text) = "" Then
             Call LoadSupplierOnPayment(adoConn, "")
             Exit Sub
        End If
        Call LoadSupplierOnPayment(adoConn, "SupplierName Like '%" & tempstr & "%'")
   End If
   
   
   If Index >= 0 And Index <= 2 Then X = 1
   If Index >= 3 And Index <= 5 Then X = 2
   If Index = 0 Or Index = 3 Then
      For i = 1 To flxSupplier(X).Rows - 1
         If UCase(txtAccountSearch(Index).text) = UCase(Left(flxSupplier(X).TextMatrix(i, 0), Len(txtAccountSearch(Index).text))) Then
            flxSupplier(X).RowHeight(i) = 240
         Else
            flxSupplier(X).RowHeight(i) = 0
         End If
      Next i
   End If
'   If Index = 1 Then
'   'written by anol 20160524
'      For i = flxSupplier(X).Rows - 1 To 1 Step -1
'         'If UCase(txtAccountSearch(Index).text) = UCase(Left(flxSupplier(X).TextMatrix(i, 1), Len(txtAccountSearch(Index).text))) Then
'         flxSupplier(X).RowHeight(i) = 240
'         If InStr(1, UCase(flxSupplier(X).TextMatrix(i, 1)), UCase(txtAccountSearch(Index).text), vbTextCompare) = 0 Then
'            flxSupplier(X).RowHeight(i) = 0
'         End If
'      'Next i
'          If flxSupplier(X).RowHeight(i) = 240 Then
'                flxSupplier(X).row = i
'          End If
'
'       Next i
'   End If
   
   
   If Index = 4 Then
      For i = 1 To flxSupplier(X).Rows - 1
         If UCase(txtAccountSearch(Index).text) = UCase(Left(flxSupplier(X).TextMatrix(i, 1), Len(txtAccountSearch(Index).text))) Then
            flxSupplier(X).RowHeight(i) = 240
         Else
            flxSupplier(X).RowHeight(i) = 0
         End If
      Next i
   End If
   If Index = 2 Then
   'written by anol 20160524
'      For i = flxSupplier(X).Rows - 1 To 1
'         flxSupplier(X).RowHeight(i) = 240
'         If InStr(1, UCase(flxSupplier(X).TextMatrix(i, 1)), UCase(txtAccountSearch(Index).text), vbTextCompare) = 0 Then
'         'If UCase(txtAccountSearch(Index).text) = UCase(Left(flxSupplier(X).TextMatrix(i, 2), Len(txtAccountSearch(Index).text))) Then
'
'         'Else
'            flxSupplier(X).RowHeight(i) = 0
'         End If
'         If flxSupplier(X).RowHeight(i) = 240 Then
'                flxSupplier(X).row = i
'          End If
'      Next i
 'written by anol 201600909
        For i = flxSupplier(X).Rows - 1 To 1 Step -1
            flxSupplier(X).RowHeight(i) = 240
            
            If InStr(1, UCase(flxSupplier(X).TextMatrix(i, 2)), UCase(txtAccountSearch(2).text), vbTextCompare) = 0 Then
                  flxSupplier(X).RowHeight(i) = 0
            End If
            If flxSupplier(X).RowHeight(i) = 240 Then
                  flxSupplier(X).row = i
            End If
       Next i
   End If
   If Index = 5 Then
      For i = 1 To flxSupplier(X).Rows - 1
         If UCase(txtAccountSearch(Index).text) = UCase(Left(flxSupplier(X).TextMatrix(i, 2), Len(txtAccountSearch(Index).text))) Then
            flxSupplier(X).RowHeight(i) = 240
         Else
            flxSupplier(X).RowHeight(i) = 0
         End If
      Next i
   End If
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub txtAccountSearch_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown And Index = 0 Then
        flxSupplier(1).SetFocus
    End If
     If Index = 5 And KeyCode = 13 Then
        FocusControl flxSupplier(2)
    End If
    If Index = 7 And KeyCode = 13 Then
        flxSupplier(1).SetFocus
    End If
End Sub

Private Sub txtClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdClientSerc.SetFocus
    End If
End Sub



Private Sub txtClientIdlist_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
          cmdOClientList.SetFocus
    End If
End Sub

Private Sub txtClientIDPurPay_Change_made()
    On Error GoTo Err
   If txtClientIDPurPay.text = "" Then Exit Sub

   Dim iRow As Integer

   If txtClientIDPurPay.Value = "ALL" Then
      For iRow = 1 To flxSPayment.Rows - 1
         flxSPayment.RowHeight(iRow) = 240
         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
            iRow = iRow + 1
            While flxSPayment.TextMatrix(iRow, 0) = "-"
               flxSPayment.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      Next iRow
   
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   txtBankCode.text = ""
   txtBankAc.text = ""
   adoConn.Close
   Set adoConn = Nothing
   'added by anol 20 July 2015 issue 571
   'The details displayed in purchase payment for 1/ Bank Account 2/ Payment Type, 3/ Reference and 4/ Total Amt £ should be cleared
'every time a user changes client.
      txtPayAmtType.text = ""
      txtPayAmtType.Tag = ""
      txtSPReference.text = ""
      txtSPaymentTotal.text = "0.00"
   'added by anol 09 July 2015
   'issue 571
   '11. Purchase payment debit is not filtering by client. It is not clearing.
   For iRow = 1 To flxSCrPoA.Rows - 1
      If flxSCrPoA.TextMatrix(iRow, 18) <> txtClientIDPurPay.text And _
            flxSCrPoA.TextMatrix(iRow, 18) <> "" Then
         flxSCrPoA.RowHeight(iRow) = 0
      Else
         flxSCrPoA.RowHeight(iRow) = 240

      End If
   Next iRow
'Error Occurs here 02 Dec 2014
'Issue No 508
'outstanding- anol
   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 23) <> txtClientIDPurPay.text And _
            flxSPayment.TextMatrix(iRow, 23) <> "" Then
         flxSPayment.RowHeight(iRow) = 0
         If flxSPayment.TextMatrix(iRow, 0) <> "" Then
            iRow = iRow + 1
            While flxSPayment.TextMatrix(iRow, 0) = "-"
               flxSPayment.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      Else
         flxSPayment.RowHeight(iRow) = 240
         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
            iRow = iRow + 1
            While flxSPayment.TextMatrix(iRow, 0) = "-"
               flxSPayment.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      End If
   Next iRow
   
   Exit Sub
Err:
End Sub

Private Sub txtClientIDPurPay_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If KeyAscii = 13 Then
        FocusControl cmdOpenClient
    End If
End Sub

'Private Sub txtClientIDPurPay_LostFocus()
'   'issue 571 Validation
'   'Added by anol 21 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim adoConn    As New ADODB.Connection
'   If Trim(txtClientIDPurPay.text) <> "" Then
'        adoConn.Open getConnectionString
'        szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'                "FROM CLIENT " & _
'                "where CLIENTNAME='" & txtClientIDPurPay.text & "';"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'            'It shall give message if client is wrong
'            'and cleared down the property
'            tabPurExp.Tab = 1
'            MsgBox "You must select the client", vbInformation, "Select a client"
'            txtClientIDPurPay.SetFocus
'        Else
'             'if client is not wrong it shall load proper property
'            txtClientIDPurPay_Change
'        End If
'        adoConn.Close
'        Set adoConn = Nothing
'    End If
'End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
   'Modified by BOSL
   'Issue 468
   'Modified by Anol 25 Mar 2015
   lblPostingDate.ToolTipText = txtDate.text
   If Left(fraLay(1).Caption, 15) = "Transaction ID:" Then
        flxPI.Tag = "EditedOrAdded"
        cmdSavePI.Enabled = True
    End If
End Sub

Private Sub txtDate_GotFocus()
   If txtDate.text = "dd/mm/yyyy" Then
      txtDate.text = ""
      Exit Sub
   End If
   If Len(txtDate.text) < 10 Then
      txtDate.text = Format(Date, "dd/mm/yyyy")
'      lblPostingDate.ToolTipText = txtDate.text
   End If
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          txtDueDate.SetFocus
    End If
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   On Error Resume Next
         'Resolved by BOSL
         'Issue 468
         'Modified by Anol 03 Dec 2014
          If txtDate.text <> "" Then TextBoxFormatDate txtDate
          If IsDate(lblPostingDate.ToolTipText) = True Then
              Dim adoConn As New ADODB.Connection
              Dim szSQL As String
              If Not isValidClient Then
                  ShowMsgInTaskBar "Please select a client", "Y", "N"
                  cmdClientSerc.SetFocus
                  Exit Sub
              End If
              adoConn.Open getConnectionString
              If IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 0 Then
                  ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
                  adoConn.Close
                  Exit Sub
              ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, txtClientID.text, adoConn) = 9 Then
                  ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
                  adoConn.Close
                  Exit Sub
              End If
           End If
              
   
   'If txtDate.text <> "" And cmdEdit(1).Enabled Then txtDueDate.text = DateAdd("d", iDayTerms, txtDate.text)
   If txtDate.text <> "" Then txtDueDate.text = DateAdd("d", iDayTerms, txtDate.text)
End Sub

Private Sub txtDept_Change(Index As Integer)
    If Len(txtDept(0).text) > 0 Then
        cmdUpdate(1).Enabled = True
   End If
End Sub

Private Sub txtDept_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdDeptList_Click
   End If
End Sub

Private Sub txtDetails__Change(Index As Integer)
    If Val(txtDetails_(0).text) > 0 Then
        cmdUpdate(1).Enabled = True
    End If
End Sub

Private Sub txtDetails__KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 0 Then
        txtNet_(0).SetFocus
    End If
End Sub

Private Sub txtDueDate_Change()
   TextBoxChangeDate txtDueDate
   If Left(fraLay(1).Caption, 15) = "Transaction ID:" Then
        flxPI.Tag = "EditedOrAdded"
        cmdSavePI.Enabled = True
    End If
End Sub

Private Sub txtDueDate_GotFocus()
   If txtDueDate.text = "dd/mm/yyyy" Then
      txtDueDate.text = ""
      Exit Sub
   End If
   If Len(txtDueDate.text) < 10 Then txtDueDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDueDate
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdClientSerc.SetFocus
    End If
   TextBoxKeyPrsDate txtDueDate, KeyAscii
End Sub

Private Sub txtDueDate_LostFocus()
   If txtDueDate.text <> "" Then TextBoxFormatDate txtDueDate
End Sub

Private Sub txtInv_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 0 Then
            txtDate.SetFocus
    End If
End Sub

Private Sub txtJobNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
            cmdSchedules(0).SetFocus
           'SendKeys vbTab
    End If
End Sub

Private Sub txtJobNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
          cmdSchedules(0).SetFocus
              'SendKeys vbTab
    End If
End Sub

Private Sub txtJobNo_LostFocus()
'issue 0000571: PRESTIGE VALIDATION ALL MODULES AND FORMS
'Note 1097
  If Trim(txtJobNo.text) <> "" Then
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Dim adoRST As New ADODB.Recordset
        Dim szSQL As String
        szSQL = "SELECT ID, Job_DiaryName " & _
           "FROM PropertyMaintHistory " & _
           "WHERE ID = '" & txtJobNo.text & "' " & _
           "ORDER BY ID;"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If adoRST.EOF Then
            MsgBox "Please select a valid Job Number to proceed.", vbCritical + vbOKOnly, "Select a job"
            txtJobNo.SetFocus
            Exit Sub
        End If
        adoRST.Close
        adoConn.Close
   End If
End Sub

Private Sub txtNet__Change(Index As Integer)
    If Index = 0 And Val(txtNet_(0).text) > 0 Then
        cmdUpdate(1).Enabled = True
    End If
End Sub

Private Sub txtNet__GotFocus(Index As Integer)
   txtNet_(Index).SelStart = 0
   txtNet_(Index).SelLength = Len(txtNet_(Index).text)
End Sub

Private Sub txtPaymentEntered_Change()
   txtDiffPay.text = Format(Val(txtPaymentTotal.text) - Val(txtPaymentEntered.text), "0.00")
End Sub

Private Sub txtPaymentTotal_Change()
   txtDiffPay.text = Format(Val(txtPaymentTotal.text) - Val(txtPaymentEntered.text), "0.00")
End Sub

Private Sub txtNC_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdNCList_Click
   End If
End Sub

Private Sub txtNet__KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And Index = 0 Then
        If cmdUpdate(1).Enabled = True Then
            cmdUpdate(1).SetFocus
        End If
   End If
   If KeyAscii = 13 Or KeyAscii = 10 Then txtNet__LostFocus (tabPurExp.Tab)

   DigitTextKeyPress txtNet_(0), KeyAscii
End Sub

Private Sub txtNet__LostFocus(Index As Integer)
  'Modifed by BOSL
    'Issue 463:Manually Typing in the VAT Value
   ' If vatOptionEnabled = True Then
        txtVat_(0).text = Format(IIf(txtNet_(0).text = "", 0, Val(txtNet_(0).text)) * (nTaxCode / 100), "0.00")
        txtNet_(0).text = Format(txtNet_(0).text, "0.00")
        'below code is wrong
'    Else
'        If Val(txtNet_(0).text) = 0 Then
'            txtVat_(0).text = ""
'        Else
'            txtVat_(0).text = "0.00"
'        End If
    'End If
    txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
    txtTotal.text = Format(txtTotal.text, "0.00")
    
End Sub

Private Sub txtProperty_Change()
    'Added by anol 13 Aug 2015
    If Left(fraLay(1).Caption, 15) = "Transaction ID:" Then
'        flxPI.Tag = "EditedOrAdded" this line is remmed beacseu if you enable it you shall have problem saving button on reference
        cmdSavePI.Enabled = True
    End If
End Sub

Private Sub txtRecoverable_GotFocus(Index As Integer)
   If Index = 0 Then
      SelTxtInCtrl txtRecoverable(0)
   End If
End Sub

Private Sub txtRecoverable_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   DigitTextKeyPress txtRecoverable(0), CInt(KeyAscii), 2
End Sub

Private Sub txtSchedules_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtSchedules.text = ""
    End If
    If KeyCode = 13 Then
            FocusControl txtDetails_(0)
        End If
End Sub

Private Sub txtSchedules_KeyPress(KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 13 Then
            txtDetails_(0).SetFocus
        End If
End Sub

Private Sub txtSearch1_Change()
    'following method is not selecting the correct row on enter press
    'this line  flxSupplier(0).row = i does not work
'   Dim i As Integer
'
   If Len(txtSearch1.text) > 0 Then
      txtSearch2 = ""
   End If
'
'   For i = flxSupplier(0).Rows - 1 To 1 Step -1
'      flxSupplier(0).RowHeight(i) = 240
'      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtSearch1.text))) <> UCase(txtSearch1.text) Then
'         flxSupplier(0).RowHeight(i) = 0
'      End If
''      If flxSupplier(0).RowHeight(i) = 240 Then
''        flxSupplier(0).row = i
''      End If
'   Next i
'   For i = 1 To flxSupplier(0).Rows - 1
'   Debug.Print flxSupplier(0).RowHeight(i)
'       If flxSupplier(0).RowHeight(i) = 240 Then
'          flxSupplier(0).row = i
'          Exit For
'       End If
'   Next
    Dim tempstr As String
    If sSearchSwitch = "ID" Then
            tempstr = txtSearch1.text
            tempstr = Replace(tempstr, "'", "''")
            If sTextBox = "A/C" Then
                        If Trim(txtSearch1.text) = "" Then
                             Call LoadSupplierAccount("")
                             Exit Sub
                        End If
               
               
                        If cmbSC.text = "Supplier" Then
                            Call LoadSupplierAccount("SupplierID Like '%" & tempstr & "%'")
                        ElseIf cmbSC.text = "Client" Then
                            Call LoadSupplierAccount("clientID Like '%" & tempstr & "%'")
                        ElseIf cmbSC.text = "Landlord" Then
                            Call LoadSupplierAccount("SupplierID Like '%" & tempstr & "%'")
                        ElseIf cmbSC.text = "Managing Agent" Then
                             Call LoadSupplierAccount("SupplierID Like '%" & tempstr & "%'")
                        End If
                
             ElseIf sTextBox = "PILIST" Or sTextBox = "PIHIST" Or sTextBox = "PAYHIST" Then
                         If Trim(txtSearch1.text) = "" Then
                             Call LoadSupplierAccount("")
                             Exit Sub
                         End If
                         Call LoadSupplierAccount("SupplierID Like '%" & tempstr & "%'")
               
            ElseIf sTextBox = "PROPERTY" Or sTextBox = "PROPERTYFILTER" Or sTextBox = "PROPERTYHIST" Then
                    If Trim(txtSearch1.text) = "" Then
                         Call LoadPropertyList("")
                         Exit Sub
                    End If
                    LoadPropertyList ("propertyID Like '%" & tempstr & "%'")
                    
            ElseIf sTextBox = "NC" Then
                    If Trim(txtSearch1.text) = "" Then
                         Call LoadNominalCode("")
                         Exit Sub
                    End If
                    LoadNominalCode ("Code Like '%" & tempstr & "%'")
            ElseIf sTextBox = "Fund" Then
                    If Trim(txtSearch1.text) = "" Then
                         Call LoadFund("")
                         Exit Sub
                    End If
                    LoadFund ("FundCode Like '%" & tempstr & "%'")
             ElseIf sTextBox = "job" Then
                    If Trim(txtSearch1.text) = "" Then
                         Call LoadJobSheet("")
                         Exit Sub
                    End If
                    LoadJobSheet ("ID Like '%" & tempstr & "%'")
            ElseIf sTextBox = "Schedules" Then
                    
                    If Trim(txtSearch1.text) = "" Then
                         Call LoadSchedules("")
                         Exit Sub
                    End If
                    LoadSchedules ("ScheduleID Like '%" & tempstr & "%'")
            End If
    End If
End Sub

Private Sub txtSearch1_KeyPress(KeyAscii As MSForms.ReturnInteger)
         'Resolved by BOSL
         'Issue 553 PRESTIGE GUI IMPROVEMENT
         'Added by Anol 25 Mar 2015
         
    If KeyAscii = 13 Then
        If Len(txtSearch1.text) Then
                flxSupplier(0).SetFocus
        End If
    End If
    If KeyAscii = 27 Then
          flxSupplier(0).Clear
          flxSupplier(0).Clear
          flxSupplier(0).Cols = 2
          flxSupplier(0).Rows = 2
          fraList.Visible = False
          tabPurExp.Enabled = True
          'Resolved by BOSL
          'Below line are modified by anol 29 Mar 2015
          'issue 553 : PRESTIGE GUI IMPROVEMENT
          If tabPurExp.Tab = 0 Then
                 If fraLay(0).Top = 360 Then
                        If sTextBox = "A/C" Then cmdACList(tabPurExp.Tab).SetFocus
                        If sTextBox = "UNIT" Then cmdUnitList.SetFocus
                        If sTextBox = "PROPERTY" Then cmdTypeList.SetFocus
                        If sTextBox = "NC" Then txtNC(0).SetFocus
                        If sTextBox = "job" Then cmdSchedules(0).SetFocus 'cmdJobNo(0).SetFocus
                        If sTextBox = "Schedules" Then txtDetails_(0).SetFocus 'cmdSchedules(0).SetFocus
                        If sTextBox = "Fund" Then cmdDeptList.SetFocus
                        If sTextBox = "VAT" Then cmdTaxList(tabPurExp.Tab).SetFocus
                Else
                         If sTextBox = "A/C" Then cmdAccSel.SetFocus
                         If sTextBox = "PROPERTY" Then cmdOpProperty.SetFocus
                         If sTextBox = "PROPERTYHIST" Then cmdOpPropertyHist.SetFocus
                End If
          End If
    End If
End Sub

Private Sub txtSearch2_Change()
    'following method is not selecting the correct row on enter press
    'this line  flxSupplier(0).row = i does not work
'   Dim i As Integer
'
   If Len(txtSearch2.text) > 0 Then
      txtSearch1.text = ""
   End If
'
'   For i = flxSupplier(0).Rows - 1 To 1 Step -1
'      flxSupplier(0).RowHeight(i) = 240
'
'      If UCase(Left(flxSupplier(0).TextMatrix(i, 1), Len(txtSearch2.text))) <> UCase(txtSearch2.text) Then
'         flxSupplier(0).RowHeight(i) = 0
'      End If
'      If flxSupplier(0).RowHeight(i) = 240 Then
'        flxSupplier(0).row = i
'      End If
'   Next i

    Dim tempstr As String
    If sSearchSwitch = "Name" Then
            tempstr = txtSearch2.text
            tempstr = Replace(tempstr, "'", "''")
            If Trim(txtSearch2.text) <> "" Then
                    txtSearch1.text = ""
            End If
           
            If sTextBox = "A/C" Then
                 If Trim(txtSearch2.text) = "" Then
                     Call LoadSupplierAccount("")
                     Exit Sub
                 End If
                    If cmbSC.text = "Supplier" Then
                        Call LoadSupplierAccount("SupplierName Like '%" & tempstr & "%'")
                    ElseIf cmbSC.text = "Client" Then
                        Call LoadSupplierAccount("clientName Like '%" & tempstr & "%'")
                    ElseIf cmbSC.text = "Landlord" Then
                        Call LoadSupplierAccount("SupplierName Like '%" & tempstr & "%'")
                    ElseIf cmbSC.text = "Managing Agent" Then
                         Call LoadSupplierAccount("SupplierName Like '%" & tempstr & "%'")
                    End If
             ElseIf sTextBox = "PILIST" Or sTextBox = "PIHIST" Or sTextBox = "PAYHIST" Then
                         If Trim(txtSearch2.text) = "" Then
                             Call LoadSupplierAccount("")
                             Exit Sub
                         End If
                         Call LoadSupplierAccount("SupplierName Like '%" & tempstr & "%'")
               
            ElseIf sTextBox = "PROPERTY" Or sTextBox = "PROPERTYFILTER" Or sTextBox = "PROPERTYHIST" Then
                    If Trim(txtSearch2.text) = "" Then
                         Call LoadPropertyList("")
                         Exit Sub
                    End If
                    LoadPropertyList ("propertyName Like '%" & tempstr & "%'")
        
            ElseIf sTextBox = "NC" Then
                    If Trim(txtSearch2.text) = "" Then
                         Call LoadNominalCode("")
                         Exit Sub
                    End If
                    LoadNominalCode ("Name Like '%" & tempstr & "%'")
            ElseIf sTextBox = "Fund" Then
                    If Trim(txtSearch2.text) = "" Then
                         Call LoadFund("")
                         Exit Sub
                    End If
                    LoadFund ("FundName Like '%" & tempstr & "%'")
            ElseIf sTextBox = "job" Then
                    If Trim(txtSearch2.text) = "" Then
                         Call LoadJobSheet("")
                         Exit Sub
                    End If
                    LoadJobSheet ("Job_DiaryName Like '%" & tempstr & "%'")
             ElseIf sTextBox = "Schedules" Then
                    
                    If Trim(txtSearch2.text) = "" Then
                         Call LoadSchedules("")
                         Exit Sub
                    End If
                    LoadSchedules ("ScheduleName Like '%" & tempstr & "%'")
            End If
       End If
End Sub

Private Sub txtSearch2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        flxSupplier(0).SetFocus
       ' SendKeys vbKeyDown
    End If
'    'Resolved by BOSL
'   'Issue 553 PRESTIGE GUI IMPROVEMENT
'   'Added by Anol 25 Mar 2015
'    Dim i As Integer
'    Dim k As Integer
'    If KeyCode = 13 Then
'       For i = 1 To flxSupplier(0).Rows - 1
'            If flxSupplier(0).RowHeight(i) = 240 Then
'                k = k + 1
'            End If
'       Next i
'       If k = 1 And txtSearch2.text <> "" Then
'            'MsgBox "One Record"
'            For i = 1 To flxSupplier(0).Rows - 1
'                If flxSupplier(0).RowHeight(i) = 240 Then
'                    flxSupplier(0).row = i
'                    flxSupplier_Click (0)
''                    If sTextBox = "A/C" Then cmdACList(tabPurExp.Tab).SetFocus
''                    If sTextBox = "UNIT" Or sTextBox = "PROP" Then cmdNCList.SetFocus
''                    If sTextBox = "NC" Then txtNC(0).SetFocus
''                    If sTextBox = "Dept" Then txtDept(0).SetFocus
''                    If sTextBox = "VAT" Then cmdTaxList(tabPurExp.Tab).SetFocus
'                    fraList.Visible = False
'                    Exit For
'                End If
'            Next i
'            fraList.Visible = False
'       End If
'       If k > 1 And txtSearch2.text = "" Then
'            If fraList.Visible = True Then
'                 flxSupplier(0).SetFocus
'            End If
'       End If
'
'    End If
End Sub

Private Sub txtSearch2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Resolved by BOSL
   'Issue 553 PRESTIGE GUI IMPROVEMENT
   'Added by Anol 25 Mar 2015
    If KeyAscii = 27 Then
      flxSupplier(0).Clear
      flxSupplier(0).Clear

      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2
      fraList.Visible = False
      tabPurExp.Enabled = True
      If sTextBox = "A/C" Then cmdACList(tabPurExp.Tab).SetFocus
      If sTextBox = "UNIT" Or sTextBox = "PROP" Then cmdNCList.SetFocus
      If sTextBox = "NC" Then txtNC(0).SetFocus
      If sTextBox = "Fund" Then txtDept(0).SetFocus
      If sTextBox = "VAT" Then cmdTaxList(tabPurExp.Tab).SetFocus
   End If
'    If KeyAscii = 13 Then
'        MsgBox "Hi"
'    End If
    
   'End of addition
End Sub

Private Sub txtSearchClientID_Change()
    If bSearchClientNameFocus = False Then
          txtSearchClientName.text = ""
          Dim tempstr As String
          If sTextBox = "BankAcPay" Then
             If Trim(txtSearchClientID.text) = "" Then
                  Call LoadflxBankAC("")
                  Exit Sub
             End If
             tempstr = txtSearchClientID.text
             tempstr = Replace(tempstr, "'", "''")
             'Call LoadflxBankAC("tlbClientBanks.NominalCode Like '%" & tempstr & "%'")
             Call LoadflxBankAC("BNC Like '%" & tempstr & "%'")
             Exit Sub
         End If
        
         If Trim(txtSearchClientID.text) = "" Then
              Call LoadflxClient("")
              Exit Sub
         End If
         tempstr = txtSearchClientID.text
         tempstr = Replace(tempstr, "'", "''")
         Call LoadflxClient("ClientID Like '%" & tempstr & "%'")
    End If
           
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyDown Then
            flxClient.SetFocus
     End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
         txtSearchClientName.SetFocus
    End If
    If KeyAscii = 27 Then
          flxClient.Clear
          flxClient.Cols = 2
          flxClient.Rows = 2
          picClient.Visible = False
          tabPurExp.Enabled = True
          tabPayment.Enabled = True
          'Resolved by BOSL
          'Below line are modified by anol 29 Mar 2015
          'issue 553 : PRESTIGE GUI IMPROVEMENT
          If tabPurExp.Tab = 0 Then
                 If fraLay(0).Top = 360 Then
                     cmdClientSerc.SetFocus
                 Else
                     cmdOpClient.SetFocus
                 End If
          End If
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
'   Dim i As Integer
'
'   If Len(txtSearchClientName.text) > 0 Then
'        txtSearchClientID.text = ""
'   End If
'
'   For i = flxClient.Rows - 1 To 1 Step -1
'      flxClient.RowHeight(i) = 240
'      If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
'            flxClient.RowHeight(i) = 0
'      End If
'      If flxClient.RowHeight(i) = 240 Then
'            flxClient.row = i
'      End If
'   Next i
    If bSearchClientNameFocus Then
        txtSearchClientID.text = ""
        Dim tempstr As String
         If sTextBox = "BankAcPay" Then
            If Trim(txtSearchClientName.text) = "" Then
                 Call LoadflxBankAC("")
                 Exit Sub
            End If
            tempstr = txtSearchClientName.text
            tempstr = Replace(tempstr, "'", "''")
            Call LoadflxBankAC("BNN Like '%" & tempstr & "%'")
            Exit Sub
        End If
        
        If Trim(txtSearchClientName.text) = "" Then
             Call LoadflxClient("")
             Exit Sub
        Else
             txtSearchClientID.text = ""
        End If
        tempstr = txtSearchClientName.text
        tempstr = Replace(tempstr, "'", "''")
        Call LoadflxClient("ClientName Like '%" & tempstr & "%'")
    End If
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        If flxClient.Rows > 1 And flxClient.row = 0 Then
            flxClient.row = 1
        End If
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            flxClient.SetFocus
        End If
    End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
         flxClient.SetFocus
    End If
End Sub

Private Sub txtSPayment_Click()
   SelTxtInCtrl txtSPayment
End Sub

Private Sub txtSPayment_GotFocus()
   If Not bTotalPayTyped Then _
      txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) - CCur(txtSPayment.text), "0.00")

   SelTxtInCtrl txtSPayment
   iCurRow = flxSPayment.row
   cGridSPTotal = cGridSPTotal - CCur(txtSPayment.text)
End Sub

Private Sub txtSPayment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then flxSPayment.SetFocus
End Sub

Private Sub txtSPayment_KeyPress(KeyAscii As Integer)
    
   DigitTextKeyPress txtSPayment, KeyAscii
End Sub

Private Sub txtSPayment_LostFocus()
   Dim i As Integer

   txtSPayment.text = Format(Val(txtSPayment.text), "0.00")
   flxSPayment.TextMatrix(flxSPayment.row, 15) = "A"

   If Val(flxSPayment.TextMatrix(iCurRow, 9)) < Val(txtSPayment.text) Then
      ShowMsgInTaskBar "Payment amount exceeds amount outstanding.", , "N"
      txtSPayment.text = "0.00"
      txtSPayment.SetFocus
      txtSPayment_GotFocus
      Exit Sub
   End If

   flxSPayment.TextMatrix(iCurRow, 10) = Format(CCur(txtSPayment.text), "0.00")

   If bTotalPayTyped Then
      If Val(txtSPaymentTotal.text) < cGridSPTotal + CCur(txtSPayment.text) Then
         txtSPayment.text = "0.00"
         flxSPayment.TextMatrix(iCurRow, 10) = Format(CCur(txtSPayment.text), "0.00")
         ShowMsgInTaskBar "Payment amount entered would exceed total payment amount.", , "N"
         txtSPayment.SetFocus
         SelTxtInCtrl txtSPayment
         Exit Sub
      End If
   Else
      If Not cmdAutoAllocSel.Visible Then
         txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + CCur(txtSPayment.text), "0.00")
      End If
      If Val(txtSPayment.text) > 0 Then flxSCrPoA.Enabled = False
   End If
   cGridSPTotal = cGridSPTotal + CCur(txtSPayment.text)

   baChangesMade(iCurRow) = IIf(CCur(flxSPayment.TextMatrix(iCurRow, 10)) > 0, True, False)
   txtSPayment.text = "0.00"

   txtSPayment.Visible = False
   flxSPayment.Enabled = True

   Call SumUpHeaderBySplits
   Call SpreadHeaderInSplits

   If Not cmdAutoAllocSel.Visible Then
      txtPaymentEntered.text = Format(TotalPaymentEntered, "0.00")
   End If
End Sub

Private Sub SumUpHeaderBySplits()
   Dim i As Integer, cSumSplits As Currency, j As Integer

   If flxSPayment.TextMatrix(iCurRow, 0) = "-" Then
      i = iCurRow
      Do
         i = i - 1
         If flxSPayment.TextMatrix(i, 0) = ">" Then Exit Do
      Loop While i > 0
      baChangesMade(i) = True
      j = i + 1
      cSumSplits = 0

      Do
         cSumSplits = cSumSplits + Val(flxSPayment.TextMatrix(j, 10))
         j = j + 1
         If j = flxSPayment.Rows Then Exit Do
      Loop While flxSPayment.TextMatrix(j, 0) = "-"

      flxSPayment.TextMatrix(i, 10) = cSumSplits
      flxSPayment.TextMatrix(i, 15) = "A"
   End If
End Sub

Private Sub txtSPaymentTotal_LostFocus()
   Dim iRow As Integer
   If txtSPaymentTotal.text = "-" Then Exit Sub
   If Trim(txtSPaymentTotal.text) = "" Then txtSPaymentTotal.text = "0.00"

   If Val(txtSPaymentTotal.text) >= 0 Then
      If (cGridSPTotal - CCur(txtSPaymentTotal.text) <> 0) And (cGridSPTotal + CCur(txtSPaymentTotal.text) <> 0) Then
         If CCur(txtSPaymentTotal.text) < cGridSPTotal Then
            ShowMsgInTaskBar "Total payment amount can not be changed to less than the analysis total.", , "N"
            txtSPaymentTotal.text = CStr(Format(cGridSPTotal, "0.00"))
            Exit Sub
         End If
      End If
      txtSPaymentTotal.text = Format(txtSPaymentTotal.text, "0.00")
      If cTempReceiptAmt - CCur(txtSPaymentTotal.text) <> 0 Then flxSCrPoA.Enabled = False
   Else                       '############ REFUND AMOUNT WHICH IS -ve #################
      If bChangesMade Then
         For iRow = 1 To flxSPayment.Rows - 1
            If Val(flxSPayment.TextMatrix(iRow, 10)) > 0 Then
               flxSPayment.TextMatrix(iRow, 10) = "0.00"
            End If
         Next iRow
         txtPaymentEntered.text = "0.00"
         cGridSPTotal = 0
         bTotalPayTyped = False
      End If
   End If
    txtSPaymentTotal.text = Format(txtSPaymentTotal.text, "0.00")
End Sub

Private Sub txtSPDate_Change()
   'Resolved by BOSL
    'Issue 468
    'Resolved by Anol 03 Sep 2014
   TextBoxChangeDate txtSPDate
   lblPayPostingDate.ToolTipText = txtSPDate.text
End Sub

Private Sub txtSPDate_GotFocus()
   If txtSPDate.text = "dd/mm/yyyy" Then
      txtSPDate.text = ""
      Exit Sub
   End If
   If Len(txtSPDate.text) < 10 Then
      txtSPDate.text = Format(Date, "dd/mm/yyyy")
      If lblPayPostingDate.ToolTipText = "" Then lblPayPostingDate.ToolTipText = txtSPDate.text
   End If
   SelTxtInCtrl txtSPDate
End Sub

Private Sub ConfigFlxPI()
'issue 469
'modified by anol 28 Dec 2014
'On Error GoTo ERR
   With flxPI
      .Clear
      .Cols = 27
      .Rows = 2
      .RowHeight(0) = 0     '                           Row Number of line
      .ColWidth(0) = 700 ' SL NO
      .ColAlignment(0) = vbLeftJustify
      .ColWidth(1) = 0 '
      .ColWidth(2) = 0 '
      .ColWidth(3) = 0 '
      .ColWidth(4) = 0 '
      .ColWidth(5) = Label7(33).Left - Label7(32).Left  '  Unit No
      .ColWidth(6) = Label7(34).Left - Label7(33).Left    'NominalCode
      .ColWidth(7) = Label7(35).Left - Label7(34).Left    'Fund Code
      .ColWidth(8) = 0  'Fund ID
      .ColWidth(9) = Label7(36).Left - Label7(35).Left   'Job No
      .ColWidth(10) = 0                     '"Cost Code"
      .ColWidth(11) = Label7(37).Left - Label7(36).Left - 300 'Details
      .ColAlignment(11) = vbLeftJustify
      .ColWidth(12) = Label7(38).Left - Label7(37).Left   'Net
      .ColWidth(13) = Label7(39).Left - Label7(38).Left   'T/C
      .ColAlignment(13) = vbRightJustify
      .ColWidth(14) = Label7(40).Left - Label7(39).Left   'VAT
      .ColWidth(15) = 1150 '   'Total
      .ColWidth(16) = 0                     '"Sage"
      .ColWidth(17) = 0           'Stores PI Id hidenly
      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
      .ColWidth(19) = 0           'keep value 0 or 1 for edit
      .ColWidth(20) = 0 'Label7(13).Left - Label7(12).Left           'Stores ScheduleId
      .ColWidth(21) = 0           'Stores Unit ID
      .ColWidth(22) = 0           '% Recoverable
      .ColWidth(23) = 0           'ID
      .ColWidth(24) = 0 '.Width - Label7(1).Left - 120           'FundCode
      .ColWidth(25) = 0           'FundName
      .ColWidth(26) = 0           'PO
      
      
'      .ColWidth(0) = Label7(4).Left - .Left '"TransactionID" SL NO
'      .ColWidth(1) = Label7(5).Left - Label7(4).Left '"A/C" SupplierId
'      .ColWidth(2) = Label7(6).Left - Label7(5).Left '"Date"
'      .ColWidth(3) = Label7(7).Left - Label7(6).Left '"Type" Property
'      .ColWidth(4) = Label7(8).Left - Label7(7).Left '"Trans"
'      .ColWidth(5) = Label7(9).Left - Label7(8).Left '"Unit ID + Name"
'      .ColWidth(6) = Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
'      .ColWidth(7) = 0                      '"N/C"
'      .ColWidth(8) = 0                      '"Fund"
'      .ColWidth(9) = 0                      '"Job No"
'      .ColWidth(10) = 0                     '"Cost Code"
'      .ColWidth(11) = Label7(11).Left - Label7(10).Left '"Details"
'      .ColWidth(12) = Label7(12).Left - Label7(11).Left '"Net"
'      .ColWidth(13) = Label7(13).Left - Label7(12).Left '"T/C"
'      .ColWidth(14) = Label7(14).Left - Label7(13).Left '"VAT"
'      .ColWidth(15) = .Width - Label7(14).Left - 120 '"Total"
'      .ColWidth(16) = 0                     '"Sage"
'      .ColWidth(17) = 0           'Stores PI Id hidenly
'      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
'      .ColWidth(19) = 0           'keep value 0 or 1 for edit
'      .ColWidth(20) = 0           'Stores ScheduleId
'      .ColWidth(21) = 0           'Stores Unit ID
'      .ColWidth(22) = 0           '% Recoverable
'      .ColWidth(23) = 0           'ID
'      .ColWidth(24) = 0           'FundCode
'      .ColWidth(25) = 0           'FundName
'      .ColWidth(26) = 0           'PO
      .row = 0
   End With

'   txtPICNNet.Left = Label7(11).Left
'   txtPICNNet.Width = flxPI.ColWidth(12)
'   txtPICNVat.Left = Label7(13).Left
'   txtPICNVat.Width = flxPI.ColWidth(14)
'   txtPICNTotal.Left = Label7(14).Left
'   txtPICNTotal.Width = flxPI.ColWidth(15)
''   txtPICNNet.Left = Label7(8).Left - 20
''   txtPICNNet.Width = flxPI.ColWidth(12)
''   txtPICNVat.Left = Label7(10).Left - 20
''   txtPICNVat.Width = flxPI.ColWidth(14)
''   txtPICNTotal.Left = Label7(11).Left - 20
''   txtPICNTotal.Width = flxPI.ColWidth(15)
   Exit Sub
Err:
End Sub

Private Sub BankFlxGridConfigure(conFlxGrid As Control)

   conFlxGrid.Clear
   conFlxGrid.Cols = 16
   
   conFlxGrid.ColWidth(0) = 700             '"Bank"
   conFlxGrid.ColWidth(1) = 700             '"Type"
   conFlxGrid.ColWidth(2) = 1000              '"Date"
   conFlxGrid.ColWidth(3) = 0              '"Client"
   conFlxGrid.ColWidth(4) = 0              '"Property"
   conFlxGrid.ColWidth(5) = 1000              '"Unit ID"
   conFlxGrid.ColWidth(6) = 700              '"N/C"
   conFlxGrid.ColWidth(7) = 3000             '"Ref"
   conFlxGrid.ColWidth(8) = 700             '"Fund"
   conFlxGrid.ColWidth(9) = 3000             '"Details"
   conFlxGrid.ColWidth(10) = 700             '"Net"
   conFlxGrid.ColWidth(11) = 500             '"T/C"
   conFlxGrid.ColWidth(12) = 800             '"TAX"

   conFlxGrid.ColWidth(13) = 1           'Stores BkPay Id hidenly
   conFlxGrid.ColWidth(14) = 1           'Marked X when row will be selected
   conFlxGrid.ColWidth(15) = 1           'keep value 0 or 1 for edit, 1=edit

   conFlxGrid.RowHeight(0) = 0
End Sub

Private Sub txtSPDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And flxSPayment.Enabled = True Then
       
        flxSPayment.col = 10
        'Fixed by anol 31 Aug 2015 issue 571 note 1156
        If flxSPayment.Rows > 1 Then
            flxSPayment.row = 1
             flxSPayment.SetFocus
        Else
            If cmdSPSave.Enabled Then
                cmdSPSave.SetFocus
            End If
        End If
    End If
   TextBoxKeyPrsDate txtSPDate, KeyAscii
End Sub

Private Sub txtSPDate_LostFocus()
   If txtSPDate.text <> "" Then
            If TextBoxFormatDate(txtSPDate) Then
               If lblPayPostingDate.ToolTipText = "" And lblPayPostingDate.ToolTipText <> txtSPDate.text Then
                  lblPayPostingDate.ToolTipText = txtSPDate.text
               End If
            End If
        End If
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
        If txtClientIDPurPay.text = "" Then
            ShowMsgInTaskBar "Please select a Client.", "Y", "N"
            'cboClient.SetFocus
            Exit Sub
        End If
        If frmMMain.IsRibbonVersion Then
        Dim adoConn As New ADODB.Connection
        Dim szSQL As String
        adoConn.Open getConnectionString
        If Trim(txtClientIDPurPay.text) = "" Then
            txtClientIDPurPay.SetFocus
            ShowMsgInTaskBar "Please select a Client.", "Y", "N"
            Exit Sub
        End If
        If IsDate(txtSPDate.text) = False Then Exit Sub
        If IsPeriodStatus(txtSPDate.text, txtClientIDPurPay.text, adoConn) = 0 Then
           ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtSPDate.text, txtClientIDPurPay.text, adoConn) = 9 Then
           ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
    End If
        'End of modification
End Sub

Private Sub txtProperty_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 47 Then cmdTypeList_Click
'   If KeyAscii = vbKeyTab Then
'        FocusControl cmdTypeList
'   End If
   If KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSPReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtBankAc.text <> "" Then
        txtSPaymentTotal.SetFocus
    End If
End Sub

Private Sub txtSupplierSearc_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdOpenSupp.SetFocus
    End If
End Sub



Private Sub txtSupSearchHis_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdOpSupSearch.SetFocus
    End If
End Sub



Private Sub txtUnit_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then cmdUnitList_Click
End Sub

Private Sub txtNC_Change(Index As Integer)
'  This change event is required when a supplier has a default NC.
'  System put the value of NC in the text box then this change event
'  look for the NC Name for the name text box.
   Dim i As Integer
   
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   txtNCName.text = ""
   LoadNominalCode ""
   For i = 1 To flxSupplier(0).Rows - 1
      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtNC(0).text))) = UCase(txtNC(0).text) Then
         txtNCName.text = flxSupplier(0).TextMatrix(i, 1)
         Exit For
      End If
   Next i
   If Len(txtNC(0).text) > 0 And txtNC(0).text <> "0" Then
        
        cmdUpdate(1).Enabled = True
   End If
End Sub

Private Sub txtVat__GotFocus(Index As Integer)
    txtVat_(Index).SelStart = 0
    txtVat_(Index).SelLength = Len(txtVat_(Index).text)
End Sub

Private Sub txtVat__KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtVat__LostFocus(Index As Integer)
   txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
   txtTotal.text = Format(txtTotal.text, "0.00")
   txtVat_(0).text = Format(Val(txtVat_(0).text), "0.00")
End Sub

Private Sub LoadSupplier()
   Dim iRow As Integer
   iRow = 1

   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 700

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(tabPayment.Tab).Width = 700
   lblSearch0(tabPayment.Tab).Left = 50
   
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(tabPayment.Tab).Left + flxSupplier(0).ColWidth(0)
   lblSearch2.Width = 400
   

   txtSearch1.Width = 800
   txtSearch1.Left = 50

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   Dim rdoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL1, szSQL2, szSQL3 As String

   ' Error Handler
   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   rdoConn.Open getConnectionString

   szSQL1 = "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientID;"
           
   szSQL2 = "SELECT AgentID, AgentName " & _
           "FROM Agent " & _
           "ORDER BY AgentID;"

   szSQL3 = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' " & _
           "ORDER BY SupplierID;"
     
   rstRst.Open szSQL1, rdoConn, adOpenStatic, adLockReadOnly

   If rstRst.EOF Then GoTo NoRes
   
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).RowHeight(0) = 0
    
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(tabPayment.Tab).Caption = "Cod"
   lblSearch1.Caption = "Name"
   lblSearch2.Caption = "Type"
      
   While Not rstRst.EOF
      flxSupplier(0).TextMatrix(iRow, 0) = rstRst!ClientID
      flxSupplier(0).TextMatrix(iRow, 1) = rstRst!ClientName
      flxSupplier(0).TextMatrix(iRow, 2) = "Client"
      rstRst.MoveNext
      flxSupplier(0).AddItem ""
      iRow = iRow + 1
   Wend
      
   Set rstRst = Nothing
   rstRst.Open szSQL2, rdoConn, adOpenStatic, adLockReadOnly

   While Not rstRst.EOF
      flxSupplier(0).TextMatrix(iRow, 0) = rstRst!AgentID
      flxSupplier(0).TextMatrix(iRow, 1) = rstRst!AgentName
      flxSupplier(0).TextMatrix(iRow, 2) = "Agent"
      rstRst.MoveNext
      flxSupplier(0).AddItem ""
      iRow = iRow + 1
   Wend
   Set rstRst = Nothing
   rstRst.Open szSQL3, rdoConn, adOpenStatic, adLockReadOnly

   While Not rstRst.EOF
      flxSupplier(0).TextMatrix(iRow, 0) = rstRst!SupplierID
      flxSupplier(0).TextMatrix(iRow, 1) = rstRst!SupplierName
      flxSupplier(0).TextMatrix(iRow, 2) = "Supplier"
      rstRst.MoveNext
      If Not rstRst.EOF Then flxSupplier(0).AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   rstRst.Close
   rdoConn.Close
   Set rstRst = Nothing
   Set rdoConn = Nothing
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
   
   rstRst.Close
   rdoConn.Close
   Set rstRst = Nothing
   Set rdoConn = Nothing

End Sub

Private Sub SuppInCombo(adoConn As ADODB.Connection, cboC As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** SUPPLIER COMBO ******************************************
   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' " & _
           "ORDER BY SupplierName;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Suppliers"
   For i = 1 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboC.Column() = Data()
   cboC.ListIndex = 0

NoRes:
   adoRST.Close
   Set adoRST = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Sub

'Resolved by BOSL
'Issue No: 0000467
'Added By: Asif. 04 Sep 2014

Private Sub LoadPropertyDropDown(cboP As Control)

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler
   
   adoConn.Open getConnectionString
   
   cboP.Clear
'*************************************** PROPERTY ******************************************
   If txtIDClient.text <> "ALL" Then
   
        szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE Property.ClientID = '" & txtIDClient.text & "' " & _
           "ORDER BY PropertyID;"
   Else
    
        szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"
           
   End If
   
'   Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   
   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

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

   cboP.Column() = Data()
   
   'Resolved by BOSL
   'Issue No: 0000467
   'Added By: Asif. 04 Sep 2014
   If cboP.ListCount > 0 Then
        cboP.ListIndex = 0
   End If
   
NoRes:
   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

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
   cboC.Column() = Data()
   cboC.ListIndex = 0
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

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
   
   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRST.Close
   Set adoRST = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Sub

'Private Sub PrepareCboClient(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
''Debug.Print szSQL
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
'   Dim Data() As String
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
'   cboClient.Column() = Data()
'   cboClient.ListIndex = -1
'   adoRst.Close
'
'NoRes:
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

Private Sub cmdDelete_Click()
   Dim strSQL As String
   Dim My_ID As String
   If flxPI.TextMatrix(1, 1) = "" Then Exit Sub

   If iSelected = 0 Then
      ShowMsgInTaskBar "Please select a record from the grid", , "N"
      Exit Sub
   End If

   If flxPI.row = 0 Then
        ShowMsgInTaskBar "Select a record in the grid to delete.", "Y", "N"
        Exit Sub
   End If
   If flxPI.Rows = iCurEditRow Then
         iCurEditRow = iCurEditRow - 1
   End If
   If MsgBox("Do you wish to delete line " & flxPI.TextMatrix(iCurEditRow, 0) & "?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub

   Dim iRow    As Integer
   Dim iCol    As Integer
   Dim iGrids  As Integer
' then there is only one row you need to code for that as well.
   
   If cmdEdit(1).Enabled = False Then 'this is the left pane edit button. If this button is disabled that means invoice in edit mode
       'Insert zero values row in collection, later this will be saved during save button
'        Call PurchaseInvoice.Add(fraLay(1).Tag, flxPI.TextMatrix(iCurEditRow, 7), IIf(CCur(flxPI.TextMatrix(iCurEditRow, 14)) > 0, 1, 0), flxPI.TextMatrix(iCurEditRow, 23), _
'                flxPI.TextMatrix(iCurEditRow, 21), flxPI.TextMatrix(iCurEditRow, 8), flxPI.TextMatrix(iCurEditRow, 11), PurchaseInvoice.Count + 1)
       
   End If
   If flxPI.Rows = 2 And flxPI.row = 1 Then
      ConfigFlxPI
   End If
   If flxPI.Rows > 2 Then
   'Below loop is shifting a line information to upper level
            For iRow = iCurEditRow To flxPI.Rows - 2
               For iCol = 1 To flxPI.Cols - 1
                  flxPI.TextMatrix(iRow, iCol) = flxPI.TextMatrix(iRow + 1, iCol)
               Next iCol
      Next iRow

      flxPI.RemoveItem flxPI.Rows - 1
     
      ' by anol 2018/02/01 issue 520
      'Insert into NLposting only when the invoice is in edit mode
      
      
      
      ''garbage
''      'search in the NLposting for this row
''      adoconn.Execute "UPDATE NLPosting AS N " & _
''                      "SET    N.DeleteFlag = TRUE " & _
''                      "WHERE  N.TRANSACTION_TYPE = " & IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7) & " AND " & _
''                             "N.TRANS_ID = '" & lSlNumber & "';"
''      strSQL = "Select * from  NLPosting AS N " & _
''      "WHERE  N.TRANSACTION_TYPE = " & IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7) & " AND " & _
''                             "N.TRANS_ID = '" & lSlNumber & "' AND N.DeleteFlag = FALSE AND this row identification;"
''      'if this item exists in the NLPOSTING
      'Copy those 2 lines and make their amount 0 and mark as deleted.
      ''Newidea Just insert two rows with zero amount and the description the row has
      
      
   End If
   UpdateTotalPICN
   'Added by anol 13 Aug 2015
    flxPI.Tag = "EditedOrAdded"
    cmdSavePI.Enabled = True
End Sub
Private Sub InsertZeroValues(adoConn As ADODB.Connection, SlNumber As String, Nominal_code As String, HaveVat As Integer, PARENT_RECORD As String, UNIT_ID As String, _
        TRANSACTION_DESCRIPTION As String, FUND_ID As String)
   Dim szSQL As String
   'Dim adoConn As New ADODB.Connection
   Dim adoSrc As New ADODB.Recordset
   Dim adoDst As New ADODB.Recordset
   Dim adoRST As New ADODB.Recordset

   Dim PurchaseLedgerControl As String
   Dim InputVAT As String
   
   PurchaseLedgerControl = ""
   InputVAT = ""
   
   'On Error GoTo Errhndler
    'adoConn.Open getConnectionString
'**************===============================***********************
'##############    Nominal Ledger Positing    #######################
'**************===============================***********************

   

'   adoConn.Execute "DELETE * FROM NLPosting;"
   szSQL = "SELECT * FROM NLPosting;"
   adoDst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

    With adoDst
      
         'Resolved by BOSL
         'Issue No: 0000476
         'Check if the Purchase Ledger Control and Input VAT Account exist in the system.
         'If not found, then exit the function before it updates or post any transsaction to the database.
         'Would prefer to validate this at the top before start posting the transaction, the nominal code for
         'Purchase Ledger Control and Input VAT Account may vary depending on the clients as thats how its been set.'
         'Modified By: Asif. 20 Sep 2014

         InputVAT = GetNominalCodeForControlAccount(adoConn, "Input VAT", txtClientID.text)
         If (InputVAT = "") Then
            GoTo Errhndler
         End If
         
         PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn, "Purchase Ledger Control", txtClientID.text)
         If (PurchaseLedgerControl = "") Then
            GoTo Errhndler
         End If
         
         '''''''''''''''''''''''''''''''''''''''''''''''
'                    AMOUNT TRANSACTION
         .AddNew
         .Fields.Item("THIS_RECORD").Value = UniqueID()
         .Fields.Item("PARENT_RECORD").Value = PARENT_RECORD
         .Fields.Item("TRANS_ID").Value = CInt(SlNumber)
         .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "DD MMMM YYYY")
         .Fields.Item("TRANSACTION_DATE").Value = Format(txtDate.text, "DD MMMM YYYY")
         .Fields.Item("TRANSACTION_TYPE").Value = IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7)
         .Fields.Item("ACCOUNT_NUMBER").Value = txtSupplierID.text 'this is the supplier code nad it is set after you click flxpurchase. No need to get it from grid
         .Fields.Item("PROPERTY_ID").Value = txtProperty.text
         .Fields.Item("UNIT_ID").Value = UNIT_ID 'Unit ID this is a detail Item
         .Fields.Item("FUND_ID").Value = FUND_ID 'Fund ID this is a detail Item
         'Exit Sub
         .Fields.Item("Amount").Value = 0
         
         .Fields.Item("REFERENCE").Value = txtReference.text
         .Fields.Item("NOMINAL_CODE").Value = Nominal_code
         .Fields.Item("TRANSACTION_DESCRIPTION").Value = TRANSACTION_DESCRIPTION
         .Fields.Item("AMOUNT_TYPE").Value = "A"
         .Fields.Item("USER_NUMBER").Value = "F"
         .Fields.Item("ClientID").Value = txtClientID.text
         
         'Nominal Ledger Issue 0000476. By BOSL Asif 18 Oct 2014
         .Fields.Item("Amount").Value = ConvertAmountToDRCR(0, IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7), .Fields.Item("AMOUNT_TYPE").Value)
         
         'Nominal Ledger Issue 0000476. By BOSL Asif 19 Nov 2014
         .Fields.Item("TRANSACTION_REF").Value = CInt(SlNumber)
         .Fields.Item("DeleteFlag").Value = True
         .Update

'                    VAT TRANSACTION
         If HaveVat Then
            .AddNew
            .Fields.Item("THIS_RECORD").Value = UniqueID()
            .Fields.Item("PARENT_RECORD").Value = PARENT_RECORD
            .Fields.Item("TRANS_ID").Value = CInt(SlNumber)
            .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "DD MMMM YYYY")
            .Fields.Item("TRANSACTION_DATE").Value = Format(txtDate.text, "DD MMMM YYYY")
            .Fields.Item("TRANSACTION_TYPE").Value = IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7)
            .Fields.Item("ACCOUNT_NUMBER").Value = txtSupplierID.text
            .Fields.Item("PROPERTY_ID").Value = txtProperty.text
            .Fields.Item("UNIT_ID").Value = UNIT_ID 'Unit ID this is a detail Item
            .Fields.Item("FUND_ID").Value = FUND_ID 'Fund ID this is a detail Item
            .Fields.Item("Amount").Value = 0
            .Fields.Item("REFERENCE").Value = txtReference.text
            
            'Resolved by BOSL
            'Issue No: 0000476
            'Assigning the Nominal Code for the Input VAT Account.
            'Modified By: Asif. 20 Sep 2014
            
            '.Fields.Item("NOMINAL_CODE").Value = "2201"
            
            .Fields.Item("NOMINAL_CODE").Value = InputVAT
            
            ''''''''''''''''''''''''''''''''
            
            .Fields.Item("TRANSACTION_DESCRIPTION").Value = TRANSACTION_DESCRIPTION
            .Fields.Item("AMOUNT_TYPE").Value = "V"
            .Fields.Item("USER_NUMBER").Value = "G"
            .Fields.Item("ClientID").Value = txtClientID.text
            
            'Nominal Ledger Issue 0000476. By BOSL Asif 18 Oct 2014
            .Fields.Item("Amount").Value = ConvertAmountToDRCR(0, .Fields.Item("TRANSACTION_TYPE").Value, .Fields.Item("AMOUNT_TYPE").Value)
            
            'Nominal Ledger Issue 0000476. By BOSL Asif 19 Nov 2014
            .Fields.Item("TRANSACTION_REF").Value = CInt(SlNumber)
            .Fields.Item("DeleteFlag").Value = True
            .Update
         End If

'                    TOTAL TRANSACTION - CREDITOR ACCOUNT
         .AddNew
         .Fields.Item("THIS_RECORD").Value = UniqueID()
         .Fields.Item("PARENT_RECORD").Value = PARENT_RECORD
         .Fields.Item("TRANS_ID").Value = CInt(SlNumber)
         
         
         .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "DD MMMM YYYY")
         .Fields.Item("TRANSACTION_DATE").Value = Format(txtDate.text, "DD MMMM YYYY")
         .Fields.Item("TRANSACTION_TYPE").Value = IIf(InStr(txtTransType.text, "Invoice") > 0, 6, 7)
         .Fields.Item("ACCOUNT_NUMBER").Value = txtSupplierID.text
         .Fields.Item("PROPERTY_ID").Value = txtProperty.text
         .Fields.Item("UNIT_ID").Value = UNIT_ID 'Unit ID this is a detail Item
         .Fields.Item("FUND_ID").Value = FUND_ID 'Fund ID this is a detail Item
         
         .Fields.Item("Amount").Value = 0
         .Fields.Item("REFERENCE").Value = txtReference.text
         
         'Resolved by BOSL
         'Issue No: 0000476
         'Assigning the Nominal Code for the Purchase Ledger Control Account.
         'Modified By: Asif. 20 Sep 2014
         
         '.Fields.Item("NOMINAL_CODE").Value = "2100"
         
         .Fields.Item("NOMINAL_CODE").Value = PurchaseLedgerControl
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
         .Fields.Item("TRANSACTION_DESCRIPTION").Value = TRANSACTION_DESCRIPTION
         .Fields.Item("AMOUNT_TYPE").Value = "C"
         .Fields.Item("USER_NUMBER").Value = "H"
         .Fields.Item("ClientID").Value = txtClientID.text
         
         'Nominal Ledger Issue 0000476. By BOSL Asif 18 Oct 2014
         .Fields.Item("Amount").Value = ConvertAmountToDRCR(0, .Fields.Item("TRANSACTION_TYPE").Value, .Fields.Item("AMOUNT_TYPE").Value)
         
         'Nominal Ledger Issue 0000476. By BOSL Asif 19 Nov 2014
         .Fields.Item("TRANSACTION_REF").Value = CInt(SlNumber)
         .Fields.Item("DeleteFlag").Value = True
         .Update

         'adoSrc.MoveNext
      End With
   

   adoDst.Close
   'adoConn.Close
   Exit Sub
Errhndler:
End Sub
Public Sub LoadFlxPurchaseFilter(adoConn As ADODB.Connection, Filter As String)
   Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWhere As String
   Dim strWhereClient As String
   Dim strWhereProperty As String
   Dim tempstr As String
   ConfigFlxPurchase
   ConfigFlxPurchaseSplit
    If Filter = "3" Then
         If txtSearchFromD.text <> "" Then
           strWhere = " AND PI.PostingDate =#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "#"
            If Len(txtSearchFromD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
    If txtPropID.text <> "ALL" Then
        strWhereProperty = " AND PI.PropertyID='" & txtPropID.text & "'"
    End If
    If txtIDClient.text <> "ALL" Then
         strWhereClient = " AND PI.CL_ID='" & txtIDClient.text & "'"
    End If
    If Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND PI.PostingDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND PI.PostingDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#"
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
   szSQL = "SELECT DISTINCT PI.MY_ID, (MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)& PI.SlNumber) AS INVNO, PI.SlNumber,PI.TransactionType, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID,PI.isManagementFee, " & _
               "Pt.OSAmount, QQ.PO, QQ.PO_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE  " & _
               "tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION,Supplier.Type,Pt.TransactionID,Pt.UserSessionID,Pt.Module,Pt.WindowsUserName,Pt.MachineName " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "LEFT JOIN (" & _
                  "SELECT Q2.MY_ID, Q1.SLNumber AS PO, Q2.PO AS PO_ID " & _
                  "FROM ( " & _
                  "SELECT MY_ID, SLNumber " & _
                  "From tblPurInv " & _
                  "WHERE TransactionType = 25) AS Q1 INNER JOIN " & _
                  "(SELECT MY_ID, PO " & _
                  "From tblPurInv " & _
                  "WHERE PO <> '') AS Q2 ON Q1.MY_ID = Q2.PO " & _
               ") AS QQ ON PI.MY_ID = QQ.MY_ID " & _
           "Where PI.History = False " & strWhere & strWhereProperty & strWhereClient & " AND (PI.TransactionType = 6 OR " & _
               "PI.TransactionType = 7) " & _
           "ORDER BY 3 Desc, 2;"
'Debug.Print szSQL
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'adoInv.Close
'Exit Sub
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoInv.Filter = "INVNO Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoInv.Filter = "SupplierName Like '%" & tempstr & "%'"
        End If
    End If

    If Not adoInv.EOF Then
        fmeLoading.Visible = True
        fmeLoading.Refresh
    End If
   iKount = 1
   colTransactionIDOtherPIGrid = ""
   With flxPurchase
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVNO").Value
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 13) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 16) = adoInv.Fields.Item("PostingDate").Value
         .TextMatrix(iKount, 17) = IIf(IsNull(adoInv.Fields.Item("PO").Value), "", adoInv.Fields.Item("PO").Value)
         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PO_ID").Value), "", adoInv.Fields.Item("PO_ID").Value)
         .TextMatrix(iKount, 19) = IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 20) = IIf(IsNull(adoInv.Fields.Item("Type").Value), "", adoInv.Fields.Item("Type").Value)
         .TextMatrix(iKount, 21) = IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value)
         
         
         .TextMatrix(iKount, 22) = IIf(IsNull(adoInv.Fields.Item("UserSessionID").Value), "", adoInv.Fields.Item("UserSessionID").Value)
         If .TextMatrix(iKount, 22) <> "" Then
            colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value) & ","
         End If
         
         .TextMatrix(iKount, 23) = IIf(IsNull(adoInv.Fields.Item("WindowsUserName").Value), "", adoInv.Fields.Item("WindowsUserName").Value)
         .TextMatrix(iKount, 24) = IIf(IsNull(adoInv.Fields.Item("MachineName").Value), "", adoInv.Fields.Item("MachineName").Value)
         .TextMatrix(iKount, 25) = IIf(IsNull(adoInv.Fields.Item("Module").Value), "", adoInv.Fields.Item("Module").Value)
         .TextMatrix(iKount, 26) = IIf(IsNull(adoInv.Fields.Item("isManagementFee").Value), False, adoInv.Fields.Item("isManagementFee").Value)

         If .TextMatrix(iKount, 22) <> "" Then
            .col = 1
            .row = iKount
            .CellBackColor = vbRed
         End If
         
         'issue 316 by anol 20170221
         If iKount = 10 Then
            frmPurchaseExpense.Refresh
            lblLoading.Caption = "Please wait while loading."
            flxPurchase.Refresh
         End If
         If iKount = 17 Then
             lblLoading.Caption = "Please wait while loading.."
             lblLoading.Refresh
            flxPurchase.Refresh
         End If

        .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
         adoInv.MoveNext
         iKount = iKount + 1
         If Not adoInv.EOF Then .AddItem ""
      Wend
      
   End With
   If Len(colTransactionIDOtherPIGrid) > 0 Then
            colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
   End If
XX:
   adoInv.Close
   Set adoInv = Nothing
End Sub
Public Sub LoadFlxPurchase(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWherePropertyId As String
   Dim strWhereClient As String
   ConfigFlxPurchase
   ConfigFlxPurchaseSplit
   If txtPropID.text <> "ALL" Then
        strWherePropertyId = " AND PI.PropertyID ='" & txtPropID.text & "'"
   End If
   If txtIDClient.text <> "ALL" Then
         strWhereClient = " AND PI.CL_ID='" & txtIDClient.text & "'"
   End If
   szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType,isManagementFee, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, " & _
               "Pt.OSAmount, PI.PO,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE  " & _
               "tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION,Supplier.Type,Pt.TransactionID,Pt.UserSessionID,Pt.WindowsUserName,Pt.MachineName ,Pt.Module  " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "Where PI.History = False " & strWherePropertyId & strWhereClient & " AND (PI.TransactionType = 6 OR " & _
               "PI.TransactionType = 7) " & _
           "ORDER BY 3 , 2 Desc;"
           ' I am fine tuning the SQL anol 20181118
'            szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType, " & _
'               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
'               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
'               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, " & _
'               "Pt.OSAmount, QQ.PO, QQ.PO_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION,SuppliEr.Type " & _
'           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
'               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
'               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
'               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
'               "LEFT JOIN (" & _
'                  "SELECT Q2.MY_ID, Q1.SLNumber AS PO, Q2.PO AS PO_ID " & _
'                  "FROM ( " & _
'                  "SELECT MY_ID, SLNumber " & _
'                  "From tblPurInv " & _
'                  "WHERE TransactionType = 25) AS Q1 INNER JOIN " & _
'                  "(SELECT MY_ID, PO " & _
'                  "From tblPurInv " & _
'                  "WHERE PO <> '') AS Q2 ON Q1.MY_ID = Q2.PO " & _
'               ") AS QQ ON PI.MY_ID = QQ.MY_ID " & _
'           "Where PI.History = False " & strWherePropertyId & strWhereClient & " AND (PI.TransactionType = 6 OR " & _
'               "PI.TransactionType = 7) " & _
'           "ORDER BY 3, 2;"
'Debug.Print time
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Debug.Print time
    If Not adoInv.EOF Then
        fmeLoading.Visible = True
        fmeLoading.Refresh
    End If
   iKount = 1
   colTransactionIDOtherPIGrid = ""
   With flxPurchase
      .Rows = adoInv.RecordCount + 1
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 13) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 16) = adoInv.Fields.Item("PostingDate").Value
         .TextMatrix(iKount, 17) = IIf(IsNull(adoInv.Fields.Item("PO").Value), "", adoInv.Fields.Item("PO").Value)
'         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PO_ID").Value), "", adoInv.Fields.Item("PO_ID").Value)'No use I can see anol 20181118
         .TextMatrix(iKount, 19) = IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 20) = IIf(IsNull(adoInv.Fields.Item("Type").Value), "", adoInv.Fields.Item("Type").Value)
         .TextMatrix(iKount, 21) = IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value)
         .TextMatrix(iKount, 22) = IIf(IsNull(adoInv.Fields.Item("UserSessionID").Value), "", adoInv.Fields.Item("UserSessionID").Value)
         .TextMatrix(iKount, 23) = IIf(IsNull(adoInv.Fields.Item("WindowsUserName").Value), "", adoInv.Fields.Item("WindowsUserName").Value)
         .TextMatrix(iKount, 24) = IIf(IsNull(adoInv.Fields.Item("MachineName").Value), "", adoInv.Fields.Item("MachineName").Value)
         .TextMatrix(iKount, 25) = IIf(IsNull(adoInv.Fields.Item("Module").Value), "", adoInv.Fields.Item("Module").Value)
         .TextMatrix(iKount, 26) = IIf(IsNull(adoInv.Fields.Item("isManagementFee").Value), "", adoInv.Fields.Item("isManagementFee").Value)
         If .TextMatrix(iKount, 22) <> "" Then
            .col = 1
            .row = iKount
            .CellBackColor = vbRed
            colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoInv.Fields.Item("TransactionID").Value), "", adoInv.Fields.Item("TransactionID").Value) & ","
         End If
         
         'Debug.Print .TextMatrix(iKount, 20)
         'issue 316 by anol 20170221
         If iKount = 10 Then
            'frmPurchaseExpense.Refresh
            lblLoading.Caption = "Please wait while loading...."
            flxPurchase.Refresh
         End If
         If iKount = 17 Then
             lblLoading.Caption = "Please wait while loading....."
             lblLoading.Refresh
            flxPurchase.Refresh
         End If
'         If iKount = 500 Then
'             lblLoading.Caption = "Please wait while loading..."
'             lblLoading.Refresh
'             flxPurchase.Refresh
''             GoTo XX
'         End If
'          If iKount = 1000 Then
'             lblLoading.Caption = "Please wait while loading...."
'             lblLoading.Refresh
'            flxPurchase.Refresh
'         End If
'          If iKount = 1500 Then
'             lblLoading.Caption = "Please wait while loading....."
'             lblLoading.Refresh
'            flxPurchase.Refresh
'         End If
'''######################################################################################################################
'''         Adding description of the header from the first split
''         szSQL = "SELECT DISTINCT * " & _
''                 "FROM tblPurInvSRec " & _
''                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
''                 "ORDER BY TRAN_ID;"
'''Debug.Print szSQL
''         adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
''
''         bFirstSp = True
''         If Not adoInvSp.EOF Then _
''            .TextMatrix(iKount, 8) = IIf(IsNull(adoInvSp.Fields.Item("DESCRIPTION").Value), "", adoInvSp.Fields.Item("DESCRIPTION").Value)
''
''         adoInvSp.Close
        .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
         adoInv.MoveNext
         iKount = iKount + 1
         'If Not adoInv.EOF Then .AddItem ""
      Wend
      'Debug.Print time
   End With
   If Len(colTransactionIDOtherPIGrid) > 0 Then
            colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
   End If
XX:
   adoInv.Close
   Set adoInv = Nothing
End Sub

Private Sub LoadFlxPurchPPHistory(adoConn As ADODB.Connection, Filter As String)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoInv As New ADODB.Recordset
   Dim strWhereTop As String
   Dim strWhere As String
   Dim strWhereSupplier As String
   Dim tempstr As String
   fmeLoading.Visible = True
   fmeLoading.Refresh
   ConfigflxPurchPPHistory flxPurchPPHistory, 50
   ConfigFlxSplit flxPurchPPHistorySplit, 39
   strWhereTop = ""
   If Filter = "" And Val(txtDisplayMaxPurchPayHist.text) > 0 Then
        strWhereTop = "Top " & Val(txtDisplayMaxPurchPayHist.text)
   End If
'   If Filter = "3" Then
'         If txtSearchFromD.text <> "" Then
'           strWhere = " AND P.TRAN_DATE =#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "#"
'            If Len(txtSearchFromD.text) > 0 Then
'                 cmdSearch.Caption = "Clear Sea&rch"
'            Else
'                 cmdSearch.Caption = "Sea&rch"
'            End If
'        End If
'    End If
   If txtSupSearchHis.text <> "ALL" Then
        strWhereSupplier = " AND P.SageAccountNumber='" & txtSupSearchHis.text & "'"
   End If
   If Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND P.PDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# "
'            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
'                 cmdSearchPurchPayHistory.Caption = "Clear Sea&rch"
'            Else
'                 cmdSearchPurchPayHistory.Caption = "Sea&rch"
'            End If
        End If
   End If
   'SWITCH(P.Type ='8',P.Amount,P.Type ='9',P.Amount,P.Type ='24',-P.Amount) AS TOTAL_AMOUNT
   'SWITCH(P.Type ='8',P.OSAmount,P.Type ='9',P.OSAmount,P.Type ='24',-P.OSAmount) as OSAmount
   If txtPurchasePaymentHistory.text = "ALL" Then
            szSQL = "SELECT " & strWhereTop & " P.TransactionID AS MY_ID, P.SlNumber, P.Type AS TransactionType, P.PDate AS TRAN_DATE, " & _
               "P.SageAccountNumber AS SUPP_AC, S.SupplierName, P.Amount AS TOTAL_AMOUNT, " & _
               "P.OSAmount as OSAmount,  " & _
               "P.ExtRef AS INV_NO, P.Details, T.DESCRIPTION, " & _
               "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)& P.SlNumber) AS INVNO,P.BankCode,P.ClientID,P.PayAmtType,P.POSTINGDATE,P.FundID " & _
           "FROM ((tlbPayment AS P INNER JOIN Supplier AS S ON P.SageAccountNumber = S.SupplierID) " & _
               "INNER JOIN tlbTransactionTypes AS T ON P.Type = T.TYPE_ID) " & _
           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) " & strWhere & strWhereSupplier & _
           "ORDER BY 1 Desc;"
    Else
        szSQL = "SELECT " & strWhereTop & " P.TransactionID AS MY_ID, P.SlNumber, P.Type AS TransactionType, P.PDate AS TRAN_DATE, " & _
               "P.SageAccountNumber AS SUPP_AC, S.SupplierName, P.Amount AS TOTAL_AMOUNT," & _
               "P.OSAmount as OSAmount,  " & _
               "P.ExtRef AS INV_NO, P.Details, T.DESCRIPTION, " & _
               "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)& P.SlNumber) AS INVNO,P.BankCode,P.ClientID,P.PayAmtType,P.POSTINGDATE,P.FundID  " & _
           "FROM ((tlbPayment AS P INNER JOIN Supplier AS S ON P.SageAccountNumber = S.SupplierID) " & _
               "INNER JOIN tlbTransactionTypes AS T ON P.Type = T.TYPE_ID) " & _
           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.ClientID='" & txtPurchasePaymentHistory.text & "'" & strWhere & strWhereSupplier & _
           "ORDER BY 1 Desc;"
    End If
'Debug.Print szSQL
'MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoInv.Filter = "INVNO Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoInv.Filter = "INV_NO Like '%" & tempstr & "%'"
        End If
    End If
    'Exit Sub
'   iKount = 1
'   Debug.Print time
'   With flxPurchPPHistory
'      While Not adoInv.EOF
''         Adding the header of the invoice
'         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
'         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVNO").Value 'adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
'         .TextMatrix(iKount, 3) = adoInv.Fields.Item("DESCRIPTION").Value
'         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
'         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
'         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
'         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value) 'this is reference
'         .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("Details").Value), "", adoInv.Fields.Item("Details").Value)
'         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
'         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("FundID").Value), "", adoInv.Fields.Item("FundID").Value)
'         .TextMatrix(iKount, 12) = IIf(IsNull(adoInv.Fields.Item("POSTINGDATE").Value), "", adoInv.Fields.Item("POSTINGDATE").Value)
'         .TextMatrix(iKount, 13) = IIf(IsNull(adoInv.Fields.Item("PayAmtType").Value), "", adoInv.Fields.Item("PayAmtType").Value)
'         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
'         .TextMatrix(iKount, 15) = adoInv.Fields.Item("BankCode").Value
'
'         adoInv.MoveNext
'         iKount = iKount + 1
'         If Not adoInv.EOF Then .AddItem ""
'      Wend
'   End With
'Debug.Print time
'
'        'If Not adoInv.EOF Then
'                adoInv.MoveFirst
'        'End If
'        flxPurchPPHistory.Clear
' for 3000 records .AddItem "" took 3 sec time . By excluding this I have save some loading time here
    With flxPurchPPHistory
      .Rows = adoInv.RecordCount + 1
      iKount = 1
      Dim dblAmount As Double
      Dim dblOSAmount As Double
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVNO").Value 'adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = adoInv.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value) 'this is reference
         .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("Details").Value), "", adoInv.Fields.Item("Details").Value)
         If adoInv.Fields.Item("TransactionType").Value = 24 Then
            .TextMatrix(iKount, 9) = Format(-adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         Else
            .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         End If
         dblAmount = dblAmount + .TextMatrix(iKount, 9)
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("FundID").Value), "", adoInv.Fields.Item("FundID").Value)
         .TextMatrix(iKount, 12) = IIf(IsNull(adoInv.Fields.Item("POSTINGDATE").Value), "", adoInv.Fields.Item("POSTINGDATE").Value)
         .TextMatrix(iKount, 13) = IIf(IsNull(adoInv.Fields.Item("PayAmtType").Value), "", adoInv.Fields.Item("PayAmtType").Value)
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = adoInv.Fields.Item("BankCode").Value
          If adoInv.Fields.Item("TransactionType").Value = 24 Then
                .TextMatrix(iKount, 16) = Format(-adoInv.Fields.Item("OSAMOUNT").Value, "0.00")
          Else
                .TextMatrix(iKount, 16) = Format(adoInv.Fields.Item("OSAMOUNT").Value, "0.00")
          End If
          dblOSAmount = dblOSAmount + .TextMatrix(iKount, 16)
         adoInv.MoveNext
         iKount = iKount + 1
         'If Not adoInv.EOF Then .AddItem ""
      Wend
   End With
   'Debug.Print time
   txtRctTotal.text = Format(dblAmount, "0.00")
   txtTotalOSAmount.text = Format(dblOSAmount, "0.00")
   adoInv.Close
   Set adoInv = Nothing
   fmeLoading.Visible = False
End Sub

Private Sub LoadFlxPurchHistory(adoConn As ADODB.Connection, Filter As String) 'Load purchase history
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset
   Dim strWhereProperty As String
   Dim strTopWhere As String
   Dim strWhere As String
   Dim strWhereSupplier As String
   Dim tempstr As String
   fmeLoading.Visible = True
   fmeLoading.Refresh
   ConfigFlxPurchHeader flxPurchHistory, 0
   ConfigFlxSplit flxPurchHistorySplit, 29
'   If Filter = "3" Then
'         If txtSearchFromD.text <> "" Then
'           strWhere = " AND PI.TRAN_DATE =#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "#"
'            If Len(txtSearchFromD.text) > 0 Then
'                 cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'            Else
'                 cmdSearchPurchaseHistory.Caption = "Sea&rch"
'            End If
'        End If
'    End If
    If txtSupplierSearc.text <> "ALL" Then
        strWhereSupplier = " AND Supplier.SupplierID='" & txtSupplierSearc.text & "'"
    End If
    If txtPropertyIDHist.text = "ALL" Then
    Else
        strWhereProperty = " AND PI.PropertyID='" & txtPropertyIDHist.text & "'"
    End If
    If Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND PI.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND PI.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#"
'            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
'                 cmdSearchPurchaseHistory.Caption = "Clear Sea&rch"
'            Else
'                 cmdSearchPurchaseHistory.Caption = "Sea&rch"
'            End If
        End If
   End If
   If Filter = "" And Val(txtDisplayMaxPurchaseHist.text) > 0 Then
        strTopWhere = " Top " & txtDisplayMaxPurchaseHist.text
   End If
'   If txtClientIdlist.text = "ALL" Then
'        szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
'                    "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
'                    "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
'                "FROM ((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
'                    "INNER JOIN tblPurInvSRec AS S ON PI.MY_ID = S.ParentID) " & _
'                    "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
'                "WHERE History = YES " & strWhere & " AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
'                "ORDER BY PI.MY_ID DESC;"
'   Else
'        szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
'                    "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
'                    "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & PI.SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
'                "FROM ((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
'                    "INNER JOIN tblPurInvSRec AS S ON PI.MY_ID = S.ParentID) " & _
'                    "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
'                "WHERE History = YES " & strWhere & " AND PI.CL_ID='" & txtClientIdlist.text & "' AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
'                "ORDER BY PI.MY_ID DESC;"
'   End If
    ' I am removing purchase split table that  iss needed for showing zero values in the invoice issue 520
    If txtClientIdlist.text = "ALL" Then
         szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
                     "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
                     "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
                 "FROM (tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
                     "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
                 "WHERE History = YES " & strWhereProperty & strWhere & strWhereSupplier & " AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
                 "ORDER BY PI.transactionType ASC,PI.slnumber DESC;"
    Else
         szSQL = "SELECT " & strTopWhere & " PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
                     "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
                     "(MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3)  & PI.SlNumber) as INVPur, PI.CL_ID,(SELECT top 1 DESCRIPTION FROM tblPurInvSRec WHERE tblPurInvSRec.ParentID = PI.MY_ID) as DESCRIPTION " & _
                 "FROM (tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
                     "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
                 "WHERE History = YES " & strWhereProperty & strWhere & strWhereSupplier & " AND PI.CL_ID='" & txtClientIdlist.text & "' AND PI.TransactionType <> 25 AND PI.TransactionType <> 26 " & _
                 "ORDER BY PI.transactionType ASC,PI.slnumber DESC;"
    End If
'   Debug.Print szSQL
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoInv.Filter = "INVPur Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoInv.Filter = "INV_NO Like '%" & tempstr & "%'"
        End If
    End If
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iKount = 1
   With flxPurchHistory
      .Rows = adoInv.RecordCount + 1
      While Not adoInv.EOF
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("INVPur").Value 'invoice number '''adoInv.Fields.Item("pf").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 6, "Invoice", "Credit Note")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value) ' this is reference
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("CL_ID").Value), "", adoInv.Fields.Item("CL_ID").Value)
'######################################################################################################################
''         Adding the split of the header
'         szSQL = "SELECT DISTINCT * " & _
'                 "FROM tblPurInvSRec " & _
'                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
'                 "ORDER BY TRAN_ID;"
''Debug.Print szSQL
'         adoInvSp.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'         If Not adoInvSp.EOF Then
            .TextMatrix(iKount, 8) = IIf(IsNull(adoInv.Fields.Item("DESCRIPTION").Value), "", adoInv.Fields.Item("DESCRIPTION").Value)
'         adoInvSp.Close

         adoInv.MoveNext
         iKount = iKount + 1
         'If Not adoInv.EOF Then .AddItem ""
      Wend
   End With

   adoInv.Close
   Set adoInv = Nothing
   fmeLoading.Visible = False
End Sub

Private Sub ConfigFlxPurchHeader(ctrHeader As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrHeader.Clear
   ctrHeader.Cols = 15
   ctrHeader.Rows = 2
   ctrHeader.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID"

   ctrHeader.FormatString = szHeader$
   ctrHeader.ColWidth(0) = 0
   ctrHeader.ColWidth(1) = Label20(1 + iLabel).Left - ctrHeader.Left
   For iCol = 2 To ctrHeader.Cols - 7
      ctrHeader.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrHeader.ColWidth(iCol) = ctrHeader.Width + ctrHeader.Left - Label20(iCol - 1 + iLabel).Left - 340
   ctrHeader.ColWidth(iCol + 1) = 0
   ctrHeader.ColWidth(iCol + 2) = 0
   ctrHeader.ColWidth(iCol + 3) = 0                   'OS Amt
   ctrHeader.ColWidth(iCol + 4) = 0                   'Due Date
   ctrHeader.ColWidth(iCol + 5) = 0                   'Client ID
End Sub
Private Sub ConfigflxPurchPPHistory(ctrHeader As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrHeader.Clear
   ctrHeader.Cols = 17 'i have increased it by 1 20180212 for purchase payment history bank code
   ctrHeader.Rows = 2
   ctrHeader.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID"

   ctrHeader.FormatString = szHeader$
   ctrHeader.ColWidth(0) = 0
   ctrHeader.ColWidth(1) = Label20(1 + iLabel).Left - ctrHeader.Left
   For iCol = 2 To 8
      ctrHeader.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrHeader.ColWidth(iCol) = 1300 'ctrHeader.Width + ctrHeader.Left - Label20(iCol - 1 + iLabel).Left - 340
   ctrHeader.ColWidth(iCol + 1) = 0
   ctrHeader.ColWidth(iCol + 2) = 0
   ctrHeader.ColWidth(iCol + 3) = 0                   'OS Amt
   ctrHeader.ColWidth(iCol + 4) = 0                   'Payment type
   ctrHeader.ColWidth(iCol + 5) = 0                   'Client ID
   ctrHeader.ColWidth(iCol + 6) = 0                   'col no 15 for Bank Code
   ctrHeader.ColWidth(iCol + 7) = 1300                    'col no 16 for Nominal Code
End Sub
Private Sub ConfigFlxPurchase()
   Dim szHeader As String, iCol As Integer

   flxPurchase.Clear
   'flxPurchase.Cols = 19
    ' I am adding 1 col . I shall use that for slnumber
   flxPurchase.Cols = 27
   flxPurchase.Rows = 2
   flxPurchase.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID|>Outstanding|PostingDate|PO|PO_ID"

   flxPurchase.FormatString = szHeader$
   flxPurchase.ColWidth(0) = 0
   flxPurchase.ColWidth(1) = Label20(10).Left - flxPurchase.Left - 10
   '19-10=9
   '20-10=10
   '20-11=9
   For iCol = 2 To 9
      flxPurchase.ColWidth(iCol) = Abs(Label20(iCol + 9).Left - Label20(iCol + 8).Left) - 50
   Next iCol
   
   'iCol = 10
   flxPurchase.ColWidth(iCol) = 0                        'Client
   flxPurchase.ColWidth(iCol + 1) = 0                    'Property
   flxPurchase.ColWidth(iCol + 2) = 0                    'OS Amt
   flxPurchase.ColWidth(iCol + 3) = 0                    'Due Date
   flxPurchase.ColWidth(iCol + 4) = 0                    'Client ID
   flxPurchase.ColWidth(iCol + 5) = flxPurchase.Width + flxPurchase.Left - Label20(18).Left - 340  'Outstanding
   flxPurchase.ColWidth(iCol + 6) = 0                    'Posting Date
   flxPurchase.ColWidth(iCol + 7) = 0                    'PO
   flxPurchase.ColWidth(iCol + 8) = 0                    'PO_ID
   flxPurchase.ColWidth(iCol + 9) = 0                    'slnumber flxPurchase.col(19) shall be used for slnumber
   flxPurchase.ColWidth(iCol + 10) = 0                   'Type col=20 type of supplier
   flxPurchase.ColWidth(iCol + 11) = 0                   ' col=21 for tlbpayment transaction ID
   flxPurchase.ColWidth(iCol + 12) = 0                   'col=22 for tlbpayment UserSessionID
   flxPurchase.ColWidth(iCol + 13) = 0                   'col=22 for tlbpayment WindowsUserName
   flxPurchase.ColWidth(iCol + 14) = 0                   'col=22 for tlbpayment MachineName
   flxPurchase.ColWidth(iCol + 15) = 0                   'col=22 for tlbpayment Module
   flxPurchase.ColWidth(26) = 0 'isManagementFee
   
End Sub

Private Sub ResizeFlxPurchase()
   Dim iCol As Integer

   flxPurchase.ColWidth(0) = 0
   flxPurchase.ColWidth(1) = Label20(10).Left - flxPurchase.Left
   For iCol = 2 To flxPurchase.Cols - 7
      flxPurchase.ColWidth(iCol) = Label20(iCol + 9).Left - Label20(iCol + 8).Left
   Next iCol
   flxPurchase.ColWidth(iCol) = 0                        'Client
   flxPurchase.ColWidth(iCol + 1) = 0                    'Property
   flxPurchase.ColWidth(iCol + 2) = 0                    'OS Amt
   flxPurchase.ColWidth(iCol + 3) = 0                    'Due Date
   flxPurchase.ColWidth(iCol + 4) = 0                    'Client ID
   flxPurchase.ColWidth(iCol + 5) = flxPurchase.Width + flxPurchase.Left - Label20(18).Left - 340  'Outstanding
End Sub

Private Sub ConfigFlxSplit(ctrSplit As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrSplit.Clear
   ctrSplit.Cols = 11
   ctrSplit.Rows = 2
   ctrSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"
   ctrSplit.FormatString = szHeader$

   ctrSplit.ColWidth(0) = 0
   ctrSplit.ColWidth(1) = Label20(1 + iLabel).Left - ctrSplit.Left

   For iCol = 2 To ctrSplit.Cols - 2
      ctrSplit.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrSplit.ColWidth(iCol) = ctrSplit.Width + ctrSplit.Left - Label20(iCol - 1 + iLabel).Left - 340
End Sub

Private Sub ConfigFlxPurchaseSplit()
   Dim szHeader As String, iCol As Integer
   Dim iLabel As Integer
   
   iLabel = 19

   flxPurchaseSplit.Clear
   flxPurchaseSplit.Cols = 13
   flxPurchaseSplit.Rows = 2
   flxPurchaseSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<ClientID|<Property ID|< Unit Name|<N/C" & _
               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount|>Recoverable"
   flxPurchaseSplit.FormatString = szHeader$

   flxPurchaseSplit.ColWidth(0) = 0
   'flxPurchaseSplit.ColWidth(1) = lblPurchaseSplit(1 + iLabel).Left - flxPurchaseSplit.Left

   For iCol = 1 To flxPurchaseSplit.Cols - 2
      flxPurchaseSplit.ColWidth(iCol) = lblPurchaseSplit(iCol + 19).Left - lblPurchaseSplit(iCol + 18).Left
   Next iCol
   flxPurchaseSplit.ColWidth(iCol) = flxPurchaseSplit.Width + flxPurchaseSplit.Left - lblPurchaseSplit(iCol - 1 + iLabel).Left - 340
End Sub

'  CheckSavedBtPay function returns TRUE if there any saved not generated Batch Payment found
'----------------------------------------------------------------------------------------------
Private Function CheckSavedBtPay(adoConn As ADODB.Connection) As Boolean
   On Error GoTo NoBP
   Dim adoBP As New ADODB.Recordset

   adoBP.Open "SELECT * " & _
              "FROM tblBatchTransaction " & _
              "WHERE SupplierID = '" & txtSPSupplier.Tag & "' AND " & _
              "BP IN (SELECT BP.BP FROM tblBatchPayment AS BP " & _
                        "WHERE BP.Generated = FALSE);", adoConn, adOpenStatic, adLockReadOnly

   If Not adoBP.EOF Then
      CheckSavedBtPay = True
   Else
      CheckSavedBtPay = False
   End If
   adoBP.Close
   Set adoBP = Nothing
   Exit Function
   
NoBP:
   CheckSavedBtPay = False
   Set adoBP = Nothing
End Function

Private Sub cmbSPSupplier_Click()
'   If cmbSPSupplier.text = "" Then Exit Sub
'
'   Frame5(5).Enabled = True
'   cboBC.ListIndex = -1
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   If Not CheckSavedBtPay(adoConn) Then
'      LoadFlxSPayment adoConn
'      LoadFlxSCrPoA adoConn
'   Else
'      MsgBox "This supplier is locked by Batch Payment." & Chr(13) & _
'             "Please clear down batch payment to book manual payment.", vbExclamation + vbOKOnly, "Batch Payment"
'      adoConn.Close
'      cmbSPSupplier.ListIndex = iSupplier
'      Set adoConn = Nothing
'      Exit Sub
'   End If
'
'   flxSCrPoA.row = 0
'   txtSupAcBal.text = flxSupplier(1).TextMatrix(flxSupplier(1).row, 3) ' Format(AccountBalance(adoConn), "0.00")
'   If Val(txtSupAcBal.text) >= 0 Then
'      txtSupAcBal.ForeColor = vbBlack
'   Else
'      txtSupAcBal.ForeColor = vbRed
'   End If
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   ReDim baChangesMade(flxSPayment.Rows) As Boolean
'   txtSPDate_GotFocus
End Sub
Private Sub ChangeSupplier()
'written by anol 20160524
   If txtSPSupplier.text = "" Then Exit Sub

   Frame5(5).Enabled = True
   'cboBC.ListIndex = -1

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not CheckSavedBtPay(adoConn) Then
      'Unlock previous locked item
      adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
      fmeLoading.Visible = True
      fmeLoading.Refresh
      LoadFlxSPayment adoConn
      DisplaylockScreen adoConn 'issue 749 done by anol 20190804
      LoadFlxSCrPoA adoConn
      fmeLoading.Visible = False
      
   Else
'      MsgBox "This supplier is locked by Batch Payment." & Chr(13) & _
'             "Please clear down batch payment to book manual payment.", vbExclamation + vbOKOnly, "Batch Payment"
        MsgBox "This supplier is locked by Batch Payments screen." & Chr(13) & _
             "Please clear down your batch payments before attempting to enter a manual purchase payment.", vbExclamation + vbOKOnly, "Batch Payment"
      adoConn.Close
      'cmbSPSupplier.ListIndex = iSupplier
      
      Set adoConn = Nothing
      Exit Sub
   End If

   flxSCrPoA.row = 0
   txtSupAcBal.text = flxSupplier(1).TextMatrix(flxSupplier(1).row, 3) ' Format(AccountBalance(adoConn), "0.00")
   If Val(txtSupAcBal.text) >= 0 Then
      txtSupAcBal.ForeColor = vbBlack
   Else
      txtSupAcBal.ForeColor = vbRed
   End If

   adoConn.Close
   Set adoConn = Nothing

   
   txtSPDate_GotFocus
End Sub
Private Sub cmdSPAmtType_Click()
   If txtSupplierType.text = "" Then Exit Sub
   If txtBankAc.text = "" Then Exit Sub
   
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString

   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

'   LoadRptAmtType "RECEIPT AMOUNT TYPE", adoConn, cmbSPAmtType
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

Private Sub txtSPaymentTotal_Change()
   txtPaymentTotal.text = txtSPaymentTotal.text
   If Val(txtSPaymentTotal.text) = 0 Then
      bChangesMade = False     'there some no change has made in the form
      cmdSPSave.Enabled = False
      Exit Sub
   Else
      bChangesMade = True     'there some changes have made in the form
      cmdSPSave.Enabled = True
   End If
   If Val(txtBankAc.text) > 0 Then cmdBankAc.Enabled = False
End Sub

Private Sub txtSPaymentTotal_GotFocus()
   If txtSPaymentTotal.text = "-" Then Exit Sub
   If txtBankAc.text = "" Then
      MsgBox "Please select the supplier's bank account", vbInformation, "Warning"
      cmdBankAc.SetFocus
      Exit Sub
   End If

   cTempPaymentAmt = CCur(txtSPaymentTotal.text)
   SelTxtInCtrl txtSPaymentTotal
End Sub

Private Sub txtSPaymentTotal_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtSPDate.SetFocus
'End If
'   If KeyAscii <> 45 Then
'      DigitTextKeyPress txtSPaymentTotal, KeyAscii
'   Else
'      txtSPaymentTotal.text = Format(Val(txtSPaymentTotal.text) * (-1), "0.00")
'      If Val(txtSPaymentTotal.text) < 0 Then KeyAscii = 0
'   End If
   
    If KeyAscii = 13 Then
        txtSPDate.SetFocus
    End If
   If KeyAscii <> 45 Then
      DigitTextKeyPress txtSPaymentTotal, KeyAscii
   Else
'      txtSPaymentTotal.text = Format(Val(txtSPaymentTotal.text) * (-1), "0.00")
'      If Val(txtSPaymentTotal.text) < 0 Then KeyAscii = 0
       txtSPaymentTotal.text = ""
   End If
   
End Sub

Private Sub txtSPaymentTotal_KeyUp(KeyCode As Integer, Shift As Integer)
   bTotalPayTyped = IIf(Val(txtSPaymentTotal.text) > 0, True, False)
End Sub

Private Sub flxSCrPoA_Click()
   sEditPPR = 2
'   If flxSCrPoA.Rows > 1 Then
'    '   'Added by anol
''   'issue 571 Validation
'        cmdEditPayment.Enabled = True
'     'End of modification
'
'   End If
   If Left(flxSCrPoA.TextMatrix(flxSCrPoA.row, 0), 2) = "PP" Or _
      Left(flxSCrPoA.TextMatrix(flxSCrPoA.row, 0), 2) = "PA" Then
      cmdEditPayment.Enabled = True
   Else
      cmdEditPayment.Enabled = False
   End If
   If flxSCrPoA.RowSel = 0 Then Exit Sub
   If flxSCrPoA.TextMatrix(1, 0) = "" Then Exit Sub
   If cmdPayAllocate.Caption = "All&ocation Only" Then Exit Sub

   iCrPoARowSel = IIf(flxSCrPoA.TextMatrix(flxSCrPoA.RowSel, 8) > 0, flxSCrPoA.RowSel, 0) 'Selected row of second grid

   Dim i As Integer, iFlxCrPoACol As Integer

   iFlxCrPoACol = 9
   flxSCrPoA.col = iFlxCrPoACol

   If flxSCrPoA.TextMatrix(flxSCrPoA.row, 2) = "" Then Exit Sub

   txtCrPayment.BackColor = vbWhite
   txtCrPayment.text = Format(flxSCrPoA.TextMatrix(flxSCrPoA.row, 7), "0.00")

   Label10(1).Caption = flxSCrPoA.row

   iCurRow = flxSCrPoA.row
   HighLightRowFlxGrid flxSCrPoA, iCurRow

   flxSCrPoA.TextMatrix(iCurRow, 9) = txtCrPayment.text
   If Val(txtCrPayment.text) > 0 Then
        cmdPayAutomatic.Enabled = True
   End If
   txtAllocatedDiff(1).text = txtCrPayment.text
'   Label10(2).Caption = flxSCrPoA.TextMatrix(iCurRow, 10)
   flxSCrPoA.Enabled = False
   Label10(5).Caption = flxSCrPoA.TextMatrix(iCurRow, 10) 'TransactionID for payment Line
   lblAllocating(1).Caption = "Allocating...                      " & flxSCrPoA.TextMatrix(iCrPoARowSel, 1)
   lblAllocating(1).Visible = True
   Frame5(5).Enabled = False                     'Payment - Saving
   Frame5(1).Enabled = True                      'Allocation - Saving

   flxSPayment.Enabled = True
   '   'Added by anol
'   'issue 571 Validation
   cmdEditPayment.Enabled = True
     'End of modification

End Sub

Private Sub txtCrPayment_GotFocus()
   SelTxtInCtrl txtCrPayment
   If Not lblAllocating(1).Visible Then
      iCurRow = flxSCrPoA.row
      HighLightRowFlxGrid flxSCrPoA, iCurRow
   Else
      iCurRow = flxSPayment.row
   End If
End Sub

Private Sub txtCrPayment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtAllocatedDiff(1).SetFocus
End Sub

Private Sub txtCrPayment_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtCrPayment, KeyAscii
End Sub

Private Sub txtCrPayment_LostFocus()
   txtCrPayment.text = Format(IIf(txtCrPayment.text = "", 0, txtCrPayment.text), "0.00")

   If Not lblAllocating(1).Visible Then
      If Val(flxSCrPoA.TextMatrix(iCurRow, 8)) < Val(txtCrPayment.text) Then
         ShowMsgInTaskBar "Payment amount exceeds amount outstanding.", , "N"
         txtCrPayment.text = "0.00"
         flxSCrPoA.RowSel = iCurRow
         flxSCrPoA.row = iCurRow
         txtCrPayment.SetFocus
         Exit Sub
      End If

      If Val(txtCrPayment.text) > 0 Then
         flxSCrPoA.TextMatrix(iCurRow, 9) = txtCrPayment.text
         txtAllocatedDiff(1).text = txtCrPayment.text
         Label10(5).Caption = flxSCrPoA.TextMatrix(iCurRow, 10)
         flxSCrPoA.Enabled = False

         lblAllocating(1).Caption = "Allocating...                      " & flxSCrPoA.TextMatrix(iCrPoARowSel, 1)
         lblAllocating(1).Visible = True

         Frame5(5).Enabled = False                     'Payment - Saving
         Frame5(1).Enabled = True                      'Allocation - Saving
      End If
   Else
'     Allocating in the invoice grid, $ txtCrPayment $ text box is in the Upper grid
'      If Val(txtCrPayment.text) = 0 Then Exit Sub

      If Val(txtAllocatedDiff(1).text) < Val(txtCrPayment.text) Then
         ShowMsgInTaskBar "Allocated amount cannot exceed allocation difference amount.", , "N"
         txtCrPayment.text = "0.00"
         txtCrPayment.SetFocus
         flxSPayment.row = Label10(7).Caption
         Exit Sub
      End If
      If Val(flxSPayment.TextMatrix(iCurRow, 9)) < Val(txtCrPayment.text) Then
         ShowMsgInTaskBar "Allocated amount exceeds amount outstanding.", , "N"
         txtCrPayment.text = "0.00"
         txtCrPayment.SetFocus
         flxSPayment.row = Label10(7).Caption
         Exit Sub
      End If

      If Val(txtCrPayment.text) > 0 Then
         flxSPayment.TextMatrix(iCurRow, 10) = txtCrPayment.text
         'txtAllocatedDiff(1).text = Format(Val(txtAllocatedDiff(1).text) - Val(txtCrPayment.text), "0.00")
         
          
         SumUpHeaderBySplits 'added by anol 20170515
         SpreadHeaderInSplits
            'added by anol 20170216 difference was not calculating correclty
         txtAllocatedDiff(1).text = Format(CalculateDiff, "0.00")
         If Val(txtAllocatedDiff(1).text) = 0 Then
            cmdPayAllocateSave.Enabled = True
            cmdPayAllocateSave.SetFocus
         Else
            cmdPayAllocateSave.Enabled = False
         End If

'         flxSPayment.TextMatrix(iCurRow, 15) = "A"
'         flxSPayment.TextMatrix(iCurRow, 16) = Label10(5).Caption
          If flxSPayment.TextMatrix(iCurRow, 0) <> "-" Then
            flxSPayment.TextMatrix(iCurRow, 15) = "A"
            flxSPayment.TextMatrix(iCurRow, 16) = Label10(5).Caption
         Else
            Dim iHeader As Integer

            iHeader = iCurRow
            Do
               iHeader = iHeader - 1
            Loop While (flxSPayment.TextMatrix(iHeader, 0) = "-")

            flxSPayment.TextMatrix(iHeader, 15) = "A"
            flxSPayment.TextMatrix(iHeader, 16) = Label10(5).Caption
         End If
      End If
      flxSPayment.Enabled = True
   End If

   txtCrPayment.Visible = False
End Sub

Private Function SpreadHeaderInSplits() As Integer       'Return the number of splits
   Dim i As Integer, cSumSplits As Currency, j As Integer

   If flxSPayment.TextMatrix(iCurRow, 0) = "+" Or flxSPayment.TextMatrix(iCurRow, 0) = ">" Then
      j = iCurRow + 1

      If Val(flxSPayment.TextMatrix(iCurRow, 10)) = Val(flxSPayment.TextMatrix(iCurRow, 9)) Then
         Do
            flxSPayment.TextMatrix(j, 10) = Format(flxSPayment.TextMatrix(j, 9), "0.00")
            If lblAllocating(1).Visible Then flxSPayment.TextMatrix(j, 15) = "A"
            baChangesMade(j) = True

            j = j + 1
            If j = flxSPayment.Rows Then Exit Do
         Loop While flxSPayment.TextMatrix(j, 0) = "-"
      Else
         cSumSplits = flxSPayment.TextMatrix(iCurRow, 10)

         Do
            flxSPayment.TextMatrix(j, 10) = IIf(cSumSplits > flxSPayment.TextMatrix(j, 9), flxSPayment.TextMatrix(j, 9), cSumSplits)
            flxSPayment.TextMatrix(j, 10) = Format(flxSPayment.TextMatrix(j, 10), "0.00")
            If lblAllocating(1).Visible Then flxSPayment.TextMatrix(j, 15) = "A"
            cSumSplits = cSumSplits - flxSPayment.TextMatrix(j, 10)

            baChangesMade(j) = True

            j = j + 1
            If j = flxSPayment.Rows Then Exit Do
         Loop While flxSPayment.TextMatrix(j, 0) = "-" 'And cSumSplits <> 0
      End If
   End If
   SpreadHeaderInSplits = j - iCurRow
   If flxSPayment.TextMatrix(iCurRow, 0) = "-" Then
      For i = iCurRow - 1 To 1 Step -1
         If flxSPayment.TextMatrix(i, 0) = "+" Or flxSPayment.TextMatrix(i, 0) = ">" Then Exit For
      Next i
      
      cSumSplits = 0
      For j = i + 1 To flxSPayment.Rows - 1
         If flxSPayment.TextMatrix(j, 0) = "+" Or _
            flxSPayment.TextMatrix(j, 0) = ">" Or _
            flxSPayment.TextMatrix(j, 0) = "" Then Exit For
         cSumSplits = cSumSplits + Val(flxSPayment.TextMatrix(j, 10))
      Next j
      
      flxSPayment.TextMatrix(i, 10) = Format(cSumSplits, "0.00")
      flxSPayment.TextMatrix(i, 15) = "A"
   End If
End Function

Public Sub LoadFlxSPayment(adoConn As ADODB.Connection)
   Dim adoRST     As New ADODB.Recordset
   Dim rdoSplits  As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer
   Dim szDataPath As String
   Dim iSpRow     As Integer
   Dim szWhProp   As String
   Dim cSplit     As Currency
   Dim cHeader    As Currency
   
   Dim cHeaderOsAmountPPR    As Currency
   Dim cSplitOsAmountPPR    As Currency
   Dim tempstr As String
   Dim colTransactionID As String
   Dim rsUserSessionID As String
   Dim rsPPRCheck As New ADODB.Recordset
'   On Error GoTo ErrHandler

   ConfigFlxSPayment


  szSQL = "SELECT Pt.TransactionID AS HeaderID,Pt.PI,  Pt.SlNumber, Pt.PI, Pt.AdjTag,I.isRentPayable,I.isManagementFee, " & _
                  "Pt.SageAccountNumber, Pt.UnitID, Pt.DDate, Pt.ExtRef AS Ref, Pt.Details, " & _
                  "Pt.Amount AS T_Amt, Pt.OSAmount AS T_OSAmt, Pt.Type, TT.DESCRIPTION, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, " & _
                  "Pt.ReconNow, Pt.PDate, Pt.PayAmtType, Pt.BankCode," & _
                  " iif(isnull(pt.CLientID ),I.CL_ID,pt.CLientID) as CLientID, Pt.UserSessionID,Pt.WindowsUserName,Pt.MachineName,Pt.Module, " & _
                  " (Select StatementID from RentSummaryStatementDetails  R where R.PINumber= cstr(pt.TransactionID) and I.transactionType=6 ) as CSID " & _
             "FROM (((tlbPayment AS Pt INNER JOIN tlbTransactionTypes AS TT ON Pt.Type = TT.TYPE_ID )) LEFT JOIN " & _
                  "Property AS P ON Pt.UnitID = P.PropertyID) LEFT JOIN tblPurInv AS I ON I.MY_ID = Pt.PI " & _
             "WHERE Pt.SageAccountNumber = '" & txtSPSupplier.Tag & "' And " & _
                   "Pt.PaymentView=True AND (TT.TYPE_ID=6 OR TT.TYPE_ID=24) AND Pt.OSAmount >0 " & _
             "ORDER BY Pt.TransactionID;"
             
             
             
   colTransactionIDOtherPayGrid = ""
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   'added this filer by anol 20160524
   tempstr = Replace(txtClientIDPurPay.text, "'", "''")
   adoRST.Filter = "ClientID='" & tempstr & "'"
   Dim rsPISplit As New ADODB.Recordset
    
    Dim invAmount As Double
    Dim invSplitAmount As Double

   iRow = 1
   While Not adoRST.EOF
     ' If adoRst.Fields.Item("T_Amt").Value <> adoRst.Fields.Item("Amount").Value Then 'comparing header amount with split amount to add split in grid
         flxSPayment.AddItem ""
         flxSPayment.TextMatrix(iRow, 0) = "+"
         flxSPayment.TextMatrix(iRow, 19) = adoRST!HeaderID
         flxSPayment.TextMatrix(iRow, 20) = IIf(IsNull(adoRST!ReconNow), "", adoRST!ReconNow)
         flxSPayment.TextMatrix(iRow, 21) = IIf(IsNull(adoRST!BankCode), "", adoRST!BankCode)
         flxSPayment.TextMatrix(iRow, 22) = IIf(IsNull(adoRST!PayAmtType), "", adoRST!PayAmtType)
         flxSPayment.TextMatrix(iRow, 23) = IIf(IsNull(adoRST!ClientID), "", adoRST!ClientID) 'IIf(IsNull(adoRst!CL_ID), "", adoRst!CL_ID)
         flxSPayment.TextMatrix(iRow, 1) = adoRST!PF & adoRST!SlNumber
         If InStr(adoRST!description, "Invoice") > 0 Then
            flxSPayment.TextMatrix(iRow, 2) = IIf(adoRST!AdjTag = "Y", "ADJI", adoRST!description)
         Else
            flxSPayment.TextMatrix(iRow, 2) = adoRST!description
         End If

         'flxSPayment.TextMatrix(iRow, 3) = adoRst!ChildID
         flxSPayment.TextMatrix(iRow, 4) = IIf(IsNull(adoRST!unitid), "", adoRST!unitid) 'Actually this is property ID
         If adoRST!Type = 24 Then
            flxSPayment.TextMatrix(iRow, 5) = IIf(Not IsNull(adoRST!PDate), Format(adoRST!PDate, "dd/mm/yyyy"), "")
         Else
            flxSPayment.TextMatrix(iRow, 5) = IIf(Not IsNull(adoRST!dDate), Format(adoRST!dDate, "dd/mm/yyyy"), "")
         End If
         flxSPayment.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!ref), "", adoRST!ref)
         flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
         flxSPayment.TextMatrix(iRow, 8) = Format(adoRST!T_Amt, "0.00")
         cHeader = adoRST!T_Amt
         cHeaderOsAmountPPR = adoRST!T_OSAmt
         flxSPayment.TextMatrix(iRow, 9) = Format(adoRST!T_OSAmt, "0.00")
         flxSPayment.TextMatrix(iRow, 10) = "0.00" 'Payment that you are inputting later when allocating
         flxSPayment.TextMatrix(iRow, 12) = IIf(IsNull(adoRST!Pi), "", adoRST!Pi)
         'flxSPayment.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!SpFund), "", adoRst!SpFund)
         flxSPayment.TextMatrix(iRow, 14) = adoRST!Type
         'insert the code for locking by anol 20190408 issue 749
         rsUserSessionID = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
         If Len(rsUserSessionID) > 0 And rsUserSessionID <> UserSessionID Then 'this means it is locked by other screen and now mark it red
               flxSPayment.col = 0
               flxSPayment.row = iRow
               flxSPayment.CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
               colTransactionIDOtherPayGrid = colTransactionIDOtherPayGrid & IIf(IsNull(adoRST!HeaderID), "", adoRST!HeaderID) & ","
         Else 'collect the transaction ID which needs to be locked
               colTransactionID = colTransactionID & IIf(IsNull(adoRST!HeaderID), "", adoRST!HeaderID) & ","
         End If
         flxSPayment.TextMatrix(iRow, 24) = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
         flxSPayment.TextMatrix(iRow, 25) = IIf(IsNull(adoRST!WindowsUserName), "", adoRST!WindowsUserName)
         flxSPayment.TextMatrix(iRow, 26) = IIf(IsNull(adoRST!MachineName), "", adoRST!MachineName)
         flxSPayment.TextMatrix(iRow, 27) = IIf(IsNull(adoRST!Module), "", adoRST!Module)
         flxSPayment.TextMatrix(iRow, 28) = IIf(IsNull(adoRST!Pi), "", adoRST!Pi)
'         flxSPayment.TextMatrix(iRow, 29) = IIf(IsNull(adoRst!splitID), "", adoRst!splitID)
         flxSPayment.TextMatrix(iRow, 30) = IIf(IsNull(adoRST!isRentPayable), False, adoRST!isRentPayable)
         flxSPayment.TextMatrix(iRow, 31) = IIf(IsNull(adoRST!isManagementFee), False, adoRST!isManagementFee)
         flxSPayment.TextMatrix(iRow, 32) = IIf(IsNull(adoRST!CSID), "", adoRST!CSID)
         
         'Add Splits
         cSplit = 0
         cSplitOsAmountPPR = 0
         rsPISplit.Open "Select * from tlbPaymentSplit where PayHeader =" & IIf(IsNull(adoRST!HeaderID), "", adoRST!HeaderID) & "", adoConn, adOpenStatic, adLockReadOnly
         While Not rsPISplit.EOF ' split amount less than header amount
            iRow = iRow + 1
            flxSPayment.AddItem ""
            flxSPayment.TextMatrix(iRow, 0) = "-"
            flxSPayment.TextMatrix(iRow, 1) = rsPISplit!splitID
            
            flxSPayment.TextMatrix(iRow, 3) = rsPISplit!TransactionID 'This is the transaction ID of split
            flxSPayment.TextMatrix(iRow, 4) = IIf(IsNull(adoRST!unitid), "", adoRST!unitid) 'Actually this is property ID
            flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRST!description), "", adoRST!description)
            flxSPayment.TextMatrix(iRow, 8) = Format(rsPISplit!amount, "0.00")
            cSplit = cSplit + rsPISplit!amount 'split amount
            flxSPayment.TextMatrix(iRow, 9) = Format(rsPISplit!OSAmount, "0.00")
            cSplitOsAmountPPR = cSplitOsAmountPPR + rsPISplit!OSAmount
            flxSPayment.TextMatrix(iRow, 10) = "0.00"
            flxSPayment.TextMatrix(iRow, 12) = IIf(IsNull(adoRST!Pi), "", adoRST!Pi)
            flxSPayment.TextMatrix(iRow, 13) = IIf(IsNull(rsPISplit!fundID), "", rsPISplit!fundID)
            flxSPayment.TextMatrix(iRow, 14) = adoRST!Type
            flxSPayment.TextMatrix(iRow, 19) = adoRST!HeaderID
            flxSPayment.TextMatrix(iRow, 28) = IIf(IsNull(adoRST!Pi), "", adoRST!Pi)
            flxSPayment.TextMatrix(iRow, 29) = IIf(IsNull(rsPISplit!splitID), "", rsPISplit!splitID)
            flxSPayment.TextMatrix(iRow, 32) = IIf(IsNull(adoRST!CSID), "", adoRST!CSID)
            flxSPayment.RowHeight(iRow) = 0
             'insert the code for locking by anol 20190408 issue 749
           

            rsPISplit.MoveNext
         Wend
         rsPISplit.Close
         If cHeader <> cSplit Then
            MsgBox "There is a mismatch in header amount and split amount '" & IIf((adoRST!Type = 24), "PPR", "PI") & adoRST!SlNumber & "'. You need to edit this transactiona and make it correct.", vbInformation, "Warning"
         End If
         If flxSPayment.TextMatrix(iRow, 14) = 24 Then 'For PPR check os amount is consistent else fix the problem
                    If cHeaderOsAmountPPR <> cSplitOsAmountPPR Then
                                 rsPPRCheck.Open "Select * from PayTransactions where FromTran=" & flxSPayment.TextMatrix(iRow, 19) & " and DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
                                 If rsPPRCheck.EOF Then
                                    rsPPRCheck.Close
                                    adoConn.Execute "Update tlbPaymentSplit set OsAmount=amount where PayHeader=" & flxSPayment.TextMatrix(iRow, 19) & ""
                                    flxSPayment.TextMatrix(iRow, 9) = flxSPayment.TextMatrix(iRow, 8)
                                 End If
                    End If
         End If
            rsUserSessionID = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
            If Len(rsUserSessionID) > 0 And rsUserSessionID <> UserSessionID Then 'this means it is locked by other screen and now mark it red
                  flxSPayment.col = 0
                  flxSPayment.row = iRow
                  flxSPayment.CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
                  colTransactionIDOtherPayGrid = colTransactionIDOtherPayGrid & IIf(IsNull(adoRST!HeaderID), "", adoRST!HeaderID) & ","
            Else 'collect the transaction ID which needs to be locked
                 colTransactionID = colTransactionID & IIf(IsNull(adoRST!HeaderID), "", adoRST!HeaderID) & ","
            End If
            
            adoRST.MoveNext
            iRow = iRow + 1
            'for the next PI do next loop
   Wend
   If Len(colTransactionIDOtherPayGrid) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherPayGrid = Left(colTransactionIDOtherPayGrid, Len(colTransactionIDOtherPayGrid) - 1)
   End If
   
      
   adoRST.Close
   Set adoRST = Nothing
   'added by anol 25 aug 2015
   'issue 571
   'Call filteringbyClient
   ReDim baChangesMade(flxSPayment.Rows) As Boolean
   If Len(colTransactionID) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionID = Left(colTransactionID, Len(colTransactionID) - 1)
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                   "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionID & ")"
    End If
   Exit Sub
ErrHandler:
   MsgBox "Purchase splits do not match with header. Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data Corruption"
   If adoRST.State = 1 Then
    adoRST.Close
    Set adoRST = Nothing
   End If
End Sub
Private Sub filteringbyClient()
    'added by anol 25 aug 2015
    Dim iRow As Integer
   'issue 571
  On Error GoTo Err
   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 23) <> txtClientIDPurPay.text And _
            flxSPayment.TextMatrix(iRow, 23) <> "" Then
         flxSPayment.RowHeight(iRow) = 0
         If flxSPayment.TextMatrix(iRow, 0) <> "" Then
            iRow = iRow + 1
            While flxSPayment.TextMatrix(iRow, 0) = "-"
               flxSPayment.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      Else
         flxSPayment.RowHeight(iRow) = 240
         If flxSPayment.TextMatrix(iRow, 0) = "+" Then
            iRow = iRow + 1
            While flxSPayment.TextMatrix(iRow, 0) = "-"
               flxSPayment.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      End If
   Next iRow
   Exit Sub
Err:
End Sub
Public Sub LoadFlxSCrPoA(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim SQLStr1 As String
   Dim iRow As Integer, szDataPath As String
   Dim colTransactionID As String
   Dim rsUserSessionID As String

   flxSCrPoA.Rows = 2
   ConfigFlxSCrPoA
'Modified by Anol 08 Sep 2014
'Issue 470
'Rsolved by BOSL
'   Get the details for the demand type selected
'"PYT.OSAmount > 0 "
'PYT.Amount > 0 AND
'cmmet out by anol 31 aug 2015
   SQLStr1 = "SELECT PYT.TransactionID, PYT.SageAccountNumber, " & _
                  "PYT.UnitID, PYT.PDate AS Dt, PYT.Ref, PYT.Details, " & _
                  "PYT.Amount, PYT.OSAmount as OS, PYT.PI as DR, " & _
                  "PYT.AdjTag, TT.DESCRIPTION, PYT.SlNumber, " & _
                  "TT.TYPE_ID, PYT.Type as TT, PYT.FundID, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF, " & _
                  "PYT.BankCode, PYT.ExtRef, PYT.PayAmtType,PYT.Reconnow,  iif(isnull(pyt.CLientID ),P.ClientID,pyt.CLientID) as CLientID,  " & _
                  "PYT.UserSessionID,PYT.WindowsUserName,PYT.MachineName,PYT.Module "
   SQLStr1 = SQLStr1 + _
             "FROM ((tlbPayment AS PYT INNER JOIN tlbTransactionTypes AS TT  ON PYT.Type = TT.TYPE_ID) LEFT JOIN Property AS P ON Pyt.UnitID = P.PropertyID) " & _
             "WHERE PYT.SageAccountNumber = '" & txtSPSupplier.Tag & "' And PYT.ClientID='" & txtClientIDPurPay.text & "' AND " & _
                   "PYT.PaymentView = True And PYT.Amount > 0 AND " & _
                   "(TT.TYPE_ID = 7 OR TT.TYPE_ID = 9 OR TT.TYPE_ID = 8) " & _
             "Order By TransactionID;"
'Debug.Print SQLStr1 'And PYT.ClientID='" & txtClientIDPurPay.text & "' added by anol 2019-05-22
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly
   cmdEditPayment.Enabled = False
   ReDim baBankRecon(adoRST.RecordCount) As Boolean
   iRow = 1
   While Not adoRST.EOF
      flxSCrPoA.TextMatrix(iRow, 10) = adoRST!TransactionID
      flxSCrPoA.TextMatrix(iRow, 0) = adoRST!PF & adoRST!SlNumber
      If InStr(adoRST!description, "Credit") > 0 Then
         flxSCrPoA.TextMatrix(iRow, 1) = IIf(adoRST!AdjTag = "Y", "ADJC", adoRST!description)
      Else
         flxSCrPoA.TextMatrix(iRow, 1) = adoRST!description
      End If
        flxSCrPoA.TextMatrix(iRow, 2) = adoRST!SageAccountNumber
        flxSCrPoA.TextMatrix(iRow, 3) = IIf(IsNull(adoRST!unitid), "", adoRST!unitid)
        flxSCrPoA.TextMatrix(iRow, 4) = Format(adoRST!dt, "dd/mm/yyyy")
        flxSCrPoA.TextMatrix(iRow, 5) = IIf(IsNull(adoRST!ExtRef), "", adoRST!ExtRef)
        flxSCrPoA.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!Details), "", adoRST!Details)
        flxSCrPoA.TextMatrix(iRow, 7) = Format(adoRST!amount, "0.00")
        flxSCrPoA.TextMatrix(iRow, 8) = Format(adoRST!OS, "0.00")
        flxSCrPoA.TextMatrix(iRow, 9) = "0.00"
        flxSCrPoA.TextMatrix(iRow, 11) = IIf(IsNull(adoRST!DR), "", adoRST!DR)
        flxSCrPoA.TextMatrix(iRow, 12) = IIf(Val(flxSCrPoA.TextMatrix(iRow, 12)) = -1, "P/ADJ/CR", Format(flxSCrPoA.TextMatrix(iRow, 12), "0.00"))
        flxSCrPoA.TextMatrix(iRow, 13) = adoRST!TYPE_ID
        flxSCrPoA.TextMatrix(iRow, 14) = IIf(IsNull(adoRST!fundID), "", adoRST!fundID)
        flxSCrPoA.TextMatrix(iRow, 15) = IIf(IsNull(adoRST!BankCode), "", adoRST!BankCode)
        flxSCrPoA.TextMatrix(iRow, 16) = IIf(IsNull(adoRST!ExtRef), "", adoRST!ExtRef)
        flxSCrPoA.TextMatrix(iRow, 17) = IIf(IsNull(adoRST!PayAmtType), "", adoRST!PayAmtType)
        flxSCrPoA.TextMatrix(iRow, 18) = IIf(IsNull(adoRST!ClientID), "", adoRST!ClientID)
        flxSCrPoA.TextMatrix(iRow, 19) = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
        flxSCrPoA.TextMatrix(iRow, 20) = IIf(IsNull(adoRST!WindowsUserName), "", adoRST!WindowsUserName)
        flxSCrPoA.TextMatrix(iRow, 21) = IIf(IsNull(adoRST!MachineName), "", adoRST!MachineName)
        flxSCrPoA.TextMatrix(iRow, 22) = IIf(IsNull(adoRST!Module), "", adoRST!Module)
        'insert the code for locking by anol 20190408 issue 749
        rsUserSessionID = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
        If Len(rsUserSessionID) > 0 And rsUserSessionID <> UserSessionID Then 'this means it is locked by other screen and now mark it red
            flxSCrPoA.col = 0
            flxSCrPoA.row = iRow
            flxSCrPoA.CellBackColor = RGB(255, 0, 0) ' 'Mark that as red so that user cannot process
            colTransactionIDOtherPayGrid = colTransactionIDOtherPayGrid & IIf(IsNull(adoRST!TransactionID), "", adoRST!TransactionID) & ","
        Else 'collect the transaction ID which needs to be locked
            colTransactionID = colTransactionID & IIf(IsNull(adoRST!TransactionID), "", adoRST!TransactionID) & ","
        End If
         
         
      'issue 470
      'added by anol 22 Sep 2014
      baBankRecon(iRow) = IsNull(adoRST.Fields.Item("ReconNow").Value)
      'end of modification
      adoRST.MoveNext
      If Not adoRST.EOF Then flxSCrPoA.AddItem ""
      iRow = iRow + 1
   Wend
    If Len(colTransactionIDOtherPayGrid) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionIDOtherPayGrid = Left(colTransactionIDOtherPayGrid, Len(colTransactionIDOtherPayGrid) - 1)
   End If
   If Len(colTransactionID) > 0 Then 'UserSessionID<>'" & UserSessionID & "' and
        colTransactionID = Left(colTransactionID, Len(colTransactionID) - 1)
        adoConn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                   "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where TransactionID in (" & colTransactionID & ")"
    End If
   adoRST.Close
   Set adoRST = Nothing
   Call filteringbyClientDebit
End Sub
Private Sub filteringbyClientDebit()
    'added by anol 25 aug 2015
    Dim iRow As Integer
   'issue 571
  On Error GoTo Err
   For iRow = 1 To flxSCrPoA.Rows - 1
      If flxSCrPoA.TextMatrix(iRow, 18) <> txtClientIDPurPay.text And _
            flxSCrPoA.TextMatrix(iRow, 18) <> "" Then
         flxSCrPoA.RowHeight(iRow) = 0
         If flxSCrPoA.TextMatrix(iRow, 0) <> "" Then
            iRow = iRow + 1
            While flxSCrPoA.TextMatrix(iRow, 0) = "-"
               flxSCrPoA.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      Else
         flxSCrPoA.RowHeight(iRow) = 240
         If flxSCrPoA.TextMatrix(iRow, 0) = "+" Then
            iRow = iRow + 1
            While flxSCrPoA.TextMatrix(iRow, 0) = "-"
               flxSCrPoA.RowHeight(iRow) = 0
               iRow = iRow + 1
            Wend
            iRow = iRow - 1
         End If
      End If
   Next iRow
   Exit Sub
Err:
End Sub
Private Function DateValidation() As Boolean
        Dim iRow As Integer
        For iRow = 1 To flxSPayment.Rows - 1
            If (flxSPayment.TextMatrix(iRow, 0) = "+" Or flxSPayment.TextMatrix(iRow, 0) = ">") And _
                    Val(flxSPayment.TextMatrix(iRow, 10)) <> 0 Then
                        If (flxSPayment.TextMatrix(iRow, 30) = True Or flxSPayment.TextMatrix(iRow, 31) = True) Then
                            If DateDiff("d", flxSPayment.TextMatrix(iRow, 5), flxSCrPoA.TextMatrix(Label10(1).Caption, 4)) <> 0 Then
                                    Exit Function
                            End If
                        End If
            End If
        Next iRow
        DateValidation = True
End Function
Private Sub cmdPayAllocateSave_Click()
   Dim adoConn       As New ADODB.Connection
   Dim szSQL As String
   Dim rsBankCheck As New ADODB.Recordset
   Dim iRow As Integer
   If Not DateValidation Then
        If MsgBox("The payment date entered is different from the due date of the selected invoice(s). Do you wish to continue? ", vbYesNo, "Please confirm") = vbYes Then
        Else
            Exit Sub
        End If
   End If
   adoConn.Open getConnectionString
   'Here is the logic if you are selecting wrong bank account in which differes from Rentpayable bank account.
   'We should check from the client statement , the Bank code selected for this Rent Payable an d compare with selected bank account for paying current PI
    'written by anol 20230728
    For iRow = 1 To flxSPayment.Rows - 1
        If flxSPayment.TextMatrix(iRow, 30) = "" Then 'isrentpayable flag is  empty then do not check anything
        Else
             ' Here the first grid is allocating with second grid
                If cmdPayAllocate.Caption <> "All&ocation Only" Then
                    If flxSPayment.TextMatrix(iRow, 30) = True And flxSPayment.TextMatrix(iRow, 14) = "6" Then
                          szSQL = "SELECT R.BankCode FROM tblpurinv AS T INNER JOIN rentsummarystatement AS R ON " & _
                          "(T.inv_no = 'SS'& R.StatementID) WHERE MY_ID='" & flxSPayment.TextMatrix(iRow, 12) & "'"
                          rsBankCheck.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                          If RecordCount(rsBankCheck) > 0 Then
                              If rsBankCheck("BankCode").Value <> flxSCrPoA.TextMatrix(iCrPoARowSel, 15) Then
                                  rsBankCheck.Close
                                  adoConn.Close
                                  MsgBox "Please select a bank account that matches the bank account on the client statement used to generate this rent payable invoice", vbInformation, "Warning"
                                  Exit Sub
                              End If
                          End If
                     End If
               End If
        End If
    Next
    adoConn.Close
   cmdPayAllocateSave.Enabled = False
   adoConn.Open getConnectionString
   adoConn.BeginTrans
   If chkSettleAll.Value = 0 Then
      If SavingAllocationS(adoConn) = True Then
            adoConn.CommitTrans
            MsgBox "Allocations have been saved successfully.", vbInformation, "Saved"
      Else
            adoConn.RollbackTrans
            MsgBox "An error occurred while Allocating, transaction rollbacked.", vbInformation, "Transaction rollbacked"
      End If
   Else
    '  right now I have disable this part settle all anol 20170214 new codes need to be implemnted
'      If SettleAllCrAllocation(adoConn) = True Then
'            adoConn.CommitTrans
'            ShowMsgInTaskBar "Allocations have been saved successfully."
'      Else
'            adoConn.RollbackTrans
'            MsgBox "An error occurred while Allocating,transaction rollbacked.", vbInformation, "Transaction rollbacked"
'      End If
   End If
   'Unlock  previous locked item
    adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
   cmdPayAllocateSave.Enabled = True
   LoadFlxSPayment adoConn
   LoadFlxSCrPoA adoConn
   adoConn.Close
   Set adoConn = Nothing
  
   txtPaymentEntered.text = "0.00"
   txtSPaymentTotal.text = "0.00"

   flxSCrPoA.Enabled = True
   flxSPayment.Enabled = True
   flxSCrPoA.Enabled = True
   'comment out 20170218 issue 306
'   cmdPayAllocateSave.Enabled = False
'   lblAllocating(1).Visible = False
'   Frame5(5).Enabled = True                     'Payment - Saving
'   Frame5(1).Enabled = False                    'Allocation - Saving
'   txtAllocatedDiff(1).text = "0.00"
'   ConfigFlxSPayment
'   cmdPayAutomatic.Enabled = True
'
'   Frame5(5).Visible = True
'   cmeRevereseAllocation.Visible = Frame5(5).Visible
'   Frame5(5).Enabled = True
'   Frame5(1).Visible = False
'   cmdPayAllocate.Caption = "All&ocation Only"
'   cmdAccounts.Enabled = True
'   Label3(5).Visible = False
'   txtAllocatedDiff(1).Visible = False
'   Label3(1).Visible = False
 'comment out 20170218 issue 306
   lblAllocating(1).Visible = False
   cmdPayAllocateSave.Enabled = False
End Sub

Private Function SettleAllCrAllocation(adoConn As ADODB.Connection) As Boolean
   Dim iRow    As Integer
   Dim szSQL   As String
   Dim i       As Integer
   Dim j       As Integer
   Dim lRT_ID  As Long
   Dim dAlloc  As Double                         'Total     Allocated  amount
   Dim dPA     As Double                            'Process   Allocation amount
   Dim dCA     As Double                            'Currently Allocating amount
   Dim dCAd    As Double                           'Currently Allocated  amount
   'Dim adoConn As New ADODB.Connection
   Dim rstRst  As New ADODB.Recordset

   dPA = MinAmtAlloc(flxSPayment, flxSCrPoA)

   'adoConn.Open getConnectionString

   j = 1

   While dPA > dAlloc
'  Update the credit side first.
'  First determin the amount to be allocated. Then allocate that amount in the Invoice side.
      If dPA >= flxSCrPoA.TextMatrix(j, 8) Then
         szSQL = "UPDATE tlbPayment " & _
                 "SET OSAmount = 0, " & _
                     "PaymentView = False " & _
                 "WHERE TransactionID = " & CLng(flxSCrPoA.TextMatrix(j, 10)) & ";"
         dAlloc = dAlloc + flxSCrPoA.TextMatrix(j, 8)
         flxSCrPoA.TextMatrix(j, 8) = Format(0, "0.00")
         dCA = flxSCrPoA.TextMatrix(j, 9)
      Else
         szSQL = "UPDATE tlbPayment " & _
                 "SET OSAmount = OSAmount - " & dPA & " " & _
                 "WHERE TransactionID = " & CLng(flxSCrPoA.TextMatrix(j, 10)) & ";"
         dAlloc = dAlloc + dPA
         dCA = dPA
         flxSCrPoA.TextMatrix(j, 8) = Format(flxSCrPoA.TextMatrix(j, 8) - dPA, "0.00")
      End If

      adoConn.Execute szSQL

'  dCA amount to be allocated in the invoice side.
      i = 0
      While dCA > 0
         If flxSPayment.TextMatrix(dSortIndex(i), 9) > 0 Then
            If dCA >= flxSPayment.TextMatrix(dSortIndex(i), 9) Then
                szSQL = "UPDATE tlbPayment " & _
                        "SET tlbPayment.OSAmount = 0, " & _
                           "PaymentView = False " & _
                        "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(dSortIndex(i), 19)) & ";"
               dCA = dCA - CDbl(flxSPayment.TextMatrix(dSortIndex(i), 9))
               dCAd = flxSPayment.TextMatrix(dSortIndex(i), 9)
               flxSPayment.TextMatrix(dSortIndex(i), 9) = Format(0, "0.00")
            Else
               szSQL = "UPDATE tlbPayment " & _
                       "SET tlbPayment.OSAmount = " & Val(flxSPayment.TextMatrix(dSortIndex(i), 9)) - dCA & " " & _
                       "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(dSortIndex(i), 19)) & ";"
               dCAd = dCA
               flxSPayment.TextMatrix(dSortIndex(i), 9) = Format(flxSPayment.TextMatrix(dSortIndex(i), 9) - dCA, "0.00")
               dCA = 0
            End If
            adoConn.Execute szSQL
   
            szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM PayTransactions;"
            rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            lRT_ID = CLng(IIf(IsNull(rstRst!TID), 1, rstRst!TID))
            rstRst.Close
   
   '  Now this allocation is needed to save in the PayTransactions table
            szSQL = "SELECT * FROM PayTransactions;"
            rstRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
            With rstRst
               .AddNew
               !TranType = IIf(UCase(flxSCrPoA.TextMatrix(j, 1)) = "ADJC", "ADJAL", "AL")
               !TransactionID = lRT_ID
               !Alloc_Unalloc = 1
               !FromTran = CLng(flxSCrPoA.TextMatrix(j, 10))
                'added by anol 2021-10-16 . Here the fix was for payment was not writing unitID that means propertyID in tlbpayment which was causing trouble in find
'               adoconn.Execute "Update tlbpaymentSplit  SET UNIT_ID='" & flxSPayment.TextMatrix(iRow, 4) _
'               & "',PayTransactionID=" & lRT_ID & " where transactionID='" & CLng(flxSCrPoA.TextMatrix(j, 10)) & "'  and SplitID =" & CLng(flxSCrPoA.TextMatrix(j, 1)) & ""
               !ToTran = CLng(flxSPayment.TextMatrix(dSortIndex(i), 19))
               !AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
               'Issue 437 Allocation date
               'we are not using this sub procedure, if we use it we neeed to use next line to pick up allocation date
               '!AllocDate = Format(flxSCrPoA.TextMatrix(iRow, 4), "DD MMMM YYYY")
               !PaymentAmount = CCur(dCAd)
               !Discount = 0
               !IsSageUpdate = IIf(UCase(flxSCrPoA.TextMatrix(j, 1)) = "ADJC", False, True)
               !UpdateSage = False
               !BankCode = txtBankCode.text
               !nominalCode = rstRst!BankCode
               !deleteFlag = False
               .Update
               .Close
            End With
         End If
         i = i + 1
      Wend
      j = j + 1
   Wend
   
   ConfigFlxSCrPoA
   ConfigFlxSPayment
   txtSPSupplier.text = ""
   txtSPSupplier.Tag = ""
   txtBankCode.text = ""
   txtBankAc.text = ""
   txtPaymentEntered.text = "0.00"
   txtSPaymentTotal.text = "0.00"

   flxSCrPoA.Enabled = True
   ReDim baChangesMade(flxSPayment.Rows) As Boolean

   'adoConn.Close
   Set rstRst = Nothing
   SettleAllCrAllocation = True
   'Set adoConn = Nothing
End Function

Private Function SavingAllocationS(adoConn As ADODB.Connection) As Boolean
   Dim iRow          As Integer
   Dim szSQL         As String ', iResponse As Integer
   Dim lRT_ID        As Long
   'Dim adoConn       As New ADODB.Connection
   Dim rstPayTrans   As New ADODB.Recordset
   Dim rstSplit      As New ADODB.Recordset
   Dim rsSSR As New ADODB.Recordset
   Dim rsInvoice As New ADODB.Recordset
   Dim szPurchaseTranID As String
   Dim cSumSplits As Double
   Dim rsPaytransaction As New ADODB.Recordset
   Dim lSPTran_ID As Long
   Dim rstSet As New ADODB.Recordset
   Dim lSp_ID As Long
   Dim splitID As Long
' I added FundID field to record FundId for splited PP.
' Later on it has been decided that PP should be booked fully against PI and PP's fund will not be changed.
' Therefore, I don’t need to save FUNDID in the mapping table.
' If user wants to change the fund of the PP then they should unallocate the PP and edit the PP to amend the fundid.
  


   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM PayTransactions;"
   rstPayTrans.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lRT_ID = CLng(IIf(IsNull(rstPayTrans!TID), 1, rstPayTrans!TID))
   rstPayTrans.Close
   
   
   

'  Update the credit Out-Standing amount by receipt amount
   If Val(flxSCrPoA.TextMatrix(Label10(1).Caption, 8)) - _
         Val(flxSCrPoA.TextMatrix(Label10(1).Caption, 9)) = 0 Then
      szSQL = "UPDATE tlbPayment " & _
              "SET OSAmount = " & Round(CCur(flxSCrPoA.TextMatrix(Label10(1).Caption, 8)) - _
                                  CCur(flxSCrPoA.TextMatrix(Label10(1).Caption, 9)), 2) & ", " & _
                  "PaymentView = False, DateTimeStamp = '',  Module = '', UserSessionID = '', WindowsUserName = '', MachineName ='', PrestigeUserName = '', ServerIPaddress ='' " & _
              "WHERE TransactionID = " & CLng(Label10(5).Caption) & ";"
   Else
      szSQL = "UPDATE tlbPayment " & _
              "SET OSAmount = " & Round(CCur(flxSCrPoA.TextMatrix(Label10(1).Caption, 8)) - _
                                  CCur(flxSCrPoA.TextMatrix(Label10(1).Caption, 9)), 2) & " " & _
              "WHERE TransactionID = " & CLng(Label10(5).Caption) & ";"
   End If
   adoConn.Execute szSQL

    szSQL = "SELECT * FROM PayTransactions;"
    rstPayTrans.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
   
   For iRow = 1 To flxSPayment.Rows - 1
   'what is this A in flxSPayment.TextMatrix(iRow, 15) = "A" ? alwys Inputting A while inserting some amount
      If flxSPayment.TextMatrix(iRow, 15) = "A" And flxSPayment.TextMatrix(iRow, 0) <> "-" And Val(flxSPayment.TextMatrix(iRow, 10)) <> 0 Then
      'Val(flxSPayment.TextMatrix(iRow, 10)) <> 0 was addedd by anol 20170221 issue 317
      'because it was adding 0 valued paytrans in the table
      'Update the OSAmt of Invoices/Debit transactions by PAYMENET amount
         If CCur(flxSPayment.TextMatrix(iRow, 9)) - _
             CCur(flxSPayment.TextMatrix(iRow, 10)) = 0 Then                'OutStanding Amount = 0
             szSQL = "UPDATE tlbPayment " & _
                     "SET tlbPayment.OSAmount = " & _
                        Round(CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                         CCur(flxSPayment.TextMatrix(iRow, 10)), 2) & ", " & _
                        "PaymentView = False " & _
                     "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(iRow, 19)) & ";"
         Else                                                               'OutStanding Amount > 0 ,UNITID='" & flxSPayment.TextMatrix(iRow, 4) & "'
            szSQL = "UPDATE tlbPayment " & _
                    "SET tlbPayment.OSAmount = " & _
                        Round(CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                         CCur(flxSPayment.TextMatrix(iRow, 10)), 2) & " " & _
                    "WHERE tlbPayment.TransactionID = " & CLng(flxSPayment.TextMatrix(iRow, 19)) & ";" 'AND UNITID='" & flxSPayment.TextMatrix(iRow, 4) & "'
         End If
         adoConn.Execute szSQL
'  Saving the cross referencing in the PayTransacitons table
'         ~CREDIT NOTE~ & ~PAYMENT ON ACCOUNT~

    'Type 7 Purchase Credit
    'Type 8 Purchase Payment
    'Type 9 Purchase Payment on Account
    'Label10(1) what is this?  this is active row (iActiveRow = IIf(flxSCrPoA.row = 0, 1, flxSCrPoA.row))
    If flxSCrPoA.TextMatrix(Label10(1).Caption, 13) >= 7 Or _
            flxSCrPoA.TextMatrix(Label10(1).Caption, 13) <= 9 Then
            With rstPayTrans
               .AddNew
               !TranType = IIf(UCase(flxSCrPoA.TextMatrix(Label10(1).Caption, 1)) = "ADJC", "ADJAL", "AL")
               !TransactionID = lRT_ID
               !Alloc_Unalloc = 1
               !FromTran = CLng(Label10(5).Caption) 'Transaction ID for payment line second grid when we are selecting this is holding value
               !ToTran = CLng(flxSPayment.TextMatrix(iRow, 19)) 'PI transaction ID it is coming when we are loading flxSPayment
               '!AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
                'Issue 437 Allocation date
               !AllocDate = Format(flxSCrPoA.TextMatrix(Label10(1).Caption, 4), "DD MMMM YYYY")
               
               !PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 10))
               !Discount = 0
               !IsSageUpdate = IIf(UCase(flxSCrPoA.TextMatrix(Label10(1).Caption, 1)) = "ADJC", False, True)
               !UpdateSage = False
               'update bank code by anol 2019 05 23 it was taking bank code frm the combo but the value adlready existas in grid,which is correct as input'
               !BankCode = flxSCrPoA.TextMatrix(Label10(1).Caption, 15)
               !nominalCode = !BankCode
               !SlNumber = Mid(flxSCrPoA.TextMatrix(Label10(1).Caption, 0), 3)
               !fundID = CLng(flxSPayment.TextMatrix(iRow + 1, 13)) 'there is fundID in the next  row. because next row contains the split
               lRT_ID = lRT_ID + 1
               .Update
            End With
         End If
'

      End If
   Next iRow
   rstPayTrans.Close

''   Next iRow
'  Delete all splits of PA, PC, PP except the first one
   szSQL = "SELECT MAX(TRANSACTIONID)+1 AS TID FROM PayTransactionsSplit;"
   rstPayTrans.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID = CLng(IIf(IsNull(rstPayTrans!TID), 1, rstPayTrans!TID))
   rstPayTrans.Close
   
   adoConn.Execute "DELETE * FROM tlbPaymentSplit " & _
                   "WHERE PayHeader = " & CLng(Label10(5).Caption) 'Transaction ID for payment line second grid when we are selecting this is holding value

   rstSplit.Open "SELECT * FROM tlbPaymentSplit;", adoConn, adOpenDynamic, adLockOptimistic
    
    szSQL = "SELECT MAX(TransactionID)+1 AS TID FROM PayTransactionsSplit;"
   rstSet.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lSPTran_ID = CLng(IIf(IsNull(rstSet!TID), 1, rstSet!TID))
   rstSet.Close
   
   
   For iRow = 1 To flxSPayment.Rows - 1
      If flxSPayment.TextMatrix(iRow, 0) <> "+" And _
               flxSPayment.TextMatrix(iRow, 0) <> ">" And _
            Val(flxSPayment.TextMatrix(iRow, 10)) <> 0 Then
'  Saving the allocation splits in table tlbPaymentSplit
         'With rstSplit
            rstSplit.AddNew
            rstSplit.Fields.Item("TransactionID").Value = UniqueID()
            rstSplit.Fields.Item("PayHeader").Value = CLng(Label10(5).Caption)
            lSp_ID = CLng(Label10(5).Caption)
            rstSplit.Fields.Item("FundID").Value = flxSPayment.TextMatrix(iRow, 13)
            rstSplit.Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 10)
            'rstSplit.Fields.Item("OSAmount").Value = flxSPayment.TextMatrix(iRow, 10)
            flxSCrPoA.TextMatrix(flxSCrPoA.row, 9) = Val(flxSCrPoA.TextMatrix(flxSCrPoA.row, 9)) - _
                                                   Val(rstSplit.Fields.Item("Amount").Value)
            If flxSPayment.TextMatrix(iRow, 0) = "" Then
               rstSplit.Fields.Item("SplitID").Value = iRow
            Else
               rstSplit.Fields.Item("SplitID").Value = flxSPayment.TextMatrix(iRow, 1)
            End If
             splitID = rstSplit.Fields.Item("SplitID").Value
            rstSplit.Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
            rstSplit.Fields.Item("Description").Value = flxSPayment.TextMatrix(iRow, 7)
            rstSplit.Fields.Item("AllocTranID").Value = flxSPayment.TextMatrix(iRow, 3)
            rstSplit.Fields.Item("PayTransactionIDSplit").Value = lSPTran_ID
            rstSplit.Fields.Item("propertyID").Value = flxSPayment.TextMatrix(iRow, 4)
            rstSplit.Fields.Item("ClientStatementID").Value = IIf(flxSPayment.TextMatrix(iRow, 32) = "", Null, flxSPayment.TextMatrix(iRow, 32))
            
            rstSplit.Update
            'now write allocation details here 2021-10-22
            'transactionID has not no relation ship with PayTransactions transactionID
            'and also PayTransactionsSplit has no relationship via a key with PayTransactions
             With rsPaytransaction
                szSQL = "SELECT * FROM PayTransactionsSplit;"
               rsPaytransaction.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                .AddNew
                !TranType = "AL"
                !TransactionID = lSPTran_ID
                !Alloc_Unalloc = 1
                !FromTran = lSp_ID                   'Payment transaction ID
                'added by anol 2021-10-16 . Here the fix was for payment was not writing unitID that means propertyID in tlbpayment which was causing trouble in find
                'supplier OS in Rent summary statement
                'you cannot touch this AllocTranID  field you are using this field while unallocation
                'adoconn.Execute "Update tlbpaymentSplit SET UNIT_ID='" & flxSPayment.TextMatrix(iRow, 4) & "',PayTransactionID=" & lSPTran_ID & " where transactionID='" & lSp_ID & "' and splitID=" & rsInvoice!SplitId & " "
                !ToTran = CLng(flxSPayment.TextMatrix(iRow, 19)) 'PI transaction ID
               ' !AllocDate = Format(Date, "DD MMMM YYYY") 'Issue 437 Allocation Date modification

                !AllocDate = Format(txtSPDate.text, "dd mmmm yyyy")
                !PaymentAmount = CCur(flxSPayment.TextMatrix(iRow, 10))

                !BankCode = txtBankCode.text
                !nominalCode = !BankCode
                '!SlNumber = lSlNumber
                 'Need to write fund ID 2021-09-13
                !fundID = CLng(flxSPayment.TextMatrix(iRow, 13))
                If flxSPayment.TextMatrix(iRow, 14) = 24 Then
                     !VATAMOUNT = 0
                     !NetAmount = !PaymentAmount
                Else
                    !VATAMOUNT = CalculateVatAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 19)), splitID, !PaymentAmount)
                    !NetAmount = CalculateNetAmountFormPIsplit(adoConn, CLng(flxSPayment.TextMatrix(iRow, 19)), splitID, !PaymentAmount)
                 End If

                !VAT_PERIOD_END_DATE = Null ' I am not sure about it, need to ask
                !SplitIDofPI = splitID
                !deleteFlag = False
                .Update
            End With
           
            adoConn.Execute "UPDATE tlbPaymentSplit " & _
                            "SET PayTransactionIDSplit=" & rsPaytransaction!TransactionID & ",OSAmount = " & Round(CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                                  CCur(flxSPayment.TextMatrix(iRow, 10)), 2) & " " & _
                            "WHERE TransactionID = '" & flxSPayment.TextMatrix(iRow, 3) & "';"
          rsPaytransaction.Close
           lSPTran_ID = lSPTran_ID + 1
'         End With
      End If
   Next iRow

'   ConfigFlxSCrPoA
'   ConfigFlxSPayment
'   cmbSPSupplier.text = ""
'   cboBC.text = ""
  
'   ReDim baChangesMade(flxSPayment.Rows) As Boolean

   
   rstSplit.Close
   'adoConn.Close
   Set rstPayTrans = Nothing
   'Set adoConn = Nothing
   szTran2Fix = "" 'Need to clear this before it stores any value to this will cause a problem
   If PI_Check(adoConn, szTran2Fix) = False Then
        MsgBox "An error occurred while allocating, transaction has been rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked on allocation"
        Exit Function
   End If
   If PayAllocation_check(adoConn, szTran2Fix) = False Then
       'MsgBox "An error occurred while allocating, transaction has been rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked on allocation"
        Exit Function
   End If
   SavingAllocationS = True
End Function
Private Function PayAllocation_check(adoConn As ADODB.Connection, ByRef szTran2Fix As String) As Boolean
    'this function is written by anol 2019 05 23 when found that (issue 673 )Updated OS amount  extra incorrectly
    'This function shall prevent saving the data if when outstading amount on payment is not updated.
    'This functionshall compare allocation with payment amount and outstanding amount
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRST As New ADODB.Recordset
    Dim i As Integer
    Dim strWhere As String
    
    
    strWhere = " AND R.sageaccountnumber ='" & txtSPSupplier.Tag & "'"
    
    '        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
    '                    "Where a.ToTran = r.TransactionID  " & StrWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
     'Problem is here you are saving by split and compare it by header amount which is problematic
    rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
      " ToTran from PayTransactions  where DeleteFlag=False group By ToTran ) as A where A.ToTran=R.transactionID " & strWhere & " AND round((amount-amt),2)<>round(osamount,2)", adoConn, adOpenStatic, adLockReadOnly
    
    While Not rsChecksum.EOF
        szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
        rsChecksum.MoveNext
    Wend
    
    rsChecksum.Close
    Set rsChecksum = Nothing
    
    If szTran2Fix = "" Then
        PayAllocation_check = True
    Else
        MsgBox "A problem occurred while creating this transaction: " & _
             Chr(13) & szTran2Fix & "." & _
             "Please contact PCM Consulting. This transaction has not been saved.", _
             vbInformation + vbOKOnly, "PI Payment Allocation rollbacked!"
    End If
End Function
Private Function PI_Check(adoConn As ADODB.Connection, ByRef szTran2Fix As String) As Boolean
       Dim szSQL      As String
       Dim adoRST     As New ADODB.Recordset
      
        
    
       
          szSQL = "SELECT  P.TransactionID,ROUND(P.Amount - P.OSAmount, 2) AS amt , Q.T " & _
                   "FROM tlbPayment AS P, (" & _
                         "SELECT PayHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                         "From tlbPaymentSplit " & _
                         "Group by PayHeader " & _
                         ") AS Q " & _
                   "WHERE P.TransactionID = Q.PayHeader AND P.Amount <> P.OSAmount AND " & _
                         "ROUND(P.Amount - P.OSAmount, 2) <> Q.T;"
    'Debug.Print szSQL
          adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    
          While Not adoRST.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRST.Fields.Item("TransactionID").Value)
    
             adoRST.MoveNext
          Wend
    
          adoRST.Close
      
        szSQL = "SELECT P.SlNumber, P.TRAN_DATE, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID,  " & _
        "tblPurInvSRec.TRAN_ID, tblPurInvSRec.NOMINAL_CODE, tblPurInvSRec.DESCRIPTION " & _
        "FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID " & _
        "WHERE (((P.TransactionType)=6 Or (P.TransactionType)=7) AND ((P.NLPost)=False))  AND P.TOTAL_AMOUNT<>0 AND  tran_ID is null; "
    'Debug.Print szSQL
    'This means there is PI without splits
          adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    
          While Not adoRST.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRST.Fields.Item("SlNumber").Value)
            'if you found this problem exists then there is some transaction in tblPurInv which are inconsitence
            'Now by comparing with the NLposting this transaction will be zerorized
            'To prevent happening this again there is check I have implemented  in PI
             adoRST.MoveNext
          Wend
    
          adoRST.Close
          szSQL = "SELECT  P.SlNumber,P.TOTAL_AMOUNT, Q.T FROM tblPurinv AS P, (SELECT ParentID, Sum(ROUND(TOTAL_AMOUNT, 2)) AS T From tblPurInvSRec Group by ParentID ) AS Q  " & _
            "WHERE P.MY_ID = Q.ParentID AND round(P.TOTAL_AMOUNT,2) <> round(Q.T,2);"
           adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
          While Not adoRST.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRST.Fields.Item("SlNumber").Value)
             adoRST.MoveNext
          Wend
          
          adoRST.Close
          Set adoRST = Nothing
          'This part shall check for the dupplicate serial number in the tblPurInv
          szSQL = "SELECT slnumber,transactiontype, COUNT(*) From tblPurInv GROUP BY slnumber, transactiontype HAVING COUNT(*) > 1"
          adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
          While Not adoRST.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRST.Fields.Item("SlNumber").Value)
             adoRST.MoveNext
          Wend
          
          adoRST.Close
          Set adoRST = Nothing
     
       If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)
    
       If Len(szTran2Fix) > 0 Then
            PI_Check = False
       Else
            PI_Check = True
            'MsgBox "HI"
       End If
      
End Function
Private Sub cmdPayAutomatic_Click()
   Frame4(1).Left = tabPurExp.Left + tabPayment.Left + Frame5(1).Left + Frame5(1).Width - 60
   Frame4(1).Top = tabPurExp.Top + tabPayment.Top + Frame5(1).Top - Frame4(1).Height + Frame5(1).Height - 20
   Frame4(1).Visible = True
   cmdAutoAllocSel.SetFocus

   tabPurExp.Enabled = False
   tabPayment.Enabled = False
End Sub

Private Sub cmdPayAllocate_Click()
   If txtSPSupplier.text = "" Then
      ShowMsgInTaskBar "Please select a supplier.", , "N"
      FocusControl cmdSPSupplier
      Exit Sub
   End If

   If Val(txtSPaymentTotal.text) < 0 Then Exit Sub
   If (flxSCrPoA.Rows = 0) Then Exit Sub
   'If flxSPayment.Rows < 2 Then Exit Sub issue
   'If flxSPayment.TextMatrix(1, 1) = "" Then Exit Sub

   Call SumUpDrCr

'   If dDrOS = 0 Or dCrOS = 0 Then
'      MsgBox "There is no " & IIf(dDrOS = 0, "Debit", "Credit") & " statement." & Chr(13) & _
'             "Therefore, allocation is not possible here.", vbCritical + vbOKOnly, "Auto Allocation"
'      Exit Sub
'   End If

   cGridSPTotal = 0

   Dim iRow As Integer

   If cmdPayAllocate.Caption = "All&ocation Only" Then                        'Allocation Only
       '      added by anol 20170118 sattle all feature is not for allocation
            If dDrOS = 0 Or dCrOS = 0 Then
                MsgBox "There is no " & IIf(dDrOS = 0, "Debit", "Credit") & " statement." & Chr(13) & _
                "Therefore, allocation is not possible here.", vbCritical + vbOKOnly, "Auto Allocation"
            Exit Sub
            End If
            chkSettleAll.Value = 0
            chkSettleAll.Visible = False
            txtSPaymentTotal.text = "0.00" 'added by anol 2019 05 23 that means all payments shall be made by 2nd grid all work is done by All&ocation Only button
            cmdSPSupplier.Enabled = False
            Frame8(1).Enabled = False
            Frame5(5).Visible = False
            cmeRevereseAllocation.Visible = Frame5(5).Visible
            Frame5(1).Left = Frame5(5).Left
            Frame5(1).Top = Frame5(5).Top
            Frame5(1).Visible = True
            cmdPayAllocate.Caption = "&Payment Only"
            Label3(5).Visible = True
            txtAllocatedDiff(1).Visible = True
            Label3(1).Visible = True
            txtPaymentEntered.text = "0.00"
            flxSCrPoA.Enabled = True
            
            cTotalAdjI = 0
            cTotalSI = 0

      For iRow = 1 To flxSPayment.Rows - 1
         If (flxSPayment.TextMatrix(iRow, 2) = "AdjI") Then
            cTotalAdjI = cTotalAdjI + CCur(flxSPayment.TextMatrix(iRow, 9))
         Else
            txtSPaymentTotal.text = Val(txtSPaymentTotal.text) - Val(flxSPayment.TextMatrix(iRow, 10))
            flxSPayment.TextMatrix(iRow, 10) = "0.00"
            cTotalSI = cTotalSI + CCur(flxSPayment.TextMatrix(iRow, 9))
         End If
      Next iRow

      Frame5(1).Enabled = True
   Else                                                                       'Receipt Only
      If cmdPayAllocateSave.Enabled Then
'         If MsgBox("Do like to cancel the allocation?", vbYesNo + vbQuestion, "Allocation") = vbNo Then
'            cmdPayAllocateSave.SetFocus
'            Exit Sub
'         End If

         Frame4(1).Visible = False
         tabPurExp.Enabled = True
         tabPayment.Enabled = True

         cmdPayAllocateSave.Enabled = False
         cmdPayAutomatic.Enabled = True
      End If
      Frame8(1).Enabled = True
      Frame5(5).Visible = True
      cmeRevereseAllocation.Visible = Frame5(5).Visible
      Frame5(5).Enabled = True
      Frame5(1).Visible = False
      cmdPayAllocate.Caption = "All&ocation Only"
      cmdSPSupplier.Enabled = True
      Label3(5).Visible = False
      txtAllocatedDiff(1).Visible = False
      Label3(1).Visible = False
      lblAllocating(1).Visible = False

      AllocDiscard
   End If
End Sub

Private Sub LoadSupplierOnPayment(ByVal adoConn As ADODB.Connection, Filter As String)
   Dim adoRST  As New ADODB.Recordset

   Dim adoC    As New ADODB.Recordset
   Dim adoMA   As New ADODB.Recordset

   Dim szSQL      As String
   Dim iTotalRow  As Integer
   Dim j          As Integer
   Dim i          As Integer
   Dim iTotalCol  As Integer
   Dim Data()     As String
    
   Call ConfigFlxSupplier
   'On Error GoTo ErrorHandler

'Modified by anol 30 Aug 2015
'issue 571 note 1156
'   szSQL = "SELECT SupplierID, SupplierName  " & _
'           "FROM Supplier " & _
'           "WHERE TYPE = 'SUPPLIER' " & _
'           "ORDER BY SupplierName;"
    If txtSupplierType.text = "" Then Exit Sub
    If txtSupplierType.text = "All Categories" Then
'           szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
'              "FROM Supplier " & _
'              "ORDER BY SupplierName ;"
              szSQL = "SELECT s.SupplierID, s.SupplierName, s.TYPE, p.Amt,q.amt1 " & _
              "FROM ((SELECT SupplierID, SupplierName, TYPE FROM Supplier ORDER BY SupplierName) AS s " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt " & _
              "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' GROUP BY SageAccountNumber) AS p ON s.SupplierID = p.SageAccountNumber) " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt1 " & _
              "FROM tlbPayment GROUP BY SageAccountNumber) AS Q ON s.SupplierID = Q.SageAccountNumber " & _
              "ORDER BY s.SupplierName;"
              
    ElseIf txtSupplierType.text = "Supplier" Then
'            szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
'              "FROM Supplier " & _
'              "WHERE TYPE = 'SUPPLIER' " & _
'              "ORDER BY SupplierName;"
            szSQL = "SELECT s.SupplierID, s.SupplierName, s.TYPE,  p.Amt,q.amt1 " & _
              "FROM (SELECT SupplierID, SupplierName, TYPE FROM Supplier where TYPE = 'SUPPLIER' ORDER BY SupplierName) AS s " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt " & _
              "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' GROUP BY SageAccountNumber) AS p ON s.SupplierID = p.SageAccountNumber " & _
             "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt1 " & _
              "FROM tlbPayment GROUP BY SageAccountNumber) AS Q ON s.SupplierID = Q.SageAccountNumber " & _
              "ORDER BY s.SupplierName;"
              
    ElseIf txtSupplierType.text = "Client" Then
'            szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
'              "FROM Supplier " & _
'              "WHERE TYPE = 'CLIENT' " & _
'              "ORDER BY SupplierName;"
          szSQL = "SELECT s.SupplierID, s.SupplierName, s.TYPE,  p.Amt,q.amt1  " & _
              "FROM (SELECT SupplierID, SupplierName, TYPE FROM Supplier where TYPE = 'CLIENT' ORDER BY SupplierName) AS s " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt " & _
              "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' GROUP BY SageAccountNumber) AS p ON s.SupplierID = p.SageAccountNumber " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt1 " & _
              "FROM tlbPayment GROUP BY SageAccountNumber) AS Q ON s.SupplierID = Q.SageAccountNumber " & _
              "ORDER BY s.SupplierName;"
              
    ElseIf txtSupplierType.text = "Managing Agent" Then
'            szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
'              "FROM Supplier " & _
'              "WHERE TYPE = 'AGENT' " & _
'              "ORDER BY SupplierName;"
        szSQL = "SELECT s.SupplierID, s.SupplierName, s.TYPE,  p.Amt,q.amt1  " & _
              "FROM (SELECT SupplierID, SupplierName, TYPE FROM Supplier where TYPE = 'AGENT' ORDER BY SupplierName) AS s " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt " & _
              "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' GROUP BY SageAccountNumber) AS p ON s.SupplierID = p.SageAccountNumber " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt1 " & _
              "FROM tlbPayment GROUP BY SageAccountNumber) AS Q ON s.SupplierID = Q.SageAccountNumber " & _
              "ORDER BY s.SupplierName;"
              
    ElseIf txtSupplierType.text = "Landlord" Then
'            szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
'              "FROM Supplier " & _
'              "WHERE TYPE = 'LLORD' " & _
'              "ORDER BY SupplierName;"
              szSQL = "SELECT s.SupplierID, s.SupplierName, s.TYPE,  p.Amt,q.amt1  " & _
              "FROM (SELECT SupplierID, SupplierName, TYPE FROM Supplier where TYPE = 'LLORD' ORDER BY SupplierName) AS s " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt " & _
              "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' GROUP BY SageAccountNumber) AS p ON s.SupplierID = p.SageAccountNumber " & _
              "LEFT JOIN (SELECT SageAccountNumber,  SUM(Switch(type=6, Amount, type=24, Amount, type=7, -Amount, type=8, -Amount, type=9, -Amount)) AS Amt1 " & _
              "FROM tlbPayment GROUP BY SageAccountNumber) AS Q ON s.SupplierID = Q.SageAccountNumber " & _
              "ORDER BY s.SupplierName;"
              
    End If
    'adoRst.Close
    If szSQL = "" Then Exit Sub
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Filter <> "" Then
        adoRST.Filter = Filter
    End If
'
'   szSQL = "SELECT DISTINCT L.LandlordID, L.LandlordName " & _
'           "FROM   PropertyLandlord AS PL, Landlord AS L " & _
'           "WHERE  PL.LandlordID = L.LandlordID;"
''Debug.Print szSQL
'   adoLL.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szSQL = "SELECT   DISTINCT C.ClientID, C.ClientName " & _
'           "FROM     Client AS C  where 1=2 " & _
'           "ORDER BY C.ClientName;"
'Debug.Print szSQL
'   adoC.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT   DISTINCT A.AgentID, A.AgentName " & _
'           "FROM     Agent AS A where 1=2 " & _
'           "ORDER BY A.AgentName;"


    flxSupplier(1).Rows = adoRST.RecordCount + 2
    i = 1
'    Debug.Print time
    While Not adoRST.EOF
              flxSupplier(1).TextMatrix(i, 0) = adoRST.Fields("Type").Value
              flxSupplier(1).TextMatrix(i, 1) = adoRST.Fields("SupplierID").Value
              flxSupplier(1).TextMatrix(i, 2) = adoRST.Fields("SupplierName").Value
              'flxSupplier(1).TextMatrix(i, 3) = Format(GetSupplierBalance(adoRst.Fields("SupplierID").Value), "0.00")
              'flxSupplier(1).TextMatrix(i, 4) = Format(GetSuppBalByClient(adoRst.Fields("SupplierID").Value), "0.00")
              flxSupplier(1).TextMatrix(i, 3) = Format(IIf(IsNull(adoRST.Fields("Amt1").Value), 0, adoRST.Fields("Amt1").Value), "0.00")
              flxSupplier(1).TextMatrix(i, 4) = Format(IIf(IsNull(adoRST.Fields("Amt").Value), 0, adoRST.Fields("Amt").Value), "0.00") 'Format((adoRst.Fields("Amt").Value), "0.00")
              adoRST.MoveNext
              i = i + 1
    Wend
'    Debug.Print time
'    Debug.Print 1
'anol 06 July 2015
'flxSupplier(1).ColWidth(0) = 0
''   While Not adoRst.EOF
''        'rem out by anol 20160523
''        flxSupplier(1).TextMatrix(i, 0) = "Supplier"
''        flxSupplier(1).TextMatrix(i, 1) = adoRst.Fields("SupplierID").Value
''        flxSupplier(1).TextMatrix(i, 2) = adoRst.Fields("SupplierName").Value
''        If cmdACType.text = "Client" Then
''            flxSupplier(1).TextMatrix(i, 3) = Format(GetClientBalance(adoRst.Fields.Item(0).Value), "0.00")
''        ElseIf cmdACType.text = "Landlord" Then
''            flxSupplier(1).TextMatrix(i, 3) = Format(GetLandlordBalance(adoRst.Fields.Item(0).Value), "0.00")
''        ElseIf cmdACType.text = "Managing Agent" Then
''            flxSupplier(1).TextMatrix(i, 3) = Format(GetAgentBalance(adoRst.Fields.Item(0).Value), "0.00")
''        Else 'supplier
''            flxSupplier(1).TextMatrix(i, 3) = Format(GetSupplierBalance(adoRst.Fields.Item(0).Value), "0.00")
''            flxSupplier(1).TextMatrix(i, 4) = Format(GetSuppBalByClient(adoRst.Fields.Item(0).Value), "0.00")
''        End If
''      adoRst.MoveNext
''      'If Not adoRst.EOF Then flxSupplier(1).AddItem ""
''      i = i + 1
''   Wend
'
'   While Not adoLL.EOF
'      If Not adoLL.EOF Then flxSupplier(1).AddItem ""
'      flxSupplier(1).TextMatrix(i, 0) = "Landlord"
'      flxSupplier(1).TextMatrix(i, 1) = adoLL.Fields.Item(0).Value
'      flxSupplier(1).TextMatrix(i, 2) = adoLL.Fields.Item(1).Value
'      flxSupplier(1).TextMatrix(i, 3) = Format(GetLandlordBalance(adoLL.Fields.Item(0).Value), "0.00")
'      adoLL.MoveNext
'      i = i + 1
'   Wend

'Below two para commneted by anol 20 July 2015
'   While Not adoC.EOF
'      If Not adoC.EOF Then flxSupplier(1).AddItem ""
'      flxSupplier(1).TextMatrix(i, 0) = "Client"
'      flxSupplier(1).TextMatrix(i, 1) = adoC.Fields.Item(0).Value
'      flxSupplier(1).TextMatrix(i, 2) = adoC.Fields.Item(1).Value
'      flxSupplier(1).TextMatrix(i, 3) = Format(GetClientBalance(adoC.Fields.Item(0).Value), "0.00")
'      adoC.MoveNext
'      i = i + 1
'   Wend
'
'   While Not adoMA.EOF
'      If Not adoMA.EOF Then flxSupplier(1).AddItem ""
'      flxSupplier(1).TextMatrix(i, 0) = "Managing Agent"
'      flxSupplier(1).TextMatrix(i, 1) = adoMA.Fields.Item(0).Value
'      flxSupplier(1).TextMatrix(i, 2) = adoMA.Fields.Item(1).Value
'      flxSupplier(1).TextMatrix(i, 3) = Format(GetAgentBalance(adoMA.Fields.Item(0).Value), "0.00")
'      adoMA.MoveNext
'      i = i + 1
'   Wend

NoRes:
'   adoLL.Close
   adoRST.Close
'   adoC.Close
'   adoMA.Close
   Set adoRST = Nothing
'   Set adoLL = Nothing
'   Set adoC = Nothing
'   Set adoMA = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

'   Set adoRst = Nothing
''   Set adoLL = Nothing
'   Set adoC = Nothing
'   Set adoMA = Nothing
End Sub

Private Sub LoadflxSupplier(ByVal adoConn As ADODB.Connection)
   ConfigFlxSupplier2

   Dim adoRST  As New ADODB.Recordset
'   Dim adoLL   As New ADODB.Recordset
   Dim adoC    As New ADODB.Recordset
   Dim adoMA   As New ADODB.Recordset

   Dim szSQL      As String
   Dim iTotalRow  As Integer
   Dim j          As Integer
   Dim i          As Integer
   Dim iTotalCol  As Integer
   Dim Data()     As String

   'On Error GoTo ErrorHandler


         szSQL = "SELECT SupplierID, SupplierName,TYPE  " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' Or TYPE = 'AGENT'  Or TYPE = 'Client' Or TYPE = 'LLORD' " & _
           "ORDER BY TYPE,SupplierName;"
 
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT DISTINCT L.LandlordID, L.LandlordName " & _
'           "FROM   PropertyLandlord AS PL, Landlord AS L " & _
'           "WHERE  PL.LandlordID = L.LandlordID;"
''Debug.Print szSQL
'   adoLL.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szSQL = "SELECT   DISTINCT C.ClientID, C.ClientName " & _
'           "FROM     Client AS C Where 1=2 " & _
'           "ORDER BY C.ClientName;"
''Debug.Print szSQL
'   adoC.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT   DISTINCT A.AgentID, A.AgentName " & _
'           "FROM     Agent AS A Where 1=2 " & _
'           "ORDER BY A.AgentName;"
''Debug.Print szSQL
'   adoMA.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


'   iTotalRow = adoRST.RecordCount + adoLL.RecordCount + adoC.RecordCount + adoMA.RecordCount
   iTotalRow = adoRST.RecordCount '+ adoC.RecordCount + adoMA.RecordCount
   If iTotalRow = 0 Then GoTo NoRes
   
   iTotalCol = adoRST.Fields.Count

'   ReDim Data(iTotalCol, iTotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Supplier"
'   If adoRst.RecordCount > 0 Then
'      For i = 1 To iTotalRow
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'          Next j
'          Data(j, i) = Format(GetSupplierBalance(adoRst.Fields.Item(0).Value), "0.00")
'          adoRst.MoveNext
'          If adoRst.EOF Then Exit For
'      Next i
'   End If
''
''   If adoLL.RecordCount > 0 Then
''      For i = i + 1 To iTotalRow
''          For j = 0 To iTotalCol - 1
''              Data(j, i) = IIf(IsNull(adoLL.Fields.Item(j).Value), "", adoLL.Fields.Item(j).Value)
''          Next j
''          adoLL.MoveNext
''          If adoLL.EOF Then Exit For
''      Next i
''   End If
''Commented out by anol on 30 Aug 2015
''issue 571 note 1156
'
'
'
'   If adoC.RecordCount > 0 Then
'      For i = i + 1 To iTotalRow
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoC.Fields.Item(j).Value), "", adoC.Fields.Item(j).Value)
'          Next j
'          adoC.MoveNext
'          If adoC.EOF Then Exit For
'      Next i
'   End If
'
'   If adoMA.RecordCount > 0 Then
'      For i = i + 1 To iTotalRow
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoMA.Fields.Item(j).Value), "", adoMA.Fields.Item(j).Value)
'          Next j
'          adoMA.MoveNext
'          If adoMA.EOF Then Exit For
'      Next i
'   End If
'
'   cboAccount.Column() = Data()
'   cboAccount.Value = "ALL"

'  LoadFlxSupplier ---->>
   If adoRST.RecordCount > 0 Then adoRST.MoveFirst
'   If adoLL.RecordCount > 0 Then adoLL.MoveFirst
'   If adoC.RecordCount > 0 Then adoC.MoveFirst
'   If adoMA.RecordCount > 0 Then adoMA.MoveFirst
   i = 1
   If Not adoRST.EOF Then
         flxSupplier(2).TextMatrix(i, 0) = "ALL"
         flxSupplier(2).TextMatrix(i, 1) = "ALL"
         flxSupplier(2).TextMatrix(i, 2) = "All Supplier"
         flxSupplier(2).AddItem ""
         i = i + 1
   End If
   While Not adoRST.EOF
      flxSupplier(2).TextMatrix(i, 0) = adoRST.Fields.Item(2).Value
      flxSupplier(2).TextMatrix(i, 1) = adoRST.Fields.Item(0).Value
      flxSupplier(2).TextMatrix(i, 2) = adoRST.Fields.Item(1).Value
      adoRST.MoveNext
      If Not adoRST.EOF Then flxSupplier(2).AddItem ""
      i = i + 1
   Wend
'
'   While Not adoLL.EOF
'      If Not adoLL.EOF Then flxSupplier(2).AddItem ""
'      flxSupplier(2).TextMatrix(i, 0) = "Landlord"
'      flxSupplier(2).TextMatrix(i, 1) = adoLL.Fields.Item(0).Value
'      flxSupplier(2).TextMatrix(i, 2) = adoLL.Fields.Item(1).Value
'      adoLL.MoveNext
'      i = i + 1
'   Wend

'   While Not adoC.EOF
'      If Not adoC.EOF Then flxSupplier(2).AddItem ""
'      flxSupplier(2).TextMatrix(i, 0) = "Client"
'      flxSupplier(2).TextMatrix(i, 1) = adoC.Fields.Item(0).Value
'      flxSupplier(2).TextMatrix(i, 2) = adoC.Fields.Item(1).Value
'      adoC.MoveNext
'      i = i + 1
'   Wend

'   While Not adoMA.EOF
'      If Not adoMA.EOF Then flxSupplier(2).AddItem ""
'      flxSupplier(2).TextMatrix(i, 0) = "Managing Agent"
'      flxSupplier(2).TextMatrix(i, 1) = adoMA.Fields.Item(0).Value
'      flxSupplier(2).TextMatrix(i, 2) = adoMA.Fields.Item(1).Value
'      adoMA.MoveNext
'      i = i + 1
'   Wend

NoRes:
'   adoLL.Close
   adoRST.Close
'   adoC.Close
'   adoMA.Close
   Set adoRST = Nothing
'   Set adoLL = Nothing
'   Set adoC = Nothing
'   Set adoMA = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   Set adoRST = Nothing
'   Set adoLL = Nothing
   Set adoC = Nothing
   Set adoMA = Nothing
End Sub

'Private Function GetClientBalance(szClientID As String) As Currency
'   Dim j As Integer
'
'   For j = 0 To UBound(szaClientBal, 2) - 1
'      If szClientID = szaClientBal(0, j) Then
'         GetClientBalance = Format(szaClientBal(1, j), "0.00")
'         Exit For
'      End If
'   Next j
'   If j = UBound(szaClientBal, 2) Then GetClientBalance = 0
'End Function

'Private Function GetAgentBalance(szAgentID As String) As Currency
'   Dim j As Integer
'
'   For j = 0 To UBound(szaAgentBal, 2) - 1
'      If szAgentID = szaAgentBal(0, j) Then
'         GetAgentBalance = Format(szaAgentBal(1, j), "0.00")
'         Exit For
'      End If
'   Next j
'   If j = UBound(szaAgentBal, 2) Then GetAgentBalance = 0
'End Function

Private Function GetSupplierBalance(szSuppID As String) As Currency
   On Error GoTo Err
   Dim j As Integer

   For j = 0 To UBound(szaSupplierBal, 2) - 1
      If szSuppID = szaSupplierBal(0, j) Then
         GetSupplierBalance = Format(szaSupplierBal(1, j), "0.00")
         Exit For
      End If
   Next j
   If j = UBound(szaSupplierBal, 2) Then GetSupplierBalance = 0
   Exit Function
Err:
End Function
Private Function GetSuppBalByClient(szSuppID As String) As Currency
'written by anol 22 Aug 2016
   Dim j As Integer

   For j = 0 To UBound(szaSuppBalbyClient, 2) - 1
      If szSuppID = szaSuppBalbyClient(0, j) Then
         GetSuppBalByClient = Format(szaSuppBalbyClient(1, j), "0.00")
         Exit For
      End If
   Next j
   If j = UBound(szaSuppBalbyClient, 2) Then GetSuppBalByClient = 0
End Function
'
'Private Function GetLandlordBalance(szLLID As String) As Currency
'   Dim j As Integer
'
'   For j = 0 To UBound(szaLandlordBal, 2) - 1
'      If szLLID = szaLandlordBal(0, j) Then
'         GetLandlordBalance = Format(szaLandlordBal(1, j), "0.00")
'         Exit For
'      End If
'   Next j
'   If j = UBound(szaLandlordBal, 2) Then GetLandlordBalance = 0
'End Function
'
''  Build up LANDLORDs' Account BALANCE
Private Sub LandlordAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer

   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

'-----------------      Purchase Side    -----------------------------------
   szSqlPI = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "FROM   tlbPayment, Landlord " & _
             "WHERE  tlbPayment.SageAccountNumber = Landlord.LandlordID " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSqlPI, adoConn, adOpenStatic, adLockReadOnly

   iPI = IIf(adoPayDr.EOF, 0, adoPayDr.Fields.Item(0).Value)

   adoPayDr.Close

'-----------------      Sales Side    -----------------------------------
   szSQLSI = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT R.SageAccountNumber  " & _
             "FROM   tlbReceipt AS R, Landlord " & _
             "WHERE  R.SageAccountNumber = Landlord.LandlordID " & _
             "GROUP BY R.SageAccountNumber" & _
            ");"
   adoRptDr.Open szSQLSI, adoConn, adOpenStatic, adLockReadOnly

   iSI = IIf(adoRptDr.EOF, 0, adoRptDr.Fields.Item(0).Value)

   adoRptDr.Close

   If iSI + iPI = 0 Then
      ReDim szaLandlordBal(0, 0) As String
      Set adoPayDr = Nothing
      Set adoRptDr = Nothing
      Exit Sub
   End If

   ReDim szaLandlordBal(1, iSI + iPI) As String

'--------------------------------------------------------------  PURCHASE SIDE   -----------------------------------
   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Supplier " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND P.SageAccountNumber = Supplier.SupplierID AND Supplier.Type='LLORD' " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaLandlordBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaLandlordBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Supplier " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber =  = Supplier.SupplierID AND Supplier.Type='LLORD' " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaLandlordBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaLandlordBal(1, i) = szaLandlordBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaLandlordBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaLandlordBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing

'--------------------------------------------------------------  SALES SIDE   -----------------------------------

   szSQL = "SELECT   R.SageAccountNumber, SUM(R.Amount) AS Dr " & _
           "FROM     tlbReceipt AS R, Landlord " & _
           "WHERE   (R.Type = 1 OR R.Type = 23) AND " & _
                    "R.SageAccountNumber = Landlord.LandlordID " & _
           "GROUP BY R.SageAccountNumber;"

   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


   While Not adoRptDr.EOF
      For i = 0 To iIndex - 1
         If szaLandlordBal(0, i) = adoRptDr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaLandlordBal(1, i) = szaLandlordBal(1, i) + adoRptDr.Fields.Item("Dr").Value
      Else
         szaLandlordBal(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
         szaLandlordBal(1, iIndex) = adoRptDr.Fields.Item("Dr").Value
         iIndex = iIndex + 1
      End If
      adoRptDr.MoveNext
   Wend

   adoRptDr.Close

   szSQL = "SELECT   R.SageAccountNumber, SUM(R.Amount) AS Cr " & _
           "FROM     tlbReceipt AS R, Landlord " & _
           "WHERE    R.Type <> 1 AND R.Type <> 23 AND " & _
                    "R.SageAccountNumber = Landlord.LandlordID " & _
           "GROUP BY R.SageAccountNumber;"

   adoRptCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRptCr.EOF
      For i = 0 To iIndex - 1
         If szaLandlordBal(0, i) = adoRptCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaLandlordBal(1, i) = szaLandlordBal(1, i) - Val(adoRptCr.Fields.Item("Cr").Value)
      Else
         szaLandlordBal(0, iIndex) = adoRptCr.Fields.Item("SageAccountNumber").Value
         szaLandlordBal(1, iIndex) = adoRptCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoRptCr.MoveNext
   Wend

   adoRptCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

'Build Supplier AC balance by client by anol 22 Aug 2016
Private Sub SupplierAcBalByClient(adoConn As ADODB.Connection)
    
'Build Supplier AC balance by client by anol 22 Aug 2016
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Dr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type = 6 Or P.Type = 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSuppBalbyClient(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSuppBalbyClient(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSuppBalbyClient(1, iIndex) = 0
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close
   '6
   'New section1
   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Dr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P,tblPurInv AS I " & _
                     "Where  I.MY_ID=P.PI AND P.Type = 6 AND I.CL_ID='" & txtClientIDPurPay.text & "' " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
'adoPayCr.Close
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
      'Debug.Print Val(adoPayCr.Fields.Item("Dr").Value)
         szaSuppBalbyClient(1, i) = szaSuppBalbyClient(1, i) + Val(adoPayCr.Fields.Item("Dr").Value)
      Else
         szaSuppBalbyClient(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSuppBalbyClient(1, iIndex) = adoPayCr.Fields.Item("Dr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close
'24
   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Dr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P,tblPurInv AS I " & _
                     "Where I.MY_ID=P.PI AND P.Type = 24 AND I.CL_ID='" & txtClientIDPurPay.text & "' " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaSuppBalbyClient(1, i) = szaSuppBalbyClient(1, i) + Val(adoPayCr.Fields.Item("Dr").Value)
      Else
         szaSuppBalbyClient(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSuppBalbyClient(1, iIndex) = adoPayCr.Fields.Item("Dr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close
   '8
   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Cr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
                     "FROM tlbPayment AS P " & _
                     "Where P.Type = 8 AND P.ClientID='" & txtClientIDPurPay.text & "' " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaSuppBalbyClient(1, i) = szaSuppBalbyClient(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaSuppBalbyClient(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSuppBalbyClient(1, iIndex) = -adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close
   
   
   'End new section 1

  szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Cr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
                     "FROM tlbPayment AS P,tblPurInv I " & _
                     "Where I.MY_ID=P.PI AND  P.Type <> 6 And P.Type <> 8 And P.Type <> 24 AND I.CL_ID='" & txtClientIDPurPay.text & "' " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"
'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaSuppBalbyClient(1, i) = szaSuppBalbyClient(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaSuppBalbyClient(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSuppBalbyClient(1, iIndex) = -adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

'######################################      CLIENT         ##############################################
'''   iIndex = UBound(szaSuppBalbyClient, 2)
'''
'''   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
'''           "FROM ( " & _
'''               "SELECT S.ClientID, P.Dr " & _
'''               "FROM Client AS S LEFT JOIN ( " & _
'''                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
'''                     "FROM tlbPayment AS P " & _
'''                     "Where (P.Type = 6 Or P.Type = 24) " & _
'''                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'''                           "P.SageAccountNumber = S.ClientID " & _
'''               ") AS X;"
'''
''''Debug.Print szSQL
'''   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'''
'''   ReDim Preserve szaSuppBalbyClient(1, iIndex + adoPayDr.RecordCount) As String
'''
'''   While Not adoPayDr.EOF
'''      szaSuppBalbyClient(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'''      szaSuppBalbyClient(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
'''      iIndex = iIndex + 1
'''      adoPayDr.MoveNext
'''   Wend
'''
'''   adoPayDr.Close
'''
'''   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
'''           "FROM ( " & _
'''               "SELECT S.ClientID, P.Cr " & _
'''               "FROM Client AS S LEFT JOIN ( " & _
'''                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
'''                     "FROM tlbPayment AS P " & _
'''                     "Where (P.Type <> 6 And P.Type <> 24) " & _
'''                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'''                           "P.SageAccountNumber = S.ClientID " & _
'''               ") AS X;"
'''
''''Debug.Print szSQL
'''   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'''
'''   While Not adoPayCr.EOF
'''      For i = 0 To iIndex - 1
'''         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'''            Exit For
'''         End If
'''      Next i
'''      If i <= iIndex - 1 Then
'''         szaSuppBalbyClient(1, i) = szaSuppBalbyClient(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
'''      Else
'''         szaSuppBalbyClient(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
'''         szaSuppBalbyClient(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
'''         iIndex = iIndex + 1
'''      End If
'''      adoPayCr.MoveNext
'''   Wend
'''
'''   adoPayCr.Close

End Sub
Private Sub SupplierAcBalByClient2(adoConn As ADODB.Connection)
    
'Build Supplier AC balance by client by anol 22 Aug 2016
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT SupplierID AS SageAccountNumber " & _
           "FROM Supplier ;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSuppBalbyClient(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSuppBalbyClient(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      If "MARIACED" = adoPayDr.Fields.Item("SageAccountNumber").Value Then
'                    MsgBox adoPayDr.Fields.Item("SageAccountNumber").Value
'       End If
      szaSuppBalbyClient(1, iIndex) = 0
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close
   '6
   'New section1
   szSQL = "SELECT Type, SageAccountNumber,Sum(Amount) AS Amt " & _
                 "FROM tlbPayment where ClientID='" & txtClientIDPurPay.text & "' group by  Type, SageAccountNumber ;"

'Debug.Print szSQL
'adoPayCr.Close
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSuppBalbyClient(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            If "EON" = adoPayCr.Fields.Item("SageAccountNumber").Value Then
                    Debug.Print adoPayCr.Fields.Item("SageAccountNumber").Value
            End If
            If adoPayCr.Fields.Item("Type").Value = 6 Or adoPayCr.Fields.Item("Type").Value = 24 Then
                 szaSuppBalbyClient(1, i) = Val(szaSuppBalbyClient(1, i)) + adoPayCr.Fields.Item("Amt").Value
            End If
            If adoPayCr.Fields.Item("Type").Value = 8 Or adoPayCr.Fields.Item("Type").Value = 7 Or adoPayCr.Fields.Item("Type").Value = 9 Then
                 szaSuppBalbyClient(1, i) = Val(szaSuppBalbyClient(1, i)) - adoPayCr.Fields.Item("Amt").Value
            End If
         End If
      Next i
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

End Sub
'  Build up SUPPLIER'S Account BALANCE
Private Sub SupplierAccountBalance(adoConn As ADODB.Connection) 'Written by anol 20181112
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   lblLoading.Caption = "Please wait while building balances..."
   fmeLoading.Refresh
  szSQL = "SELECT Supplier.SupplierID  " & _
             "From Supplier " & _
             "order by SupplierID ;"



   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSupplierBal(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SupplierID").Value
      szaSupplierBal(1, iIndex) = 0
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

  szSQL = "SELECT SageAccountNumber, Type, SUM(Amount) AS Amt " & _
           "FROM tlbPayment " & _
           "GROUP BY SageAccountNumber,Type;"
   
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            If adoPayCr.Fields.Item("Type").Value = 6 Or adoPayCr.Fields.Item("Type").Value = 24 Then
                szaSupplierBal(1, i) = szaSupplierBal(1, i) + Val(adoPayCr.Fields.Item("Amt").Value)
            End If
            If adoPayCr.Fields.Item("Type").Value = 7 Or adoPayCr.Fields.Item("Type").Value = 8 Or adoPayCr.Fields.Item("Type").Value = 9 Then
                szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Amt").Value)
            End If
         End If
         
      Next i
      adoPayCr.MoveNext
   Wend
   'Debug.Print time
   adoPayCr.Close
'   If chkShowBal.Value = 1 Then
'       UpdateBalance
'   End If

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
   fmeLoading.Visible = False
   lblLoading.Caption = "Please wait while loading data..."
End Sub
'Private Sub SupplierAccountBalance(adoConn As ADODB.Connection)
'   Dim szSQL As String, i As Integer, iIndex As Integer
'   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
'
'   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
'           "FROM ( " & _
'               "SELECT S.SupplierID, P.Dr " & _
'               "FROM Supplier AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type = 6 Or P.Type = 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.SupplierID " & _
'               "WHERE S.TYPE = 'SUPPLIER' " & _
'           ") AS X;"
'
''Debug.Print szSQL
'   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim szaSupplierBal(1, adoPayDr.RecordCount) As String
'
'   iIndex = 0
'   While Not adoPayDr.EOF
'      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
'      iIndex = iIndex + 1
'      adoPayDr.MoveNext
'   Wend
'
'   adoPayDr.Close
'
'   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
'           "FROM ( " & _
'               "SELECT S.SupplierID, P.Cr " & _
'               "FROM Supplier AS S LEFT JOIN ( " & _
'                  "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
'                  "FROM tlbPayment AS P " & _
'                  "Where P.Type <> 6 And P.Type <> 24 " & _
'                  "GROUP BY P.SageAccountNumber) AS P ON P.SageAccountNumber = S.SupplierID " & _
'               "WHERE TYPE = 'SUPPLIER' " & _
'           ") AS X;"
'
''Debug.Print szSQL
'   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoPayCr.EOF
'      For i = 0 To iIndex - 1
'         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'            Exit For
'         End If
'      Next i
'      If i <= iIndex - 1 Then
'         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
'      Else
'         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
'         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
'         iIndex = iIndex + 1
'      End If
'      adoPayCr.MoveNext
'   Wend
'
'   adoPayCr.Close
'
''######################################      CLIENT         ##############################################
'   iIndex = UBound(szaSupplierBal, 2)
'
'   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
'           "FROM ( " & _
'               "SELECT S.ClientID, P.Dr " & _
'               "FROM Client AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type = 6 Or P.Type = 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.ClientID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String
'
'   While Not adoPayDr.EOF
'      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
'      iIndex = iIndex + 1
'      adoPayDr.MoveNext
'   Wend
'
'   adoPayDr.Close
'
'   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
'           "FROM ( " & _
'               "SELECT S.ClientID, P.Cr " & _
'               "FROM Client AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type <> 6 And P.Type <> 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.ClientID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoPayCr.EOF
'      For i = 0 To iIndex - 1
'         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'            Exit For
'         End If
'      Next i
'      If i <= iIndex - 1 Then
'         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
'      Else
'         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
'         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
'         iIndex = iIndex + 1
'      End If
'      adoPayCr.MoveNext
'   Wend
'
'   adoPayCr.Close
''
'''######################################      LANDLORD       ##############################################
''   iIndex = UBound(szaSupplierBal, 2)
''
''   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
''           "FROM ( " & _
''               "SELECT S.LandlordID, P.Dr " & _
''               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
''                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
''                     "FROM tlbPayment AS P " & _
''                     "Where (P.Type = 6 Or P.Type = 24) " & _
''                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
''                           "P.SageAccountNumber = S.LandlordID " & _
''               ") AS X;"
''
'''Debug.Print szSQL
''   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String
''
''   While Not adoPayDr.EOF
''      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
''      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
''      iIndex = iIndex + 1
''      adoPayDr.MoveNext
''   Wend
''
''   adoPayDr.Close
''
''   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
''           "FROM ( " & _
''               "SELECT S.LandlordID, P.Cr " & _
''               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
''                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
''                     "FROM tlbPayment AS P " & _
''                     "Where (P.Type <> 6 And P.Type <> 24) " & _
''                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
''                           "P.SageAccountNumber = S.LandlordID " & _
''               ") AS X;"
''
'''Debug.Print szSQL
''   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   While Not adoPayCr.EOF
''      For i = 0 To iIndex - 1
''         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
''            Exit For
''         End If
''      Next i
''      If i <= iIndex - 1 Then
''         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
''      Else
''         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
''         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
''         iIndex = iIndex + 1
''      End If
''      adoPayCr.MoveNext
''   Wend
''
''   adoPayCr.Close
'
''######################################      AGENT       ##############################################
'   iIndex = UBound(szaSupplierBal, 2)
'
'   szSQL = "SELECT DISTINCT X.AgentID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
'           "FROM ( " & _
'               "SELECT S.AgentID, P.Dr " & _
'               "FROM Agent AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type = 6 Or P.Type = 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.AgentID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String
'
'   While Not adoPayDr.EOF
'      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
'      iIndex = iIndex + 1
'      adoPayDr.MoveNext
'   Wend
'
'   adoPayDr.Close
''
''   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
''           "FROM ( " & _
''               "SELECT S.LandlordID, P.Cr " & _
''               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
''                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
''                     "FROM tlbPayment AS P " & _
''                     "Where (P.Type <> 6 And P.Type <> 24) " & _
''                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
''                           "P.SageAccountNumber = S.LandlordID " & _
''               ") AS X;"
''
'''Debug.Print szSQL
''   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   While Not adoPayCr.EOF
''      For i = 0 To iIndex - 1
''         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
''            Exit For
''         End If
''      Next i
''      If i <= iIndex - 1 Then
''         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
''      Else
''         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
''         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
''         iIndex = iIndex + 1
''      End If
''      adoPayCr.MoveNext
''   Wend
''
''   adoPayCr.Close
'
'   Set adoPayDr = Nothing
'   Set adoPayCr = Nothing
'End Sub
'  Build up CLIENTs' Account BALANCE
Private Sub ClientAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer

   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

'-----------------      Purchase Side    -----------------------------------
   szSqlPI = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "FROM   tlbPayment, Client " & _
             "WHERE  tlbPayment.SageAccountNumber = Client.ClientID " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSqlPI, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaClientBal(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Client " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND P.SageAccountNumber = Client.ClientID " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaClientBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaClientBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Client " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber = Client.ClientID " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaClientBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaClientBal(1, i) = szaClientBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaClientBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaClientBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

'  Build up AGENTs' Account BALANCE
Private Sub AgentAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer

   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

'-------------------------------------------------------      Purchase Side       -----------------------------------
   szSqlPI = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "FROM   tlbPayment,supplier " & _
             "where supplier.SupplierID=tlbPayment.SageAccountNumber AND supplier.Type='AGENT'" & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSqlPI, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaAgentBal(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

'--------------------------------------------------------------  PURCHASE SIDE   -----------------------------------
   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P,supplier " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND P.SageAccountNumber = supplier.SupplierID AND supplier.Type='AGENT' " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaAgentBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaAgentBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Agent " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber = Agent.AgentID " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaAgentBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaAgentBal(1, i) = szaAgentBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaAgentBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaAgentBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

'Private Sub PrepareCboBC(ByVal adoConn As ADODB.Connection)
'   On Error GoTo Error_Handler
'
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, Data() As String, j As Integer
'   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer
'   Dim iDefaultBankAC As Integer
'   iDefaultBankAC = -1
'
''   If cboClient.Value = "ALL" Then
''      szSQL = "SELECT C.NominalCode AS BNC, " & _
''                  "C.Bank_AC_Name AS BNN, C.CLIENT_ID " & _
''              "FROM tlbClientBanks AS C " & _
''              "WHERE C.CLIENT_ID <> '' order by C.NominalCode;"
''   Else
''      szSQL = "SELECT C.NominalCode AS BNC, " & _
''                  "C.Bank_AC_Name AS BNN, C.CLIENT_ID " & _
''              "FROM tlbClientBanks AS C " & _
''              "WHERE C.CLIENT_ID = '" & cboClient.Value & "' order by C.NominalCode;"
''   End If
'   'Modified by anol 05 July 2015
'   'issue 571 Note 1121
''   Bank account needs to show the nominal bank account description and not the client bank description same as
''   it does in demand receipts.
'
'If txtClientIDPurPay.text = "ALL" Then
'      szSQL = "SELECT C.NominalCode AS BNC, " & _
'                  "N.Name AS BNN, C.CLIENT_ID,DEFAULT_AC " & _
'              "FROM tlbClientBanks AS C,NominalLedger as N " & _
'              "WHERE C.NominalCode = N.Code AND C.CLIENT_ID=N.ClientID AND C.CLIENT_ID <> '' and 1=2 order by C.NominalCode;"
'   Else
'      szSQL = "SELECT C.NominalCode AS BNC, " & _
'                  "N.Name AS BNN, C.CLIENT_ID,DEFAULT_AC " & _
'              "FROM tlbClientBanks AS C,NominalLedger as N " & _
'              "WHERE C.NominalCode = N.Code AND C.CLIENT_ID=N.ClientID AND C.CLIENT_ID = '" & txtClientIDPurPay.text & "' order by C.NominalCode;"
'   End If
'
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
'   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String
'   ReDim szNominal(iTotalCol - 1, iTotalRow - 1) As String
'   For i = 0 To iTotalRow
'      For j = 0 To iTotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'      Next j
'      If CBool(adoRst.Fields.Item("DEFAULT_AC").Value) = True Then
'            iDefaultBankAC = i
'       End If
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
''   cboBC.Column() = Data()
'   If iDefaultBankAC < 0 Then
'         MsgBox "Please set a default Client Bank Account for: " & txtClientIDPurPay.text & "", vbInformation, "Warning"
'         cboBC.ListIndex = 0
'   Else
'         cboBC.ListIndex = iDefaultBankAC
'   End If
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'Error_Handler:
'   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
'
'   Set adoRst = Nothing
'End Sub

'Private Sub LoadPayAmtType(szValue As String, adoConn As ADODB.Connection)
'   Dim SQLStr1 As String, szaData() As String, i As Integer
'   Dim adoRst As New ADODB.Recordset
'
'   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
'             "FROM PrimaryCode, SecondaryCode " & _
'             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
'                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
'             "ORDER BY SecondaryCode.Value;"
'
'   adoRst.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      adoRst.Close
'      Set adoRst = Nothing
'      Exit Sub
'   End If
'
'   ReDim szaData(1, adoRst.RecordCount - 1) As String
'
'   cmbSPAmtType.Clear
'   i = 0
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst!c
'      szaData(1, i) = adoRst!V
'      adoRst.MoveNext
'      i = i + 1
'   Wend
'   adoRst.Close
'   Set adoRst = Nothing
'
'   cmbSPAmtType.Column() = szaData()
'End Sub

Private Sub ConfigFlxSPayment()
   Dim szHeader As String

   With flxSPayment
      .Clear
      .Cols = 33
      .Rows = 1
      .RowHeight(0) = 0

      szHeader$ = "|<No.|<Type|<Tenant A/C|<Unit ID|<Due Date" & _
                  "|<Ref|<Details|>Amount £|>O/S Amt. £|>Receipt £|>Discount" & _
                  "|<DemandID|<FundID £|<RptNo.|ClientID"

      .ColAlignment(0) = vbCenter
      .ColWidth(0) = Label1(1).Left - .Left    'Sign
      .ColWidth(1) = Label1(2).Left - Label1(1).Left    'No
      .ColWidth(2) = Label1(3).Left - Label1(2).Left    'Type
      .ColWidth(3) = 0      'Split ID
      .ColWidth(4) = Label1(4).Left - Label1(3).Left    'Unit ID
      .ColWidth(5) = Label1(5).Left - Label1(4).Left    'Date
      .ColAlignment(5) = vbLeftJustify
      .ColWidth(6) = Label1(6).Left - Label1(5).Left    'Ref
      .ColAlignment(6) = vbLeftJustify
      .ColWidth(7) = Label1(7).Left - Label1(6).Left    'Details
      .ColAlignment(7) = vbLeftJustify
      .ColWidth(8) = Label1(8).Left - Label1(7).Left    'Amount
      .ColWidth(9) = Label1(9).Left - Label1(8).Left    'O/S Amount
      .ColWidth(10) = 1000 '.Width + .Left - Label1(9).Left - 600 'Payment that you are inputting
      .ColWidth(11) = 0     'Discount
      .ColWidth(12) = 0     'DemandID
      .ColWidth(13) = 0     'Fund ID
      .ColWidth(14) = 0     'Transaction Type - linked with column 1 Type
      .ColWidth(15) = 0     'R/A; R -> receipt, A -> allocation
      .ColWidth(16) = 0     'allocation ref
      .ColWidth(17) = 0     'allocation amount
      .ColWidth(18) = 0     'Sage Department
      .ColWidth(19) = 0     'Payment No
      .ColWidth(20) = 0     'Reconciliation
      .ColWidth(21) = 0     'Bank Code
      .ColWidth(22) = 0     'RptAmtType
      .ColWidth(23) = 0     'Client ID
      
      .ColWidth(24) = 0     'UserSessionID
      .ColWidth(25) = 0     'WindowsUserName
      .ColWidth(26) = 0     'MachineName
      .ColWidth(27) = 0     'Module
      
       .ColWidth(28) = 0     'PI number
       .ColWidth(29) = 0     'Split ID
       .ColWidth(30) = 0     'isRentPayable
       .ColWidth(31) = 0     'isManagementFee
       .ColWidth(32) = 0     'CSID
       

      .FormatString = szHeader$
   End With
End Sub

Private Sub ResizeFlxSPayment()
   Dim szHeader As String

   With flxSPayment
      .ColWidth(0) = Label1(1).Left - .Left    'Sign
      .ColWidth(1) = Label1(2).Left - Label1(1).Left    'No
      .ColWidth(2) = Label1(3).Left - Label1(2).Left    'Type
      .ColWidth(3) = 0      'Split ID
      .ColWidth(4) = Label1(4).Left - Label1(3).Left    'Unit ID
      .ColWidth(5) = Label1(5).Left - Label1(4).Left    'Date
      .ColAlignment(5) = vbLeftJustify
      .ColWidth(6) = Label1(6).Left - Label1(5).Left    'Ref
      .ColAlignment(6) = vbLeftJustify
      .ColWidth(7) = Label1(7).Left - Label1(6).Left    'Details
      .ColAlignment(7) = vbLeftJustify
      .ColWidth(8) = Label1(8).Left - Label1(7).Left    'Amount
      .ColWidth(9) = Label1(9).Left - Label1(8).Left    'O/S Amount
      .ColWidth(10) = .Width + .Left - Label1(9).Left - 600 'Payment
      .ColWidth(11) = 0     'Discount
      .ColWidth(12) = 0     'DemandID
      .ColWidth(13) = 0     'Fund ID
      .ColWidth(14) = 0     'Transaction Type - linked with column 1 Type
      .ColWidth(15) = 0     'R/A; R -> receipt, A -> allocation
      .ColWidth(16) = 0     'allocation ref
      .ColWidth(17) = 0     'allocation amount
      .ColWidth(18) = 0     'Sage Department
      .ColWidth(19) = 0     'Payment No
   End With
End Sub


Private Sub ConfigFlxSCrPoA()
   Dim szHeader As String

   With flxSCrPoA
      .Clear
      .Cols = 24
      .Rows = 2
      .RowHeight(0) = 0
      szHeader$ = "<No.|<Type|<Tenant A/C|<Unit ID|<Date|<Ref|<Details|>Amount £|>O/S Amt. £" & _
                  "|>Receipt £|TransactionID|<DemandID|>SAGE O/S £|Type|Fund|Bank|ExtRef|AmtType|ClientID"
      .FormatString = szHeader$

      .ColWidth(0) = Label19(3).Left - Label19(2).Left    'Serial No
      .ColWidth(1) = Label19(4).Left - Label19(3).Left    'Type
      .ColWidth(2) = 0      'Tenant A/c - no need to show it in the grid, its already in the header part
      .ColWidth(3) = Label19(5).Left - Label19(4).Left    'Unit ID
      .ColWidth(4) = Label19(6).Left - Label19(5).Left    'Date
      .ColWidth(5) = Label19(7).Left - Label19(6).Left    'Ref
      .ColWidth(6) = Label19(8).Left - Label19(7).Left    'Details
      .ColWidth(7) = Label19(9).Left - Label19(8).Left    'Amount
      .ColWidth(8) = Label19(10).Left - Label19(9).Left   'O/S Amount
      .ColWidth(9) = .Left + .Width - Label19(10).Left - 400 'Receipt
      .ColWidth(10) = 0     'Transaction ID
      .ColWidth(11) = 0     'DemandID
      .ColWidth(12) = 0     'SAGE O/S £
      .ColWidth(13) = 0     'Type ID    ID of col 1
      .ColWidth(14) = 0     'Fund id
      .ColWidth(15) = 0     'BankCode
      .ColWidth(16) = 0     'ExtRef
      .ColWidth(17) = 0     'PayAmtType
      .ColWidth(18) = 0     'ClientID 'added by anol 09 July 2015
      .ColWidth(19) = 0     'SessionID
      .ColWidth(20) = 0
      .ColWidth(21) = 0
      .ColWidth(22) = 0
      .ColWidth(23) = 0
      .ColWidth(24) = 0
      .row = 0
      .col = 0

      txtAllocatedDiff(1).Left = Label19(10).Left
      txtAllocatedDiff(1).Width = .ColWidth(9)
   End With
   Label3(5).Left = txtAllocatedDiff(1).Left - Label3(5).Width - 100
End Sub

Private Function TotalPaymentEntered() As Currency
   Dim i As Integer

   For i = 1 To flxSPayment.Rows - 1
      TotalPaymentEntered = TotalPaymentEntered + IIf(Val(flxSPayment.TextMatrix(i, 10)) > 0 And _
                                                  flxSPayment.TextMatrix(i, 0) <> "+" And _
                                                  flxSPayment.TextMatrix(i, 0) <> ">", _
                                                  Val(flxSPayment.TextMatrix(i, 10)), 0)
   Next i
End Function

Private Sub LoadRptAmtType(szValue As String, adoConn As ADODB.Connection, conCombo As Control)
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

   conCombo.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   conCombo.Column() = szaData()
End Sub

Private Sub cmdPaymentDiscard_Click()
   If MsgBox("Do you wish to discard all changes?", vbQuestion + vbYesNo, "Payment") = vbNo Then Exit Sub

   Dim iRow As Integer

   If bChangesMade Then
      For iRow = 1 To flxSPayment.Rows - 1
         If Val(flxSPayment.TextMatrix(iRow, 10)) > 0 Then
            flxSPayment.TextMatrix(iRow, 10) = "0.00"
         End If
      Next iRow
      txtPaymentEntered.text = "0.00"
      flxSCrPoA.Enabled = True
      txtSPaymentTotal.text = "0.00"
      bChangesMade = False
      cGridSPTotal = 0
      bTotalPayTyped = False
   End If
End Sub

Private Sub SumUpDrCr()
   Dim i As Integer

   dDrOS = 0
   dCrOS = 0
   dCrOS_Adj = 0
   dDrOS_Adj = 0

   For i = 1 To flxSPayment.Rows - 1
      If (flxSPayment.TextMatrix(i, 2) = "ADJC") Then
         If flxSPayment.TextMatrix(i, 0) <> "-" Then
            dDrOS_Adj = dDrOS_Adj + Val(flxSPayment.TextMatrix(i, 9))
         End If
      Else
         If flxSPayment.TextMatrix(i, 0) <> "-" Then
            dDrOS = dDrOS + Val(flxSPayment.TextMatrix(i, 9))
         End If
      End If
   Next i
   For i = 1 To flxSCrPoA.Rows - 1
      If (flxSCrPoA.TextMatrix(i, 1) = "ADJC") Then
         If flxSCrPoA.TextMatrix(i, 0) <> "-" Then
            dCrOS_Adj = dCrOS_Adj + Val(flxSCrPoA.TextMatrix(i, 9))
         End If
      Else
         If flxSCrPoA.TextMatrix(i, 0) <> "-" Then
            dCrOS = dCrOS + Val(flxSCrPoA.TextMatrix(i, 8))
         End If
      End If
   Next i
End Sub

Private Sub cmdAutoAllocSel_Click()
   If tabPayment.Tab = 0 Then
      'Resolved by BOSL
      'issue 477  Purchase Payment Allocation inconsistency
      'solution : I have put here a check point
      'Modified by Anol 23 Sep 2014
      Dim invoiceAmt As Double
      Dim ReceiptAmt As Double
      Dim iRow As Integer
      For iRow = 1 To flxSCrPoA.Rows - 1
         If flxSCrPoA.RowHeight(iRow) > 0 Then
'            ReceiptAmt = ReceiptAmt + Val(flxSCrPoA.TextMatrix(iRow, 8))
        'modified by anol 2021-10-22
            ReceiptAmt = ReceiptAmt + Val(flxSCrPoA.TextMatrix(iRow, 9))
         End If
      Next iRow
      invoiceAmt = 0
      For iRow = 1 To flxSPayment.Rows - 1
         If flxSPayment.RowHeight(iRow) > 0 Then
            invoiceAmt = invoiceAmt + Val(flxSPayment.TextMatrix(iRow, 9))
         End If
         'added by anol 20170215 automatic allocation was not clearing manual entry and the difference
           flxSPayment.TextMatrix(iRow, 10) = "0.00"
           If chkSettleAll.Value = 1 Then
                txtAllocatedDiff(1).text = "0.00"
           End If
      Next iRow
      If (ReceiptAmt - invoiceAmt) > 0 Then
         MsgBox "The credit amount to be allocated cannot be greater than the available debits", vbInformation, "Warning"
         Exit Sub
      End If
      
      If Not lblAllocating(1).Visible And flxSPayment.TextMatrix(flxSPayment.row, 2) <> "ADJI" And _
                 Val(txtSPaymentTotal.text) >= 0 And Val(txtAllocatedDiff(1).text) = 0 And cmdPayAllocate.Caption <> "All&ocation Only" And chkSettleAll.Value = 0 Then
                 MsgBox " You must first select a credit amount from the grid below before you can enter any amounts in this grid.", vbInformation + vbOKOnly, "Allocation"
                 Frame4(1).Visible = False
                 tabPurExp.Enabled = True
                 tabPayment.Enabled = True
                 FocusControl flxSCrPoA
                 Exit Sub
      End If
      txtAllocatedDiff(1).text = "0.00"
      If chkSettleAll.Value = 0 Then
         AutomaticAllocationTR
      Else
         AutoAllocSA                               'Automatic allocation - Settle All

         Frame4(1).Visible = False
         tabPurExp.Enabled = True
         tabPayment.Enabled = True

         cmdPayAllocateSave.Enabled = True
         cmdPayAllocateSave.SetFocus

         cmdPayAutomatic.Enabled = False
      End If
   End If

   If tabPayment.Tab = 1 Then AutomaticAllocationSP
End Sub

'**   This procedue will be called when user will choose 'Select All'.
'**   All Cr transactions will be allocated against Dr.
'**   The mapping between Dr and Cr will not be shown here.
'**   It will be mapped in at the time of saving.
Private Sub AutoAllocSA()                          'Automatic allocation - Settle All
   Dim iRow As Integer, iInd As Integer, j As Integer, valx As Integer
   Dim dProcessAmount As Double, iGridRow As Integer
   Dim bTrue As Boolean, iCount As Integer, iActiveRow As Integer, cAllocTotal As Currency

   If (flxSCrPoA.Rows = 0) Then Exit Sub

   If (optOIF.Value) Then                                      'Old Invoice First - is selected
      ReDim dSortIndex(flxSPayment.Rows - 1)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSPayment.Rows - 1
         If flxSPayment.TextMatrix(iGridRow, 0) <> "-" Then
            bTrue = False
            For iInd = 0 To iCount - 1
               If (Not IsNull(dSortIndex(iInd)) And dSortIndex(iInd) <> 0) Then
                  If CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) > (CDate(flxSPayment.TextMatrix(iGridRow, 5))) Then
                     valx = dSortIndex(iInd)
                     dSortIndex(iInd) = iGridRow

                     For j = iCount To iInd + 1 Step -1
                        dSortIndex(j) = dSortIndex(j - 1)
                     Next j

                     dSortIndex(iInd + 1) = valx
                     bTrue = True
                  End If
               End If
            Next iInd
            If Not bTrue Then
               dSortIndex(iCount) = iGridRow
            End If
            iCount = iCount + 1
         End If
      Next iGridRow
   Else                                                  'Recent Invoice First - is selected
      ReDim dSortIndex(flxSPayment.Rows - 1)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSPayment.Rows - 1
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) < (CDate(flxSPayment.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd
         If Not bTrue Then
            dSortIndex(iCount) = iGridRow
         End If
         iCount = iCount + 1
      Next iGridRow
   End If

   dProcessAmount = MinAmtAlloc(flxSPayment, flxSCrPoA)

   If (flxSCrPoA.TextMatrix(1, 1) = "ADJC") Then
      For j = 0 To iCount - 1
         If (flxSPayment.TextMatrix(dSortIndex(j), 2) = "ADJI") Then
            If (dProcessAmount > flxSPayment.TextMatrix(dSortIndex(j), 9)) Then
               flxSPayment.TextMatrix(dSortIndex(j), 10) = flxSPayment.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               dProcessAmount = dProcessAmount - flxSPayment.TextMatrix(dSortIndex(j), 9)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
            Else
               flxSPayment.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
               Exit For
            End If
         End If
      Next j

      flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cAllocTotal, "0.00")
   Else
      cAllocTotal = 0
      j = 0
      While cAllocTotal < dProcessAmount
         If flxSPayment.TextMatrix(dSortIndex(j), 2) <> "AdjI" Then
            If (dProcessAmount - cAllocTotal > flxSPayment.TextMatrix(dSortIndex(j), 9)) Then
               flxSPayment.TextMatrix(dSortIndex(j), 10) = flxSPayment.TextMatrix(dSortIndex(j), 9)
            Else
               flxSPayment.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount - cAllocTotal, "0.00")
            End If
            cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
            j = j + 1
         End If
      Wend

      cAllocTotal = 0
      j = 1

      While cAllocTotal < dProcessAmount
         If flxSCrPoA.TextMatrix(j, 1) <> "AdjC" Then
            If (dProcessAmount - cAllocTotal > flxSCrPoA.TextMatrix(j, 8)) Then
               flxSCrPoA.TextMatrix(j, 9) = flxSCrPoA.TextMatrix(j, 8)
            Else
               flxSCrPoA.TextMatrix(j, 9) = Format(dProcessAmount - cAllocTotal, "0.00")
            End If
            cAllocTotal = cAllocTotal + CCur(flxSCrPoA.TextMatrix(j, 9))
            j = j + 1
         End If
      Wend
   End If
End Sub

Private Function MinAmtAlloc(conFlxSPayment As MSHFlexGrid, conFlxSCrPoA As MSHFlexGrid) As Double
   Dim iRow As Integer, dDr As Double, dCr As Double

   MinAmtAlloc = 0

   For iRow = 1 To conFlxSPayment.Rows - 1
      MinAmtAlloc = MinAmtAlloc + CDbl(conFlxSPayment.TextMatrix(iRow, 9))
   Next iRow
   For iRow = 1 To conFlxSCrPoA.Rows - 1
      dCr = dCr + CDbl(conFlxSCrPoA.TextMatrix(iRow, 8))
   Next iRow

   If MinAmtAlloc > dCr Then MinAmtAlloc = dCr
End Function

Private Sub AutomaticAllocationSP()
   Dim iRow As Integer, iInd As Integer, j As Integer, valx As Integer, dOSAmount As Double
   Dim dSortIndex() As Integer, dProcessAmount As Double, iGridRow As Integer
   Dim bTrue As Boolean, iCount As Integer, iActiveRow As Integer

   If (flxSCrPoA.Rows = 0) Then Exit Sub

   If (optOIF.Value) Then
      ReDim dSortIndex(flxSCrPoA.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSCrPoA.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxSCrPoA.TextMatrix(dSortIndex(iInd), 5)) > (CDate(flxSCrPoA.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd

         If Not bTrue Then dSortIndex(iCount) = iGridRow

         iCount = iCount + 1
      Next iGridRow
   Else
      ReDim dSortIndex(flxSCrPoA.Rows - 3)
      dSortIndex(0) = 1
      iCount = 1
      For iGridRow = 2 To flxSCrPoA.Rows - 2
         bTrue = False
         For iInd = 0 To iCount - 1
            If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
               If CDate(flxSCrPoA.TextMatrix(dSortIndex(iInd), 5)) < (CDate(flxSCrPoA.TextMatrix(iGridRow, 5))) Then
                  valx = dSortIndex(iInd)
                  dSortIndex(iInd) = iGridRow

                  For j = iCount To iInd + 1 Step -1
                     dSortIndex(j) = dSortIndex(j - 1)
                  Next j

                  dSortIndex(iInd + 1) = valx
                  bTrue = True
               End If
            End If
         Next iInd
         If Not bTrue Then
            dSortIndex(iCount) = iGridRow
         End If
         iCount = iCount + 1
      Next iGridRow
   End If

   iActiveRow = IIf(flxSCrPoA.row = 0, 1, flxSCrPoA.row)

   Label10(1).Caption = iActiveRow
   Label10(5).Caption = flxSCrPoA.TextMatrix(iActiveRow, 10)

   dOSAmount = CDbl(flxSCrPoA.TextMatrix(iActiveRow, 8))
   dProcessAmount = IIf(flxSCrPoA.TextMatrix(iActiveRow, 9) = "0.00", CDbl(flxSCrPoA.TextMatrix(iActiveRow, 8)), CDbl(flxSCrPoA.TextMatrix(iActiveRow, 9)))

   If (flxSCrPoA.TextMatrix(iActiveRow, 3) = "ADJC") Then
      If dProcessAmount > cTotalAdjI Then
         dProcessAmount = cTotalAdjI
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(dProcessAmount, "0.00")
      End If

      For j = 0 To flxSCrPoA.Rows - 3
         If (flxSCrPoA.TextMatrix(dSortIndex(j), 2) = "ADJI") Then
            If (dProcessAmount > flxSCrPoA.TextMatrix(dSortIndex(j), 9)) Then
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               dProcessAmount = dProcessAmount - flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
            Else
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               Exit For
            End If
         End If
      Next j
   Else
      If dProcessAmount > cTotalSI Then
         dProcessAmount = cTotalAdjI
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cTotalAdjI, "0.00")
      Else
         flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(dProcessAmount, "0.00")
      End If

      For j = 0 To flxSCrPoA.Rows - 3
         If (Not flxSCrPoA.TextMatrix(dSortIndex(j), 2) = "AdjI") Then
            If (dProcessAmount > flxSCrPoA.TextMatrix(dSortIndex(j), 9)) Then
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               dProcessAmount = dProcessAmount - flxSCrPoA.TextMatrix(dSortIndex(j), 9)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
            Else
               flxSCrPoA.TextMatrix(dSortIndex(j), 10) = Format(dProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSCrPoA.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSCrPoA.TextMatrix(dSortIndex(j), 15) = "A"
               flxSCrPoA.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               Exit For
            End If
         End If
       Next j
   End If

   Frame4(1).Visible = False
   tabPurExp.Enabled = True
   tabPayment.Enabled = True

   cmdPayAllocateSave.Enabled = True
   cmdPayAllocateSave.SetFocus

   cmdPayAutomatic.Enabled = False
End Sub

Private Sub AutomaticAllocationTR()
   Dim iRow As Integer, iInd As Integer, j As Integer, valx As Integer
   Dim dSortIndex() As Integer, ProcessAmount As Double, iGridRow As Integer
   Dim bTrue As Boolean, iCount As Integer, iActiveRow As Integer, cAllocTotal As Currency

   If (flxSCrPoA.Rows = 0) Then Exit Sub

   ReDim dSortIndex(flxSPayment.Rows - 2)
   
   If (optOIF.Value) Then                                      'Old Invoice First - is selected
      dSortIndex(0) = 1
      iCount = 1
     
      For iGridRow = 2 To flxSPayment.Rows - 1
       'Below code is rem by anol  06 June 2016 and below part is copied from demand form
'         If flxSPayment.TextMatrix(iGridRow, 0) <> "-" Then
'            bTrue = False
'            For iInd = 0 To iCount - 1
'               If (Not IsNull(dSortIndex(iInd)) And dSortIndex(iInd) <> 0) Then
'                  If CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) > (CDate(flxSPayment.TextMatrix(iGridRow, 5))) Then
'                     valx = dSortIndex(iInd)
'                     dSortIndex(iInd) = iGridRow
'
'                     For j = iCount To iInd + 1 Step -1
'                        dSortIndex(j) = dSortIndex(j - 1)
'                     Next j
'
'                     dSortIndex(iInd + 1) = valx
'                     bTrue = True
'                  End If
'               End If
'            Next iInd
'            If Not bTrue Then
'               dSortIndex(iCount) = iGridRow
'            End If
'            iCount = iCount + 1
'         End If
         ' New replaced code starts from here
             If flxSPayment.TextMatrix(iGridRow, 0) <> "-" Then
                    For iInd = 0 To iCount - 1
                       If (CDate(flxSPayment.TextMatrix(iGridRow, 5))) < CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) Then
                          j = iCount
                          Do
                             dSortIndex(j) = dSortIndex(j - 1)
                             j = j - 1
                          Loop While j > iInd

                          dSortIndex(iInd) = iGridRow
                          Exit For

                       End If
                    Next iInd
                    If iInd = iCount Then dSortIndex(iCount) = iGridRow
                    iCount = iCount + 1
              End If
            'End of replaced code
     Next iGridRow

     
   Else                                                  'Recent Invoice First - is selected
      dSortIndex(0) = 1
      iCount = 1
      'Below code is rem by anol  06 June 2016 and below part is copied from demand form
'      For iGridRow = 2 To flxSPayment.Rows - 1
'
'         If flxSPayment.TextMatrix(iGridRow, 0) <> "-" Then
'            bTrue = False
'            For iInd = 0 To iCount - 1
'               If (Not IsNull(dSortIndex(iInd)) And Not dSortIndex(iInd) = 0) Then
'                  If CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) < (CDate(flxSPayment.TextMatrix(iGridRow, 5))) Then
'                     valx = dSortIndex(iInd)
'                     dSortIndex(iInd) = iGridRow
'
'                     For j = iCount To iInd + 1 Step -1
'                        dSortIndex(j) = dSortIndex(j - 1)
'                     Next j
'
'                     dSortIndex(iInd + 1) = valx
'                     bTrue = True
'                  End If
'               End If
'            Next iInd
'            If Not bTrue Then
'               dSortIndex(iCount) = iGridRow
'            End If
'            iCount = iCount + 1
'         End If
'        Next iGridRow
' New replaced code starts from here
            For iGridRow = 2 To flxSPayment.Rows - 1
                 If flxSPayment.TextMatrix(iGridRow, 0) <> "-" Then
                    For iInd = 0 To iCount - 1
                       If (CDate(flxSPayment.TextMatrix(iGridRow, 5))) > CDate(flxSPayment.TextMatrix(dSortIndex(iInd), 5)) Then
                          j = iCount
                          Do
                             dSortIndex(j) = dSortIndex(j - 1)
                             j = j - 1
                          Loop While j > iInd
        
                          dSortIndex(iInd) = iGridRow
                          Exit For
        
                       End If
                    Next iInd
                    If iInd = iCount Then dSortIndex(iCount) = iGridRow
                    iCount = iCount + 1
                 End If
              Next iGridRow
              'End of modification
      End If
   iActiveRow = IIf(flxSCrPoA.row = 0, 1, flxSCrPoA.row)

   Label10(1).Caption = iActiveRow
   Label10(5).Caption = flxSCrPoA.TextMatrix(iActiveRow, 10)

   ProcessAmount = IIf(CCur(flxSCrPoA.TextMatrix(iActiveRow, 9)) = 0, CDbl(flxSCrPoA.TextMatrix(iActiveRow, 8)), CDbl(flxSCrPoA.TextMatrix(iActiveRow, 9)))

   If (flxSCrPoA.TextMatrix(iActiveRow, 3) = "ADJC") Then
      For j = 0 To iCount - 1
         If (flxSPayment.TextMatrix(dSortIndex(j), 2) = "ADJI") Then
            If (ProcessAmount > flxSPayment.TextMatrix(dSortIndex(j), 9)) Then
               flxSPayment.TextMatrix(dSortIndex(j), 10) = flxSPayment.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               ProcessAmount = ProcessAmount - flxSPayment.TextMatrix(dSortIndex(j), 9)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               flxSPayment.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
            Else
               flxSPayment.TextMatrix(dSortIndex(j), 10) = Format(ProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               flxSPayment.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
               Exit For
            End If
         End If
      Next j

      flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cAllocTotal, "0.00")
   Else
      For j = 0 To iCount - 1
         If flxSPayment.TextMatrix(dSortIndex(j), 2) <> "AdjI" Then
            If (ProcessAmount > flxSPayment.TextMatrix(dSortIndex(j), 9)) Then
            'added on 20170529
               txtSPayment.text = Format(flxSPayment.TextMatrix(dSortIndex(j), 9), "0.00")
               flxSPayment.row = dSortIndex(j)
               iCurRow = dSortIndex(j)
               txtSPayment_LostFocus
             'end of addition
'               flxSPayment.TextMatrix(dSortIndex(j), 10) = flxSPayment.TextMatrix(dSortIndex(j), 9)
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               ProcessAmount = ProcessAmount - flxSPayment.TextMatrix(dSortIndex(j), 9)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               flxSPayment.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
            Else
               txtSPayment.text = Format(ProcessAmount, "0.00")
               flxSPayment.row = dSortIndex(j)
               iCurRow = dSortIndex(j)
               txtSPayment_LostFocus
'               flxSPayment.TextMatrix(dSortIndex(j), 10) = Format(ProcessAmount, "0.00")
               baChangesMade(dSortIndex(j)) = IIf(Val(flxSPayment.TextMatrix(dSortIndex(j), 10)) > 0, True, False)
               flxSPayment.TextMatrix(dSortIndex(j), 15) = "A"
               flxSPayment.TextMatrix(dSortIndex(j), 16) = Label10(5).Caption
               cAllocTotal = cAllocTotal + CCur(flxSPayment.TextMatrix(dSortIndex(j), 10))
               Exit For
            End If
         End If
      Next j

      flxSCrPoA.TextMatrix(iActiveRow, 9) = Format(cAllocTotal, "0.00")
   End If

   Frame4(1).Visible = False
   tabPurExp.Enabled = True
   tabPayment.Enabled = True

   cmdPayAllocateSave.Enabled = True
   cmdPayAllocateSave.SetFocus

   cmdPayAutomatic.Enabled = False
End Sub

Private Sub cmdAutoAllocSelCancel_Click()
   Frame4(1).Visible = False
   tabPurExp.Enabled = True
   tabPayment.Enabled = True
   cmdPayAutomatic.SetFocus
End Sub

Public Sub AfterEditPayment(adoConn As ADODB.Connection)
   If sEditPPR = 2 Then GoTo EditPayment

   flxSPayment.Clear
   flxSPayment.Rows = 2
   'Unloack prevoius locked item
   adoConn.Execute "Update tlbPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                   "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & UserSessionID & "'"
   LoadFlxSPayment adoConn

   Exit Sub
EditPayment:

   flxSCrPoA.Clear
   flxSCrPoA.Rows = 2

   LoadFlxSCrPoA adoConn
End Sub

Private Sub ConfigFlxSupplier()
   Dim szHeader As String

   flxSupplier(1).Clear
   flxSupplier(1).Rows = 2
   flxSupplier(1).Cols = 5

   szHeader$ = "<|<|<|>"
   flxSupplier(1).FormatString = szHeader$

   flxSupplier(1).RowHeight(0) = 0
   flxSupplier(1).ColWidth(0) = 100 'lblSearch0(2).Left - lblSearch0(1).Left
   flxSupplier(1).ColWidth(1) = 1200 'lblSearch0(3).Left - lblSearch0(2).Left
   flxSupplier(1).ColWidth(2) = 2600 'lblSearch0(4).Left - lblSearch0(3).Left
   flxSupplier(1).ColWidth(3) = 1000 'flxSupplier(1).Width - lblSearch0(4).Left - 300
   flxSupplier(1).ColWidth(4) = 1000
End Sub

Private Sub ConfigFlxSupplier2()
   Dim szHeader As String

   flxSupplier(2).Clear
   flxSupplier(2).Rows = 2
   flxSupplier(2).Cols = 3

   szHeader$ = "<|<|<"
   flxSupplier(2).FormatString = szHeader$

   flxSupplier(2).RowHeight(0) = 0
   flxSupplier(2).ColWidth(0) = lblSearch0(7).Left - lblSearch0(6).Left
   txtAccountSearch(3).Width = lblSearch0(7).Left - lblSearch0(6).Left
   flxSupplier(2).ColWidth(1) = lblSearch0(8).Left - lblSearch0(7).Left
   txtAccountSearch(4).Width = lblSearch0(8).Left - lblSearch0(7).Left
   flxSupplier(2).ColWidth(2) = flxSupplier(2).Width - lblSearch0(8).Left - 300
   txtAccountSearch(5).Width = flxSupplier(2).Width - lblSearch0(8).Left - 300
   txtAccountSearch(3).Left = lblSearch0(6).Left
   txtAccountSearch(4).Left = lblSearch0(7).Left
   txtAccountSearch(5).Left = lblSearch0(8).Left
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

Private Sub txtStartDate_Change()
    TextBoxChangeDate txtStartDate
End Sub
Private Sub txtStartDate_GotFocus()
    SelTxtInCtrl txtStartDate
End Sub
Private Sub txtStartDate_LostFocus()
    If txtStartDate.text <> "" Then
        TextBoxFormatDate txtStartDate
'        txtEndDate.text = txtStartDate.text
        SelTxtInCtrl txtEndDate
     End If
End Sub
Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEndDate.SetFocus
    End If
    TextBoxKeyPrsDate txtStartDate, KeyAscii
End Sub
Private Sub txtEndDate_Change()
     TextBoxChangeDate txtEndDate
End Sub
Private Sub txtEndDate_GotFocus()
    SelTxtInCtrl txtEndDate
End Sub
Private Sub txtEndDate_LostFocus()
    If txtEndDate.text <> "" Then TextBoxFormatDate txtEndDate
End Sub
Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPrintHistOK.SetFocus
    End If
    TextBoxKeyPrsDate txtEndDate, KeyAscii
End Sub



