VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLandLord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Landlord"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLandlord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleMode       =   0  'User
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Height          =   1095
      Left            =   120
      TabIndex        =   139
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton cmdSupplier 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   240
      End
      Begin VB.CommandButton cmdSupplierType 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   9855
         TabIndex        =   4
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit:"
         Height          =   195
         Index           =   6
         Left            =   10440
         TabIndex        =   145
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
         Height          =   195
         Index           =   5
         Left            =   10440
         TabIndex        =   144
         Top             =   660
         Width           =   975
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   143
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Landlord A/C:"
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
         TabIndex        =   142
         Top             =   210
         Width           =   945
      End
      Begin MSForms.TextBox txtSupplierID 
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   210
         Width           =   3660
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "6456;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSupplierName 
         Height          =   315
         Left            =   1695
         TabIndex        =   2
         Top             =   645
         Width           =   3945
         VariousPropertyBits=   746604571
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "6959;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboSupplierType 
         Height          =   315
         Left            =   7620
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3836;556"
         TextColumn      =   2
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
         Caption         =   "Landlord Type:"
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
         Left            =   6000
         TabIndex        =   141
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax / VAT Number:"
         Height          =   195
         Index           =   21
         Left            =   6000
         TabIndex        =   140
         Top             =   660
         Width           =   1425
      End
      Begin MSForms.TextBox txtCreditLimit 
         Height          =   315
         Left            =   11565
         TabIndex        =   6
         Top             =   240
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   15
         BorderStyle     =   1
         Size            =   "2672;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSupplierACBal 
         Height          =   315
         Left            =   11565
         TabIndex        =   7
         Top             =   660
         Width           =   1515
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2672;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTaxVatNumber 
         Height          =   315
         Left            =   7620
         TabIndex        =   5
         Top             =   660
         Width           =   2460
         VariousPropertyBits=   746604571
         MaxLength       =   15
         BorderStyle     =   1
         Size            =   "4339;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdEditBank_ 
      Caption         =   "&Edit"
      Height          =   360
      Left            =   6720
      TabIndex        =   68
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveBank_ 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6720
      TabIndex        =   67
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelBank_ 
      Caption         =   "Canc&el"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6720
      TabIndex        =   66
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox fraList 
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
      Height          =   2535
      Index           =   0
      Left            =   7440
      ScaleHeight     =   2505
      ScaleWidth      =   4785
      TabIndex        =   53
      Top             =   6960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdGridUnitClose 
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   1935
         Index           =   0
         Left            =   40
         TabIndex        =   55
         Top             =   540
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3413
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
            Name            =   "Arial"
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
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   20
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
         Index           =   0
         Left            =   840
         TabIndex        =   62
         Top             =   20
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
         Index           =   0
         Left            =   1800
         TabIndex        =   61
         Top             =   20
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch3 
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   60
         Top             =   20
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch4 
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   59
         Top             =   20
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1560
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
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
         Index           =   0
         Left            =   2160
         TabIndex        =   64
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   45
         Top             =   30
         Width           =   4500
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   260
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
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   57
         Top             =   260
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
      Begin MSForms.TextBox txtSearch3 
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   56
         Top             =   260
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
   End
   Begin VB.PictureBox picSupList 
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
      Height          =   3025
      Left            =   720
      ScaleHeight     =   3000
      ScaleWidth      =   5625
      TabIndex        =   44
      Top             =   6960
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdGridUnitClose2 
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
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   15
         Width           =   255
      End
      Begin VB.TextBox txtSupplierSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   900
      End
      Begin VB.TextBox txtSupplierSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1100
         TabIndex        =   47
         Top             =   300
         Width           =   2800
      End
      Begin VB.TextBox txtSupplierSearch 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4000
         TabIndex        =   48
         Top             =   300
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierList 
         Height          =   2370
         Left            =   40
         TabIndex        =   50
         Top             =   600
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   5
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Balance"
         Height          =   210
         Index           =   2
         Left            =   4005
         TabIndex        =   52
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Landlord Name"
         Height          =   210
         Index           =   1
         Left            =   1095
         TabIndex        =   51
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   75
         Width           =   165
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   45
         Top             =   60
         Width           =   5280
      End
   End
   Begin TabDlg.SSTab tabSupplier 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmLandlord.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSaveSupplierDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdEditSupplierDetails"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelSupplierDetails"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCloseSupplier"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSaveSupplier"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDeleteSupplier"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdEditSupplier"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCancelSupplier"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAddNewSupplier"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Payments"
      TabPicture(1)   =   "frmLandlord.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(1)"
      Tab(1).Control(1)=   "Shape1(3)"
      Tab(1).Control(2)=   "Label1(7)"
      Tab(1).Control(3)=   "txtPaymentTerms"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Label1(17)"
      Tab(1).Control(6)=   "Label1(18)"
      Tab(1).Control(7)=   "Label1(19)"
      Tab(1).Control(8)=   "Label1(20)"
      Tab(1).Control(9)=   "Shape1(0)"
      Tab(1).Control(10)=   "Shape1(2)"
      Tab(1).Control(11)=   "Label3"
      Tab(1).Control(12)=   "Label4"
      Tab(1).Control(13)=   "cboPayType"
      Tab(1).Control(14)=   "cboAccType"
      Tab(1).Control(15)=   "Label1(22)"
      Tab(1).Control(16)=   "Label1(3)"
      Tab(1).Control(17)=   "Label18(0)"
      Tab(1).Control(18)=   "lblVatCode(0)"
      Tab(1).Control(19)=   "txtCodeVat"
      Tab(1).Control(20)=   "Label1(31)"
      Tab(1).Control(21)=   "Label1(32)"
      Tab(1).Control(22)=   "Label1(25)"
      Tab(1).Control(23)=   "Label1(33)"
      Tab(1).Control(24)=   "Label1(26)"
      Tab(1).Control(25)=   "Label1(27)"
      Tab(1).Control(26)=   "Label1(28)"
      Tab(1).Control(27)=   "Label1(29)"
      Tab(1).Control(28)=   "Label1(30)"
      Tab(1).Control(29)=   "txtSortCode"
      Tab(1).Control(30)=   "txtAcNo"
      Tab(1).Control(31)=   "txtAcName"
      Tab(1).Control(32)=   "txtBankPayRef"
      Tab(1).Control(33)=   "cmdAccType"
      Tab(1).Control(34)=   "cmdPayType"
      Tab(1).Control(35)=   "cmdCancelPayments"
      Tab(1).Control(36)=   "cmdEditPayments"
      Tab(1).Control(37)=   "cmdSavePayments"
      Tab(1).Control(38)=   "cmdTaxList(0)"
      Tab(1).Control(39)=   "txtFLX"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "Account History"
      TabPicture(2)   =   "frmLandlord.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblGridCaption(1)"
      Tab(2).Control(1)=   "lblGridCaption(0)"
      Tab(2).Control(2)=   "Label11(8)"
      Tab(2).Control(3)=   "Label11(7)"
      Tab(2).Control(4)=   "Label11(4)"
      Tab(2).Control(5)=   "Label11(3)"
      Tab(2).Control(6)=   "Label11(2)"
      Tab(2).Control(7)=   "Label11(1)"
      Tab(2).Control(8)=   "Label11(5)"
      Tab(2).Control(9)=   "Label11(6)"
      Tab(2).Control(10)=   "Label11(9)"
      Tab(2).Control(11)=   "Label11(10)"
      Tab(2).Control(12)=   "Label11(11)"
      Tab(2).Control(13)=   "Label11(12)"
      Tab(2).Control(14)=   "Label11(13)"
      Tab(2).Control(15)=   "Label11(21)"
      Tab(2).Control(16)=   "Label11(19)"
      Tab(2).Control(17)=   "Label11(14)"
      Tab(2).Control(18)=   "Label11(20)"
      Tab(2).Control(19)=   "Label11(15)"
      Tab(2).Control(20)=   "Label11(16)"
      Tab(2).Control(21)=   "Label11(17)"
      Tab(2).Control(22)=   "Label11(18)"
      Tab(2).Control(23)=   "flxACHistorySplit"
      Tab(2).Control(24)=   "flxACHistory"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Memo/Attachment"
      TabPicture(3)   =   "frmLandlord.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdUnitMemoEdit"
      Tab(3).Control(1)=   "cmdUnitMemoSave"
      Tab(3).Control(2)=   "cmdUnitMemoCancel"
      Tab(3).Control(3)=   "Frame17"
      Tab(3).Control(4)=   "txtNote"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Contacts"
      TabPicture(4)   =   "frmLandlord.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblGridCaption(2)"
      Tab(4).Control(1)=   "Label11(22)"
      Tab(4).Control(2)=   "Label11(23)"
      Tab(4).Control(3)=   "Label11(24)"
      Tab(4).Control(4)=   "Label11(25)"
      Tab(4).Control(5)=   "Label11(26)"
      Tab(4).Control(6)=   "Label11(27)"
      Tab(4).Control(7)=   "Label11(28)"
      Tab(4).Control(8)=   "Label11(29)"
      Tab(4).Control(9)=   "flxContacts"
      Tab(4).Control(10)=   "cmdAddNewContacts"
      Tab(4).Control(11)=   "cmdEditContacts"
      Tab(4).Control(12)=   "txtAddress"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Job Maintenance"
      TabPicture(5)   =   "frmLandlord.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "Job Maintenance"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   159
         Top             =   360
         Width           =   12945
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "View &Job"
            Height          =   355
            Left            =   3120
            TabIndex        =   168
            Top             =   4395
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   6840
            TabIndex        =   167
            Top             =   4395
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   11400
            TabIndex        =   166
            Top             =   4395
            Width           =   1395
         End
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "View &Diary Entry"
            Height          =   355
            Left            =   4680
            TabIndex        =   165
            Top             =   4395
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdEmailJS_PO 
            Caption         =   "Email"
            Height          =   355
            Left            =   9000
            TabIndex        =   164
            Top             =   4395
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   160
            Top             =   4260
            Width           =   2775
            Begin VB.OptionButton optDiary 
               Caption         =   "Diary Entries"
               Height          =   255
               Left            =   1440
               TabIndex        =   161
               Top             =   160
               Width           =   1215
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "Jobs"
               Height          =   255
               Left            =   720
               TabIndex        =   162
               Top             =   160
               Width           =   735
            End
            Begin VB.OptionButton optAll 
               Caption         =   "All"
               Height          =   255
               Left            =   120
               TabIndex        =   163
               Top             =   160
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   3555
            Left            =   120
            TabIndex        =   169
            Top             =   690
            Width           =   12675
            _ExtentX        =   22357
            _ExtentY        =   6271
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
            Caption         =   "Date Reported"
            Height          =   435
            Index           =   2
            Left            =   2385
            TabIndex        =   173
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   7800
            TabIndex        =   180
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   6600
            TabIndex        =   179
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Name / Diary Entry"
            Height          =   495
            Index           =   4
            Left            =   4680
            TabIndex        =   178
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   9000
            TabIndex        =   177
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9840
            TabIndex        =   176
            Top             =   255
            Width           =   435
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   175
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   174
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   3240
            TabIndex        =   172
            Top             =   255
            Width           =   1275
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   10560
            TabIndex        =   171
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   11640
            TabIndex        =   170
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   -68640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   158
         Text            =   "frmLandlord.frx":0972
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdEditContacts 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   -63240
         TabIndex        =   157
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNewContacts 
         Caption         =   "&New"
         Height          =   345
         Left            =   -64800
         TabIndex        =   156
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtFLX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72960
         TabIndex        =   128
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdTaxList 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -71040
         TabIndex        =   73
         Top             =   2365
         Width           =   320
      End
      Begin VB.TextBox txtNote 
         Enabled         =   0   'False
         Height          =   3255
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   125
         Top             =   460
         Width           =   12675
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   120
         Top             =   4255
         Width           =   12735
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete"
            Height          =   435
            Left            =   11280
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add"
            Height          =   435
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open"
            Height          =   435
            Left            =   9540
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   240
            Width           =   1200
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   5490
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9684;503"
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
         Height          =   385
         Left            =   -63645
         TabIndex        =   119
         Top             =   3770
         Width           =   1200
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
         Enabled         =   0   'False
         Height          =   385
         Left            =   -65370
         TabIndex        =   118
         Top             =   3770
         Width           =   1200
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         Height          =   385
         Left            =   -67080
         TabIndex        =   117
         Top             =   3770
         Width           =   1200
      End
      Begin VB.CommandButton cmdSavePayments 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66240
         TabIndex        =   80
         Top             =   4525
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditPayments 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -68280
         TabIndex        =   79
         Top             =   4525
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelPayments 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64200
         TabIndex        =   81
         Top             =   4525
         Width           =   1215
      End
      Begin VB.CommandButton cmdPayType 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68985
         TabIndex        =   72
         Top             =   1885
         Width           =   315
      End
      Begin VB.CommandButton cmdAccType 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68985
         TabIndex        =   70
         Top             =   1465
         Width           =   315
      End
      Begin VB.TextBox txtBankPayRef 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65280
         TabIndex        =   78
         Top             =   2845
         Width           =   2415
      End
      Begin VB.TextBox txtAcName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65280
         TabIndex        =   77
         Top             =   2365
         Width           =   2415
      End
      Begin VB.TextBox txtAcNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65280
         MaxLength       =   8
         TabIndex        =   76
         Top             =   1885
         Width           =   2415
      End
      Begin VB.TextBox txtSortCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65280
         MaxLength       =   6
         TabIndex        =   75
         Top             =   1405
         Width           =   2415
      End
      Begin VB.CommandButton cmdAddNewSupplier 
         Caption         =   "&New"
         Height          =   345
         Left            =   4920
         TabIndex        =   24
         Top             =   4780
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelSupplier 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   8952
         TabIndex        =   27
         Top             =   4780
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditSupplier 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   6264
         TabIndex        =   25
         Top             =   4780
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteSupplier 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   10296
         TabIndex        =   28
         Top             =   4780
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveSupplier 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   7608
         TabIndex        =   26
         Top             =   4780
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseSupplier 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   11640
         TabIndex        =   29
         Top             =   4780
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Landlord Address:"
         Enabled         =   0   'False
         Height          =   3975
         Index           =   0
         Left            =   375
         TabIndex        =   36
         Top             =   465
         Width           =   4095
         Begin VB.TextBox txtSupplierAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   12
            Top             =   1239
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   2700
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   17
            Top             =   3090
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   11
            Top             =   906
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   1572
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   10
            Top             =   573
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   43
            Top             =   1920
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            Height          =   210
            Index           =   14
            Left            =   120
            TabIndex        =   42
            Top             =   3090
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   41
            Top             =   2700
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Web:"
            Height          =   210
            Index           =   13
            Left            =   120
            TabIndex        =   40
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   39
            Top             =   2310
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   210
            Index           =   9
            Left            =   120
            TabIndex        =   38
            Top             =   1575
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Landlord Alternative Address :"
         Enabled         =   0   'False
         Height          =   3735
         Left            =   7380
         TabIndex        =   33
         Top             =   585
         Width           =   4695
         Begin VB.TextBox txtSupplierOfficeAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   22
            Top             =   1950
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   20
            Top             =   1050
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   23
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   21
            Top             =   1530
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   19
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   34
            Top             =   2520
            Width           =   750
         End
      End
      Begin VB.CommandButton cmdCancelSupplierDetails 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   9960
         TabIndex        =   31
         Top             =   3580
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditSupplierDetails 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   7560
         TabIndex        =   32
         Top             =   4185
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveSupplierDetails 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   8760
         TabIndex        =   30
         Top             =   3580
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   92
         Top             =   580
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4577
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
         TabIndex        =   93
         Top             =   3340
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxContacts 
         Height          =   4035
         Left            =   -74760
         TabIndex        =   146
         Top             =   720
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   10
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
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Email"
         Height          =   195
         Index           =   29
         Left            =   -64080
         TabIndex        =   154
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         Height          =   195
         Index           =   28
         Left            =   -64800
         TabIndex        =   153
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Office Mobile"
         Height          =   195
         Index           =   27
         Left            =   -65880
         TabIndex        =   152
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office Email"
         Height          =   210
         Index           =   26
         Left            =   -67560
         TabIndex        =   151
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office Tel"
         Height          =   210
         Index           =   25
         Left            =   -69360
         TabIndex        =   150
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office Address"
         Height          =   210
         Index           =   24
         Left            =   -71760
         TabIndex        =   149
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   23
         Left            =   -73320
         TabIndex        =   148
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         Height          =   195
         Index           =   22
         Left            =   -74700
         TabIndex        =   147
         Top             =   480
         Width           =   1185
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   -62520
         TabIndex        =   138
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   -68520
         TabIndex        =   137
         Top             =   2880
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   -68520
         TabIndex        =   136
         Top             =   2400
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   -69120
         TabIndex        =   135
         Top             =   2400
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   -68520
         TabIndex        =   134
         Top             =   1920
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   -62520
         TabIndex        =   133
         Top             =   2880
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   -68520
         TabIndex        =   132
         Top             =   1560
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   -62520
         TabIndex        =   131
         Top             =   2400
         Visible         =   0   'False
         Width           =   435
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   -62520
         TabIndex        =   130
         Top             =   1920
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSForms.TextBox txtCodeVat 
         Height          =   285
         Left            =   -71655
         TabIndex        =   129
         Top             =   2365
         Width           =   615
         VariousPropertyBits=   746604569
         BorderStyle     =   1
         Size            =   "1085;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVatCode 
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   127
         Top             =   2365
         Width           =   375
         BackColor       =   14737632
         Size            =   "661;503"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         Height          =   195
         Index           =   0
         Left            =   -74280
         TabIndex        =   126
         Top             =   2365
         Width           =   690
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   18
         Left            =   -67830
         TabIndex        =   114
         Top             =   3150
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   17
         Left            =   -69000
         TabIndex        =   113
         Top             =   3145
         Width           =   945
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop No"
         Height          =   195
         Index           =   16
         Left            =   -69840
         TabIndex        =   112
         Top             =   3145
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No"
         Height          =   195
         Index           =   15
         Left            =   -70560
         TabIndex        =   111
         Top             =   3145
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
         TabIndex        =   110
         Top             =   3145
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   14
         Left            =   -71400
         TabIndex        =   109
         Top             =   3145
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
         TabIndex        =   108
         Top             =   3145
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
         TabIndex        =   107
         Top             =   3145
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   13
         Left            =   -72600
         TabIndex        =   106
         Top             =   3150
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   12
         Left            =   -73560
         TabIndex        =   105
         Top             =   3145
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   11
         Left            =   -74520
         TabIndex        =   104
         Top             =   3145
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   103
         Top             =   3145
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
         TabIndex        =   102
         Top             =   340
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
         TabIndex        =   101
         Top             =   340
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   5
         Left            =   -69840
         TabIndex        =   100
         Top             =   340
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   99
         Top             =   340
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   -74040
         TabIndex        =   98
         Top             =   340
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   -72720
         TabIndex        =   97
         Top             =   340
         Width           =   345
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   4
         Left            =   -71640
         TabIndex        =   96
         Top             =   340
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
         TabIndex        =   95
         Top             =   340
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
         TabIndex        =   94
         Top             =   340
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         Caption         =   "Bank Account Details:"
         Height          =   210
         Index           =   3
         Left            =   -67680
         TabIndex        =   85
         Top             =   1060
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         Caption         =   "Payment Details:"
         Height          =   210
         Index           =   22
         Left            =   -74400
         TabIndex        =   84
         Top             =   1060
         Width           =   1185
      End
      Begin MSForms.ComboBox cboAccType 
         Height          =   285
         Left            =   -72000
         TabIndex        =   69
         Top             =   1465
         Width           =   3015
         VariousPropertyBits=   746604569
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5318;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPayType 
         Height          =   285
         Left            =   -72000
         TabIndex        =   71
         Top             =   1885
         Width           =   3015
         VariousPropertyBits=   746604569
         MaxLength       =   10
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5318;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   -74280
         TabIndex        =   91
         Top             =   1885
         Width           =   1935
         Caption         =   "Payment Type"
         Size            =   "3413;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   -74280
         TabIndex        =   90
         Top             =   1465
         Width           =   2055
         Caption         =   "Account Type"
         Size            =   "3625;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         Height          =   2295
         Index           =   2
         Left            =   -74520
         Top             =   1165
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         Height          =   2295
         Index           =   0
         Left            =   -67800
         Top             =   1165
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Payment Ref:"
         Height          =   195
         Index           =   20
         Left            =   -67560
         TabIndex        =   89
         Top             =   2845
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name:"
         Height          =   195
         Index           =   19
         Left            =   -67560
         TabIndex        =   88
         Top             =   2365
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number:"
         Height          =   195
         Index           =   18
         Left            =   -67560
         TabIndex        =   87
         Top             =   1885
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Code:"
         Height          =   195
         Index           =   17
         Left            =   -67560
         TabIndex        =   86
         Top             =   1465
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Days"
         Height          =   210
         Left            =   -70800
         TabIndex        =   83
         Top             =   2845
         Width           =   735
      End
      Begin MSForms.TextBox txtPaymentTerms 
         Height          =   285
         Left            =   -72000
         TabIndex        =   74
         Top             =   2845
         Width           =   1095
         VariousPropertyBits=   746604569
         BorderStyle     =   1
         Size            =   "1931;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Terms:"
         Height          =   210
         Index           =   7
         Left            =   -74280
         TabIndex        =   82
         Top             =   2845
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   2295
         Index           =   3
         Left            =   -74520
         Top             =   1165
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   2295
         Index           =   1
         Left            =   -67800
         Top             =   1165
         Width           =   5295
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   115
         Top             =   340
         Width           =   12735
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   116
         Top             =   3100
         Width           =   12735
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   -74760
         TabIndex        =   155
         Top             =   480
         Width           =   12735
      End
   End
End
Attribute VB_Name = "frmLandLord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FORM_STATUS        As String
Private VAT_CODE           As String

Dim bSortingCol1           As Boolean
Dim bSortingCol2           As Boolean
Dim bSortingCol3           As Boolean
Dim szaSupplierBalance()   As String
Dim szaChoice()            As String
Dim bVatCodeLoaded         As Boolean
Dim szaAddresses()         As String

Public Sub ConfigGridMaintenanceHistory(ByVal rstMHistory_ As ADODB.Recordset)
   Dim iColumn As Integer
   Dim oColumn As ADODB.Field

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = rstMHistory_.Fields.count + 1

   For iColumn = 1 To 9
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn).Left - 240

   For iColumn = 11 To rstMHistory_.Fields.count
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

Public Sub LoadGridMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection)
   Dim rstMHistory_ As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY'), S.Value, " & _
                "H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.ReportedBy, " & _
                "H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent, " & _
                "H.MaintenanceType " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
           "WHERE H.AssignedTo = '" & txtSupplierID.text & "' " & _
               "AND S.Code = H.MaintenanceType " & _
               "AND S.PrimaryCode = 'MTYP' " & _
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

Private Sub cboPayType_Change()
   If txtSupplierName.text = "" Or _
      cboPayType.text = "" Or _
      cboPayType.text = "Cheque" Then Exit Sub

   If txtSortCode.text = "" Then
      cboPayType.text = ""
      MsgBox "Please update Landlord's Bank details.", vbCritical + vbOKOnly, "BACS setting"
      Exit Sub
   End If
End Sub

Private Sub cboSupplierType_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdAccType_Click()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim selType As String

   selType = IIf(cboAccType.text = "", "", cboAccType.text)
   frmSecondaryCode.PRIMARY_CODE_SHOW = "ACCT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1
   
   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'ACCT'"
   populateCombo adoConn, sSQLQuery, cboAccType
   cboAccType.text = selType

   adoConn.Close
   Set adoConn = Nothing
End Sub

'  Build up lessee's Account History
Private Sub SupplierAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "From tlbPayment " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaSupplierBalance(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type = 6 OR Type = 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalance(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBalance(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalance(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaSupplierBalance(1, i) = szaSupplierBalance(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaSupplierBalance(0, iIndex) = adoPayCr.Fields.Item("Cr").Value
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Sub

Private Sub cmdAddDiary_Click()
   If txtSupplierID.text = "" Then Exit Sub

   With frmMaintananceDairy
      .CallingForm = "S"          'Calling from lessee form
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

Private Sub cmdAddNewContacts_Click()
   Load frmContacts
   frmContacts.WHOS_CONTACT = "S"
   frmContacts.LOADING_MODE = "NEW"
   frmContacts.Show
   Me.Enabled = False
End Sub

Private Sub cmdCancelBank_Click()
   txtSortCode.Enabled = False
   txtAcName.Enabled = False
   txtAcNo.Enabled = False
   txtBankPayRef.Enabled = False

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub

Private Sub cmdCancelPayments_Click()
   cboAccType.ListIndex = Val(Label1(25).Caption)
   cboPayType.ListIndex = Val(Label1(26).Caption)
   lblVatCode(0).Caption = Label1(27).Caption
   txtCodeVat.text = Label1(28).Caption
   txtPaymentTerms.text = Label1(29).Caption
   txtSortCode.text = Label1(30).Caption
   txtAcNo.text = Label1(31).Caption
   txtAcName.text = Label1(32).Caption
   txtBankPayRef.text = Label1(33).Caption

   cmdAccType.Enabled = False
   cmdPayType.Enabled = False
   cboAccType.Enabled = False
   cboPayType.Enabled = False
   txtPaymentTerms.Enabled = False

   cmdSavePayments.Enabled = False
   cmdCancelPayments.Enabled = False
   cmdEditPayments.Enabled = True

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True

   txtSortCode.Enabled = False
   txtAcName.Enabled = False
   txtAcNo.Enabled = False
   txtBankPayRef.Enabled = False

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub

Private Sub cmdCancelSupplier_Click()
   If (FORM_STATUS = "Edit") Then
      FORM_STATUS = txtSupplierID.text
      SetControls DefaultMode
      txtSupplierID.text = FORM_STATUS

      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString
      loadValues FORM_STATUS, adoConn
      adoConn.Close
   Else
      SetControls DefaultMode
   End If
   cmdTaxList(0).Enabled = False

   cmdSaveSupplierDetails.Enabled = False
   cmdCancelSupplierDetails.Enabled = False
   cmdEditSupplierDetails.Enabled = True

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
   Me.Caption = "Landlord"
End Sub

Private Sub cmdCancelSupplierDetails_Click()
   cmdTaxList(0).Enabled = False

   cmdSaveSupplierDetails.Enabled = False
   cmdCancelSupplierDetails.Enabled = False
   cmdEditSupplierDetails.Enabled = True

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachmentInCombo cmbFiles, "Supplier", txtSupplierID.text
   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdCloseSupplier_Click()
   Unload Me
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtSupplierID.text, "Landlord"
   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdEditBank_Click()
   If txtSupplierID.text = "" Then Exit Sub

   txtSortCode.Enabled = True
   txtAcNo.Enabled = True
   txtAcName.Enabled = True
   txtBankPayRef.Enabled = True
   txtSortCode.SetFocus

   cmdAddNewSupplier.Enabled = False
   cmdEditSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
End Sub

Private Sub cmdEditDefaults_Click()
   If txtSupplierID.text = "" Then Exit Sub

   cmdTaxList(0).Enabled = True
   cmdTaxList(0).SetFocus

   cmdAddNewSupplier.Enabled = False
   cmdEditSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = False
End Sub

Private Sub cmdEditContacts_Click()
   If flxContacts.TextMatrix(flxContacts.row, 0) = "" Then Exit Sub

   Load frmContacts
   frmContacts.WHOS_CONTACT = "S"
   frmContacts.LOADING_MODE = "EDIT"
   frmContacts.LOADING_ID = flxContacts.TextMatrix(flxContacts.row, 0)
   frmContacts.Show
   Me.Enabled = False
   
End Sub

Private Sub cmdEditSupplier_Click()
   If txtSupplierID.text = "" Then Exit Sub
   SetControls EditMode
   
   FORM_STATUS = "Edit"
   cmdEditSupplierDetails_Click
End Sub

Private Sub cmdEditSupplierDetails_Click()
   If txtSupplierID.text = "" Then Exit Sub

   Frame1(0).Enabled = True
   Frame2.Enabled = True
   cmdSaveSupplierDetails.Enabled = True
   cmdCancelSupplierDetails.Enabled = True
   cmdEditSupplierDetails.Enabled = False
   txtSupplierAddressLine1.SetFocus

   cmdAddNewSupplier.Enabled = False
   cmdEditSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = False
   cmdCancelSupplier.Enabled = True
End Sub

Private Sub cmdGridUnitClose_Click(Index As Integer)
   fraList(0).Visible = False
End Sub

Private Sub cmdGridUnitClose2_Click(Index As Integer)
   picSupList.Visible = False
End Sub

Private Sub cmdNewMHistory_Click()
   If txtSupplierID.text = "" Then Exit Sub

   Load frmMaintenanceJob
   With frmMaintenanceJob
   'added by anol 23 Jun 2015 issue 566
      '.UpdateRow = gridMaintenanceHistory.row
      .CallingForm = "S"          'Calling from lessee form
      .RecordType = "J"
      .lblJobName.Caption = "Job Name"
      .Label1.Caption = "Job No."
      .txtRef.Enabled = True
   'modified by anol 23 Jun 2015 issue 566
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
      ShowMsgInTaskBar "File has been moved from original location."

   MousePointer = vbDefault
End Sub

Private Sub cmdPayType_Click()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim selType As String

   selType = IIf(cboPayType.text = "", "", cboPayType.text)
   frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'RAT';"
   populateCombo adoConn, sSQLQuery, cboPayType
   cboPayType.text = selType

   adoConn.Close
   Set adoConn = Nothing
End Sub
'
'Private Sub cmdPLC_Click()
'   LoadNC
'   txtFLX.text = "PLC"
'
'   txtSearch1(0).text = ""
'   txtSearch2(0).text = ""
'   txtSearch3(0).text = ""
'   fraList(0).Width = 3520
'   flxSupplier(0).Width = 3400
'   cmdGridUnitClose(0).Left = fraList(0).Width - cmdGridUnitClose(0).Width - 60
'   Shape4(0).Width = fraList(0).Width - 200
'
'   fraList(0).Left = tabSupplier.Left + txtPLControl.Left
'   fraList(0).Top = tabSupplier.Top + txtPLControl.Top + txtPLControl.Height + 10
'   fraList(0).Visible = True
'   fraList(0).ZOrder 0
'   txtSearch1(0).SetFocus
'End Sub

Private Sub cmdSaveBank_Click()
   If txtSortCode.text = "" Then
      MsgBox "Please enter the sort code.", vbExclamation + vbOKOnly, "Bank Details"
      txtSortCode.SetFocus
      Exit Sub
   End If
   If txtAcName.text = "" Then
      MsgBox "Please enter the account name.", vbExclamation + vbOKOnly, "Bank Details"
      txtAcName.SetFocus
      Exit Sub
   End If
   If txtAcNo.text = "" Then
      MsgBox "Please enter the account number.", vbExclamation + vbOKOnly, "Bank Details"
      txtAcNo.SetFocus
      Exit Sub
   End If
   If txtBankPayRef.text = "" Then
      MsgBox "Please enter the payment reference.", vbExclamation + vbOKOnly, "Bank Details"
      txtBankPayRef.SetFocus
      Exit Sub
   End If

   FORM_STATUS = "DetEdit"
   cmdSaveSupplier_Click

   txtSortCode.Enabled = False
   txtAcName.Enabled = False
   txtAcNo.Enabled = False
   txtBankPayRef.Enabled = False

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub
'
'Private Sub cmdSaveDefaults_Click()
'   FORM_STATUS = "DetEdit"
'
'   cmdSaveSupplier_Click
'
'   cmdTaxList(0).Enabled = False
'   cmdNC.Enabled = False
'
'   cmdAddNewSupplier.Enabled = True
'   cmdEditSupplier.Enabled = True
'   cmdCancelSupplier.Enabled = False
'   cmdCloseSupplier.Enabled = True
'   cmdEditDefaults.Enabled = True
'   cmdCancelDefaults.Enabled = False
'   cmdSaveDefaults.Enabled = False
'End Sub

Private Sub cmdSavePayments_Click()
   If cboPayType.text = "BACS" And txtSortCode.text = "" Then
      MsgBox "Please enter the sort code.", vbExclamation + vbOKOnly, "Bank Details"
      txtSortCode.SetFocus
      Exit Sub
   End If
   If cboPayType.text = "BACS" And txtAcName.text = "" Then
      MsgBox "Please enter the account name.", vbExclamation + vbOKOnly, "Bank Details"
      txtAcName.SetFocus
      Exit Sub
   End If
   If cboPayType.text = "BACS" And txtAcNo.text = "" Then
      MsgBox "Please enter the account number.", vbExclamation + vbOKOnly, "Bank Details"
      txtAcNo.SetFocus
      Exit Sub
   End If
   If cboPayType.text = "BACS" And txtBankPayRef.text = "" Then
      MsgBox "Please enter the payment reference.", vbExclamation + vbOKOnly, "Bank Details"
      txtBankPayRef.SetFocus
      Exit Sub
   End If
'    'Resolved by BOSL
'    '0000446: Error creating new supplier record
'    'Description While entering new supplier records Austin Chambers
'    'received this error. The error occurred when Dharshy hit the save button. See Screenshot 1- Error saving supplier record.
'    'Modified by Anol 04 Aug 2014
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString
'   Dim rsSuppcheck As New ADODB.Recordset
'   rsSuppcheck.Open "SELECT * FROM Supplier WHERE SupplierID = '" & txtSupplierID.text & "'", adoConn, adOpenDynamic, adLockOptimistic
'   If rsSuppcheck.RecordCount > 1 Then
        FORM_STATUS = "DetEdit"
'   Else
'        FORM_STATUS = "New"
'   End If
'   rsSuppcheck.Close
'   Set rsSuppcheck = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
   
   cmdSaveSupplier_Click

   cmdAccType.Enabled = False
   cmdPayType.Enabled = False
   cboAccType.Enabled = False
   cboPayType.Enabled = False
   txtPaymentTerms.Enabled = False

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True

   cmdEditPayments.Enabled = True
   cmdCancelPayments.Enabled = False
   cmdSavePayments.Enabled = False
   
   txtSortCode.Enabled = False
   txtAcName.Enabled = False
   txtAcNo.Enabled = False
   txtBankPayRef.Enabled = False

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub

Private Sub cmdSaveSupplier_Click()
   If txtSupplierName.text = "" Then
       ShowMsgInTaskBar "Please enter a Landlord Name to continue.", , "N"
       txtSupplierName.text = ""
       txtSupplierName.SetFocus
       Exit Sub
   End If
   If txtSupplierID.text = "" Then
       ShowMsgInTaskBar "Please enter a Landlord ID to continue.", , "N"
       txtSupplierID.text = ""
       txtSupplierID.SetFocus
       Exit Sub
   End If
   If Len(txtSortCode.text) > 0 And Len(txtSortCode.text) < 6 Then
      ShowMsgInTaskBar "Please enter six digit bank sort code to continue.", , "N"
      txtSortCode.SetFocus
      Exit Sub
   End If
   If Len(txtAcNo.text) > 0 And Len(txtAcNo.text) < 8 Then
      ShowMsgInTaskBar "Please enter eight digit bank account number to continue.", , "N"
      txtAcNo.SetFocus
      Exit Sub
   End If
   If UCase(cboSupplierType.text) <> UCase("Landlord") Then
         ShowMsgInTaskBar "Please select 'Landlord' as supplier type to continue.", , "N"
        Exit Sub
   End If
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   '-----------------------------------------Add the record to DB
   Dim rstMHistory_ As New ADODB.Recordset
   Dim rstID As New ADODB.Recordset
   Dim sSQLQuery_ As String, sSQLDelete As String, sSQLFilter As String, iRowIndex As Integer
   Dim lTableID As Long

   sSQLFilter = ""

   If (FORM_STATUS = "New") Then
      sSQLFilter = ""
   Else
      sSQLFilter = "WHERE SupplierID = '" & txtSupplierID.text & "'"
   End If

   sSQLQuery_ = "SELECT * FROM Supplier " & sSQLFilter

   rstMHistory_.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic

   If (FORM_STATUS = "New") Then
      rstMHistory_.AddNew
      rstMHistory_!SupplierID = txtSupplierID.text
   End If

   rstMHistory_!SupplierName = txtSupplierName.text
   rstMHistory_!VATReg = txtTaxVatNumber.text
   rstMHistory_!VatCode = lblVatCode(0).Caption
   rstMHistory_!AcBalance = IIf(txtSupplierACBal.text = "", "0.00", Format(txtSupplierACBal.text, "0.00"))
   rstMHistory_!SupplierAddressLine1 = txtSupplierAddressLine1.text
   rstMHistory_!SupplierAddressLine2 = txtSupplierAddressLine2.text
   rstMHistory_!SupplierAddressLine3 = txtSupplierAddressLine3.text
   rstMHistory_!SupplierAddressLine4 = txtSupplierAddressLine4.text
   rstMHistory_!SupplierPostCode = txtSupplierPostCode.text
   rstMHistory_!SupplierOfficeEmail = txtSupplierOfficeEmail.text
   rstMHistory_!SupplierPersonalEmail = txtSupplierPersonalEmail.text
   rstMHistory_!SupplierHomeTel = txtSupplierHomeTel.text
   rstMHistory_!SupplierMobile = txtSupplierMobile.text
   rstMHistory_!SupplierOfficeAddressLine1 = txtSupplierOfficeAddressLine1.text
   rstMHistory_!SupplierOfficeAddressLine2 = txtSupplierOfficeAddressLine2.text
   rstMHistory_!SupplierOfficeAddressLine3 = txtSupplierOfficeAddressLine3.text
   rstMHistory_!SupplierOfficeAddressLine4 = txtSupplierOfficeAddressLine4.text
   rstMHistory_!SupplierOfficePostCode = txtSupplierOfficePostCode.text
   rstMHistory_!SupplierOfficeTel = txtSupplierOfficeTel.text
   rstMHistory_!SupplierMemo = txtNote.text
   rstMHistory_!SupplierType = IIf(cboSupplierType.text = "", "", cboSupplierType.Value)
   rstMHistory_!CreditLimit = IIf(txtCreditLimit.text = "", 0, txtCreditLimit.text)
   If cboAccType.text = "" Or cboAccType.ListCount = 0 Then
      rstMHistory_!AccountType = ""
   Else
      rstMHistory_!AccountType = cboAccType.Column(0)
   End If
   'Fixed by anol 02 June 2015
   If cboPayType.text = "" Then
        rstMHistory_!PaymentType = ""
   Else
        rstMHistory_!PaymentType = cboPayType.Column(0)
   End If
   rstMHistory_!PaymentTerms = IIf(txtPaymentTerms.text = "", 0, txtPaymentTerms.text)
   rstMHistory_!SortCode = IIf(txtSortCode.text = "", "", txtSortCode.text)
   rstMHistory_!AcNo = IIf(txtAcNo.text = "", "", txtAcNo.text)
   rstMHistory_!AcName = IIf(txtAcName.text = "", "", txtAcName.text)
   rstMHistory_!BPR = IIf(txtBankPayRef.text = "", Null, txtBankPayRef.text)
   rstMHistory_!Type = "LLORD"

   rstMHistory_.Update

   rstMHistory_.Close
   Set rstMHistory_ = Nothing
  

   ShowMsgInTaskBar "Landlord entry has been saved successfully."

   SetControls DefaultMode
   FORM_STATUS = "DetEdit"
   Me.Caption = "Landlord"

   cmdAccType.Enabled = False
   cmdPayType.Enabled = False
   cboAccType.Enabled = False
   cboPayType.Enabled = False
   txtPaymentTerms.Enabled = False
    'Resolved by BOSL
    'Issue 465
    'Modified by Anol 02 Sep 2014
    cboSupplierType.Enabled = False
    cmdSupplierType(0).Enabled = False
    'End of modification
   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True

   cmdEditPayments.Enabled = True
   cmdCancelPayments.Enabled = False
   cmdSavePayments.Enabled = False
   'I am commenting this for Landlord form
'   'Added by anol 13 May 2015
'   'Lanlord table needs to be updated from Supplier table
'   'Should be alwys syncronized with supplier table where supplier type = 'LL'
'   'So I am deleting all record from landlord table and transferring all where supplier type = 'LL' from supplier table
'   adoConn.Execute "DELETE FROM Landlord"
'   Dim rsSupplier As New ADODB.Recordset
'   Dim rsLandlord As New ADODB.Recordset
'   rsSupplier.Open "Select * from Supplier where SupplierType='LL'", adoConn, adOpenStatic, adLockReadOnly
'   rsLandlord.Open "Select * from Landlord", adoConn, adOpenDynamic, adLockOptimistic
'   While Not rsSupplier.EOF
'           rsLandlord.AddNew
'           rsLandlord!landLordID = "L-" & rsSupplier!SupplierID
'           rsLandlord!LandlordName = rsSupplier!SupplierName
'           rsLandlord!LandlordAddressLine1 = rsSupplier!SupplierAddressLine1
'           rsLandlord!LandlordAddressLine2 = rsSupplier!SupplierAddressLine2
'           rsLandlord!LandlordAddressLine3 = rsSupplier!SupplierAddressLine3
'           rsLandlord!LandlordAddressLine4 = rsSupplier!SupplierAddressLine4
'           rsLandlord!LandlordPostCode = rsSupplier!SupplierPostCode
'           rsLandlord!LandlordOfficeEmail = rsSupplier!SupplierOfficeEmail
'           rsLandlord!LandlordPersonalEmail = rsSupplier!SupplierPersonalEmail
'           rsLandlord!LandlordHomeTel = rsSupplier!SupplierHomeTel
'           rsLandlord!LandlordMobile = rsSupplier!SupplierMobile
'           rsLandlord!LandlordOfficeAddressLine1 = rsSupplier!SupplierOfficeAddressLine1
'           rsLandlord!LandlordOfficeAddressLine2 = rsSupplier!SupplierOfficeAddressLine2
'           rsLandlord!LandlordOfficeAddressLine3 = rsSupplier!SupplierOfficeAddressLine3
'           rsLandlord!LandlordOfficeAddressLine4 = rsSupplier!SupplierOfficeAddressLine4
'           rsLandlord!LandlordOfficePostCode = rsSupplier!SupplierOfficePostCode
'           rsLandlord!LandlordOfficeTel = rsSupplier!SupplierOfficeTel
'           rsLandlord!LandlordMemo = rsSupplier!SupplierMemo
'           rsLandlord!LandLordSageSuppAC = rsSupplier!SageSuppAC
'           rsLandlord!VATReg = rsSupplier!VATReg
'           rsLandlord!AcBalance = rsSupplier!AcBalance
'           rsLandlord!BacsRef = rsSupplier!BacsRef
'           rsLandlord.Update
'        rsSupplier.MoveNext
'   Wend
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function IsSupplierExist_(ByRef SupplierID As String, adoConn As ADODB.Connection) As Boolean
   Dim rstRst     As New ADODB.Recordset
   Dim szSQL      As String
   Dim szID       As String
   Dim i          As Integer
   Dim bFlag      As Boolean

   szSQL = "SELECT SupplierID FROM Supplier WHERE SupplierID = '" & SupplierID & "' AND TYPE = 'SUPPLIER';"

   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If (rstRst.RecordCount > 0) Then
      IsSupplierExist_ = True
      rstRst.Close

      i = 1
      bFlag = False

      While Not bFlag
         If Len(SupplierID) + Len(CStr(i)) > 10 Then
            szID = Left(SupplierID, 10 - Len(CStr(i)))
            szID = szID + CStr(i)
         Else
            szID = SupplierID + CStr(i)
         End If
         

         szSQL = "SELECT SupplierID FROM Supplier WHERE SupplierID = '" & szID & "' AND TYPE = 'SUPPLIER';"
         rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         If rstRst.RecordCount = 0 Then
            rstRst.Close
            SupplierID = szID
            Set rstRst = Nothing
            Exit Function
         End If
         rstRst.Close

         i = i + 1
      Wend

      Exit Function
   Else
      IsSupplierExist_ = False
      rstRst.Close
      Exit Function
   End If
End Function
'
'Private Function IsSupplierExist(ByRef SupplierID As String, adoConn As ADODB.Connection) As Boolean
'   Dim adoAllID   As New ADODB.Recordset
'   Dim szSQL      As String
'   Dim szId       As String
'   Dim i          As Integer
'   Dim bFlag      As Boolean
'
'   szSQL = "SELECT SupplierID AS ID FROM Supplier"
'   szSQL = szSQL & " UNION "
'   szSQL = szSQL & "SELECT ClientID AS ID FROM Client"
'   szSQL = szSQL & " UNION "
'   szSQL = szSQL & "SELECT SageAccountNumber AS ID FROM Tenants"
'   szSQL = szSQL & " UNION "
'   szSQL = szSQL & "SELECT AgentID AS ID FROM Agent"
''Debug.Print szSQL
'   adoAllID.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   adoAllID.Find "ID = '" & SupplierID & "' "
'
'   If Not adoAllID.EOF Then
'      IsSupplierExist = True
'
'      i = 1
'      bFlag = False
'
'      While Not bFlag
'         If Len(SupplierID) + Len(CStr(i)) > 10 Then
'            szId = Left(SupplierID, 10 - Len(CStr(i)))
'            szId = szId + CStr(i)
'         Else
'            szId = SupplierID + CStr(i)
'         End If
'
'         adoAllID.MoveFirst
'         adoAllID.Find "ID = '" & szId & "' "
'
'         If adoAllID.EOF Then
'            adoAllID.Close
'            SupplierID = szId
'            Set adoAllID = Nothing
'            Exit Function
'         End If
'
'         i = i + 1
'      Wend
'   Else
'      IsSupplierExist = False
'      adoAllID.Close
'      Set adoAllID = Nothing
'   End If
'End Function

Private Sub cmdSaveSupplierDetails_Click()
   FORM_STATUS = "DetEdit"
   cmdSaveSupplier_Click
   
   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
End Sub

Private Sub cmdSupplier_Click()
   Call PrepareList

   picSupList.Top = txtSupplierID.Top + txtSupplierID.Height + 5
   picSupList.Left = txtSupplierID.Left + 5
   picSupList.Visible = True
   picSupList.ZOrder 0
   
   txtSupplierID.Locked = True
   'added by anol 08 Apr 2015
   txtSupplierSearchID.SetFocus
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxSupplierList
   LoadAllSupplierFlxGrd
   UpdateBalance
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplierList.Rows - 1
      For j = 0 To UBound(szaSupplierBalance, 2) - 1
         If flxSupplierList.TextMatrix(i, 1) = szaSupplierBalance(0, j) Then
            flxSupplierList.TextMatrix(i, 4) = Format(szaSupplierBalance(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalance, 2) Then flxSupplierList.TextMatrix(i, 4) = "0.00"
   Next i
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 5
   conFlxGrid.Clear
   szHeader$ = "|<ManagingAgentID|<ManagingAgentName|<ManagingAgentPostCode|>AccBalance"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 0          'Solid column
   conFlxGrid.ColWidth(1) = 1000       'Landlord ID
   conFlxGrid.ColWidth(2) = 3000       'Landlord Name
   conFlxGrid.ColWidth(3) = 0          'Post Code
   conFlxGrid.ColWidth(4) = 1100       'Account Balance
   conFlxGrid.Rows = 2
'
   'conFlxGrid.RowHeightMin = 300
   conFlxGrid.RowHeight(0) = 0
End Sub

Private Sub LoadAllSupplierFlxGrd()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName, SupplierPostCode " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'LLORD'" & _
           "ORDER BY SupplierName;"

   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRst.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstRst.EOF
      flxSupplierList.TextMatrix(iRow, 1) = rstRst!SupplierID
      flxSupplierList.TextMatrix(iRow, 2) = rstRst!SupplierName
      flxSupplierList.TextMatrix(iRow, 3) = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
      rstRst.MoveNext
      If Not rstRst.EOF Then flxSupplierList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdSupplierType_Click(Index As Integer)
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim SelSupplierCode As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "LLORD"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'LLORD' AND CODE='LL'"
               
   SelSupplierCode = IIf(cboSupplierType.text = "", "", cboSupplierType.Value)
   populateCombo adoConn, sSQLQuery, cboSupplierType
   cboSupplierType.Value = SelSupplierCode

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdTaxList_Click(Index As Integer)
   If Not bVatCodeLoaded Then LoadVAT

   txtSearch1(0).text = ""
   txtSearch2(0).text = ""
   txtSearch3(0).text = ""
   fraList(0).Width = 3520
   flxSupplier(0).Width = 3400
   cmdGridUnitClose(0).Left = fraList(0).Width - cmdGridUnitClose(0).Width - 60
   Shape4(0).Width = fraList(0).Width - 200

   If Index = 0 Then
      fraList(0).Left = cmdTaxList(Index).Left + tabSupplier.Left
      fraList(0).Top = cmdTaxList(Index).Top + cmdTaxList(Index).Height + tabSupplier.Top
   Else
      fraList(0).Left = tabSupplier.Left + cmdTaxList(Index).Left - 400
      fraList(0).Top = tabSupplier.Top + cmdTaxList(Index).Top + cmdTaxList(Index).Height + 200
   End If

   fraList(0).Visible = True
   fraList(0).ZOrder 0
   txtSearch1(0).SetFocus
End Sub

Private Sub LoadVAT()
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 2000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "RATE"
   
   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 1900
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1(0).Width = 900
   txtSearch1(0).Left = 40
   
   txtSearch2(0).Width = 1900
   txtSearch2(0).Left = txtSearch1(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch3(0).Visible = False
   
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1(0).Caption = "RATE"
   lblSearch2(0).Visible = False
   lblSearch3(0).Visible = False
   lblSearch4(0).Visible = False
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString
'
   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2
      flxSupplier(0).RowHeight(0) = 0

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      flxSupplier(0).TextMatrix(0, 0) = "VAT Code"
      flxSupplier(0).TextMatrix(0, 1) = "VAT Rate"

      rRow = 1
      flxSupplier(0).AddItem ""
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
   bVatCodeLoaded = True
End Sub

Private Sub flxContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If flxContacts.MouseCol <> 4 Then
      flxContacts.ToolTipText = flxContacts.TextMatrix(flxContacts.MouseRow, flxContacts.MouseCol)
      txtAddress.Visible = False
   Else
      If flxContacts.MouseRow = 0 Then Exit Sub
      If flxContacts.TextMatrix(flxContacts.MouseRow, 0) = "" Then Exit Sub

      txtAddress.text = flxContacts.TextMatrix(flxContacts.MouseRow, flxContacts.MouseCol)
      
      txtAddress.text = szaAddresses(flxContacts.MouseRow - 1, 0) & vbCrLf & _
                        IIf(szaAddresses(flxContacts.MouseRow - 1, 1) <> "", szaAddresses(flxContacts.MouseRow - 1, 1) & vbCrLf, "") & _
                        IIf(szaAddresses(flxContacts.MouseRow - 1, 2) <> "", szaAddresses(flxContacts.MouseRow - 1, 2) & vbCrLf, "") & _
                        IIf(szaAddresses(flxContacts.MouseRow - 1, 3) <> "", szaAddresses(flxContacts.MouseRow - 1, 3) & vbCrLf, "") & _
                        IIf(szaAddresses(flxContacts.MouseRow - 1, 4) <> "", szaAddresses(flxContacts.MouseRow - 1, 4), "")
      txtAddress.Top = flxContacts.Top + (flxContacts.MouseRow * 240)
      txtAddress.Left = Label11(24).Left
      txtAddress.Visible = True
   End If
End Sub

Private Sub flxSupplier_Click(Index As Integer)
   lblVatCode(0).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
   txtCodeVat.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)

   fraList(0).Visible = False
End Sub
'
'Private Sub LoadNC()
'   flxSupplier(0).ColWidth(0) = 1000
'   flxSupplier(0).ColWidth(1) = 2000
'   flxSupplier(0).TextMatrix(0, 0) = "CODE"
'   flxSupplier(0).TextMatrix(0, 1) = "NAME"
'
'   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
'   lblSearch0(0).Width = 900
'   lblSearch0(0).Left = 50
'   lblSearch1(0).Width = 1900
'   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
'
'   txtSearch1(0).Width = 900
'   txtSearch1(0).Left = 40
'
'   txtSearch2(0).Width = 1900
'   txtSearch2(0).Left = txtSearch1(0).Left + flxSupplier(0).ColWidth(0)
'
'   txtSearch3(0).Visible = False
'
'   '~~~Added By Senthuran~~~ Code to configuer Label Caption
'   lblSearch0(0).Caption = "CODE"
'   lblSearch1(0).Caption = "NAME"
'   lblSearch2(0).Visible = False
'   lblSearch3(0).Visible = False
'   lblSearch4(0).Visible = False
'
'   flxSupplier(0).RowHeight(0) = 0
'
'   Dim rRow As Integer
'   Dim Conn2 As New ADODB.Connection
'
'   Dim szSQL As String
'   Dim rstRec As New ADODB.Recordset
'
''   Reset screen to show all the units in cboUnits.
''   Set the RDO Connections to the dataset
'   Conn2.Open getConnectionString
'
'   szSQL = "SELECT CODE, NAME " & _
'           "FROM SpareTable1;"
'   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
'
'   If Not rstRec.EOF Then
'      flxSupplier(0).Clear
'      flxSupplier(0).Cols = 2
'      flxSupplier(0).Rows = 2
'
'      rstRec.MoveFirst
'      flxSupplier(0).ColAlignment(1) = vbRightJustify
'
'      flxSupplier(0).TextMatrix(0, 0) = "Code"
'      flxSupplier(0).TextMatrix(0, 1) = "Name"
'
'      rRow = 2
'      flxSupplier(0).AddItem ""
'      While Not rstRec.EOF
'         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!Code
'         flxSupplier(0).TextMatrix(rRow, 1) = rstRec!Name
'         rstRec.MoveNext
'         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
'         rRow = rRow + 1
'      Wend
'   End If
'
'   flxSupplier(0).Sort = 1
'   rstRec.Close
'   Conn2.Close
'
'   Set rstRec = Nothing
'   Set Conn2 = Nothing
'End Sub

Private Sub cmdUnitMemoCancel_Click()
   cmdUnitMemoEdit.Enabled = True
   Frame17.Enabled = True
End Sub

Private Sub cmdUnitMemoEdit_Click()
   If txtSupplierID.text = "" Then Exit Sub

   txtNote.Enabled = True
   cmdUnitMemoCancel.Enabled = True
   cmdUnitMemoSave.Enabled = True
   cmdUnitMemoEdit.Enabled = False
End Sub

Private Sub cmdUnitMemoSave_Click()
   If (SaveMemo("Supplier", "SupplierMemo", txtSupplierID.text, "SupplierID", txtNote)) Then
      txtNote.Enabled = False
      cmdUnitMemoEdit.Enabled = True
      cmdUnitMemoSave.Enabled = False
      cmdUnitMemoCancel.Enabled = False
      ShowMsgInTaskBar "Memo has been saved successfully."
   Else
      ShowMsgInTaskBar "Data could not be saved, Please Contact with administrator", , "N"
   End If
End Sub
'
'Private Sub cmdNC_Click()
'   LoadNC
'   txtFLX.text = "NC"
'
'   txtSearch1(0).text = ""
'   txtSearch2(0).text = ""
'   txtSearch3(0).text = ""
'   fraList(0).Width = 3520
'   flxSupplier(0).Width = 3400
'   cmdGridUnitClose(0).Left = fraList(0).Width - cmdGridUnitClose(0).Width - 60
'   Shape4(0).Width = fraList(0).Width - 200
'   fraList(0).Left = tabSupplier.Left + txtNominalCode.Left
'   fraList(0).Top = tabSupplier.Top + txtNominalCode.Top + txtNominalCode.Height + 10
'   fraList(0).Visible = True
'   fraList(0).ZOrder 0
'   txtSearch1(0).SetFocus
'End Sub

Private Sub cmdEditPayments_Click()
   If txtSupplierID.text = "" Then Exit Sub

   Label1(25).Caption = cboAccType.ListIndex
   Label1(26).Caption = cboPayType.ListIndex
   Label1(27).Caption = lblVatCode(0).Caption
   Label1(28).Caption = txtCodeVat.text
   Label1(29).Caption = txtPaymentTerms.text
   Label1(30).Caption = txtSortCode.text
   Label1(31).Caption = txtAcNo.text
   Label1(32).Caption = txtAcName.text
   Label1(33).Caption = txtBankPayRef.text
   
   cmdTaxList(0).Enabled = True
   cmdAccType.Enabled = True
   cmdPayType.Enabled = True
   cboAccType.Enabled = True
   cboPayType.Enabled = True
   txtPaymentTerms.Enabled = True

   cmdSavePayments.Enabled = True
   cmdCancelPayments.Enabled = True
   cmdEditPayments.Enabled = False
   cboAccType.SetFocus

   cmdAddNewSupplier.Enabled = False
   cmdEditSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False

   txtSortCode.Enabled = True
   txtAcNo.Enabled = True
   txtAcName.Enabled = True
   txtBankPayRef.Enabled = True
   txtSortCode.SetFocus

'   cmdSaveBank.Enabled = True
'   cmdCancelBank.Enabled = True
'   cmdEditDefaults.Enabled = False

   cmdAddNewSupplier.Enabled = False
   cmdEditSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
End Sub

Private Sub flxACHistory_Click()
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
      szSQL = "SELECT S.*, F.FundName " & _
              "FROM tlbPayment AS P, tblPurInv AS I, tblPurInvSRec AS S, Fund AS F " & _
              "WHERE P.PI = I.MY_ID AND " & _
                  "I.MY_ID = S.ParentID AND S.DEPT_ID = F.FundID AND " & _
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
            .TextMatrix(iRow, 7) = adoRst.Fields.Item("FundName").Value
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
      szSQL = "SELECT P.ExtRef, S.UNIT_ID, F.FundName, S.Amount AS PaymentAmount, P.NominalCode " & _
              "FROM tlbPayment AS P, tlbPaymentSplit AS S, Fund AS F " & _
              "WHERE P.TransactionID = S.PayHeader AND P.FundID = F.FundID AND " & _
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
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("UNIT_ID").Value), "", adoRst.Fields.Item("UNIT_ID").Value)
            .TextMatrix(iRow, 7) = adoRst.Fields.Item("FundName").Value
            .TextMatrix(iRow, 8) = flxACHistory.TextMatrix(flxACHistory.row, 5)
            .TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("PaymentAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("PaymentAmount").Value, "0.00")
            adoRst.MoveNext
            If Not adoRst.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
      adoRst.Close
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
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("UnitID").Value), "", adoRst.Fields.Item("UnitID").Value)
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
      
      adoRst.Close
   End If
   

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxACHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If flxACHistory.TextMatrix(flxACHistory.MouseRow, 1) = "" Then Exit Sub

   flxACHistory.ToolTipText = flxACHistory.TextMatrix(flxACHistory.MouseRow, flxACHistory.MouseCol)
End Sub

Private Sub flxACHistorySplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxACHistorySplit.ToolTipText = flxACHistorySplit.TextMatrix(flxACHistorySplit.MouseRow, flxACHistorySplit.MouseCol)
End Sub

Private Sub flxSupplierList_Click()
   If flxSupplierList.TextMatrix(flxSupplierList.row, 1) = "" Then Exit Sub

   txtSupplierID.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)

   Dim adoConn As New ADODB.Connection
   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   loadValues txtSupplierID.text, adoConn
   LoadGridMaintenanceHistory adoConn

   adoConn.Close
   Set adoConn = Nothing

   Me.Caption = "Landlord: " & txtSupplierName.text

   picSupList.Visible = False
   cmdEditSupplierDetails.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdEditPayments.Enabled = True
   Frame17.Enabled = True
   tabSupplier.Enabled = True

'   Me.Height = 5850
'   tabSupplier.Top = 0
End Sub

Private Sub loadValues(ID As String, adoConn As ADODB.Connection)
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   Dim sSQLQuery_ As String, sFilter As String

   MousePointer = vbHourglass

   sSQLQuery_ = "SELECT S.SupplierID, S.SupplierName, S.SupplierAddressLine1, S.SupplierAddressLine2, " & _
                  "S.SupplierAddressLine3,S.SupplierAddressLine4, S.SupplierPostCode, " & _
                  "S.SupplierOfficeEmail, S.SupplierPersonalEmail, V.VAT_RATE, " & _
                  "S.SupplierHomeTel, S.SupplierMobile, S.SupplierOfficeAddressLine1, " & _
                  "S.SupplierOfficeAddressLine2, S.SupplierOfficeAddressLine3,S.SupplierOfficeAddressLine4, " & _
                  "S.SupplierOfficePostCode, S.PLControl, S.PLControlName, " & _
                  "S.SupplierOfficeTel, S.SupplierMemo, S.VATReg,S.VATCode, " & _
                  "S.BacsRef, S.SupplierType, S.creditlimit, S.nominalcode, " & _
                  "S.AccountType, S.PaymentType, S.PaymentTerms, S.SortCode, S.AcNo, S.AcName, S.BPR, " & _
                  "N.Name AS NN " & _
                "FROM (Supplier AS S LEFT OUTER JOIN NominalLedger AS N " & _
                  "ON S.NominalCode = N.Code) LEFT OUTER JOIN tlbVatCode AS V " & _
                  "ON S.VATCode = V.VAT_CODE " & _
                "WHERE S.SupplierID = '" & ID & "';"
'Debug.Print sSQLQuery_
   rstRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   txtSupplierName.text = rstRst!SupplierName
   txtSupplierACBal.text = Format(GetSupplierBalance(ID), "0.00")
   txtTaxVatNumber.text = IIf(IsNull(rstRst!VATReg), "", rstRst!VATReg)
   txtSupplierAddressLine1.text = IIf(IsNull(rstRst!SupplierAddressLine1), "", rstRst!SupplierAddressLine1)
   txtSupplierAddressLine2.text = IIf(IsNull(rstRst!SupplierAddressLine2), "", rstRst!SupplierAddressLine2)
   txtSupplierAddressLine3.text = IIf(IsNull(rstRst!SupplierAddressLine3), "", rstRst!SupplierAddressLine3)
   txtSupplierAddressLine4.text = IIf(IsNull(rstRst!SupplierAddressLine4), "", rstRst!SupplierAddressLine4)
   txtSupplierPostCode.text = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
   txtSupplierHomeTel.text = IIf(IsNull(rstRst!SupplierHomeTel), "", rstRst!SupplierHomeTel)
   txtSupplierMobile.text = IIf(IsNull(rstRst!SupplierMobile), "", rstRst!SupplierMobile)
   txtSupplierOfficeTel.text = IIf(IsNull(rstRst!SupplierOfficeTel), "", rstRst!SupplierOfficeTel)
   txtSupplierOfficeEmail.text = IIf(IsNull(rstRst!SupplierOfficeEmail), "", rstRst!SupplierOfficeEmail)
   txtSupplierPersonalEmail.text = IIf(IsNull(rstRst!SupplierPersonalEmail), "", rstRst!SupplierPersonalEmail)
   txtSupplierOfficeAddressLine1.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine1), "", rstRst!SupplierOfficeAddressLine1)
   txtSupplierOfficeAddressLine2.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine2), "", rstRst!SupplierOfficeAddressLine2)
   txtSupplierOfficeAddressLine3.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine3), "", rstRst!SupplierOfficeAddressLine3)
   txtSupplierOfficeAddressLine4.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine4), "", rstRst!SupplierOfficeAddressLine4)
   txtSupplierOfficePostCode.text = IIf(IsNull(rstRst!SupplierOfficePostCode), "", rstRst!SupplierOfficePostCode)
   cboSupplierType.Value = IIf(IsNull(rstRst!SupplierType), "", rstRst!SupplierType)
   txtCreditLimit.text = IIf(IsNull(rstRst!CreditLimit), "0.00", Format(rstRst!CreditLimit, "0.00"))
   lblVatCode(0).Caption = IIf(IsNull(rstRst!VatCode), "", rstRst!VatCode)
   cboAccType.Value = IIf(IsNull(rstRst!AccountType), "", rstRst!AccountType)
   txtPaymentTerms.text = IIf(IsNull(rstRst!PaymentTerms), 0, rstRst!PaymentTerms)
   txtSortCode.text = IIf(IsNull(rstRst!SortCode), "", rstRst!SortCode)
   txtAcNo.text = IIf(IsNull(rstRst!AcNo), "", rstRst!AcNo)
   txtAcName.text = IIf(IsNull(rstRst!AcName), "", rstRst!AcName)
   txtBankPayRef.text = IIf(IsNull(rstRst!BPR), "", rstRst!BPR)
   cboPayType.Value = IIf(IsNull(rstRst!PaymentType), "", rstRst!PaymentType)
   txtCodeVat.text = IIf(IsNull(rstRst!VAT_RATE), "", rstRst!VAT_RATE)

   RetrieveMemo "Supplier", "SupplierMemo", txtSupplierID.text, "SupplierID", txtNote

   LoadAttachmentFiles cmbFiles, txtSupplierID.text, "Supplier"

   LoadFlxACHistory adoConn
   LoadFlxContact adoConn

   MousePointer = vbDefault

   cmdAddNewSupplier.Enabled = True
   cmdEditSupplier.Enabled = True
   cmdSaveSupplier.Enabled = False
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
         
   cmdSaveSupplierDetails.Enabled = False
   cmdCancelSupplierDetails.Enabled = False
   cmdEditSupplierDetails.Enabled = True
End Sub

Private Function GetSupplierBalance(szSuppID As String) As Currency
   Dim j As Integer

   For j = 0 To UBound(szaSupplierBalance, 2) - 1
      If szSuppID = szaSupplierBalance(0, j) Then
         GetSupplierBalance = Format(szaSupplierBalance(1, j), "0.00")
         Exit For
      End If
   Next j
   If j = UBound(szaSupplierBalance, 2) Then GetSupplierBalance = 0
End Function

Private Sub Form_Load()
  'Issue 465 Supplier Record - Adding editing new records
    'Modified by Anol 20 Aug 2014
    cboSupplierType.Enabled = False
    cmdSupplierType(0).Enabled = False
    'End of resolution
   Me.BackColor = MODULEBACKCOLOR
   tabSupplier.BackColor = MODULEBACKCOLOR
   Dim adoConn As New ADODB.Connection

   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   SupplierAccountBalance adoConn
   'added by anol 22 Jun issue 566
    insertLL adoConn
    'end of modification
   loadCBOValues adoConn
   loadSupplierType adoConn

   adoConn.Close
   Set adoConn = Nothing
   
   Me.Height = 7080
   Me.Width = 13545
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   tabSupplier.Tab = 0

   bVatCodeLoaded = False
   
'   Call WheelHook(Me.hWnd)
End Sub
Private Sub insertLL(adoConn As ADODB.Connection)
'added by anol 26 Jun 2015
    Dim rsSupplier As New ADODB.Recordset
    Dim szSQL As String
   szSQL = "SELECT * " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'SCODE' and CODE='LL'"
   rsSupplier.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   If rsSupplier.EOF Then
       rsSupplier.AddNew
       rsSupplier!Code = "LL"
       rsSupplier!PrimaryCode = "SCODE"
       rsSupplier!Value = "LandLord"
       rsSupplier.Update
       rsSupplier.Close
       Set rsSupplier = Nothing
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   frmMMain.fraCmdButton.Enabled = True
   'I am commening the below Code for Landlord
'   'Added by anol 13 May 2015
'   'Lanlord table needs to be updated from Supplier table
'   'Should be alwys syncronized with supplier table where supplier type = 'LL'
'   'So I am deleting all record from landlord table and transferring all where supplier type = 'LL' from supplier table
'   Dim adoConn As New ADODB.Connection
'   If adoConn.State = 0 Then
'        adoConn.Open getConnectionString
'   End If
'   adoConn.Execute "DELETE FROM Landlord"
'   Dim rsSupplier As New ADODB.Recordset
'   Dim rsLandlord As New ADODB.Recordset
'   rsSupplier.Open "Select * from Supplier where SupplierType='LL'", adoConn, adOpenStatic, adLockReadOnly
'   rsLandlord.Open "Select * from Landlord", adoConn, adOpenDynamic, adLockOptimistic
'   While Not rsSupplier.EOF
'           rsLandlord.AddNew
'           rsLandlord!landLordID = "L-" & rsSupplier!SupplierID
'           rsLandlord!LandlordName = rsSupplier!SupplierName
'           rsLandlord!LandlordAddressLine1 = rsSupplier!SupplierAddressLine1
'           rsLandlord!LandlordAddressLine2 = rsSupplier!SupplierAddressLine2
'           rsLandlord!LandlordAddressLine3 = rsSupplier!SupplierAddressLine3
'           rsLandlord!LandlordAddressLine4 = rsSupplier!SupplierAddressLine4
'           rsLandlord!LandlordPostCode = rsSupplier!SupplierPostCode
'           rsLandlord!LandlordOfficeEmail = rsSupplier!SupplierOfficeEmail
'           rsLandlord!LandlordPersonalEmail = rsSupplier!SupplierPersonalEmail
'           rsLandlord!LandlordHomeTel = rsSupplier!SupplierHomeTel
'           rsLandlord!LandlordMobile = rsSupplier!SupplierMobile
'           rsLandlord!LandlordOfficeAddressLine1 = rsSupplier!SupplierOfficeAddressLine1
'           rsLandlord!LandlordOfficeAddressLine2 = rsSupplier!SupplierOfficeAddressLine2
'           rsLandlord!LandlordOfficeAddressLine3 = rsSupplier!SupplierOfficeAddressLine3
'           rsLandlord!LandlordOfficeAddressLine4 = rsSupplier!SupplierOfficeAddressLine4
'           rsLandlord!LandlordOfficePostCode = rsSupplier!SupplierOfficePostCode
'           rsLandlord!LandlordOfficeTel = rsSupplier!SupplierOfficeTel
'           rsLandlord!LandlordMemo = rsSupplier!SupplierMemo
'           rsLandlord!LandLordSageSuppAC = rsSupplier!SageSuppAC
'           rsLandlord!VATReg = rsSupplier!VATReg
'           rsLandlord!AcBalance = rsSupplier!AcBalance
'           rsLandlord!BacsRef = rsSupplier!BacsRef
'           rsLandlord.Update
'        rsSupplier.MoveNext
'   Wend
'   adoConn.Close
'   Set adoConn = Nothing
   Unload Me
End Sub

Private Sub SetControls(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         tabSupplier.Tab = 0

         cmdAddNewSupplier.Enabled = True
         cmdEditSupplier.Enabled = True
         cmdSaveSupplier.Enabled = False
         cmdCancelSupplier.Enabled = False
         cmdCloseSupplier.Enabled = True
         cmdSupplier.Enabled = True

         fraMain.Enabled = True

         cmdSaveSupplierDetails.Enabled = False
         cmdCancelSupplierDetails.Enabled = False
         cmdEditSupplierDetails.Enabled = False
         cmdUnitMemoEdit.Enabled = True
         Frame17.Enabled = True

         txtSupplierID.text = ""
         cboSupplierType.text = ""
         txtSupplierName.text = ""
         txtSupplierACBal.text = ""
         txtTaxVatNumber.text = ""
         txtSupplierAddressLine1.text = ""
         txtSupplierAddressLine2.text = ""
         txtSupplierAddressLine3.text = ""
         txtSupplierAddressLine4.text = ""
         txtSupplierHomeTel.text = ""
         txtSupplierMobile.text = ""
         txtSupplierOfficeEmail.text = ""
         txtSupplierPersonalEmail.text = ""
         txtSupplierOfficeAddressLine1.text = ""
         txtSupplierOfficeAddressLine2.text = ""
         txtSupplierOfficeAddressLine3.text = ""
         txtSupplierOfficeAddressLine4.text = ""
         txtSupplierPostCode.text = ""
         txtSupplierOfficeTel.text = ""
         txtSupplierOfficePostCode.text = ""
         txtCreditLimit.text = ""
         txtCodeVat.text = ""
         txtPaymentTerms.text = ""
         cboAccType.text = ""
         cboPayType.text = ""
         txtSortCode.text = ""
         txtAcNo.text = ""
         txtAcName.text = ""
         txtBankPayRef.text = ""

      Case ComponentMode.NewEntryMode
         tabSupplier.Tab = 0
         Frame1(0).Enabled = True
         Frame2.Enabled = True
        'Resolved by BOSL
        'Issue 465 Supplier Record - Adding editing new records
        'Modified by Anol 20 Aug 2014
        cboSupplierType.Enabled = True
        cmdSupplierType(0).Enabled = True
        'End of resolution
         cmdAddNewSupplier.Enabled = False
         cmdEditSupplier.Enabled = False
         cmdSaveSupplier.Enabled = True
         cmdCancelSupplier.Enabled = True
         cmdCloseSupplier.Enabled = False
         cmdSupplier.Enabled = False

         cmdSaveSupplierDetails.Enabled = False
         cmdCancelSupplierDetails.Enabled = False
         cmdEditSupplierDetails.Enabled = False

         fraMain.Enabled = True

         txtSupplierID.text = ""
         cboSupplierType.text = ""
         txtSupplierName.text = ""
         txtSupplierACBal.text = ""
         txtTaxVatNumber.text = ""
         txtSupplierAddressLine1.text = ""
         txtSupplierAddressLine2.text = ""
         txtSupplierAddressLine3.text = ""
         txtSupplierAddressLine4.text = ""
         txtSupplierHomeTel.text = ""
         txtSupplierMobile.text = ""
         txtSupplierOfficeEmail.text = ""
         txtSupplierPersonalEmail.text = ""
         txtSupplierOfficeAddressLine1.text = ""
         txtSupplierOfficeAddressLine2.text = ""
         txtSupplierOfficeAddressLine3.text = ""
         txtSupplierOfficeAddressLine4.text = ""
         txtSupplierPostCode.text = ""
         txtSupplierOfficeTel.text = ""
         txtSupplierOfficePostCode.text = ""
         txtCreditLimit.text = ""
         txtCodeVat.text = ""
         txtPaymentTerms.text = ""
         txtSortCode.text = ""
         txtAcNo.text = ""
         txtAcName.text = ""
         txtBankPayRef.text = ""

      Case ComponentMode.EditMode
         tabSupplier.Tab = 0
        'Resolved by BOSL
        'Issue 465 Supplier Record - Adding editing new records
        'Modified by Anol 20 Aug 2014
        cboSupplierType.Enabled = True
        cmdSupplierType(0).Enabled = True
        'End of resolution
         cmdAddNewSupplier.Enabled = False
         cmdEditSupplier.Enabled = False
         cmdSaveSupplier.Enabled = True
         cmdCancelSupplier.Enabled = True
         cmdCloseSupplier.Enabled = False
         cmdSupplier.Enabled = False

         fraMain.Enabled = True

         cmdSaveSupplierDetails.Enabled = False
         cmdCancelSupplierDetails.Enabled = False
         cmdEditSupplierDetails.Enabled = False

         cmdUnitMemoEdit.Enabled = False
         Frame17.Enabled = False

      Case ComponentMode.SavedMode
         tabSupplier.Tab = 0
        'Resolved by BOSL
        'Issue 465 Supplier Record - Adding editing new records
        'Modified by Anol 20 Aug 2014
        cboSupplierType.Enabled = False
        cmdSupplierType(0).Enabled = False
        'End of resolution
         cmdAddNewSupplier.Enabled = True
         cmdEditSupplier.Enabled = True
         cmdSaveSupplier.Enabled = False
         cmdCancelSupplier.Enabled = False
         cmdCloseSupplier.Enabled = True
         cmdSupplier.Enabled = True

         fraMain.Enabled = False

         cmdSaveSupplierDetails.Enabled = False
         cmdCancelSupplierDetails.Enabled = False
         cmdEditSupplierDetails.Enabled = True
         cmdUnitMemoEdit.Enabled = True

         Frame1(0).Enabled = False
         Frame2.Enabled = False
         Frame17.Enabled = True
   End Select
End Sub

Private Sub Label20_Click(Index As Integer)
   If Index = 0 Then                               ' Sort Tenant ID
      SortingGrid flxSupplierList, Index + 1, bSortingCol1
      bSortingCol1 = IIf(bSortingCol1, False, True)
      Label20(0).FontBold = True
      Label20(1).FontBold = False
      Label20(2).FontBold = False
   End If

   If Index = 1 Then                               ' Sort Tenant Name
      SortingGrid flxSupplierList, Index + 1, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      Label20(0).FontBold = False
      Label20(1).FontBold = True
      Label20(2).FontBold = False
   End If

   If Index = 2 Then                               ' Sort Unit Name
      SortingGrid flxSupplierList, Index + 1, bSortingCol3
      bSortingCol3 = IIf(bSortingCol3, False, True)
      Label20(0).FontBold = False
      Label20(1).FontBold = False
      Label20(2).FontBold = True
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

Private Sub tabSupplier_Click(PreviousTab As Integer)
   If tabSupplier.Tab = 4 Then
      ConfigFlxContacts
   End If
End Sub

Private Sub tabSupplier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tabSupplier.MousePointer = vbArrow
End Sub

Private Sub txtAcNo_GotFocus()
   SelTxtInCtrl txtAcNo
End Sub

Private Sub txtAcNo_KeyPress(KeyAscii As Integer)
   BankSCTextKeyPress txtAcNo, KeyAscii
End Sub

Private Sub txtAcNo_LostFocus()
   If Len(txtAcNo.text) < 8 And Len(txtAcNo.text) > 0 Then
      MsgBox "Account Number has to be eight digits.", vbInformation + vbOKOnly, "Supplier Account Number"
      txtAcNo.SetFocus
   End If
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 13 Or KeyAscii = 10 Then txtCreditLimit_LostFocus

   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtCreditLimit_LostFocus()
   txtCreditLimit.text = Format(txtCreditLimit.text, "0.00")
End Sub

Private Sub txtPaymentTerms_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtSortCode_GotFocus()
   SelTxtInCtrl txtSortCode
End Sub

Private Sub txtSortCode_KeyPress(KeyAscii As Integer)
   BankSCTextKeyPress txtSortCode, KeyAscii
End Sub

Private Sub txtSortCode_LostFocus()
   If Len(txtSortCode.text) < 6 And Len(txtSortCode.text) > 0 Then
      MsgBox "Sort Code has to be six digits.", vbInformation + vbOKOnly, "Supplier Sort Code"
      txtSortCode.SetFocus
   End If
End Sub

Private Sub txtSupplierID_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtSupplierID_LostFocus()
   If txtSupplierID.Locked Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim szID As String

   If (FORM_STATUS = "New") Then
      adoConn.Open getConnectionString

      szID = txtSupplierID.text

      If (IsAccountExist(szID, adoConn)) Then
         If (Not (txtSupplierID.text = szID)) Then
            MsgBox "The ID is already in use. Possible suggestion is '" & szID & "' and you may chose different ID"
            txtSupplierID.text = szID
            SelTxtInCtrl txtSupplierID
         End If
      End If
      adoConn.Close
   End If

   Set adoConn = Nothing
End Sub

Private Sub txtSupplierName_Change()
'   If cmdAddNewSupplier.Enabled Then Exit Sub
'
'   On Error GoTo ErrHanlder
'
'   Dim szChoice As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   If UBound(szaChoice) > 0 Then
'      If szaChoice(3) <> "" Then
'         If InStr(szaChoice(3), "S") > 0 Then
'            If (FORM_STATUS = "New") Then
'               txtSupplierID.text = txtSupplierName.text
'               txtSupplierID.text = Replace(txtSupplierID.text, " ", "")
'               txtSupplierID.text = UCase(Left(txtSupplierID.text, 8))
'            End If
'         End If
'      End If
'   End If
'
'   Exit Sub
'
'ErrHanlder:
'   ShowMsgInTaskBar ERR.Number & ": " & ERR.description, , "N"
End Sub

Private Sub txtSupplierOfficeEmail_LostFocus()
   Dim szErrMsg As String

   If Trim(txtSupplierOfficeEmail.text) <> "" Then
      If Not ValidateEmail(txtSupplierOfficeEmail.text, szErrMsg) Then
         MsgBox szErrMsg, vbCritical + vbOKOnly, "Supplier Email"
         SelTxtInCtrl txtSupplierOfficeEmail
         txtSupplierOfficeEmail.SetFocus
      End If
   End If
End Sub

Private Sub txtSupplierSearchID_Change()
   Dim i As Integer

   If Len(txtSupplierSearchID.text) > 0 Then
      txtSupplierSearchName.text = ""
      txtSupplierSearch.text = ""
   End If

   For i = 1 To flxSupplierList.Rows - 1
      flxSupplierList.RowHeight(i) = 240
      If UCase(Left(flxSupplierList.TextMatrix(i, 1), Len(txtSupplierSearchID.text))) <> UCase(txtSupplierSearchID.text) Then
         flxSupplierList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSupplierSearchName_Change()
   Dim i As Integer

   If Len(txtSupplierSearchName.text) > 0 Then
      txtSupplierSearchID.text = ""
      txtSupplierSearch.text = ""
   End If

   For i = 1 To flxSupplierList.Rows - 1
      flxSupplierList.RowHeight(i) = 240
      If UCase(Left(flxSupplierList.TextMatrix(i, 2), Len(txtSupplierSearchName.text))) <> UCase(txtSupplierSearchName.text) Then
         flxSupplierList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub cmdAddNewSupplier_Click()
   SetControls NewEntryMode

   FORM_STATUS = "New"
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   loadSupplierType adoConn
   
   adoConn.Close
   Set adoConn = Nothing

   txtSupplierID.Locked = False
   Me.Caption = "New Landlord"
End Sub

Private Sub loadSupplierType(adoConn As ADODB.Connection)
   Dim szSQL As String
   
   szSQL = "SELECT CODE, VALUE " & _
           "FROM SECONDARYCODE " & _
           "WHERE PRIMARYCODE = 'SCODE' AND CODE='LL'"

   populateCombo adoConn, szSQL, cboSupplierType
'Debug.Print cboSupplierType.ListCount
End Sub

Private Sub loadCBOValues(adoConn As ADODB.Connection)
   Dim sSQLQuery As String

   'Account Type
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'ACCT'"

   populateCombo adoConn, sSQLQuery, cboAccType

   'Payment Type
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'RAT'"

   populateCombo adoConn, sSQLQuery, cboPayType
   cboPayType.ListIndex = 1
End Sub

Public Sub LoadFlxContact(adoConn As ADODB.Connection)
   If txtSupplierID.text = "" Then Exit Sub

   Dim szSQL As String, iRow As Integer, iChild As Integer
   Dim adoCont As New ADODB.Recordset

   szSQL = "SELECT * FROM Contacts WHERE WhosContact = 'S' AND HeadContact = '" & txtSupplierID.text & "';"

   adoCont.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If Not adoCont.EOF Then
      ReDim szaAddresses(adoCont.RecordCount - 1, 4) As String
   End If

   flxContacts.Rows = 1
   iRow = 1
   While Not adoCont.EOF
      flxContacts.AddItem ""
      flxContacts.TextMatrix(iRow, 0) = adoCont.Fields.Item("ContactID").Value
      flxContacts.TextMatrix(iRow, 1) = adoCont.Fields.Item("WhosContact").Value
      flxContacts.TextMatrix(iRow, 2) = adoCont.Fields.Item("ContactName").Value
      flxContacts.TextMatrix(iRow, 3) = IIf(IsNull(adoCont.Fields.Item("CompanyName").Value), "", adoCont.Fields.Item("CompanyName").Value)
      flxContacts.TextMatrix(iRow, 4) = IIf(IsNull(adoCont.Fields.Item("OfficeAddressLine1").Value), "", adoCont.Fields.Item("OfficeAddressLine1").Value)
      szaAddresses(iRow - 1, 0) = flxContacts.TextMatrix(iRow, 4)
      szaAddresses(iRow - 1, 1) = IIf(IsNull(adoCont.Fields.Item("OfficeAddressLine2").Value), "", adoCont.Fields.Item("OfficeAddressLine2").Value)
      szaAddresses(iRow - 1, 2) = IIf(IsNull(adoCont.Fields.Item("OfficeAddressLine3").Value), "", adoCont.Fields.Item("OfficeAddressLine3").Value)
      szaAddresses(iRow - 1, 3) = IIf(IsNull(adoCont.Fields.Item("OfficeAddressLine4").Value), "", adoCont.Fields.Item("OfficeAddressLine4").Value)
      szaAddresses(iRow - 1, 4) = IIf(IsNull(adoCont.Fields.Item("OfficePostCode").Value), "", adoCont.Fields.Item("OfficePostCode").Value)
      flxContacts.TextMatrix(iRow, 5) = IIf(IsNull(adoCont.Fields.Item("OfficeTel").Value), "", adoCont.Fields.Item("OfficeTel").Value)
      flxContacts.TextMatrix(iRow, 6) = IIf(IsNull(adoCont.Fields.Item("OfficeEmail").Value), "", adoCont.Fields.Item("OfficeEmail").Value)
      flxContacts.TextMatrix(iRow, 7) = IIf(IsNull(adoCont.Fields.Item("OfficeMobile").Value), "", adoCont.Fields.Item("OfficeMobile").Value)
      flxContacts.TextMatrix(iRow, 8) = IIf(IsNull(adoCont.Fields.Item("Mobile").Value), "", adoCont.Fields.Item("Mobile").Value)
'Debug.Print flxContacts.Cols
      flxContacts.TextMatrix(iRow, 9) = IIf(IsNull(adoCont.Fields.Item("PersonalEmail").Value), "", adoCont.Fields.Item("PersonalEmail").Value)
      iRow = iRow + 1
      adoCont.MoveNext
   Wend

   adoCont.Close
   Set adoCont = Nothing
End Sub

Private Sub LoadFlxACHistory(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset

   ConfigFlxACHistory

   szSQL = "SELECT P.*, TT.DESCRIPTION AS TT_DES, PI.SlNumber AS INV_REF, TT.CONSTANT " & _
           "FROM (tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON  " & _
                  "P.Type = TT.TYPE_ID) LEFT JOIN tblPurInv AS PI ON P.PI = PI.MY_ID " & _
           "WHERE  P.SageAccountNumber = '" & txtSupplierID.text & "' AND P.Amount > 0 " & _
           "ORDER BY P.TransactionID;"
'Debug.Print szSQL
   adoPty.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoPty.EOF
         If adoPty!Type = 6 Or adoPty!Type = 24 Then
            szSQL = "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, " & _
                        "P.Type, P.SlNumber, MID(T.CONSTANT, 4) AS TT " & _
                    "FROM PayTransactions AS PT, tlbPayment AS P, tlbTransactionTypes AS T " & _
                    "WHERE PT.ToTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID AND P.Type = T.TYPE_ID;"
         Else
            szSQL = "SELECT SQ.*, P.SlNumber, MID(T.CONSTANT, 4) AS TT " & _
                    "FROM tlbPayment AS P, tlbTransactionTypes AS T, (" & _
                     "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type, MID(T.CONSTANT, 4) AS TT " & _
                     "FROM PayTransactions AS PT, tlbPayment AS P, tlbTransactionTypes AS T " & _
                     "WHERE PT.FromTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID AND P.Type = T.TYPE_ID) SQ " & _
                    "WHERE SQ.ToTran = P.TransactionID AND P.Type = T.TYPE_ID;"
         End If

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
               If adoPty!Type = 6 Or adoPty!Type = 24 Then
                  .TextMatrix(iChild, 5) = "Payment from: " & adoPtyDtl.Fields.Item("TT").Value & adoPtyDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 5) = "Payment to: " & adoPtyDtl.Fields.Item("TT").Value & adoPtyDtl.Fields.Item("SlNumber").Value
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
         
         If InStr(.TextMatrix(iKount, 2), "Purchase") > 0 Then .TextMatrix(iKount, 2) = Mid(.TextMatrix(iKount, 2), 10)
         If InStr(.TextMatrix(iKount, 2), "Payment") > 0 And InStr(.TextMatrix(iKount, 2), "Account") = 0 Then .TextMatrix(iKount, 2) = "Payment"
         If InStr(.TextMatrix(iKount, 2), "Account") > 0 Then .TextMatrix(iKount, 2) = "Payment on A/C"
         If InStr(.TextMatrix(iKount, 2), "Invoice") > 0 Then .TextMatrix(iKount, 2) = "Invoice"
         
         .TextMatrix(iKount, 3) = IIf(IsNull(adoPty.Fields.Item("PDate").Value), "", adoPty.Fields.Item("PDate").Value)
         'Below line has been modified by anol 08 Apr 2015
         'rollbacked 16 Apr 2015
         .TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("Ref").Value), "", adoPty.Fields.Item("Ref").Value)
         '.TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("ExtRef").Value), "", adoPty.Fields.Item("ExtRef").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoPty.Fields.Item("Details").Value), "", adoPty.Fields.Item("Details").Value)
         .TextMatrix(iKount, 6) = Format(adoPty.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoPty.Fields.Item("OSAmount").Value, "0.00")
         If adoPty!Type = 6 Or adoPty!Type = 24 Then
            .TextMatrix(iKount, 9) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 8) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Credit
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

Private Sub ConfigFlxACHistory()
   Dim szHeader As String, iCol As Integer

   szHeader$ = "Sign|<No|<Type|<Transaction Date" & _
               "|<Reference|<Desc|>Amount|>Balance|>Dr|>Cr|Transaction ID"

   With flxACHistory
      .Clear
      .Cols = 11
      .Rows = 2
      .RowHeight(0) = 0
      .FormatString = szHeader$

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

Private Sub ConfigFlxContacts()
   Dim szHeader As String, iCol As Integer

   szHeader$ = "ContactID|WhosContact|<ContactName|<CompanyName" & _
               "|<OfficeAddressLine1|<OfficeTel|<OfficeEmail|<OfficeMobile|<Mobile|<PersonalEmail"

   With flxContacts
'      .Clear
'      .Rows = 2
'      .Cols = 10
      .RowHeight(0) = 0
      .FormatString = szHeader$

      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = Label11(23).Left - Label11(22).Left
      .ColWidth(3) = Label11(24).Left - Label11(23).Left
      .ColWidth(4) = Label11(25).Left - Label11(24).Left
      .ColWidth(5) = Label11(26).Left - Label11(25).Left
      .ColWidth(6) = Label11(27).Left - Label11(26).Left
      .ColWidth(7) = Label11(28).Left - Label11(27).Left
      .ColWidth(8) = Label11(29).Left - Label11(28).Left
      .ColWidth(9) = .Width + .Left - Label11(29).Left - 360
   End With
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
