VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suppliers"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleMode       =   0  'User
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEditBank_ 
      Caption         =   "&Edit"
      Height          =   360
      Left            =   13230
      TabIndex        =   63
      Top             =   7470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveBank_ 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   13320
      TabIndex        =   62
      Top             =   8010
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelBank_ 
      Caption         =   "Canc&el"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6720
      TabIndex        =   61
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
      Height          =   3615
      Index           =   0
      Left            =   7440
      ScaleHeight     =   3585
      ScaleWidth      =   5640
      TabIndex        =   48
      Top             =   8355
      Visible         =   0   'False
      Width           =   5670
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
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3015
         Index           =   0
         Left            =   45
         TabIndex        =   50
         Top             =   540
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   5318
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   53
         Top             =   255
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
         Index           =   0
         Left            =   1440
         TabIndex        =   52
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
         TabIndex        =   51
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
      Height          =   4020
      Left            =   720
      ScaleHeight     =   3990
      ScaleWidth      =   5895
      TabIndex        =   39
      Top             =   8325
      Visible         =   0   'False
      Width           =   5925
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
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   15
         Width           =   255
      End
      Begin VB.TextBox txtSupplierSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   165
         TabIndex        =   40
         Top             =   300
         Width           =   1350
      End
      Begin VB.TextBox txtSupplierSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1590
         TabIndex        =   41
         Top             =   300
         Width           =   2895
      End
      Begin VB.TextBox txtSupplierSearch 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4515
         TabIndex        =   42
         Top             =   300
         Width           =   1200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierList 
         Height          =   3315
         Left            =   45
         TabIndex        =   43
         Top             =   675
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   5847
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
         Left            =   4320
         TabIndex        =   47
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   46
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   45
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
         Width           =   5595
      End
   End
   Begin TabDlg.SSTab tabSupplier 
      Height          =   6645
      Left            =   120
      TabIndex        =   22
      Top             =   1695
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11721
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
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
      TabPicture(0)   =   "frmSupplier.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdFcancel"
      Tab(0).Control(1)=   "cmdUpdateSuAddress"
      Tab(0).Control(2)=   "cmdSaveFirstTab"
      Tab(0).Control(3)=   "Frame1(0)"
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(5)=   "cmdCancelSupplierDetails"
      Tab(0).Control(6)=   "cmdEditSupplierDetails"
      Tab(0).Control(7)=   "cmdSaveSupplierDetails"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Payments"
      TabPicture(1)   =   "frmSupplier.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "cboAccType"
      Tab(1).Control(2)=   "Label1(31)"
      Tab(1).Control(3)=   "Label1(32)"
      Tab(1).Control(4)=   "Label1(25)"
      Tab(1).Control(5)=   "Label1(33)"
      Tab(1).Control(6)=   "Label1(26)"
      Tab(1).Control(7)=   "Label1(27)"
      Tab(1).Control(8)=   "Label1(28)"
      Tab(1).Control(9)=   "Label1(29)"
      Tab(1).Control(10)=   "Label1(30)"
      Tab(1).Control(11)=   "Label1(23)"
      Tab(1).Control(12)=   "cmdAccType"
      Tab(1).Control(13)=   "cmdCancelPayments"
      Tab(1).Control(14)=   "cmdEditPayments"
      Tab(1).Control(15)=   "cmdSavePayments"
      Tab(1).Control(16)=   "txtFLX"
      Tab(1).Control(17)=   "Frame4"
      Tab(1).Control(18)=   "Frame5"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Account History"
      TabPicture(2)   =   "frmSupplier.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblGridCaption(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblGridCaption(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label11(8)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label11(7)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label11(3)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label11(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label11(1)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11(5)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label11(6)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label11(9)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label11(10)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label11(11)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label11(12)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label11(13)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label11(21)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label11(19)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label11(14)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label11(20)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label11(15)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label11(16)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label11(17)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label11(18)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtFilterClient"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label1(3)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtACBalanceByCl"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label1(4)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtSupplierFilter"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtSearchRef1"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label1(22)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "flxACHistorySplit"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "flxACHistory"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "cmdClientFilter"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "chkShowOutstanding"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "fmeLoading"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "cmdSupplierFilter"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Command1"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Command2"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "fraSearch"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "cmdSearch"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).ControlCount=   40
      TabCaption(3)   =   "Memo/Attachment"
      TabPicture(3)   =   "frmSupplier.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame17"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Contacts"
      TabPicture(4)   =   "frmSupplier.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtAddress"
      Tab(4).Control(1)=   "cmdEditContacts"
      Tab(4).Control(2)=   "cmdAddNewContacts"
      Tab(4).Control(3)=   "flxContacts"
      Tab(4).Control(4)=   "Label11(29)"
      Tab(4).Control(5)=   "Label11(28)"
      Tab(4).Control(6)=   "Label11(27)"
      Tab(4).Control(7)=   "Label11(26)"
      Tab(4).Control(8)=   "Label11(25)"
      Tab(4).Control(9)=   "Label11(24)"
      Tab(4).Control(10)=   "Label11(23)"
      Tab(4).Control(11)=   "Label11(22)"
      Tab(4).Control(12)=   "lblGridCaption(2)"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Job Maintenance"
      TabPicture(5)   =   "frmSupplier.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   6120
         Width           =   1080
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Automatic Demand Generate:"
         ForeColor       =   &H00FF00FF&
         Height          =   2220
         Left            =   9405
         TabIndex        =   213
         Top             =   4320
         Visible         =   0   'False
         Width           =   3715
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E5E5E5&
            Height          =   2100
            Index           =   0
            Left            =   40
            ScaleHeight     =   2040
            ScaleWidth      =   3555
            TabIndex        =   214
            Top             =   90
            Width           =   3615
            Begin VB.CommandButton cmdSearchOK 
               Caption         =   "&OK"
               Height          =   375
               Left            =   120
               TabIndex        =   221
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CommandButton cmdSearchCancel 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   2055
               TabIndex        =   223
               Top             =   1635
               Width           =   1200
            End
            Begin VB.TextBox txtSearchNo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               MaxLength       =   10
               TabIndex        =   215
               Top             =   450
               Width           =   2685
            End
            Begin VB.TextBox txtSearchRef 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               MaxLength       =   20
               TabIndex        =   216
               Top             =   790
               Width           =   2685
            End
            Begin VB.TextBox txtSearchFromD 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               MaxLength       =   80
               TabIndex        =   217
               Top             =   1125
               Width           =   1290
            End
            Begin VB.TextBox txtSearchToD 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2025
               MaxLength       =   80
               TabIndex        =   219
               Top             =   1125
               Width           =   1380
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
               TabIndex        =   225
               Top             =   0
               Width           =   255
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
            Begin VB.Shape Shape3 
               BorderColor     =   &H00C0FFFF&
               FillColor       =   &H00FFC0C0&
               FillStyle       =   0  'Solid
               Height          =   30
               Left            =   0
               Top             =   260
               Width           =   3855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000002&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Range"
               Height          =   195
               Index           =   36
               Left            =   765
               TabIndex        =   224
               Top             =   45
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000002&
               BackStyle       =   0  'Transparent
               Caption         =   "No"
               Height          =   195
               Index           =   35
               Left            =   135
               TabIndex        =   222
               Top             =   495
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000002&
               BackStyle       =   0  'Transparent
               Caption         =   "Ref."
               Height          =   210
               Index           =   34
               Left            =   135
               TabIndex        =   220
               Top             =   810
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000002&
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   195
               Index           =   24
               Left            =   135
               TabIndex        =   218
               Top             =   1125
               Width           =   945
            End
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   465
         Left            =   12105
         TabIndex        =   212
         Top             =   3960
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Old"
         Height          =   465
         Left            =   11160
         TabIndex        =   211
         Top             =   3960
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSupplierFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   360
         Width           =   330
      End
      Begin VB.PictureBox fmeLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5085
         ScaleHeight     =   450
         ScaleWidth      =   2655
         TabIndex        =   201
         Top             =   1575
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label lblLoading 
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait while loading"
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
            Left            =   315
            TabIndex        =   202
            Top             =   90
            Width           =   2745
         End
      End
      Begin VB.CheckBox chkShowOutstanding 
         Caption         =   "Show Outstanding only"
         Height          =   240
         Left            =   7650
         TabIndex        =   198
         Top             =   405
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.CommandButton cmdClientFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   360
         Width           =   330
      End
      Begin VB.Frame Frame8 
         Caption         =   "Memo"
         Height          =   5190
         Left            =   -74865
         TabIndex        =   177
         Top             =   270
         Width           =   13020
         Begin VB.PictureBox fraAllMemo 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   4500
            Left            =   135
            ScaleHeight     =   4500
            ScaleWidth      =   12825
            TabIndex        =   178
            Top             =   270
            Width           =   12825
            Begin VB.CommandButton cmdCloseMemo 
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
               Left            =   12105
               Style           =   1  'Graphical
               TabIndex        =   179
               Top             =   45
               Width           =   660
            End
            Begin VB.TextBox txtMemoAll 
               Height          =   4080
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   194
               Top             =   360
               Width           =   12780
            End
            Begin MSForms.Label lblSea 
               Height          =   195
               Left            =   135
               TabIndex        =   180
               Top             =   90
               Visible         =   0   'False
               Width           =   2175
               VariousPropertyBits=   8388627
               Caption         =   "Details"
               Size            =   "3836;344"
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
               Height          =   240
               Index           =   3
               Left            =   0
               Top             =   75
               Width           =   12375
            End
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8730
            TabIndex        =   190
            Top             =   4755
            Width           =   1125
         End
         Begin VB.CommandButton cmdVAMemo 
            Caption         =   "&View All Memo"
            Height          =   315
            Left            =   3825
            TabIndex        =   183
            Top             =   4755
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.TextBox txtMemoID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11385
            TabIndex        =   182
            Top             =   135
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdUnitMemoNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   5355
            TabIndex        =   184
            Top             =   4755
            Width           =   975
         End
         Begin VB.TextBox txtUnitMemo 
            Height          =   1695
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   181
            Top             =   210
            Width           =   12825
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   6375
            TabIndex        =   186
            Top             =   4755
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   7560
            TabIndex        =   188
            Top             =   4755
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9870
            TabIndex        =   192
            Top             =   4755
            Width           =   1125
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMemoDetails 
            Height          =   2385
            Left            =   90
            TabIndex        =   185
            Top             =   2250
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            HighLight       =   2
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
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
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.Label Label10 
            Height          =   195
            Left            =   765
            TabIndex        =   193
            Top             =   1935
            Width           =   1005
            VariousPropertyBits=   8388627
            Caption         =   "Date"
            Size            =   "1773;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label8 
            Height          =   195
            Left            =   10620
            TabIndex        =   191
            Top             =   1935
            Width           =   1095
            VariousPropertyBits=   8388627
            Caption         =   "User"
            Size            =   "1931;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label6 
            Height          =   195
            Left            =   2205
            TabIndex        =   189
            Top             =   1935
            Width           =   1905
            VariousPropertyBits=   8388627
            Caption         =   "Description"
            Size            =   "3360;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label5 
            Height          =   195
            Left            =   225
            TabIndex        =   187
            Top             =   1935
            Width           =   420
            VariousPropertyBits=   8388627
            Caption         =   "No"
            Size            =   "741;344"
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
            Height          =   240
            Index           =   4
            Left            =   90
            Top             =   1935
            Width           =   12825
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Bank Account Details:"
         Height          =   4245
         Left            =   -68475
         TabIndex        =   167
         Top             =   1440
         Width           =   6540
         Begin VB.TextBox txtSortCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2685
            MaxLength       =   6
            TabIndex        =   169
            Top             =   315
            Width           =   2415
         End
         Begin VB.TextBox txtAcNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2685
            MaxLength       =   8
            TabIndex        =   170
            Top             =   795
            Width           =   2415
         End
         Begin VB.TextBox txtAcName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2685
            MaxLength       =   100
            TabIndex        =   171
            Top             =   1275
            Width           =   2415
         End
         Begin VB.TextBox txtBankPayRef 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2685
            MaxLength       =   100
            TabIndex        =   173
            Top             =   1755
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Index           =   17
            Left            =   405
            TabIndex        =   176
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Index           =   18
            Left            =   405
            TabIndex        =   175
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   195
            Index           =   19
            Left            =   405
            TabIndex        =   174
            Top             =   1275
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Payment Ref:"
            Height          =   195
            Index           =   20
            Left            =   405
            TabIndex        =   172
            Top             =   1755
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Payment Details:"
         Height          =   4245
         Left            =   -74820
         TabIndex        =   161
         Top             =   1440
         Width           =   6270
         Begin VB.CommandButton cmdPayType 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4860
            TabIndex        =   166
            Top             =   450
            Width           =   315
         End
         Begin MSForms.TextBox txtPaymentTerms 
            Height          =   285
            Left            =   1845
            TabIndex        =   168
            Top             =   945
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
         Begin VB.Label Label2 
            Caption         =   "Days"
            Height          =   210
            Left            =   3045
            TabIndex        =   165
            Top             =   915
            Width           =   735
         End
         Begin MSForms.ComboBox cboPayType 
            Height          =   285
            Left            =   1845
            TabIndex        =   164
            Top             =   450
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Terms:"
            Height          =   210
            Index           =   7
            Left            =   270
            TabIndex        =   163
            Top             =   915
            Width           =   1155
         End
         Begin MSForms.Label Label4 
            Height          =   255
            Left            =   270
            TabIndex        =   162
            Top             =   450
            Width           =   1935
            Caption         =   "Payment Type"
            Size            =   "3413;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdFcancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -65820
         TabIndex        =   21
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdateSuAddress 
         Caption         =   "&Update Supplier Address"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -73740
         TabIndex        =   159
         Top             =   765
         Width           =   2925
      End
      Begin VB.CommandButton cmdSaveFirstTab 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -67170
         TabIndex        =   20
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Job Maintenance"
         Height          =   5895
         Left            =   -74880
         TabIndex        =   130
         Top             =   360
         Width           =   12945
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "View &Job"
            Height          =   355
            Left            =   3120
            TabIndex        =   139
            Top             =   5295
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   6840
            TabIndex        =   138
            Top             =   5295
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   11400
            TabIndex        =   137
            Top             =   5295
            Width           =   1395
         End
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "View &Diary Entry"
            Height          =   355
            Left            =   4680
            TabIndex        =   136
            Top             =   5295
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdEmailJS_PO 
            Caption         =   "Email"
            Height          =   355
            Left            =   9000
            TabIndex        =   135
            Top             =   5295
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   131
            Top             =   5160
            Width           =   2775
            Begin VB.OptionButton optDiary 
               Caption         =   "Diary Entries"
               Height          =   255
               Left            =   1440
               TabIndex        =   132
               Top             =   160
               Width           =   1215
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "Jobs"
               Height          =   255
               Left            =   720
               TabIndex        =   133
               Top             =   160
               Width           =   735
            End
            Begin VB.OptionButton optAll 
               Caption         =   "All"
               Height          =   255
               Left            =   120
               TabIndex        =   134
               Top             =   160
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   4500
            Left            =   120
            TabIndex        =   140
            Top             =   690
            Width           =   12675
            _ExtentX        =   22357
            _ExtentY        =   7938
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
            TabIndex        =   144
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   7800
            TabIndex        =   151
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   6600
            TabIndex        =   150
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Name / Diary Entry"
            Height          =   495
            Index           =   4
            Left            =   4680
            TabIndex        =   149
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   9000
            TabIndex        =   148
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9840
            TabIndex        =   147
            Top             =   255
            Width           =   435
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   146
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   145
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   3240
            TabIndex        =   143
            Top             =   255
            Width           =   1275
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   10560
            TabIndex        =   142
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   11640
            TabIndex        =   141
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
         TabIndex        =   129
         Text            =   "frmSupplier.frx":0972
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdEditContacts 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   -63240
         TabIndex        =   128
         Top             =   5970
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNewContacts 
         Caption         =   "&New"
         Height          =   345
         Left            =   -64800
         TabIndex        =   127
         Top             =   5970
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
         Left            =   -64200
         TabIndex        =   100
         Top             =   495
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   95
         Top             =   5430
         Width           =   13050
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete"
            Height          =   435
            Left            =   11280
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add"
            Height          =   435
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open"
            Height          =   435
            Left            =   9540
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   240
            Width           =   1200
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   99
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
      Begin VB.CommandButton cmdSavePayments 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66240
         TabIndex        =   67
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditPayments 
         Caption         =   "&Update Payment Details"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -74685
         TabIndex        =   66
         Top             =   855
         Width           =   2610
      End
      Begin VB.CommandButton cmdCancelPayments 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64650
         TabIndex        =   68
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccType 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -65250
         TabIndex        =   65
         Top             =   660
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Caption         =   "Supplier Address:"
         Enabled         =   0   'False
         Height          =   4830
         Index           =   0
         Left            =   -74940
         TabIndex        =   31
         Top             =   1260
         Width           =   6615
         Begin VB.TextBox txtSupplierAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   70
            TabIndex        =   8
            Top             =   1239
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2085
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2085
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2085
            MaxLength       =   50
            TabIndex        =   12
            Top             =   2700
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2085
            MaxLength       =   100
            TabIndex        =   14
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2085
            MaxLength       =   100
            TabIndex        =   13
            Top             =   3090
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   70
            TabIndex        =   5
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   70
            TabIndex        =   7
            Top             =   906
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1572
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   70
            TabIndex        =   6
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
            Left            =   885
            TabIndex        =   38
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
            Left            =   885
            TabIndex        =   37
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
            Left            =   885
            TabIndex        =   36
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
            Left            =   885
            TabIndex        =   35
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
            Left            =   885
            TabIndex        =   34
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
            Left            =   885
            TabIndex        =   33
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
            Left            =   885
            TabIndex        =   32
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Supplier Alternative Address :"
         Enabled         =   0   'False
         Height          =   4815
         Left            =   -68205
         TabIndex        =   28
         Top             =   1260
         Width           =   6270
         Begin VB.TextBox txtSupplierOfficeAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   70
            TabIndex        =   18
            Top             =   1950
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   70
            TabIndex        =   16
            Top             =   1050
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   19
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   70
            TabIndex        =   17
            Top             =   1530
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   70
            TabIndex        =   15
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   2520
            Width           =   750
         End
      End
      Begin VB.CommandButton cmdCancelSupplierDetails 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -65310
         TabIndex        =   26
         Top             =   3580
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditSupplierDetails 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -67440
         TabIndex        =   27
         Top             =   4185
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveSupplierDetails 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66240
         TabIndex        =   25
         Top             =   3580
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   3135
         Left            =   90
         TabIndex        =   70
         Top             =   1035
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   5530
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
         Height          =   1605
         Left            =   30
         TabIndex        =   71
         Top             =   4425
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   2831
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
         Height          =   5160
         Left            =   -74760
         TabIndex        =   117
         Top             =   720
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   9102
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
         Index           =   23
         Left            =   -62535
         TabIndex        =   210
         Top             =   3375
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
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
         Index           =   22
         Left            =   5715
         TabIndex        =   206
         Top             =   405
         Width           =   315
      End
      Begin MSForms.TextBox txtSearchRef1 
         Height          =   315
         Left            =   6120
         TabIndex        =   205
         Top             =   360
         Width           =   1365
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "2408;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSupplierFilter 
         Height          =   315
         Left            =   3690
         TabIndex        =   204
         Top             =   360
         Width           =   1590
         VariousPropertyBits=   746604575
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "2805;556"
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
         Caption         =   "A/C Balance by Client"
         Height          =   210
         Index           =   4
         Left            =   9810
         TabIndex        =   200
         Top             =   405
         Width           =   1560
      End
      Begin MSForms.TextBox txtACBalanceByCl 
         Height          =   315
         Left            =   11520
         TabIndex        =   199
         Top             =   360
         Width           =   1515
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2672;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By Client"
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
         Left            =   135
         TabIndex        =   197
         Top             =   405
         Width           =   1050
      End
      Begin MSForms.TextBox txtFilterClient 
         Height          =   315
         Left            =   1215
         TabIndex        =   196
         Tag             =   "ALL"
         Top             =   360
         Width           =   1950
         VariousPropertyBits=   746604575
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "3440;556"
         Value           =   "ALL Clients"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Email"
         Height          =   195
         Index           =   29
         Left            =   -64080
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
         Top             =   1920
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   18
         Left            =   7170
         TabIndex        =   92
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   17
         Left            =   6000
         TabIndex        =   91
         Top             =   4185
         Width           =   945
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   16
         Left            =   5160
         TabIndex        =   90
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No"
         Height          =   195
         Index           =   15
         Left            =   4440
         TabIndex        =   89
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   20
         Left            =   10440
         TabIndex        =   88
         Top             =   4185
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   14
         Left            =   3600
         TabIndex        =   87
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   19
         Left            =   9360
         TabIndex        =   86
         Top             =   4185
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   21
         Left            =   11520
         TabIndex        =   85
         Top             =   4185
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   13
         Left            =   2400
         TabIndex        =   84
         Top             =   4185
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   12
         Left            =   1440
         TabIndex        =   83
         Top             =   4185
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   11
         Left            =   480
         TabIndex        =   82
         Top             =   4185
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   81
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   9
         Left            =   11235
         TabIndex        =   80
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   6
         Left            =   7635
         TabIndex        =   79
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   5
         Left            =   5115
         TabIndex        =   78
         Top             =   795
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   77
         Top             =   795
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   1050
         TabIndex        =   76
         Top             =   795
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   2235
         TabIndex        =   75
         Top             =   795
         Width           =   345
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   4
         Left            =   3315
         TabIndex        =   74
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   7
         Left            =   8835
         TabIndex        =   73
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   8
         Left            =   10035
         TabIndex        =   72
         Top             =   795
         Width           =   1185
      End
      Begin MSForms.ComboBox cboAccType 
         Height          =   285
         Left            =   -68265
         TabIndex        =   64
         Top             =   660
         Visible         =   0   'False
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
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   -70545
         TabIndex        =   69
         Top             =   660
         Visible         =   0   'False
         Width           =   2055
         Caption         =   "Account Type"
         Size            =   "3625;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   93
         Top             =   750
         Width           =   12915
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00C0E0FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Top             =   4185
         Width           =   12915
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   -74760
         TabIndex        =   126
         Top             =   480
         Width           =   12735
      End
   End
   Begin VB.Frame fraMain 
      Height          =   1095
      Left            =   120
      TabIndex        =   110
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton cmdTaxList 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   10800
         TabIndex        =   209
         Top             =   720
         Width           =   320
      End
      Begin VB.CommandButton cmdSupplier 
         Caption         =   ".."
         Height          =   315
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton cmdSupplierType 
         Caption         =   ".."
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   8190
         TabIndex        =   160
         Top             =   240
         Width           =   330
      End
      Begin MSForms.TextBox txtCodeVat 
         Height          =   285
         Left            =   9675
         TabIndex        =   208
         Top             =   720
         Width           =   1110
         VariousPropertyBits=   746604569
         BorderStyle     =   1
         Size            =   "1958;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
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
         Left            =   8685
         TabIndex        =   207
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit:"
         Height          =   195
         Index           =   6
         Left            =   8685
         TabIndex        =   116
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
         Height          =   195
         Index           =   5
         Left            =   10890
         TabIndex        =   115
         Top             =   315
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
         TabIndex        =   114
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier A/C:"
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
         TabIndex        =   113
         Top             =   210
         Width           =   930
      End
      Begin MSForms.TextBox txtSupplierID 
         Height          =   315
         Left            =   1245
         TabIndex        =   1
         Top             =   210
         Width           =   2760
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "4868;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSupplierName 
         Height          =   315
         Left            =   1245
         TabIndex        =   2
         Top             =   645
         Width           =   2775
         VariousPropertyBits=   746604571
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "4895;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboSupplierType 
         Height          =   315
         Left            =   5865
         TabIndex        =   3
         Top             =   240
         Width           =   2310
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4075;556"
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
         Caption         =   "Supplier Type:"
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
         Left            =   4560
         TabIndex        =   112
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax / VAT Number:"
         Height          =   195
         Index           =   21
         Left            =   4470
         TabIndex        =   111
         Top             =   660
         Width           =   1425
      End
      Begin MSForms.TextBox txtCreditLimit 
         Height          =   315
         Left            =   9630
         TabIndex        =   23
         Top             =   270
         Width           =   1155
         VariousPropertyBits=   746604571
         MaxLength       =   9
         BorderStyle     =   1
         Size            =   "2037;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtSupplierACBal 
         Height          =   315
         Left            =   11880
         TabIndex        =   24
         Top             =   270
         Width           =   1290
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2275;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtTaxVatNumber 
         Height          =   315
         Left            =   5865
         TabIndex        =   4
         Top             =   675
         Width           =   2370
         VariousPropertyBits=   746604571
         MaxLength       =   15
         BorderStyle     =   1
         Size            =   "4180;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   135
      TabIndex        =   152
      Top             =   1035
      Width           =   13200
      Begin VB.CommandButton cmdCloseSupplier 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   11520
         TabIndex        =   158
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteSupplier 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   7065
         TabIndex        =   157
         Top             =   135
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelSupplier 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5805
         TabIndex        =   156
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveSupplier 
         Caption         =   "&Save Supplier"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2520
         TabIndex        =   155
         Top             =   135
         Width           =   1980
      End
      Begin VB.CommandButton cmdEditSupplier 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4545
         TabIndex        =   154
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNewSupplier 
         Caption         =   "&New supplier"
         Height          =   345
         Left            =   1170
         TabIndex        =   153
         Top             =   135
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSupplier"
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
Dim szaSupplierBalanceCL() As String
Dim szaSupplierBalance1CL() As String
Dim szaChoice()            As String
Dim bVatCodeLoaded         As Boolean
Dim szaAddresses()         As String
Dim Memo_Save_mode  As Boolean
Dim sText As String

Dim reportingDate As String
Dim sessionID As String
Dim Rate As Double

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdSearchCancel_Click()
        fraSearch.Visible = False
End Sub

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtSearchFromD.text = ""
            txtSearchToD.text = ""
            txtSearchRef.text = ""
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If Len(txtSearchNo.text) > 0 Then
                Call LoadFlxACHistory(adoConn, "1")
            Else
                Call LoadFlxACHistory(adoConn, "")
            End If
            If Len(txtSearchNo.text) > 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
            Else
                cmdSearch.Caption = "Sea&rch"
            End If
             
            
             adoConn.Close
             Set adoConn = Nothing
             FocusControl txtSearchRef
    End If
End Sub

Private Sub txtSearchFromD_Change()
    TextBoxChangeDate txtSearchFromD
    txtSearchNo.text = ""
    txtSearchRef.text = ""
End Sub

Private Sub txtSearchFromD_GotFocus()
    SelTxtInCtrl txtSearchFromD
End Sub

Private Sub txtSearchFromD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchFromD, KeyAscii
End Sub

Private Sub txtSearchFromD_LostFocus()
    If txtSearchFromD.text <> "" Then
        TextBoxFormatDate txtSearchFromD
        txtSearchToD.text = txtSearchFromD.text
        SelTxtInCtrl txtSearchToD
     End If
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        txtSearchNo.text = ""
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            Call LoadFlxACHistory(adoConn, "2")
        Else
            Call LoadFlxACHistory(adoConn, "")
        End If
        If Len(txtSearchRef.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
            
        
        adoConn.Close
        Set adoConn = Nothing
        FocusControl txtSearchFromD
    End If
    
End Sub

Private Sub txtSearchToD_Change()
     TextBoxChangeDate txtSearchToD
     txtSearchNo.text = ""
     txtSearchRef.text = ""
End Sub

Private Sub txtSearchToD_GotFocus()
    SelTxtInCtrl txtSearchToD
End Sub

Private Sub txtSearchToD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchToD, KeyAscii
End Sub

Private Sub txtSearchToD_LostFocus()
    If txtSearchToD.text <> "" Then TextBoxFormatDate txtSearchToD
End Sub



Private Sub txtSupplierSearc_Change()
'   If Not bFormLoaded Then Exit Sub
'   SortTheGrid flxPurchHistory, txtClientIdlist, cmbPropertyHistory, txtSupplierSearc
'   flxPurchHistorySplit.Clear
'   flxPurchHistorySplit.Rows = 2
End Sub

Private Sub cmdSearchOK_Click()
    fraSearch.Visible = False
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "")
             cmdSearch.Caption = "Sea&rch"
        ElseIf Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "3")
             cmdSearch.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
             cmdSearch.Caption = "Clear Sea&rch"
             Call LoadFlxACHistory(adoConn, "4")
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

   For iColumn = 1 To 9
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn).Left - 240
   gridMaintenanceHistory.ColWidth(6) = 0
   gridMaintenanceHistory.ColWidth(11) = 900
   For iColumn = 12 To rstMHistory_.Fields.Count
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
' Comment out by anol 20161121 view job was not working
'   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY'), S.Value, " & _
'                "H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
'                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
'                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
'                "H.Detail, H.ActualCost, H.ReportedBy, " & _
'                "H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent, " & _
'                "H.MaintenanceType " & _
'           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
'           "WHERE H.AssignedTo = '" & txtSupplierID.text & "' " & _
'               "AND S.Code = H.MaintenanceType " & _
'               "AND S.PrimaryCode = 'MTYP' " & _
'           "ORDER BY H.ReportedDate DESC;"

'Debug.Print szSQL
szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID, '', P.PropertyName , '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S , Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE H.AssignedTo = '" & txtSupplierID.text & "' AND  U.UnitNumber = L.UnitNumber AND U.PropertyID= H.PropertyID AND  H.PropertyID = P.PropertyID AND H.ReportedBy=L.SageAccountNumber AND " & _
               "S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' " & _
           "ORDER BY H.ReportedDate DESC;"
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

Private Sub cboPayType_Click()
    If txtSupplierName.text = "" Or _
      cboPayType.text = "" Or _
      cboPayType.text = "Cheque" Then Exit Sub

   If txtSortCode.text = "" Then
      'cboPayType.text = ""
      MsgBox "Please update supplier's Bank details.", vbCritical + vbOKOnly, "BACS setting"
      If txtSortCode.Enabled = True Then
        txtSortCode.SetFocus
      End If
      Exit Sub
   End If
End Sub

Private Sub cboSupplierType_GotFocus()
    SelTxtInCtrl cboSupplierType
End Sub

Private Sub cboSupplierType_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 13 And txtTaxVatNumber.Enabled = True Then
        txtTaxVatNumber.SetFocus
   End If
   
End Sub

Private Sub chkShowOutstanding_Click()
'    Dim adoConn As New ADODB.Connection
'    fmeLoading.Visible = True
'    fmeLoading.Refresh
'    flxACHistory.Clear
'    flxACHistory.Rows = 2
'    flxACHistory.Refresh
'    adoConn.Open getConnectionString
'    LoadFlxACHistory adoConn
'    adoConn.Close
'    fmeLoading.Visible = False
'    fmeLoading.Refresh
'    Dim iRow As Integer
''    Dim iRow As Integer
''    If chkShowOutstanding.Value = 1 Then
''        For iRow = 1 To flxACHistory.Rows - 1
''            If Val(flxACHistory.TextMatrix(iRow, 7)) = 0 Then
''                flxACHistory.RowHeight(iRow) = 0
''            End If
''        Next
''    Else
''        For iRow = 1 To flxACHistory.Rows - 1
''            If Val(flxACHistory.TextMatrix(iRow, 7)) = 0 Then
''                If flxACHistory.TextMatrix(iRow, 0) = "-" Then
''                    flxACHistory.RowHeight(iRow) = 0
''                Else
''                    flxACHistory.RowHeight(iRow) = 280
''                End If
''
''            End If
''        Next
''    End If
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    'Debug.Print time
    Call LoadFlxACHistory(adoConn, "")
    'Debug.Print time
    adoConn.Close
    Set adoConn = Nothing
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

'  Build up Suppler's Account Balance
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
Private Sub SupplierAccountBalance1Client(adoConn As ADODB.Connection, szClientID As String)
'written by anol 20170809
'This function shall build up balance for one client For each supllier
'This function will be used upon filtering by supplier
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "From tlbPayment where ClientID='" & szClientID & "' " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaSupplierBalance1CL(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE (Type = 6 OR Type = 24) AND ClientID='" & szClientID & "' " & _
           "GROUP BY SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
    If adoPayDr.Fields.Item("SageAccountNumber").Value = "DERBYCITYC" Then
    Debug.Print adoPayDr.Fields.Item("SageAccountNumber").Value
    End If
      szaSupplierBalance1CL(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBalance1CL(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE (Type = 7 OR Type = 8 OR Type=9) AND ClientID='" & szClientID & "' " & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalance1CL(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaSupplierBalance1CL(1, i) = szaSupplierBalance1CL(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaSupplierBalance1CL(0, iIndex) = adoPayCr.Fields.Item("Cr").Value
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Sub
Private Function SupplierAccountBalanceByClient2(adoConn As ADODB.Connection, szSuppID As String) As Currency
    
'Build Supplier AC balance by client by anol 20180913 this balance is for only one supplier

   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT ClientID AS SageAccountNumber " & _
           "FROM Client;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSupplierBalanceCL(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalanceCL(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      If "MARIACED" = adoPayDr.Fields.Item("SageAccountNumber").Value Then
'                    MsgBox adoPayDr.Fields.Item("SageAccountNumber").Value
'       End If
      szaSupplierBalanceCL(1, iIndex) = 0
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close
   '6
   'New section1
   szSQL = "SELECT Type, ClientID,Type,Sum(Amount) AS Amt " & _
                 "FROM tlbPayment where SageAccountNumber='" & szSuppID & "' group by  Type, ClientID ;"

'Debug.Print szSQL
'adoPayCr.Close
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalanceCL(0, i) = adoPayCr.Fields.Item("ClientID").Value Then
'            If "N19PROPE" = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'                    Debug.Print adoPayCr.Fields.Item("SageAccountNumber").Value
'            End If
            If adoPayCr.Fields.Item("Type").Value = 6 Or adoPayCr.Fields.Item("Type").Value = 24 Then
                 szaSupplierBalanceCL(1, i) = Val(szaSupplierBalanceCL(1, i)) + adoPayCr.Fields.Item("Amt").Value
            End If
            If adoPayCr.Fields.Item("Type").Value = 8 Or adoPayCr.Fields.Item("Type").Value = 7 Or adoPayCr.Fields.Item("Type").Value = 9 Then
                 szaSupplierBalanceCL(1, i) = Val(szaSupplierBalanceCL(1, i)) - adoPayCr.Fields.Item("Amt").Value
            End If
         End If
      Next i
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

End Function
Private Function SupplierAccountBalanceByClient(adoConn As ADODB.Connection, szSuppID As String) As Currency
'function written by anol 20170724
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(ClientID) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.ClientID  " & _
             "From tlbPayment where SageAccountNumber='" & szSuppID & "'" & _
             "GROUP BY tlbPayment.ClientID" & _
            ");"
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Function
   End If

   ReDim szaSupplierBalanceCL(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT ClientID, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE (Type = 6 OR Type = 24) AND SageAccountNumber='" & szSuppID & "'" & _
           "GROUP BY ClientID;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalanceCL(0, iIndex) = adoPayDr.Fields.Item("ClientID").Value
      szaSupplierBalanceCL(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT ClientID, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 AND  SageAccountNumber='" & szSuppID & "' " & _
           "GROUP BY ClientID;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalanceCL(0, i) = adoPayCr.Fields.Item("ClientID").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaSupplierBalanceCL(1, i) = szaSupplierBalanceCL(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaSupplierBalanceCL(0, iIndex) = adoPayCr.Fields.Item("Cr").Value
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Function
Private Function SupplierAccountBalanceALLClient(adoConn As ADODB.Connection, szSuppID As String) As Currency
'I am using this function when Showing supplier balance for all client
'function written by anol 20170724
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

 

    szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE (Type = 6 OR Type = 24) and SageAccountNumber='" & szSuppID & "' " & _
           "GROUP BY SageAccountNumber;"


   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  
   While Not adoPayDr.EOF
      
      SupplierAccountBalanceALLClient = IIf(IsNull(adoPayDr.Fields.Item("Dr").Value), 0, adoPayDr.Fields.Item("Dr").Value)
      
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

  szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 and SageAccountNumber='" & szSuppID & "'" & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      
        SupplierAccountBalanceALLClient = SupplierAccountBalanceALLClient - IIf(IsNull(adoPayCr.Fields.Item("Cr").Value), 0, adoPayCr.Fields.Item("Cr").Value)
      
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Function
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
'   lblVatCode(0).Caption = Label1(27).Caption
'   lblVatCode(0).Tag = Label1(23).Caption
   txtCodeVat.text = Label1(27).Caption & " / " & Rate
   txtCodeVat.Tag = Label1(23).Caption
   
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
      LoadValues FORM_STATUS, adoConn
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
   Me.Caption = "Suppliers"
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

Private Sub cmdClientFilter_Click()
    sText = 1
    
    LoadClients 'load client and also the array of balance
    UpdateBalanceCL 'this function fills the grid with balalnce from array
    Dim Conn2 As New ADODB.Connection
    Conn2.Open getConnectionString
    flxSupplier(0).TextMatrix(1, 3) = Format(SupplierAccountBalanceALLClient(Conn2, txtSupplierID.text), "0.00")
    txtACBalanceByCl.text = flxSupplier(0).TextMatrix(1, 3)
    Conn2.Close
    fraList(0).Top = cmdClientFilter.Top + 1700
    fraList(0).Left = txtSupplierID.Left - 500
    fraList(0).Visible = True
    fraList(0).ZOrder 0
    FocusControl txtSearch1(0)
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachmentInCombo cmbFiles, "Supplier", txtSupplierID.text
   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdCloseMemo_Click()
    fraAllMemo.Visible = False
    txtUnitMemo.SetFocus
    cmdVAMemo.Visible = True
End Sub

Private Sub cmdCloseSupplier_Click()
   Unload Me
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtSupplierID.text, "Supplier"
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
     'anol 15 July 2015
  txtSupplierID.Locked = False
  txtSupplierName.Locked = False
  txtTaxVatNumber.Locked = False
  txtCreditLimit.Locked = False
  txtCodeVat.Enabled = True
  txtCodeVat.Locked = True
  'End of modification
   If txtSupplierID.text = "" Then Exit Sub
   SetControls EditMode
   
   FORM_STATUS = "Edit"
   cmdEditSupplierDetails_Click
   cmdEditPayments.Enabled = True
   cmdUpdateSuAddress.Enabled = True
   txtSupplierID.Locked = True
   cmdTaxList(0).Enabled = True
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

Private Sub cmdFcancel_Click()
    cmdSaveFirstTab.Enabled = False
    cmdUpdateSuAddress.Enabled = True
    cmdFcancel.Enabled = False
     Dim adoConn As New ADODB.Connection
   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString
   LoadValuesFirstTab txtSupplierID.text, adoConn
   adoConn.Close
   
End Sub

Private Sub cmdGridUnitClose_Click(Index As Integer)
   fraList(0).Visible = False
   If Index = 0 Then
        cmdTaxList(0).Enabled = True
   End If
End Sub

Private Sub cmdGridUnitClose2_Click(Index As Integer)
   picSupList.Visible = False
   fraMain.Enabled = True
   Frame3.Enabled = True
   tabSupplier.Enabled = True
End Sub

Private Sub cmdNewMHistory_Click()
   If txtSupplierID.text = "" Then Exit Sub

'   Load frmMaintenanceJob
'   With frmMaintenanceJob
'   'added by anol 23 Jun 2015 issue 566
'      '.UpdateRow = gridMaintenanceHistory.row
'      .CallingForm = "S"          'Calling from lessee form
'      .RecordType = "J"
'      .lblJobName.Caption = "Job Name"
'      .Label1.Caption = "Job No."
'      .txtRef.Enabled = True
'   'modified by anol 23 Jun 2015 issue 566
'      .isEdit = False
'      .Show
'      .ZOrder 0
'   End With
    If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
         With frmMaintenanceJob
           .isEdit = True
           .CallingForm = "S"          'Calling from lessee form
           .RecordType = "J"
           .lblJobName.Caption = "Job Name"
           .Label1.Caption = "Job No."
           .txtRef.Enabled = True
           .UpdateRow = gridMaintenanceHistory.row
           .Frame1.Enabled = False
           .Show
           .ZOrder 0
        End With
    Else
        ShowMsgInTaskBar "Please select a job."
    End If
'   Me.Enabled = False
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

Private Sub cmdSaveFirstTab_Click()
    cmdEditSupplier_Click
    cmdSaveSupplier_Click
    cmdSaveFirstTab.Enabled = False
     'added by anol issue 571
    Frame1(0).Enabled = True
    Frame2.Enabled = True
    txtSupplierAddressLine1.Locked = True
   txtSupplierAddressLine2.Locked = True
   txtSupplierAddressLine3.Locked = True
   txtSupplierAddressLine4.Locked = True
   txtSupplierPostCode.Locked = True
   txtSupplierOfficeTel.Locked = True
   txtSupplierHomeTel.Locked = True
   txtSupplierMobile.Locked = True
   txtSupplierOfficeEmail.Locked = True
   txtSupplierPersonalEmail.Locked = True
   txtSupplierOfficeAddressLine1.Locked = True
   txtSupplierOfficeAddressLine2.Locked = True
   txtSupplierOfficeAddressLine3.Locked = True
   txtSupplierOfficeAddressLine4.Locked = True
   txtSupplierOfficePostCode.Locked = True
   txtAcName.Locked = True
   txtBankPayRef.Locked = True
   txtSortCode.Locked = True
   txtAcNo.Locked = True
   cmdUpdateSuAddress.Enabled = True
   cmdSupplier.Enabled = True
   tabSupplier.SetFocus
  'End of addition
End Sub

Private Function IsCboPayTypeValidation() As Boolean
        On Error GoTo XX
        Dim X
        X = cboPayType.Column(0)
        If IsNull(cboPayType.Value) Then
            MsgBox "Please select a valid Payment Type.", vbInformation, "Warning"
            cboPayType.text = ""
            FocusControl cboPayType
            Exit Function
        End If
        IsCboPayTypeValidation = True
        Exit Function 'chaging month correctly'cboFreqBR.Value
XX:
        MsgBox "Please select a valid Payment Type.", vbInformation, "Warning"
        cboPayType.text = ""
        FocusControl cboPayType
End Function
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
   If cboPayType.text = "" Then
        MsgBox "Please select a valid payment type", vbInformation, "Warning!!!"
        cboPayType.SetFocus
        Exit Sub
   End If
   If IsCboPayTypeValidation = False Then
        Exit Sub
   End If
   'cboPayType
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
   cmdEditSupplier.Enabled = False
   cmdAccType.Enabled = False
   cmdPayType.Enabled = False
   cboAccType.Enabled = False
   cboPayType.Enabled = False
   txtPaymentTerms.Enabled = False

   cmdAddNewSupplier.Enabled = True
'   cmdEditSupplier.Enabled = True
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
'   cmdEditSupplier.Enabled = True
   cmdCancelSupplier.Enabled = False
   cmdCloseSupplier.Enabled = True
   cmdSupplier.Enabled = True
   cboPayType.text = "Cheque"
End Sub

Private Sub cmdSaveSupplier_Click()
   If txtSupplierName.text = "" Then
       ShowMsgInTaskBar "Please enter a Supplier Name to continue.", , "N"
       txtSupplierName.text = ""
       txtSupplierName.SetFocus
       Exit Sub
   End If
   If txtSupplierID.text = "" Then
       ShowMsgInTaskBar "Please enter a Supplier ID to continue.", , "N"
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
'   If txtCodeVat.text = "" Then
'      MsgBox "Please select a tax code.", vbInformation, "Warning"
'      cmdTaxList(0).Enabled = True
'      cmdTaxList(0).SetFocus
'      Exit Sub
'   End If
     If Trim(txtCodeVat.text) = "" Then
            If MsgBox("You have not selected a Tax Code for this supplier. Do you wish to enter a Tax code?", vbYesNo, "Select a Tax code") = vbYes Then
                cmdTaxList(0).Enabled = True
                cmdTaxList(0).SetFocus
                Exit Sub
            End If
      End If
      
      
   If cboSupplierType.ListIndex = -1 Then
        MsgBox "Please select a Supplier Type.", vbInformation, "Warning"
        FocusControl cboSupplierType
        Exit Sub
   End If
   On Error GoTo Err
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
       rstMHistory_!CreatedBy = User
       rstMHistory_!CreatedDate = Now
   End If

   rstMHistory_!SupplierName = txtSupplierName.text
   rstMHistory_!VATReg = txtTaxVatNumber.text
   rstMHistory_!VatCode = txtCodeVat.Tag  'Caption will hold vat name and tag will hold vat ID
'   rstMHistory_!OptedtoTax = IIf(chkOptedtoTax.Value = 1, True, False)
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
'   rstMHistory_!SupplierMemo = txtNote.text
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
   rstMHistory_!Type = "SUPPLIER"
   rstMHistory_!LastModifiedBy = User
   rstMHistory_!LastModifiedDate = Now

   rstMHistory_.Update

   rstMHistory_.Close
   Set rstMHistory_ = Nothing
  

   MsgBox "Supplier entry has been saved successfully.", vbOKOnly, "Saved"

   SetControls EditMode
   
   FORM_STATUS = "DetEdit"
   Me.Caption = "Suppliers"

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
   cmdSaveSupplier.Enabled = False
   tabSupplier.Enabled = True
   tabSupplier.Tab = 0
   cmdUpdateSuAddress.Enabled = True
  'Added by anol 20161014
'   cmdUpdateSuAddress.SetFocus
    cmdUpdateSuAddress_SetFocus
 
   'Added by anol 13 May 2015
   'Lanlord table needs to be updated from Supplier table
   'Should be alwys syncronized with supplier table where supplier type = 'LL'
   'So I am deleting all record from landlord table and transferring all where supplier type = 'LL' from supplier table
   'GoTo Issue 745
   adoConn.Execute "DELETE FROM Landlord"
   Dim rsSupplier As New ADODB.Recordset
   Dim rslandlord As New ADODB.Recordset
   rsSupplier.Open "Select * from Supplier where SupplierType='LL'", adoConn, adOpenStatic, adLockReadOnly
   rslandlord.Open "Select * from Landlord", adoConn, adOpenDynamic, adLockOptimistic
   While Not rsSupplier.EOF
           rslandlord.AddNew
           rslandlord!landLordID = rsSupplier!SupplierID '"L-" & why there is L- I don't know anol 20190403
           rslandlord!LandlordName = rsSupplier!SupplierName
           rslandlord!LandlordAddressLine1 = rsSupplier!SupplierAddressLine1
           rslandlord!LandlordAddressLine2 = rsSupplier!SupplierAddressLine2
           rslandlord!LandlordAddressLine3 = rsSupplier!SupplierAddressLine3
           rslandlord!LandlordAddressLine4 = rsSupplier!SupplierAddressLine4
           rslandlord!LandlordPostCode = rsSupplier!SupplierPostCode
           rslandlord!LandlordOfficeEmail = rsSupplier!SupplierOfficeEmail
           rslandlord!LandlordPersonalEmail = rsSupplier!SupplierPersonalEmail
           rslandlord!LandlordHomeTel = rsSupplier!SupplierHomeTel
           rslandlord!LandlordMobile = rsSupplier!SupplierMobile
           rslandlord!LandlordOfficeAddressLine1 = rsSupplier!SupplierOfficeAddressLine1
           rslandlord!LandlordOfficeAddressLine2 = rsSupplier!SupplierOfficeAddressLine2
           rslandlord!LandlordOfficeAddressLine3 = rsSupplier!SupplierOfficeAddressLine3
           rslandlord!LandlordOfficeAddressLine4 = rsSupplier!SupplierOfficeAddressLine4
           rslandlord!LandlordOfficePostCode = rsSupplier!SupplierOfficePostCode
           rslandlord!LandlordOfficeTel = rsSupplier!SupplierOfficeTel
           rslandlord!LandlordMemo = rsSupplier!SupplierMemo
           rslandlord!LandLordSageSuppAC = rsSupplier!SageSuppAC
           rslandlord!VATReg = rsSupplier!VATReg
           rslandlord!AcBalance = rsSupplier!AcBalance
           rslandlord!BacsRef = rsSupplier!BacsRef
           rslandlord.Update
        rsSupplier.MoveNext
   Wend
   adoConn.Close
   Set adoConn = Nothing
   cmdSupplier.Enabled = True
   txtSupplierName.Locked = True
   cboSupplierType.Enabled = False
   txtTaxVatNumber.Locked = True
   cmdTaxList(0).Enabled = False
   Exit Sub
Err:
   MsgBox Err.description
End Sub
Private Sub cmdUpdateSuAddress_SetFocus()
    On Error GoTo Err
    cmdUpdateSuAddress.SetFocus
    Exit Sub
Err:
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

Private Sub cmdSearch_Click()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        fraSearch.Left = 9404
        fraSearch.Top = 4140
        
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        If cmdSearch.Caption = "Clear Sea&rch" Then
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             fmeLoading.Visible = False
             cmdSearch.Caption = "Sea&rch"
             fraSearch.Visible = False
             Call LoadFlxACHistory(adoConn, "")
        Else
            If fraSearch.Visible = False Then
                fraSearch.Visible = True
                txtSearchNo.SetFocus
            Else
                fraSearch.Visible = False
            End If
        End If
        adoConn.Close
        Set adoConn = Nothing
End Sub

Private Sub cmdSupplier_Click()
   sText = 2
   txtSupplierSearch.text = ""
   txtSupplierSearchName.text = ""
   txtSupplierSearchID.text = ""
   'ConfigflxSupplierList
   LoadflxSupplierList
   UpdateBalance

   picSupList.Top = txtSupplierID.Top + txtSupplierID.Height + 5
   picSupList.Left = txtSupplierID.Left + 5
   picSupList.Visible = True
   picSupList.ZOrder 0
   
   txtSupplierID.Locked = True
   'added by anol 08 Apr 2015
   txtSupplierSearchID.SetFocus
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
      If j = UBound(szaSupplierBalance, 2) Then flxSupplierList.TextMatrix(i, 4) = ""
   Next i
End Sub
Private Sub UpdateBalanceCL()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplier(0).Rows - 1
      For j = 0 To UBound(szaSupplierBalanceCL, 2) - 1
         If flxSupplier(0).TextMatrix(i, 1) = szaSupplierBalanceCL(0, j) Then
            flxSupplier(0).TextMatrix(i, 3) = Format(szaSupplierBalanceCL(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalanceCL, 2) Then flxSupplier(0).TextMatrix(i, 3) = "0.00"
   Next i
End Sub
Private Sub UpdateBalance1CL()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplierList.Rows - 1
      For j = 0 To UBound(szaSupplierBalance1CL, 2) - 1
'      Debug.Print flxSupplierList.TextMatrix(i, 1)
      Debug.Print szaSupplierBalance1CL(0, j)
         If flxSupplierList.TextMatrix(i, 1) = szaSupplierBalance1CL(0, j) Then
            flxSupplierList.TextMatrix(i, 4) = Format(szaSupplierBalance1CL(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalance1CL, 2) Then flxSupplierList.TextMatrix(i, 3) = "0.00"
   Next i
End Sub
Private Sub configflxSupplierList()
   Dim szHeader As String

   flxSupplierList.Cols = 5
   flxSupplierList.Clear
   szHeader$ = "|<SupplierID|<SupplierName|<SupplierPostCode|>AccBalance"
   flxSupplierList.FormatString = szHeader$
   flxSupplierList.ColWidth(0) = 220          'Solid column
   flxSupplierList.ColWidth(1) = 1300       'Supplier ID
   flxSupplierList.ColWidth(2) = 3000       'Supplier Name
   flxSupplierList.ColWidth(3) = 0          'Post Code
   flxSupplierList.ColWidth(4) = 1000       'Account Balance
   flxSupplierList.Rows = 2
'
   'flxSupplierList.RowHeightMin = 300
   flxSupplierList.RowHeight(0) = 0
End Sub

Private Sub LoadflxSupplierList(Optional Filter As String)
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   
   adoConn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName, SupplierPostCode " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' " & _
           "ORDER BY SupplierName;"

   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Len(Filter) > 0 Then
        rstRst.Filter = Filter
   End If
   configflxSupplierList
  
   If rstRst.RecordCount = 0 Then
        flxSupplierList.Rows = 2
        'flxSupplierList.TextMatrix(1, 4) = ""
        'Exit Sub
   Else
        flxSupplierList.Rows = rstRst.RecordCount + 1
   End If
   
   If rstRst.EOF Then GoTo NoRes
   Dim iRow As Integer
   iRow = 1
   
   While Not rstRst.EOF
      flxSupplierList.TextMatrix(iRow, 1) = rstRst!SupplierID
      flxSupplierList.TextMatrix(iRow, 2) = rstRst!SupplierName
      flxSupplierList.TextMatrix(iRow, 3) = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
      flxSupplierList.TextMatrix(iRow, 4) = "0.00"
      rstRst.MoveNext
      'If Not rstRst.EOF Then flxSupplierList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstRst.Close
   
   'SupplierAccountBalance adoConn '20170724 supplier balance was not updating when comeback from the PI
   If frmMMain.frmSupplier_SupplierList_isUptoDate = False Then
        SupplierAccountBalance adoConn
        frmMMain.frmSupplier_SupplierList_isUptoDate = True
   End If
   
    If tabSupplier.Tab = 2 And frmMMain.frmSupplier_SupplierListBCL_isUptoDate = False Then 'this means you are in the history tab and working for filterred by Client and supplier
        SupplierAccountBalance1Client adoConn, txtFilterClient.text
        frmMMain.frmSupplier_SupplierListBCL_isUptoDate = True
    End If
        
   
   
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdSupplierFilter_Click()
   sText = 3
   fraMain.Enabled = False
   Frame3.Enabled = False
   tabSupplier.Enabled = False
   txtSupplierSearch.text = ""
   txtSupplierSearchName.text = ""
   txtSupplierSearchID.text = ""
   'ConfigflxSupplierList
   LoadflxSupplierList
   If txtFilterClient.text = "ALL Clients" Then
        UpdateBalance
   Else
        UpdateBalance1CL
   End If
    
   picSupList.Top = txtSupplierFilter.Top + 1500
   picSupList.Left = txtSupplierFilter.Left - 300
   picSupList.Visible = True
   picSupList.ZOrder 0
   
   txtSupplierFilter.Locked = True
   'added by anol 08 Apr 2015
   txtSupplierSearchID.SetFocus
End Sub

Private Sub cmdSupplierType_Click(Index As Integer)
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim SelSupplierCode As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "SCODE"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'SCODE' AND Code not in ('LL','MA')"
               
   SelSupplierCode = IIf(cboSupplierType.text = "", "", cboSupplierType.Value)
   populateCombo adoConn, sSQLQuery, cboSupplierType
   cboSupplierType.Value = SelSupplierCode

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdTaxList_Click(Index As Integer)
'           fraList(0).Left = 5415
'           fraList(0).Top = 705
           sText = "VAT"
           LoadVAT
           txtSearch1(0).text = ""
           txtSearch2(0).text = ""
           txtSearch3(0).text = ""
           fraList(0).Width = 3520
           flxSupplier(0).Width = 3400
           cmdGridUnitClose(0).Left = fraList(0).Width - cmdGridUnitClose(0).Width - 60
           Shape4(0).Width = fraList(0).Width - 200
        
           If Index = 0 Then
'              fraList(0).Left = cmdTaxList(Index).Left + tabSupplier.Left
'              fraList(0).Top = cmdTaxList(Index).Top + cmdTaxList(Index).Height + tabSupplier.Top
               fraList(0).Left = 6815
               fraList(0).Top = 705
           Else
              fraList(0).Left = tabSupplier.Left + cmdTaxList(Index).Left - 400
              fraList(0).Top = tabSupplier.Top + cmdTaxList(Index).Top + cmdTaxList(Index).Height + 200
           End If
            
           fraList(0).Visible = True
           cmdTaxList(0).Enabled = False
           fraList(0).ZOrder 0
           FocusControl txtSearch1(0)
 
End Sub
Private Sub LoadClients()
    flxSupplier(0).ColWidth(0) = 100
    flxSupplier(0).ColWidth(1) = 1200
    flxSupplier(0).ColWidth(2) = 3000
    flxSupplier(0).ColWidth(3) = 1000
    flxSupplier(0).Clear
    Shape4(0).Width = 5300
    cmdGridUnitClose(0).Left = 5500
    flxSupplier(0).Cols = 4
    flxSupplier(0).TextMatrix(0, 0) = ""
    flxSupplier(0).TextMatrix(0, 1) = "Client ID"
    flxSupplier(0).TextMatrix(0, 2) = "Client Name"
    flxSupplier(0).TextMatrix(0, 3) = "Balance"
    fraList(0).Width = 5800
    flxSupplier(0).Width = 5700
   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1200
   lblSearch0(0).Left = 100
   lblSearch1(0).Width = 3000
   lblSearch1(0).Left = 1300 'lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1(0).Width = 1200
   txtSearch1(0).Left = 100
   
   txtSearch2(0).Width = 3000
   txtSearch2(0).Left = 1350 'txtSearch1(0).Left + flxSupplier(0).ColWidth(0)
   txtSearch3(0).Left = 3100
  
   
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Client ID"
   lblSearch1(0).Caption = "Client Name"
   lblSearch3(0).Caption = "Balance"
   lblSearch3(0).Left = 4500
   lblSearch2(0).Visible = False
   lblSearch3(0).Visible = True
   lblSearch4(0).Visible = False
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString
   
   'SupplierAccountBalanceByClient Conn2, txtSupplierFilter.text  'load client wise balance with single supplier
   SupplierAccountBalanceByClient2 Conn2, txtSupplierFilter.text  'load client wise balance with single supplier
   
   szSQL = "SELECT ClientID, ClientName " & _
           "FROM Client order by ClientID;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
   
   If Not rstRec.EOF Then
     
     
      flxSupplier(0).Rows = 2
      flxSupplier(0).RowHeight(0) = 0

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      flxSupplier(0).TextMatrix(0, 1) = "Client ID"
      flxSupplier(0).TextMatrix(0, 2) = "Client Name"
      flxSupplier(0).TextMatrix(1, 1) = "ALL"
      flxSupplier(0).TextMatrix(1, 2) = "ALL Clients"
      
'      txtACBalanceByCl.text = Format(GetSupplierBalance(txtSupplierID.text), "0.00")
      rRow = 2
      flxSupplier(0).AddItem ""
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec!ClientID
'          flxSupplier(0).Cols = 3
         flxSupplier(0).TextMatrix(rRow, 2) = rstRec!ClientName
'         flxSupplier(0).TextMatrix(rRow, 3) = GetSupplierBalanceByClient(rstRec!clientID)
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
   szSQL = "SELECT VAT_CODE, VAT_RATE,VAT_ID " & _
           "FROM tlbVatCode where IN_USE;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 3
      flxSupplier(0).Rows = 2
      flxSupplier(0).ColWidth(2) = 0
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
         flxSupplier(0).TextMatrix(rRow, 2) = rstRec!VAT_ID
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

Private Sub cmdUnitMemoNew_Click()
    Memo_Save_mode = True
'   fmeTenant.Enabled = False
   cmdUnitMemoNew.Enabled = False
   cmdUnitMemoEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdUnitMemoSave.Enabled = True
   cmdUnitMemoCancel.Enabled = True
   flxMemoDetails.Enabled = False
   txtUnitMemo.Locked = False
   fraAllMemo.Visible = False
   txtUnitMemo.text = ""
   txtUnitMemo.SetFocus
End Sub
Private Sub cmdUnitMemoEdit_Click()
   'Modified by Anol 02 Nov 2014
   'Issue 488  Memo and attachment not saving a record of the date of the memo entry
      fraAllMemo.Visible = False
       If txtMemoID.text = "" Then
            cmdVAMemo.Enabled = True
      Else
            cmdVAMemo.Enabled = False
      End If
      cmdVAMemo.Visible = True
      If txtMemoID.text = "" Then
          ShowMsgInTaskBar "Please select the memo you would like to edit", "Y"
          If flxMemoDetails.Enabled = True Then
               flxMemoDetails.SetFocus
          End If
          Exit Sub
      End If
      
      cmdUnitMemoNew.Enabled = False
      cmdUnitMemoEdit.Enabled = False
      cmdUnitMemoSave.Enabled = True
      cmdUnitMemoCancel.Enabled = True

'      fmeTenant.Enabled = False
      txtUnitMemo.Locked = False
      Memo_Save_mode = False
   If txtUnitMemo.Enabled = True Then
      txtUnitMemo.SetFocus
   End If
End Sub
Public Sub LoadGridMemo()
   'Issue 488
   'Added by anol 03 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtSupplierID.text & "' And  MemoType='Supplier' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
   Dim iRow As Integer
   iRow = 1

   flxMemoDetails.Clear
   flxMemoDetails.Rows = 1
   flxMemoDetails.Cols = 7
   flxMemoDetails.ColWidth(0) = Label10.Left - Label5.Left   'Serial No
   flxMemoDetails.ColWidth(1) = 0
   flxMemoDetails.ColWidth(2) = 0
   flxMemoDetails.ColWidth(3) = 0
   flxMemoDetails.ColWidth(4) = Label6.Left - Label10.Left    'UpdateTime
   flxMemoDetails.ColWidth(5) = Label8.Left - Label6.Left    'MemoDescription
   flxMemoDetails.ColWidth(6) = 2000                         'UserName
   flxMemoDetails.RowHeight(0) = 0
   If rstLeaseAnalysis_.EOF = True Then
       flxMemoDetails.Rows = 2
   End If
   
   While Not rstLeaseAnalysis_.EOF
      flxMemoDetails.AddItem ""
      flxMemoDetails.TextMatrix(iRow, 0) = iRow
      flxMemoDetails.TextMatrix(iRow, 1) = rstLeaseAnalysis_!MemoID
      flxMemoDetails.TextMatrix(iRow, 2) = rstLeaseAnalysis_!MemoType 'col size 0
      flxMemoDetails.TextMatrix(iRow, 3) = rstLeaseAnalysis_!SageAccountNumber 'col size 0
      flxMemoDetails.TextMatrix(iRow, 4) = rstLeaseAnalysis_!UpdateTime
      flxMemoDetails.TextMatrix(iRow, 5) = rstLeaseAnalysis_!MemoDescription
      flxMemoDetails.TextMatrix(iRow, 6) = rstLeaseAnalysis_!UserName
      rstLeaseAnalysis_.MoveNext
      iRow = iRow + 1
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   If iRow > 0 Then
      flxMemoDetails.row = 0
   End If
End Sub

Private Sub cmdUnitMemoSave_Click()
   cmdVAMemo.Visible = False
   If Len(txtUnitMemo.text) = 0 Then
      ShowMsgInTaskBar "Please enter description of memo", "Y"
      If txtUnitMemo.Enabled = True Then
         txtUnitMemo.SetFocus
      End If
      Exit Sub
   End If
   If flxMemoDetails.row = 0 And Memo_Save_mode = False Then
       ShowMsgInTaskBar "Please select a memo from list", "Y"
       Exit Sub
   End If
   
   If SaveMemo Then
      ShowMsgInTaskBar "The memo has been saved successfully."
      LoadGridMemo
   Else
      ShowMsgInTaskBar "Could not save lease analysis", , "N"
   End If
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdUnitMemoCancel.Enabled = False
   flxMemoDetails.Enabled = True
   flxMemoDetails.row = 0
   txtUnitMemo.text = ""
   txtMemoID.text = ""
   txtUnitMemo.Locked = True
'   fmeTenant.Enabled = True
   cmdDelete.Enabled = False
   cmdVAMemo.Enabled = True
   fraAllMemo.Visible = True
   txtMemoAll.text = ""
   Call ViewMemo
   txtMemoAll.SetFocus
End Sub
Private Sub ViewMemo()
   'Issue 488
   'Added by anol 04 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtSupplierID.text & "' And  MemoType='Supplier' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
  Dim strTemp As String
   While Not rstLeaseAnalysis_.EOF
         If Len(rstLeaseAnalysis_!UpdateTime) > 0 Then
               strTemp = " -  "
         Else
               strTemp = ""
         End If
         If Len(txtMemoAll.text) > 0 Then txtMemoAll.text = txtMemoAll.text & vbCrLf & vbCrLf
         txtMemoAll.text = txtMemoAll.text & Left(rstLeaseAnalysis_!UpdateTime, 11) & strTemp & rstLeaseAnalysis_!UserName & vbCrLf & vbCrLf & IIf(IsNull(rstLeaseAnalysis_!MemoDescription) = True, "", rstLeaseAnalysis_!MemoDescription)
         rstLeaseAnalysis_.MoveNext
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   cmdCloseMemo.Refresh
End Sub
Private Function NewMemoID() As Integer
   Dim conMemo As New ADODB.Connection
   conMemo.Open getConnectionString
   Dim szSQL As String
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT MAX(MemoID) AS x   " & _
                 "FROM MemoDetails;"
   rstSet.Open szSQL, conMemo, adOpenStatic, adLockReadOnly

   NewMemoID = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
   rstSet.Close
   Set rstSet = Nothing
   conMemo.Close
End Function

Private Function SaveMemo() As Boolean
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim conMemo As New ADODB.Connection
   Dim rstLease_ As New ADODB.Recordset
   conMemo.Open getConnectionString
   Dim sSQLQuery_ As String
   Dim sSQLFilter As String
   If Not Memo_Save_mode Then
       sSQLFilter = "WHERE MemoID = " & Val(flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)) & " AND Memotype='Supplier' AND SageAccountNumber = '" & txtSupplierID.text & "'"
   Else
       sSQLFilter = ""
   End If
   sSQLQuery_ = "SELECT * " & _
                "FROM MemoDetails " & sSQLFilter
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenDynamic, adLockOptimistic
   If Memo_Save_mode Then rstLeaseAnalysis_.AddNew
   If Memo_Save_mode = False Then
      rstLeaseAnalysis_!MemoID = txtMemoID.text
   Else
      rstLeaseAnalysis_!MemoID = NewMemoID()
   End If
   
   rstLeaseAnalysis_!MemoType = "Supplier"
   rstLeaseAnalysis_!SageAccountNumber = txtSupplierID.text
   rstLeaseAnalysis_!MemoDescription = IIf(txtUnitMemo.text <> "", txtUnitMemo.text, "")
   rstLeaseAnalysis_!UpdateTime = Now
   rstLeaseAnalysis_!UserName = frmMMain.SystemUserName
   rstLeaseAnalysis_.Update
   rstLeaseAnalysis_.Close
   Set rstLease_ = Nothing
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   SaveMemo = True
End Function
Private Sub cmdDelete_Click()
      fraAllMemo.Visible = False
      If flxMemoDetails.row = 0 Then
          ShowMsgInTaskBar "Please select a memo from the list", "Y"
          If flxMemoDetails.Enabled = True Then
               flxMemoDetails.SetFocus
          End If
          Exit Sub
      End If
      If MsgBox("Are you sure to delete memo?", vbQuestion + vbYesNo, "Delete Memo") = vbNo Then Exit Sub
      Dim adoConn As New ADODB.Connection
      adoConn.Open getConnectionString
      adoConn.Execute "DELETE from MemoDetails where MemoID=" & Val(flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)) & " and sageaccountNumber='" & txtSupplierID.text & "'"
      
      adoConn.Close
      MsgBox "Memo has been deleted successfully", vbInformation + vbOKOnly, "Delete Memo"
      
      LoadGridMemo
      
      cmdUnitMemoNew.Enabled = True
      cmdUnitMemoEdit.Enabled = True
      cmdUnitMemoSave.Enabled = False
      cmdUnitMemoCancel.Enabled = False
      flxMemoDetails.Enabled = True
      flxMemoDetails.row = 0
      txtUnitMemo.text = ""
      txtMemoID.text = ""
      txtUnitMemo.Locked = True
'      fmeTenant.Enabled = True
      cmdDelete.Enabled = False
      fraAllMemo.Visible = True
      txtMemoAll.text = ""
      Call ViewMemo
      txtMemoAll.SetFocus
End Sub
Private Sub cmdUnitMemoCancel_Click()
   'Issue 488
   'Modified by anol 04 Oct 2014
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   'MemoButtonEnable False
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdDelete.Enabled = False
   txtUnitMemo.Locked = True
   flxMemoDetails.Enabled = True
   txtUnitMemo.text = ""
   fraAllMemo.Visible = True
   cmdVAMemo.Enabled = True
   cmdVAMemo.Visible = False
   txtMemoAll.SetFocus
End Sub
Private Sub cmdVAMemo_Click()
        fraAllMemo.Visible = True
        txtMemoAll.text = ""
        Call ViewMemo
        cmdVAMemo.Visible = False
        txtMemoAll.SetFocus
End Sub

Private Sub Command1_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Debug.Print time
    LoadFlxACHistory_old adoConn
    Debug.Print time
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub Command2_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Debug.Print time
    Call LoadFlxACHistory(adoConn, "")
    Debug.Print time
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub flxMemoDetails_Click()
      txtMemoID.text = flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)
      txtUnitMemo.text = flxMemoDetails.TextMatrix(flxMemoDetails.row, 5)
      cmdUnitMemoEdit.Enabled = True
      cmdDelete.Enabled = True
End Sub

Private Sub flxMemoDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     flxMemoDetails.ToolTipText = flxMemoDetails.TextMatrix(flxMemoDetails.MouseRow, flxMemoDetails.MouseCol)
End Sub
Private Sub cmdUpdateSuAddress_Click()
    cmdSaveFirstTab.Enabled = True
    cmdUpdateSuAddress.Enabled = False
    cmdFcancel.Enabled = True
    txtSupplierAddressLine1.Locked = False
    txtSupplierAddressLine2.Locked = False
    txtSupplierAddressLine3.Locked = False
    txtSupplierAddressLine4.Locked = False
    txtSupplierPostCode.Locked = False
    txtSupplierOfficeTel.Locked = False
    txtSupplierHomeTel.Locked = False
    txtSupplierMobile.Locked = False
    txtSupplierOfficeEmail.Locked = False
    txtSupplierPersonalEmail.Locked = False
    txtSupplierOfficeAddressLine1.Locked = False
    txtSupplierOfficeAddressLine2.Locked = False
    txtSupplierOfficeAddressLine3.Locked = False
    txtSupplierOfficeAddressLine4.Locked = False
    txtSupplierOfficePostCode.Locked = False
    txtAcName.Locked = False
    txtBankPayRef.Locked = False
    txtSortCode.Locked = False
    txtAcNo.Locked = False
    txtSupplierAddressLine1.SetFocus
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
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If sText = "VAT" Then
'        lblVatCode(0).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
'        txtCodeVat.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        If flxSupplier(0).TextMatrix(flxSupplier(0).row, 0) = "" Then
                 txtCodeVat.text = ""
                 txtCodeVat.Tag = ""
        Else
                txtCodeVat.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0) & " / " & flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
                txtCodeVat.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2) 'this is ID
        End If
         cmdTaxList(0).Enabled = True
    ElseIf sText = 1 Then
        txtFilterClient.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        If txtFilterClient.text = "ALL" Then
            txtFilterClient.text = "ALL Clients"
        End If
        txtFilterClient.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtACBalanceByCl.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 3)
        fmeLoading.Visible = True
        fmeLoading.Refresh
        flxACHistory.Clear
        flxACHistory.Rows = 2
        flxACHistory.Refresh
        Debug.Print time
        Call LoadFlxACHistory(adoConn, "")
        Debug.Print time
        fmeLoading.Visible = False
        fmeLoading.Refresh
    End If
    adoConn.Close
    Set adoConn = Nothing
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





'Private Sub cmdUnitMemoSave_Click()
'   If (SaveMemo("Supplier", "SupplierMemo", txtSupplierID.text, "SupplierID", txtNote)) Then
'     'szTable As String, szFieldName As String, szID As String, szIDFieldName As String, conTextBox As TextBox
'      txtNote.Enabled = False
'      cmdUnitMemoEdit.Enabled = True
'      cmdUnitMemoSave.Enabled = False
'      cmdUnitMemoCancel.Enabled = False
'      ShowMsgInTaskBar "Memo has been saved successfully."
'   Else
'      ShowMsgInTaskBar "Data could not be saved, Please Contact with administrator", , "N"
'   End If
'End Sub
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
'   Label1(27).Caption = lblVatCode(0).Caption 'It is holding the values in case of cancel
'   Label1(23).Caption = lblVatCode(0).Tag 'It is holding the values in case of cancel
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
   cboPayType.SetFocus

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
   txtAcName.Locked = False
   txtBankPayRef.Locked = False
   txtSortCode.Locked = False
   txtAcNo.Locked = False
End Sub

Private Sub flxACHistory_Click()
   Dim iCurRowHeight As Integer, iRow As Integer
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
   'HighLightRowFlxGrid flxACHistory, flxACHistory.row

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "-" Then Exit Sub

ChildGrid:
'  Displaying the splits ************************************************************************************
   
   ConfigFlxACHistorySplit
   adoConn.Open getConnectionString

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PC" Then
      szSQL = "SELECT S.*, F.FundCode " & _
              "FROM tlbPayment AS P, tblPurInv AS I, tblPurInvSRec AS S, Fund AS F " & _
              "WHERE P.PI = I.MY_ID AND " & _
                  "I.MY_ID = S.ParentID AND S.DEPT_ID = F.FundID AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & " " & _
              "ORDER BY S.MY_ID;"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NOMINAL_CODE").Value
            .TextMatrix(iRow, 5) = adoRST.Fields.Item("JOB_ID").Value
            'Modified by anol 06 July 2015
            'Supplier history split line needs to display property code as not displaying currently.
            .TextMatrix(iRow, 6) = adoRST.Fields.Item("TRANS").Value
            '.TextMatrix(iRow, 6) = adoRst.Fields.Item("UNIT_ID").Value
            'Modified by anol 06 July 2015
            'Supplier history split line needs to display fund code instead of Fund Name.
            'issue 571
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("FundCode").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRST.Close
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PP" And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "PPR" Then
      szSQL = "SELECT P.ExtRef, S.UNIT_ID, F.FundName, S.Amount AS PaymentAmount, P.NominalCode " & _
              "FROM tlbPayment AS P, tlbPaymentSplit AS S, Fund AS F " & _
              "WHERE P.TransactionID = S.PayHeader AND P.FundID = F.FundID AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRST.Fields.Item("UNIT_ID").Value), "", adoRST.Fields.Item("UNIT_ID").Value)
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("FundName").Value
            .TextMatrix(iRow, 8) = flxACHistory.TextMatrix(flxACHistory.row, 5)
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("PaymentAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("PaymentAmount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
      adoRST.Close
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then
      szSQL = "SELECT P.* " & _
              "FROM tlbPayment AS P " & _
              "WHERE P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 10) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRST.Fields.Item("UnitID").Value), "", adoRST.Fields.Item("UnitID").Value)
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("FundID").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Details").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then _
               .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Then _
               .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
      
      adoRST.Close
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
   Dim adoConn As New ADODB.Connection
   If flxSupplierList.TextMatrix(flxSupplierList.row, 1) = "" Then
        'picSupList.Visible = False
        Exit Sub
   End If
   If sText = 2 Then
            If flxSupplierList.row = 0 Then
                picSupList.Visible = False
                Exit Sub
            End If
            'anol 15 July 2015
            txtSupplierID.Locked = True
            txtSupplierName.Locked = True
            txtTaxVatNumber.Locked = True
            txtCreditLimit.Locked = True
            cmdUpdateSuAddress.Enabled = True
          'End of modification
          
          
           txtSupplierID.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
           txtSupplierFilter.text = txtSupplierID.text
'           Dim adoConn As New ADODB.Connection
           adoConn.Open getConnectionString
        
           LoadValues txtSupplierID.text, adoConn 'load supplier address ,Payment, memo contact
           Debug.Print time
           Call LoadFlxACHistory(adoConn, "")                   'load main transaction history for supplier
           Debug.Print time
           LoadGridMaintenanceHistory adoConn   'load job and diary
        
           adoConn.Close
           Set adoConn = Nothing
        
           Me.Caption = "Suppliers: " & txtSupplierName.text
        
           picSupList.Visible = False
           cmdEditSupplierDetails.Enabled = True
           cmdUnitMemoEdit.Enabled = True
           cmdEditPayments.Enabled = True
           Frame17.Enabled = True
           tabSupplier.Enabled = True
           If txtSupplierName.Enabled = True Then
                txtSupplierName.SetFocus
           End If
           Call LoadGridMemo
           Call ViewMemo
           'I do not need to come at edit mode when there is no tax code
'           If Len(txtCodeVat.text) = 0 Then
'                MsgBox "Please select a tax code", vbInformation, "Warning"
'                cmdEditSupplier_Click
'                FocusControl cmdTaxList(0)
'           End If
   ElseIf sText = 3 Then
        txtSupplierFilter.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
        txtSupplierID.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
        
        fraMain.Enabled = True
        Frame3.Enabled = True
        tabSupplier.Enabled = True
        picSupList.Visible = False
        adoConn.Open getConnectionString
        'If txtSupplierFilter.text <> txtSupplierID.text Then
            LoadValues txtSupplierID.text, adoConn
        'End If
        Call LoadFlxACHistory(adoConn, "")
        txtACBalanceByCl.text = flxSupplierList.TextMatrix(flxSupplierList.row, 4)
        adoConn.Close
        Set adoConn = Nothing
   End If
'   Me.Height = 5850
'   tabSupplier.Top = 0
End Sub
Private Sub LoadValuesFirstTab(Id As String, adoConn As ADODB.Connection)
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
                  "S.SupplierOfficeTel, S.SupplierMemo, S.VATReg,S.VATCode,S.chkOptedtoTax, " & _
                  "S.BacsRef, S.SupplierType, S.creditlimit, S.nominalcode, " & _
                  "S.AccountType, S.PaymentType, S.PaymentTerms, S.SortCode, S.AcNo, S.AcName, S.BPR, " & _
                  "N.Name AS NN " & _
                "FROM (Supplier AS S LEFT OUTER JOIN NominalLedger AS N " & _
                  "ON S.NominalCode = N.Code) LEFT OUTER JOIN tlbVatCode AS V " & _
                  "ON S.VATCode = V.VAT_CODE " & _
                "WHERE S.SupplierID = '" & Id & "';"
'Debug.Print sSQLQuery_
   rstRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

  
   txtSupplierAddressLine1.text = IIf(IsNull(rstRst!SupplierAddressLine1), "", rstRst!SupplierAddressLine1)
   txtSupplierAddressLine1.Locked = True
   txtSupplierAddressLine2.text = IIf(IsNull(rstRst!SupplierAddressLine2), "", rstRst!SupplierAddressLine2)
   txtSupplierAddressLine2.Locked = True
   txtSupplierAddressLine3.text = IIf(IsNull(rstRst!SupplierAddressLine3), "", rstRst!SupplierAddressLine3)
   txtSupplierAddressLine3.Locked = True
   txtSupplierAddressLine4.text = IIf(IsNull(rstRst!SupplierAddressLine4), "", rstRst!SupplierAddressLine4)
   txtSupplierAddressLine4.Locked = True
   txtSupplierPostCode.text = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
   txtSupplierPostCode.Locked = True
   txtSupplierHomeTel.text = IIf(IsNull(rstRst!SupplierHomeTel), "", rstRst!SupplierHomeTel)
   txtSupplierHomeTel.Locked = True
   txtSupplierMobile.text = IIf(IsNull(rstRst!SupplierMobile), "", rstRst!SupplierMobile)
   txtSupplierMobile.Locked = True
   txtSupplierOfficeTel.text = IIf(IsNull(rstRst!SupplierOfficeTel), "", rstRst!SupplierOfficeTel)
   txtSupplierOfficeTel.Locked = True
   txtSupplierOfficeEmail.text = IIf(IsNull(rstRst!SupplierOfficeEmail), "", rstRst!SupplierOfficeEmail)
   txtSupplierOfficeEmail.Locked = True
   txtSupplierPersonalEmail.text = IIf(IsNull(rstRst!SupplierPersonalEmail), "", rstRst!SupplierPersonalEmail)
   txtSupplierPersonalEmail.Locked = True
   txtSupplierOfficeAddressLine1.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine1), "", rstRst!SupplierOfficeAddressLine1)
   txtSupplierOfficeAddressLine1.Locked = True
   txtSupplierOfficeAddressLine2.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine2), "", rstRst!SupplierOfficeAddressLine2)
   txtSupplierOfficeAddressLine2.Locked = True
   txtSupplierOfficeAddressLine3.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine3), "", rstRst!SupplierOfficeAddressLine3)
   txtSupplierOfficeAddressLine3.Locked = True
   txtSupplierOfficeAddressLine4.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine4), "", rstRst!SupplierOfficeAddressLine4)
   txtSupplierOfficeAddressLine4.Locked = True
   txtSupplierOfficePostCode.text = IIf(IsNull(rstRst!SupplierOfficePostCode), "", rstRst!SupplierOfficePostCode)
   txtSupplierOfficePostCode.Locked = True
   

   MousePointer = vbDefault

'   cmdAddNewSupplier.Enabled = True
'   cmdEditSupplier.Enabled = True
'   cmdSaveSupplier.Enabled = False
'   cmdCancelSupplier.Enabled = False
'   cmdCloseSupplier.Enabled = True
'
'   cmdSaveSupplierDetails.Enabled = False
'   cmdCancelSupplierDetails.Enabled = False
'   cmdEditSupplierDetails.Enabled = True
End Sub
Private Sub LoadValues(Id As String, adoConn As ADODB.Connection)
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   Dim sSQLQuery_ As String, sFilter As String

'   MousePointer = vbHourglass

   sSQLQuery_ = "SELECT S.SupplierID, S.SupplierName, S.SupplierAddressLine1, S.SupplierAddressLine2, " & _
                  "S.SupplierAddressLine3,S.SupplierAddressLine4, S.SupplierPostCode, " & _
                  "S.SupplierOfficeEmail, S.SupplierPersonalEmail, V.VAT_RATE, " & _
                  "S.SupplierHomeTel, S.SupplierMobile, S.SupplierOfficeAddressLine1, " & _
                  "S.SupplierOfficeAddressLine2, S.SupplierOfficeAddressLine3,S.SupplierOfficeAddressLine4, " & _
                  "S.SupplierOfficePostCode, S.PLControl, S.PLControlName, " & _
                  "S.SupplierOfficeTel, S.SupplierMemo, S.VATReg,S.VATCode,optedtoTax, " & _
                  "S.BacsRef, S.SupplierType, S.creditlimit, S.nominalcode, " & _
                  "S.AccountType, S.PaymentType, S.PaymentTerms, S.SortCode, S.AcNo, S.AcName, S.BPR, " & _
                  "N.Name AS NN, V.VAT_Code,V.VAT_ID  " & _
                "FROM (Supplier AS S LEFT OUTER JOIN NominalLedger AS N " & _
                  "ON S.NominalCode = N.Code) LEFT OUTER JOIN tlbVatCode AS V " & _
                  "ON S.VATCode = cstr(V.VAT_ID) " & _
                "WHERE S.SupplierID = '" & Id & "';"
'Debug.Print sSQLQuery_
   rstRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   txtSupplierName.text = rstRst!SupplierName
   txtSupplierACBal.text = Format(GetSupplierBalance(Id), "0.00")
   txtACBalanceByCl.text = txtSupplierACBal.text
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
   If (IIf(IsNull(rstRst!VAT_CODE), "", rstRst!VAT_CODE)) = "" Then
         txtCodeVat.text = ""
         txtCodeVat.Tag = ""
   Else
        txtCodeVat.text = IIf(IsNull(rstRst!VAT_CODE), "", rstRst!VAT_CODE) & " / " & IIf(IsNull(rstRst!VAT_RATE), "", rstRst!VAT_RATE)
        txtCodeVat.Tag = IIf(IsNull(rstRst!VAT_ID), "", rstRst!VAT_ID)
   End If
   Rate = IIf(IsNull(rstRst!VAT_RATE), "0", rstRst!VAT_RATE)
   Label1(27).Caption = IIf(IsNull(rstRst!VAT_CODE), "", rstRst!VAT_CODE) 'It is holding the values in case of cancel
   Label1(23).Caption = txtCodeVat.Tag  'It is holdi
   
   'lblVatCode(0).Caption = IIf(IsNull(rstRst!VAT_CODE), "", rstRst!VAT_CODE)
'   chkOptedtoTax.Value = IIf(rstRst!OptedtoTax = True, 1, 0) 'anol 20201012
   cboAccType.Value = IIf(IsNull(rstRst!AccountType), "", rstRst!AccountType)
   txtPaymentTerms.text = IIf(IsNull(rstRst!PaymentTerms), 0, rstRst!PaymentTerms)
   txtSortCode.text = IIf(IsNull(rstRst!SortCode), "", rstRst!SortCode)
   txtAcNo.text = IIf(IsNull(rstRst!AcNo), "", rstRst!AcNo)
   txtAcName.text = IIf(IsNull(rstRst!AcName), "", rstRst!AcName)
   txtBankPayRef.text = IIf(IsNull(rstRst!BPR), "", rstRst!BPR)
   cboPayType.Value = IIf(IsNull(rstRst!PaymentType), "", rstRst!PaymentType)
   rstRst.Close
   Set rstRst = Nothing
   'txtCodeVat.text = IIf(IsNull(rstRst!VAT_RATE), "", rstRst!VAT_RATE)
Rem by anol 20161106
'   RetrieveMemo "Supplier", "SupplierMemo", txtSupplierID.text, "SupplierID", txtNote

   LoadAttachmentFiles cmbFiles, txtSupplierID.text, "Supplier"

   
   LoadFlxContact adoConn

'   MousePointer = vbDefault

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
'Private Function GetSupplierBalanceByClient(szClientID As String) As Currency
'    Dim j As Integer
'
'   For j = 0 To UBound(szaSupplierBalanceCL, 2) - 1
'      If szClientID = szaSupplierBalanceCL(0, j) Then
'         GetSupplierBalanceByClient = Format(szaSupplierBalanceCL(1, j), "0.00")
'         Exit For
'      End If
'   Next j
'   If j = UBound(szaSupplierBalanceCL, 2) Then GetSupplierBalanceByClient = 0
'
'End Function
Private Sub flxSupplierList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplierList_Click
    End If
End Sub

Private Sub Form_Load()
    If WS_Name = "PCM-DEV2" Then
        Command1.Visible = True
        Command2.Visible = True
    End If
    Frame1(0).BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Frame3.BackColor = MODULEBACKCOLOR
    fraMain.BackColor = MODULEBACKCOLOR
    Frame4.BackColor = MODULEBACKCOLOR
    Frame5.BackColor = MODULEBACKCOLOR
    Frame8.BackColor = MODULEBACKCOLOR
    Frame17.BackColor = MODULEBACKCOLOR
    Frame7.BackColor = MODULEBACKCOLOR
    Label2.BackColor = MODULEBACKCOLOR
    Label4.BackColor = MODULEBACKCOLOR
    sessionID = GetTimeStamp
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
    fraAllMemo.Top = 180
    fraAllMemo.Left = 100
   
    'Issue 465 Supplier Record - Adding editing new records
    'anol 15 July 2015
    txtSupplierID.Locked = True
    txtSupplierName.Locked = True
    txtTaxVatNumber.Locked = True
    txtCreditLimit.Locked = True
    'End of modification
    'Modified by Anol 20 Aug 2014
    cboSupplierType.Enabled = False
    cmdSupplierType(0).Enabled = False
    'End of resolution
   Me.BackColor = MODULEBACKCOLOR
   tabSupplier.BackColor = MODULEBACKCOLOR
   Dim adoConn As New ADODB.Connection

   
   adoConn.Open getConnectionString
   If frmMMain.frmSupplier_SupplierList_isUptoDate = False Then
        SupplierAccountBalance adoConn
        frmMMain.frmSupplier_SupplierList_isUptoDate = True
   End If

   loadCBOValues adoConn
   loadSupplierType adoConn

   adoConn.Close
   Set adoConn = Nothing
   
   Me.Height = 8865 '7515 ' 7080
   Me.Width = 13470
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   tabSupplier.Tab = 0

   bVatCodeLoaded = False
   'added by anol issue 571
   Frame1(0).Enabled = True
   Frame2.Enabled = True
   txtSupplierAddressLine1.Locked = True
   txtSupplierAddressLine2.Locked = True
   txtSupplierAddressLine3.Locked = True
   txtSupplierAddressLine4.Locked = True
   txtSupplierPostCode.Locked = True
   txtSupplierOfficeTel.Locked = True
   txtSupplierHomeTel.Locked = True
   txtSupplierMobile.Locked = True
   txtSupplierOfficeEmail.Locked = True
   txtSupplierPersonalEmail.Locked = True
   txtSupplierOfficeAddressLine1.Locked = True
   txtSupplierOfficeAddressLine2.Locked = True
   txtSupplierOfficeAddressLine3.Locked = True
   txtSupplierOfficeAddressLine4.Locked = True
   txtSupplierOfficePostCode.Locked = True
   txtAcName.Locked = True
   txtBankPayRef.Locked = True
   txtSortCode.Locked = True
   txtAcNo.Locked = True
  'End of addition
   
   Call WheelHook(Me.hWnd)
   cboPayType.text = "Cheque"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   'Added by anol 13 May 2015
   'Lanlord table needs to be updated from Supplier table
   'Should be alwys syncronized with supplier table where supplier type = 'LL'
   'So I am deleting all record from landlord table and transferring all where supplier type = 'LL' from supplier table
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
         cmdTaxList(0).Enabled = True

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
   If tabSupplier.Tab = 1 And cmdEditPayments.Enabled = True Then
    cmdEditPayments.SetFocus
    End If
   If tabSupplier.Tab = 1 Then
        If cboPayType.Enabled = True Then
            cboPayType.SetFocus
        End If
   End If
   If tabSupplier.Tab = 2 Then
        FocusControl cmdClientFilter
   End If
End Sub

Private Sub tabSupplier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' tabSupplier.MousePointer = vbArrow
End Sub

Private Sub txtAcName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBankPayRef.SetFocus
    End If
End Sub

Private Sub txtAcNo_GotFocus()
   SelTxtInCtrl txtAcNo
End Sub

Private Sub txtAcNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAcName.SetFocus
End If
   BankSCTextKeyPress txtAcNo, KeyAscii
End Sub

Private Sub txtAcNo_LostFocus()
   If Len(txtAcNo.text) < 8 And Len(txtAcNo.text) > 0 Then
      MsgBox "Account Number has to be eight digits.", vbInformation + vbOKOnly, "Supplier Account Number"
      txtAcNo.SetFocus
   End If
End Sub

Private Sub txtBankPayRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdSavePayments.Enabled = True Then
        cmdSavePayments.SetFocus
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

Private Sub txtSearch1_Change(Index As Integer)
     'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearch1(0).text) > 0 Then
        txtSearch2(0).text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
        flxSupplier(0).RowHeight(i) = 240
        If InStr(1, UCase(flxSupplier(0).TextMatrix(i, 1)), UCase(txtSearch1(0).text), vbTextCompare) = 0 Then
              flxSupplier(0).RowHeight(i) = 0
        End If
        If flxSupplier(0).RowHeight(i) = 240 Then
              flxSupplier(0).row = i
        End If
   Next i
End Sub

Private Sub txtSearch1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl flxSupplier(0)
    End If
End Sub

Private Sub txtSearch2_Change(Index As Integer)
      'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearch2(0).text) > 0 Then
        txtSearch1(0).text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
        flxSupplier(0).RowHeight(i) = 240
        If InStr(1, UCase(flxSupplier(0).TextMatrix(i, 2)), UCase(txtSearch2(0).text), vbTextCompare) = 0 Then
              flxSupplier(0).RowHeight(i) = 0
        End If
        If flxSupplier(0).RowHeight(i) = 240 Then
              flxSupplier(0).row = i
        End If
   Next i
End Sub

Private Sub txtSearch2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
             
        FocusControl flxSupplier(0)
   
    End If
End Sub

Private Sub txtSearchRef1_Change()
     Dim adoConn As New ADODB.Connection
     adoConn.Open getConnectionString
     Call LoadFlxACHistory(adoConn, "5")
     adoConn.Close
     Set adoConn = Nothing
End Sub

'Private Sub txtSearchRef1_Change()
'    Dim i As Integer
'    For i = flxACHistory.Rows - 1 To 1 Step -1
'        flxACHistory.RowHeight(i) = 240
'
'        If InStr(1, UCase(flxACHistory.TextMatrix(i, 4)), UCase(txtSearchRef1.text), vbTextCompare) = 0 Then
'             flxACHistory.RowHeight(i) = 0
'        End If
'
'        If flxACHistory.RowHeight(i) = 240 Then
'        flxACHistory.row = i
'      End If
'   Next i
'End Sub

Private Sub txtSortCode_GotFocus()
   SelTxtInCtrl txtSortCode
End Sub

Private Sub txtSortCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAcNo.SetFocus
    End If
   BankSCTextKeyPress txtSortCode, KeyAscii
End Sub

Private Sub txtSortCode_LostFocus()
   If Len(txtSortCode.text) < 6 And Len(txtSortCode.text) > 0 Then
      MsgBox "Sort Code has to be six digits.", vbInformation + vbOKOnly, "Supplier Sort Code"
      txtSortCode.SetFocus
   End If
End Sub

Private Sub txtSupplierAddressLine1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierAddressLine2.SetFocus
    End If
End Sub

Private Sub txtSupplierAddressLine2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierAddressLine3.SetFocus
    End If
    
End Sub

Private Sub txtSupplierAddressLine3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierAddressLine4.SetFocus
    End If
End Sub

Private Sub txtSupplierAddressLine4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierPostCode.SetFocus
    End If
End Sub

Private Sub txtSupplierHomeTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierMobile.SetFocus
    End If
End Sub

Private Sub txtSupplierID_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If (KeyAscii >= 65 And KeyAscii <= 90) Or _
         (KeyAscii >= 97 And KeyAscii <= 122) Or _
         (KeyAscii >= 48 And KeyAscii <= 57) Then
      If (KeyAscii >= 97 And KeyAscii <= 122) Then
         KeyAscii = KeyAscii - 32
      End If
   ElseIf KeyAscii = 13 Then
        txtSupplierName.SetFocus
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

Private Sub txtSupplierMobile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeEmail.SetFocus
    End If
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

Private Sub txtSupplierName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 And cboSupplierType.Enabled = True Then
        cboSupplierType.SetFocus
    End If
End Sub

Private Sub txtSupplierOfficeAddressLine1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeAddressLine2.SetFocus
    End If
End Sub

Private Sub txtSupplierOfficeAddressLine2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeAddressLine3.SetFocus
    End If
End Sub

Private Sub txtSupplierOfficeAddressLine3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeAddressLine4.SetFocus
    End If
End Sub

Private Sub txtSupplierOfficeAddressLine4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficePostCode.SetFocus
    End If
End Sub

Private Sub txtSupplierOfficeEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierPersonalEmail.SetFocus
    End If
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

Private Sub txtSupplierOfficePostCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtSupplierOfficeTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierHomeTel.SetFocus
    End If
End Sub

Private Sub txtSupplierPersonalEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeAddressLine1.SetFocus
    End If
End Sub

Private Sub txtSupplierPostCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSupplierOfficeTel.SetFocus
    End If
End Sub

Private Sub txtSupplierSearch_Change()
    Dim i As Integer

   If Len(txtSupplierSearch.text) > 0 Then
      txtSupplierSearchID.text = ""
      txtSupplierSearchName.text = ""
   End If

   For i = flxSupplierList.Rows - 1 To 1 Step -1
      flxSupplierList.RowHeight(i) = 240
      
      If Abs(Val(flxSupplierList.TextMatrix(i, 4))) < Val(txtSupplierSearch.text) Then
         flxSupplierList.RowHeight(i) = 0
      End If
      If flxSupplierList.RowHeight(i) = 240 Then
        flxSupplierList.row = i
      End If
   Next i
End Sub

Private Sub txtSupplierSearch_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtSupplierSearch, KeyAscii
End Sub

Private Sub txtSupplierSearchID_Change()
'   Dim i As Integer
'
'   If Len(txtSupplierSearchID.text) > 0 Then
'      txtSupplierSearchName.text = ""
'      txtSupplierSearch.text = ""
'   End If
'
'   For i = 1 To flxSupplierList.Rows - 1
'      flxSupplierList.RowHeight(i) = 240
'      If UCase(Left(flxSupplierList.TextMatrix(i, 1), Len(txtSupplierSearchID.text))) <> UCase(txtSupplierSearchID.text) Then
'         flxSupplierList.RowHeight(i) = 0
'      End If
'   Next i
'Updated by anol 20 July 2015
   Dim i As Integer

   If Len(txtSupplierSearchID.text) > 0 Then
        txtSupplierSearchName.text = ""
        txtSupplierSearch.text = ""
   End If

'   For i = flxSupplierList.Rows - 1 To 1 Step -1
'      flxSupplierList.RowHeight(i) = 240
'
'      If UCase(Left(flxSupplierList.TextMatrix(i, 1), Len(txtSupplierSearchID.text))) <> UCase(txtSupplierSearchID.text) Then
'         flxSupplierList.RowHeight(i) = 0
'      End If
'      If flxSupplierList.RowHeight(i) = 240 Then
'        flxSupplierList.row = i
'      End If
'   Next i
    If Trim(txtSupplierSearchID.text) = "" Then
        LoadflxSupplierList
    Else
        LoadflxSupplierList " SupplierID LIKE '%" + UCase(txtSupplierSearchID.text) + "*'"
    End If
    UpdateBalance
End Sub

Private Sub txtSupplierSearchID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        flxSupplierList.SetFocus
    End If
End Sub

Private Sub txtSupplierSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtSupplierSearchID.text) = 0 Then
            txtSupplierSearchName.SetFocus
        Else
            flxSupplierList.SetFocus
        End If
    End If
End Sub

Private Sub txtSupplierSearchName_Change()
'   Dim i As Integer
'
'   If Len(txtSupplierSearchName.text) > 0 Then
'      txtSupplierSearchID.text = ""
'      txtSupplierSearch.text = ""
'   End If
'
'   For i = 1 To flxSupplierList.Rows - 1
'      flxSupplierList.RowHeight(i) = 240
'      If UCase(Left(flxSupplierList.TextMatrix(i, 2), Len(txtSupplierSearchName.text))) <> UCase(txtSupplierSearchName.text) Then
'         flxSupplierList.RowHeight(i) = 0
'      End If
'   Next i
'Updated by anol 20 July 2015
'    Dim i As Integer

   If Len(txtSupplierSearchName.text) > 0 Then
      txtSupplierSearchID.text = ""
      txtSupplierSearch.text = ""
   End If

'   For i = flxSupplierList.Rows - 1 To 1 Step -1
'      flxSupplierList.RowHeight(i) = 240
'
'      If UCase(Left(flxSupplierList.TextMatrix(i, 2), Len(txtSupplierSearchName.text))) <> UCase(txtSupplierSearchName.text) Then
'         flxSupplierList.RowHeight(i) = 0
'      End If
'      If flxSupplierList.RowHeight(i) = 240 Then
'        flxSupplierList.row = i
'      End If
'   Next i
    If Trim(txtSupplierSearchName.text) = "" Then
        LoadflxSupplierList
    Else
        LoadflxSupplierList " SupplierName LIKE '%" + UCase(txtSupplierSearchName.text) + "*'"
    End If
    UpdateBalance
End Sub

Private Sub cmdAddNewSupplier_Click()
    'anol 15 July 2015
  txtSupplierID.Locked = False
  txtSupplierName.Locked = False
  txtTaxVatNumber.Locked = False
  txtCreditLimit.Locked = False
   cboSupplierType.Enabled = True
  'End of modification
   SetControls NewEntryMode

   FORM_STATUS = "New"
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   loadSupplierType adoConn
   
   adoConn.Close
   Set adoConn = Nothing

   txtSupplierID.Locked = False
   Me.Caption = "New Supplier"
   tabSupplier.Enabled = False
   
   txtSupplierID.SetFocus
End Sub

Private Sub loadSupplierType(adoConn As ADODB.Connection)
   Dim szSQL As String
   'Modified by anol 07 July by anol
   'Landlord was showing in the supplier type ( But it should not be shown in supplier form)
   szSQL = "SELECT CODE, VALUE " & _
           "FROM SECONDARYCODE " & _
           "WHERE PRIMARYCODE = 'SCODE' AND Code not in ('LL','MA')"

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

Private Sub LoadFlxACHistory(adoConn As ADODB.Connection, Filter As String)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset
   Dim iLoop As Long
   Dim strWhere As String
   Dim rsReportSAChistory As New ADODB.Recordset
   Dim rsDiffClients As New ADODB.Recordset
   Dim tempstr As String
   Dim ifNeedBuildingBal1 As Boolean
   Dim ifNeedBuildingBal2 As Boolean
   fmeLoading.Visible = True
   fmeLoading.Refresh
   ConfigFlxACHistory
   Dim StrWhere2 As String
   Dim strWhere3 As String
   'The ralationship between payment and invoice is One Batch payment can be spread over many Invoice to allocate
   'below method need to be updated, I cannot get the outstading amount from the table as an invoice can be paid of from diferrent client
   'Then need to to consider that payment in case user wants to see the filtered PI and Payment only from this client
   
   'I did not find any partialy allocated payments:( SELECT *FROM tlbPayment where type in (7,8,9) and osamount>0 and amount>osamount;) IN WPM
   ' I have found invoices which is partially paid IN WPM
   
   szSQL = "delete from  ReportSAChistory WHERE SessionID = '" & sessionID & "';"
   adoConn.Execute szSQL
   
   szSQL = "delete from  ReportSAChistory WHERE ReportingDate < #" & reportingDate & "# ;"
   adoConn.Execute szSQL
   
   If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            strWhere = " AND PF Like '%" & tempstr & "%'" 'PF=Inovice number/Receipt number
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            strWhere = " AND Extref Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "3" Or Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND RDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND RDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# "
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
   End If
   If Filter = "5" Then
         If txtSearchRef1.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef1.text), "'", "''")
            strWhere = " AND Extref Like '%" & tempstr & "%'"
        End If
    End If
   'ActualINV is only usefull to show to the slave records
    If txtFilterClient.Tag = "ALL" Then
        szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, '',P.transactionID,P.Type,TT.DESCRIPTION ," & _
                 "(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)) & P.SlNumber AS INVno,Pdate,Details ,extref,amount,osamount,0,0,P.ClientID  FROM " & _
                 "(tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON P.Type = TT.TYPE_ID)" & _
                 "WHERE P.SageAccountNumber = '" & txtSupplierFilter.text & "' ORDER BY P.TransactionID;"
    Else
             '# specail case 1
            '
             'we had faced a problem finding outstanding balance when there was two different clients, .i.e
             'PI in one client that has been paid from another client. So balance needs to be build again.
             'I have following SQL to check that condition if we need to build balance again or not.
             'if you select all client then there is no problem. Other you need to check this problem and call this section
             szSQL = "SELECT P.ClientID, P1.ClientID FROM (tlbPayment P INNER JOIN PayTransactions T ON P.TransactionID = T.FromTran) INNER JOIN tlbPayment " & _
                     "AS P1 ON T.ToTran = P1.TransactionID where P.ClientID<> P1.ClientID and P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P1.ClientID = '" & txtFilterClient.Tag & "';"
             rsDiffClients.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
             If Not rsDiffClients.EOF Then
                 ifNeedBuildingBal1 = True
             End If
             rsDiffClients.Close
             szSQL = "SELECT P.ClientID, P1.ClientID FROM (tlbPayment P INNER JOIN PayTransactions T ON P.TransactionID = T.FromTran) INNER JOIN tlbPayment " & _
                     "AS P1 ON T.ToTran = P1.TransactionID where P.ClientID<> P1.ClientID and P1.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "';"
             rsDiffClients.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
             If Not rsDiffClients.EOF Then
                 ifNeedBuildingBal2 = True
             End If
             rsDiffClients.Close
    
        szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, '',P.transactionID,P.Type,TT.DESCRIPTION ," & _
                 "(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)) & P.SlNumber AS INVno,Pdate,Details ,extref,amount,osamount,0,0,P.ClientID  FROM " & _
                 "(tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON P.Type = TT.TYPE_ID) " & _
                 "WHERE  P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "' ORDER BY P.TransactionID;"
    End If
    If txtFilterClient.Tag <> "ALL" Then
        strWhere3 = " AND  P.ClientID = '" & txtFilterClient.Tag & "'"
    End If
    szSQL = "insert into ReportSAChistory (reportingDate,SessionID,SIGN,transactionID,Type,Type_desc,PF,Pdate,Details ,extref ,amount,Osamount, flag ,isMaster,ClientID) " & _
          szSQL
    adoConn.Execute szSQL
'    and P.Type in (7,8,9)
 'Update allocation sign
    
    If txtFilterClient.Tag = "ALL" Then
            adoConn.Execute "Update ReportSAChistory A INNER JOIN PayTransactions R ON A.transactionID=R.FromTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (7,8,9)"
            adoConn.Execute "Update ReportSAChistory A INNER JOIN PayTransactions R ON A.transactionID=R.ToTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (6,24)"
    Else
            adoConn.Execute "UPDATE (ReportSAChistory AS A INNER JOIN PayTransactions AS R ON A.transactionID = R.FromTran) INNER JOIN tlbPayment P ON R.ToTran = P.TransactionID SET A.SIGN = '+' WHERE SessionID = '" & sessionID & "' AND (A.Type In (7,8,9))" & strWhere3
            adoConn.Execute "UPDATE tlbPayment P INNER JOIN (ReportSAChistory AS A INNER JOIN PayTransactions AS R ON A.transactionID = R.ToTran) ON P.TransactionID = R.FromTran SET A.SIGN = '+'  WHERE SessionID = '" & sessionID & "' AND (A.Type) In (6,24)" & strWhere3
    End If
    'allocation slave for invoices
    'We are not using serial number later on anywhere in report table
    szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
            "'-', R.TransactionID, P.Type,R.PF,PT.Allocdate,PT.PaymentAmount,'1',P.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO" & _
            " FROM ((ReportSAChistory R  INNER JOIN PayTransactions  PT ON R.transactionID = PT.ToTran)" & _
            " INNER JOIN tlbPayment P ON PT.FromTran = P.TransactionID) INNER JOIN tlbTransactionTypes T ON P.Type = T.TYPE_ID where SessionID = '" & sessionID & "' AND R.Type In(6,24)" & strWhere3

    szSQL = "insert into ReportSAChistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,PDate,amount,isMaster,ClientID,ActualINV) " & _
            szSQL
    adoConn.Execute szSQL
  'allocation slave for Receipts
'   Exit Sub
    szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
            "'-', R.TransactionID, P.Type, R.PF,PT.Allocdate,PT.PaymentAmount,'1',P.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO " & _
            "FROM (ReportSAChistory AS R INNER JOIN PayTransactions AS PT ON R.transactionID = PT.FromTran)" & _
            "INNER JOIN (tlbPayment AS P INNER JOIN tlbTransactionTypes AS T ON P.Type = T.TYPE_ID) ON PT.ToTran = P.TransactionID where SessionID = '" & sessionID & "' AND R.Type In(7,8,9)" & strWhere3
     szSQL = "insert into ReportSAChistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,Pdate,amount,isMaster,ClientID,ActualINV) " & _
            szSQL
    adoConn.Execute szSQL
    
    
    Dim PaymentBalance As Double
   
            
    If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
            adoPty.Open "Select * from ReportSAChistory where SessionID = '" & sessionID & "'  order by transactionID desc,ismaster desc", adoConn, adOpenDynamic, adLockOptimistic
            Debug.Print time
            While Not adoPty.EOF
                If adoPty("Sign").Value = "-" Then
                  PaymentBalance = PaymentBalance + adoPty("Amount").Value
               End If
               If adoPty("Sign").Value = "+" Then
                    adoPty("Balance") = adoPty("Amount") - PaymentBalance 'Balance
                    adoPty("Balance") = Round(adoPty("OSAmount"), 2)
                    PaymentBalance = 0
             
               End If
               If adoPty("Sign").Value = "" Then
                     adoPty("Balance") = adoPty("Amount")
                     PaymentBalance = 0
               End If
               adoPty.Update
               adoPty.MoveNext
           Wend
        '    adoPty.UpdateBatch adAffectAllChapters
            adoPty.Close
    End If
    If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
            StrWhere2 = " AND Balance>0 "
    Else
            StrWhere2 = " AND OSAmount>0 "
    End If
    If chkShowOutstanding.Value = 0 Then
        'rsReportSAChistory.Open "Select * from ReportLAChistory where 1=1 " & strWhere & "  order by transactionID,ismaster", adoconn, adOpenStatic, adLockReadOnly
        adoConn.Execute "Update  ReportSAChistory A, (Select transactionID from ReportSAChistory where  SessionID= '" & sessionID & "'" & _
                         strWhere & " order by transactionID,ismaster) As B Set flag=1 where A.transactionID=B.transactionID"
    Else
        adoConn.Execute "Update  ReportSAChistory A, (Select transactionID from ReportSAChistory where SessionID= '" & sessionID & "'" & StrWhere2 & _
                         strWhere & " order by transactionID,ismaster) As B Set flag=1 where A.transactionID=B.transactionID"
        
    End If
    rsReportSAChistory.Open "Select * from ReportSAChistory where flag=1 and SessionID= '" & sessionID & "' order by transactionID ,ismaster ", adoConn, adOpenKeyset, adLockReadOnly
    'rsReportSAChistory.Close
    Dim amtBalance As Double
    Dim amtcr As Double
   If rsReportSAChistory.RecordCount = 0 Then
        flxACHistory.Rows = 2
   Else
        flxACHistory.Rows = rsReportSAChistory.RecordCount + 1
   End If
 iKount = 1
    With flxACHistory
    While Not rsReportSAChistory.EOF
        .TextMatrix(iKount, 0) = rsReportSAChistory("SIGN").Value
       
        .TextMatrix(iKount, 2) = IIf(IsNull(rsReportSAChistory("Type_desc").Value), "", rsReportSAChistory("Type_desc").Value)
        If InStr(.TextMatrix(iKount, 2), "Purchase") > 0 Then .TextMatrix(iKount, 2) = Mid(.TextMatrix(iKount, 2), 10)
        If InStr(.TextMatrix(iKount, 2), "Payment") > 0 And InStr(.TextMatrix(iKount, 2), "Account") = 0 Then .TextMatrix(iKount, 2) = "Payment"
        If InStr(.TextMatrix(iKount, 2), "Account") > 0 Then .TextMatrix(iKount, 2) = "Payment on A/C"
        If InStr(.TextMatrix(iKount, 2), "Invoice") > 0 Then .TextMatrix(iKount, 2) = "Invoice"
        .TextMatrix(iKount, 3) = IIf(IsNull(rsReportSAChistory("PDate").Value), "", rsReportSAChistory("PDate").Value)
        If rsReportSAChistory("SIGN").Value = "-" Then
            .RowHeight(iKount) = 0
            '.TextMatrix(iKount, 4) = "Receipt to " & rsReportSAChistory("PF").Value
            '.TextMatrix(iKount, 5) = "Receipt From " & rsReportSAChistory("PF").Value
            .TextMatrix(iKount, 1) = rsReportSAChistory("ActualINV").Value
             If rsReportSAChistory("Type").Value = 6 Or rsReportSAChistory("Type").Value = 24 Then
                .TextMatrix(iKount, 5) = "Payment to: " & rsReportSAChistory.Fields.Item("ActualINV").Value 'from i sreveresrse because slave has reverse PF
             Else
                .TextMatrix(iKount, 5) = "Payment From: " & rsReportSAChistory.Fields.Item("ActualINV").Value
             End If
        Else
            .RowHeight(iKount) = 260
            .TextMatrix(iKount, 1) = rsReportSAChistory("PF").Value
            .TextMatrix(iKount, 4) = IIf(IsNull(rsReportSAChistory("Extref").Value), "", rsReportSAChistory("Extref").Value)
            .TextMatrix(iKount, 5) = IIf(IsNull(rsReportSAChistory("Details").Value), "", rsReportSAChistory("Details").Value)
           
        End If
        .TextMatrix(iKount, 6) = IIf(rsReportSAChistory("Amount").Value = 0, "", Format(rsReportSAChistory("Amount").Value, "0.00"))
        If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
            .TextMatrix(iKount, 7) = IIf(rsReportSAChistory("Balance").Value = 0, "", Format(rsReportSAChistory("Balance").Value, "0.00"))
        Else
            .TextMatrix(iKount, 7) = IIf(rsReportSAChistory("OSamount").Value = 0, "", Format(rsReportSAChistory("OSamount").Value, "0.00"))
        End If
        If rsReportSAChistory("SIGN").Value <> "-" Then 'for the allocated amount you don't need debit or credit row
            If rsReportSAChistory("Type").Value = 6 Or rsReportSAChistory("Type").Value = 24 Then
                .TextMatrix(iKount, 9) = Format(rsReportSAChistory("Amount").Value, "0.00")
            Else
                .TextMatrix(iKount, 8) = Format(rsReportSAChistory("Amount").Value, "0.00")
            End If
        End If
        .TextMatrix(iKount, 10) = rsReportSAChistory("transactionID").Value
'        iKount = iKount + 1
        rsReportSAChistory.MoveNext
        amtBalance = amtBalance + Val(flxACHistory.TextMatrix(iKount, 9))
        amtcr = amtcr + Val(flxACHistory.TextMatrix(iKount, 8))
        iKount = iKount + 1
   Wend
   End With
   rsReportSAChistory.Close
   Set rsReportSAChistory = Nothing
   fmeLoading.Visible = False
   fmeLoading.Refresh
  ' MsgBox "Debit :" & amtBalance & "Credit :" & amtcr
    adoConn.Execute "Delete from ReportSAChistory where SessionID= '" & sessionID & "'"
   ' MsgBox flxACHistory.Rows
  End Sub
Private Sub LoadFlxACHistory_old(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset
   Dim iLoop As Long
   Dim strWhere As String
   ConfigFlxACHistory
   'The ralationship between payment and invoice is One Batch payment can be spread over many Invoice to allocate
   'below method need to be updated, I cannot get the outstading amount from the table as an invoice can be paid of from diferrent client
   'Then need to to consider that payment in case user wants to see the filtered PI and Payment only from this client
'   If chkShowOutstanding.Value = 1 Then
'        strWhere = " AND P.OSAmount>0 "
'   End If
   If txtFilterClient.Tag = "ALL" Then
        szSQL = "SELECT P.*, TT.DESCRIPTION AS TT_DES, PI.SlNumber AS INV_REF, TT.CONSTANT " & _
                "FROM (tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON  " & _
                       "P.Type = TT.TYPE_ID) LEFT JOIN tblPurInv AS PI ON P.PI = PI.MY_ID " & _
                "WHERE  P.SageAccountNumber = '" & txtSupplierFilter.text & "' " & strWhere & _
                "ORDER BY P.TransactionID;"
                'rem "AND P.Amount > 0  " issue 520 show zero values on suppllier history
               
  Else
        szSQL = "SELECT P.*, TT.DESCRIPTION AS TT_DES, PI.SlNumber AS INV_REF, TT.CONSTANT " & _
                "FROM (tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON  " & _
                       "P.Type = TT.TYPE_ID) LEFT JOIN tblPurInv AS PI ON P.PI = PI.MY_ID " & _
                "WHERE  P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "' " & strWhere & _
                "ORDER BY P.TransactionID;"
                ' AND P.Amount > 0 " issue 520 show zero values on suppllier history
'                txtACBalanceByCl.text = Format(SupplierAccountBalanceByClient(adoConn, txtSupplierID.text, txtFilterClient.text), "0.00")
  End If
                
'Debug.Print szSQL
   adoPty.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoPty.EOF
        '6,24 are totrans(invoice) and fromtrans is payment
         'implentation of os amount psudo code
           'if checkbox yes
                'load where osmaount is >0
           'Else
                'load everything
           'End if  ' ' "P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "' AND " & _
      'on the fifth line I have added a filter by clinet and the sage account number
      'issue 632
      'Now need to update the balance as per my algorithm
     
                 If adoPty!Type = 6 Or adoPty!Type = 24 Then
                        If txtFilterClient.Tag <> "ALL" Then
                              strWhere = "P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "' AND "
                        Else
                              strWhere = ""
                        End If
                         szSQL = "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, " & _
                                "P.Type, P.SlNumber, MID(T.CONSTANT, 4) AS TT " & _
                            "FROM PayTransactions AS PT, tlbPayment AS P, tlbTransactionTypes AS T " & _
                            "WHERE PT.ToTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                             strWhere & _
                                "PT.FromTran = P.TransactionID AND P.Type = T.TYPE_ID;"
                 Else
                       If txtFilterClient.Tag <> "ALL" Then
                              strWhere = "AND P.SageAccountNumber = '" & txtSupplierFilter.text & "' AND  P.ClientID = '" & txtFilterClient.Tag & "' "
                        Else
                              strWhere = ""
                        End If
                        szSQL = "SELECT SQ.*, P.SlNumber, MID(T.CONSTANT, 4) AS TT " & _
                            "FROM tlbPayment AS P, tlbTransactionTypes AS T, (" & _
                             "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type, MID(T.CONSTANT, 4) AS TT " & _
                             "FROM PayTransactions AS PT, tlbPayment AS P, tlbTransactionTypes AS T " & _
                             "WHERE PT.FromTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                                "PT.FromTran = P.TransactionID AND P.Type = T.TYPE_ID) SQ " & _
                            "WHERE SQ.ToTran = P.TransactionID AND P.Type = T.TYPE_ID " & strWhere & ";"
                            
                           
                 End If
        
                 adoPtyDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                 iChild = 0
                 If adoPtyDtl.RecordCount > 0 Then 'looping by header in payment table for all types
                    .AddItem "" ' Here it is adding a header from payment table i.e iKount, it may be an invoice or payment
                    .TextMatrix(iKount, 0) = "+"
                    iChild = iKount + 1
                    While Not adoPtyDtl.EOF ' then it is going to the matching table and adding the the related payment or invoice
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
                       .RowHeight(iChild) = 280
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
                 '.TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("Ref").Value), "", adoPty.Fields.Item("Ref").Value)
                 .TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("ExtRef").Value), "", adoPty.Fields.Item("ExtRef").Value)
                 .TextMatrix(iKount, 5) = IIf(IsNull(adoPty.Fields.Item("Details").Value), "", adoPty.Fields.Item("Details").Value)
                 .TextMatrix(iKount, 6) = Format(adoPty.Fields.Item("Amount").Value, "0.00")
                 .TextMatrix(iKount, 7) = Format(adoPty.Fields.Item("Amount").Value, "0.00") 'Format(adoPty.Fields.Item("OSAmount").Value, "0.00")
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
   Dim amtBalance As Double
   Dim amtcr As Double
   Dim PaymentBalance As Double
   For iLoop = flxACHistory.Rows - 1 To 1 Step -1
        amtBalance = amtBalance + Val(flxACHistory.TextMatrix(iLoop, 9))
        amtcr = amtcr + Val(flxACHistory.TextMatrix(iLoop, 8))
        If flxACHistory.TextMatrix(iLoop, 0) = "-" Then
          PaymentBalance = PaymentBalance + flxACHistory.TextMatrix(iLoop, 6)
       End If
       If flxACHistory.TextMatrix(iLoop, 0) = "+" Then
          flxACHistory.TextMatrix(iLoop, 7) = flxACHistory.TextMatrix(iLoop, 7) - PaymentBalance 'Balance
          flxACHistory.TextMatrix(iLoop, 7) = Round(flxACHistory.TextMatrix(iLoop, 7), 2)
          If flxACHistory.TextMatrix(iLoop, 7) = "0" Then
            flxACHistory.TextMatrix(iLoop, 7) = ""
          End If
         'amtBalance = 0
          'after deduct make payment balance=0
         
          PaymentBalance = 0
       End If
        If flxACHistory.TextMatrix(iLoop, 0) = "+" Or flxACHistory.TextMatrix(iLoop, 0) = "" Then
            flxACHistory.RowHeight(iLoop) = 280
        End If
   Next
    'MsgBox "Debit :" & amtBalance & "Credit :" & amtcr
   adoPty.Close
   Set adoPty = Nothing
   Set adoPtyDtl = Nothing
   Dim iRow As Integer
    If chkShowOutstanding.Value = 1 Then
        For iRow = 1 To flxACHistory.Rows - 1
            If Val(flxACHistory.TextMatrix(iRow, 7)) = 0 Then
                flxACHistory.RowHeight(iRow) = 0
            End If
        Next
    Else
        For iRow = 1 To flxACHistory.Rows - 1
            If Val(flxACHistory.TextMatrix(iRow, 7)) = 0 Then
                If flxACHistory.TextMatrix(iRow, 0) = "-" Then
                    flxACHistory.RowHeight(iRow) = 0
                Else
                    flxACHistory.RowHeight(iRow) = 280
                End If

            End If
        Next
    End If
   flxACHistory.row = 0
  ' MsgBox flxACHistory.Rows
   'flxACHistory.row = 0
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
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            bHandled = False
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

Private Sub txtSupplierSearchName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        flxSupplierList.SetFocus
    End If
End Sub

Private Sub txtSupplierSearchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplierList.SetFocus
    End If
End Sub

Private Sub txtTaxVatNumber_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdUpdateSuAddress.SetFocus
    End If
End Sub
