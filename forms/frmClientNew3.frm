VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientNew3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Myriad Condensed Web"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientNew3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12060
   Begin TabDlg.SSTab tabMain 
      Height          =   5295
      Left            =   80
      TabIndex        =   96
      Top             =   2600
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmClientNew3.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape1(5)"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "cmdClientDetailsSave"
      Tab(0).Control(4)=   "cmdClientDetailsEdit"
      Tab(0).Control(5)=   "cmdClientDetailsCancel"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Property"
      TabPicture(1)   =   "frmClientNew3.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblImageName"
      Tab(1).Control(1)=   "cmdImgLeftMove"
      Tab(1).Control(2)=   "imgPremises"
      Tab(1).Control(3)=   "imgList"
      Tab(1).Control(4)=   "tvwLandLord"
      Tab(1).Control(5)=   "fraOccupied"
      Tab(1).Control(6)=   "fraType"
      Tab(1).Control(7)=   "cmdImgDelete"
      Tab(1).Control(8)=   "cmdUploadImageAdd"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Agreement"
      TabPicture(2)   =   "frmClientNew3.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Shape2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(21)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label4"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cboRecharge_NONEED"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(22)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cboCHARGE_BASIS"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboCHARGE_METHOD"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label5"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cboCHARGE_TYPE"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cboDEMAND_TYPE"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label7(0)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label7(1)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label7(4)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label7(5)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label7(6)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label7(7)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label7(8)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label7(11)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label7(12)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cboFund"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label7(2)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label7(3)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "cboHandling"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label7(9)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label7(10)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtNtDueDate"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "cboFrequency"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtAnnualCharge"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtSTART_DATE"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtEND_DATE"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "flxAgreement"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "cmdAgmntEdit"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "cmdAgmntSave"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "txtRRPA"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "cboProperty"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "cmdAgmntAddNew"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "cmdAgmntCancel"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "txtAGREEMENT_ID"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "cmdSecondaryCode_NONEED"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "cmdAgrTopSave"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "txtOWNERSHIP_PERCENT_NONEED"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "cmdAgrTopEdit"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "txtREVIEW_DATE"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "txtNOTICE_DAYS"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).ControlCount=   47
      TabCaption(3)   =   "Bank/Payment Details"
      TabPicture(3)   =   "frmClientNew3.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape1(1)"
      Tab(3).Control(1)=   "fraBank(0)"
      Tab(3).Control(2)=   "fraBank(1)"
      Tab(3).Control(3)=   "Frame14"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Account History"
      TabPicture(4)   =   "frmClientNew3.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Shape1(2)"
      Tab(4).Control(1)=   "MSHFlexGrid1"
      Tab(4).Control(2)=   "Picture2"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Global Settings"
      TabPicture(5)   =   "frmClientNew3.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Shape1(3)"
      Tab(5).Control(1)=   "Label83(0)"
      Tab(5).Control(2)=   "Label83(13)"
      Tab(5).Control(3)=   "tabDates"
      Tab(5).Control(4)=   "cmdGSCancel"
      Tab(5).Control(5)=   "cmdGSEdit"
      Tab(5).Control(6)=   "cmdGSSave"
      Tab(5).Control(7)=   "txtFeeIsuDays"
      Tab(5).Control(8)=   "txtPayIsuDays"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "Memo/Attachemnt"
      TabPicture(6)   =   "frmClientNew3.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Shape1(4)"
      Tab(6).Control(1)=   "txtNote"
      Tab(6).Control(2)=   "cmdUnitMemoEdit"
      Tab(6).Control(3)=   "cmdUnitMemoSave"
      Tab(6).Control(4)=   "cmdUnitMemoCancel"
      Tab(6).Control(5)=   "Frame17"
      Tab(6).ControlCount=   6
      Begin VB.TextBox txtPayIsuDays 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -66120
         TabIndex        =   180
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox txtFeeIsuDays 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72720
         TabIndex        =   179
         Top             =   840
         Width           =   915
      End
      Begin VB.CommandButton cmdGSSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67680
         TabIndex        =   178
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -69480
         TabIndex        =   177
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -65880
         TabIndex        =   176
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtNOTICE_DAYS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   43
         Top             =   945
         Width           =   1240
      End
      Begin VB.TextBox txtREVIEW_DATE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   41
         Top             =   945
         Width           =   1035
      End
      Begin VB.CommandButton cmdAgrTopEdit 
         Caption         =   "Edit"
         Height          =   345
         Left            =   10485
         TabIndex        =   45
         Top             =   885
         Width           =   1095
      End
      Begin VB.TextBox txtOWNERSHIP_PERCENT_NONEED 
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
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   450
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdAgrTopSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10485
         TabIndex        =   44
         Top             =   450
         Width           =   1095
      End
      Begin VB.CommandButton cmdSecondaryCode_NONEED 
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
         Left            =   15360
         TabIndex        =   48
         Top             =   855
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAGREEMENT_ID 
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
         Left            =   11760
         TabIndex        =   161
         Top             =   2160
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton cmdAgmntCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   360
         Left            =   10560
         TabIndex        =   60
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgmntAddNew 
         Caption         =   "Add New"
         Height          =   360
         Left            =   6120
         TabIndex        =   57
         Top             =   4800
         Width           =   1215
      End
      Begin VB.ComboBox cboProperty 
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   40
         Top             =   450
         Width           =   3015
      End
      Begin VB.TextBox txtRRPA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   450
         Width           =   1240
      End
      Begin VB.CommandButton cmdUploadImageAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -64440
         TabIndex        =   149
         Top             =   3960
         Width           =   555
      End
      Begin VB.CommandButton cmdImgDelete 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -63840
         TabIndex        =   148
         Top             =   3960
         Width           =   555
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74760
         TabIndex        =   143
         Top             =   4200
         Width           =   11535
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   435
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   435
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   145
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   435
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   147
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
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.CommandButton cmdClientDetailsCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64680
         TabIndex        =   29
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientDetailsEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68040
         TabIndex        =   27
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientDetailsSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66360
         TabIndex        =   28
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74760
         TabIndex        =   141
         Top             =   2760
         Width           =   11535
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "Canc&el"
            Enabled         =   0   'False
            Height          =   360
            Left            =   8520
            TabIndex        =   76
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditBank 
            Caption         =   "&Edit"
            Height          =   360
            Left            =   5280
            TabIndex        =   74
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteBank 
            Caption         =   "&Delete"
            Height          =   360
            Left            =   10080
            TabIndex        =   77
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   360
            Left            =   6960
            TabIndex        =   75
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddNewBank 
            Caption         =   "&Add New"
            Height          =   360
            Left            =   3720
            TabIndex        =   73
            Top             =   2020
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOtherBankDetails 
            Height          =   1785
            Left            =   120
            TabIndex        =   142
            Top             =   195
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3149
            _Version        =   393216
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
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -64680
         TabIndex        =   140
         Top             =   3780
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -66300
         TabIndex        =   139
         Top             =   3780
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         Height          =   435
         Left            =   -67920
         TabIndex        =   138
         Top             =   3780
         Width           =   1350
      End
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   137
         Top             =   480
         Width           =   11595
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   1335
         Left            =   -67440
         ScaleHeight     =   1305
         ScaleWidth      =   4185
         TabIndex        =   133
         Top             =   480
         Width           =   4215
         Begin VB.TextBox Text23 
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
            Left            =   1920
            TabIndex        =   78
            Top             =   120
            Width           =   2000
         End
         Begin VB.TextBox Text22 
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
            Left            =   1920
            TabIndex        =   79
            Top             =   480
            Width           =   2000
         End
         Begin VB.TextBox Text21 
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
            Left            =   1920
            TabIndex        =   80
            Top             =   840
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   136
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Received (YTD):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   135
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Receivable (YTD):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   134
            Top             =   840
            Width           =   1710
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Account Details:"
         Height          =   2295
         Index           =   1
         Left            =   -67800
         TabIndex        =   127
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtBacsRef 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   72
            Top             =   1800
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_AC_NUM 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   71
            Top             =   1440
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_SC 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   70
            Top             =   1080
            Width           =   2800
         End
         Begin VB.TextBox txtBank_AC_Name 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   69
            Top             =   720
            Width           =   2800
         End
         Begin MSForms.ComboBox cboPaymentMethod 
            Height          =   285
            Left            =   1560
            TabIndex        =   68
            Top             =   240
            Width           =   2800
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4939;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Condensed Web"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   132
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BACS REF:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   60
            Left            =   120
            TabIndex        =   131
            Top             =   1800
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   59
            Left            =   120
            TabIndex        =   130
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   58
            Left            =   120
            TabIndex        =   129
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   57
            Left            =   120
            TabIndex        =   128
            Top             =   720
            Width           =   1050
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Bank Details:"
         Height          =   2295
         Index           =   0
         Left            =   -74760
         TabIndex        =   121
         Top             =   480
         Width           =   5295
         Begin VB.CommandButton cmdNewBank 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtBank_ID_ 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   122
            Top             =   1920
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtBANK_POST_CODE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   67
            Top             =   1920
            Width           =   1395
         End
         Begin VB.TextBox txtBANK_ADDRESS3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   66
            Top             =   1560
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   65
            Top             =   1260
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   64
            Top             =   960
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_NAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   63
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
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   61
            Top             =   240
            Width           =   3195
            VariousPropertyBits=   1820346399
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5636;503"
            TextColumn      =   1
            ColumnCount     =   6
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Condensed Web"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   125
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   124
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   123
            Top             =   600
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdAgmntSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   9080
         TabIndex        =   59
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgmntEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   360
         Left            =   7600
         TabIndex        =   58
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame fraType 
         BackColor       =   &H80000016&
         Caption         =   "CLIENT"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -70305
         TabIndex        =   118
         Top             =   360
         Width           =   3720
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   990
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1770
            Width           =   1455
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1380
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   120
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   119
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame fraOccupied 
         BackColor       =   &H80000016&
         Caption         =   "Tenancy Details:"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -70320
         TabIndex        =   108
         Top             =   2640
         Width           =   3735
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedFr 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   35
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtPreTenancyType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtPreRentRvw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   38
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   117
            Top             =   2160
            Width           =   1260
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   115
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   114
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   113
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant ID:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   112
            Top             =   195
            Width           =   705
         End
         Begin VB.Label lblTenantIDLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantID"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   1560
            MouseIcon       =   "frmClientNew3.frx":098E
            MousePointer    =   99  'Custom
            TabIndex        =   111
            Top             =   195
            Width           =   645
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant Name:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   110
            Top             =   450
            Width           =   975
         End
         Begin VB.Label lblTenantNameLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantName"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   1560
            MousePointer    =   99  'Custom
            TabIndex        =   109
            Top             =   450
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Client Statement Address:"
         Height          =   4095
         Left            =   -68760
         TabIndex        =   105
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtClientOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtClientOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   107
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   106
            Top             =   2160
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Client Address:"
         Height          =   4575
         Left            =   -74640
         TabIndex        =   97
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            TabIndex        =   18
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            TabIndex        =   19
            Top             =   2565
            Width           =   2655
         End
         Begin VB.TextBox txtClientMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            TabIndex        =   20
            Top             =   3000
            Width           =   2655
         End
         Begin VB.TextBox txtClientPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            TabIndex        =   21
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            TabIndex        =   22
            Top             =   3960
            Width           =   2655
         End
         Begin VB.TextBox txtClientAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtClientAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtClientPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtClientAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   104
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   103
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   102
            Top             =   3000
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   101
            Top             =   3480
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   100
            Top             =   2160
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   99
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Myriad Condensed Web"
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
            TabIndex        =   98
            Top             =   600
            Width           =   570
         End
      End
      Begin MSComctlLib.TreeView tvwLandLord 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   150
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8493
         _Version        =   393217
         Indentation     =   441
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   151
         Top             =   1920
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5636
         _Version        =   393216
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
         _Band(0).Cols   =   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   -74880
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew3.frx":0C98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew3.frx":1572
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew3.frx":1E4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew3.frx":2726
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgreement 
         Height          =   2835
         Left            =   120
         TabIndex        =   162
         Top             =   1920
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5001
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   -2147483638
         ForeColorSel    =   4210752
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9
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
      Begin TabDlg.SSTab tabDates 
         Height          =   3075
         Left            =   -74040
         TabIndex        =   181
         Top             =   1200
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   5424
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Monthly Payment Dates"
         TabPicture(0)   =   "frmClientNew3.frx":2A40
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraPaymentDate(0)"
         Tab(0).Control(1)=   "fraPaymentDate(1)"
         Tab(0).Control(2)=   "fraPaymentDate(2)"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Quarterly Payment Dates"
         TabPicture(1)   =   "frmClientNew3.frx":2A5C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraPaymentDate(3)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Half Yearly payments"
         TabPicture(2)   =   "frmClientNew3.frx":2A78
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraPaymentDate(4)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Yearly payments"
         TabPicture(3)   =   "frmClientNew3.frx":2A94
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "fraPaymentDate(5)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin VB.Frame fraPaymentDate 
            Caption         =   "Yearly Payment Date"
            Enabled         =   0   'False
            Height          =   975
            Index           =   5
            Left            =   2760
            TabIndex        =   264
            Top             =   1080
            Width           =   3135
            Begin VB.ComboBox cboYDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   840
               TabIndex        =   266
               Top             =   360
               Width           =   615
            End
            Begin VB.ComboBox cboYMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               TabIndex        =   265
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Once:"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   67
               Left            =   210
               TabIndex        =   267
               Top             =   360
               Width           =   390
            End
         End
         Begin VB.Frame fraPaymentDate 
            Caption         =   "Half Yearly Payment Dates"
            Enabled         =   0   'False
            Height          =   1575
            Index           =   4
            Left            =   -72240
            TabIndex        =   257
            Top             =   960
            Width           =   3135
            Begin VB.ComboBox cboHMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1560
               TabIndex        =   261
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboHDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   840
               TabIndex        =   260
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboHMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1560
               TabIndex        =   259
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboHDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   840
               TabIndex        =   258
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "First"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   51
               Left            =   435
               TabIndex        =   263
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Second"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   52
               Left            =   225
               TabIndex        =   262
               Top             =   840
               Width           =   510
            End
         End
         Begin VB.Frame fraPaymentDate 
            Caption         =   "Quarterly Payment Dates"
            Enabled         =   0   'False
            Height          =   2295
            Index           =   3
            Left            =   -72120
            TabIndex        =   244
            Top             =   600
            Width           =   3015
            Begin VB.ComboBox cboQMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1440
               TabIndex        =   252
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboQDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   720
               TabIndex        =   251
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox cboQMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   1440
               TabIndex        =   250
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox cboQDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   720
               TabIndex        =   249
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox cboQMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1440
               TabIndex        =   248
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboQDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   720
               TabIndex        =   247
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboQMth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1440
               TabIndex        =   246
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboQDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   720
               TabIndex        =   245
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fourth"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   50
               Left            =   150
               TabIndex        =   256
               Top             =   1800
               Width           =   465
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Third"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   49
               Left            =   255
               TabIndex        =   255
               Top             =   1320
               Width           =   360
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Second"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   48
               Left            =   105
               TabIndex        =   254
               Top             =   840
               Width           =   510
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "First"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   47
               Left            =   315
               TabIndex        =   253
               Top             =   360
               Width           =   300
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Quarterly Payment Dates"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   -72180
            TabIndex        =   231
            Top             =   960
            Width           =   3015
            Begin VB.ComboBox cboM4 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   239
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboD4 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   238
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox cboM3 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   237
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox cboD3 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   236
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox cboM2 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   235
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboD2 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   234
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboM1 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   233
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboD1 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   232
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Fourth"
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
               Left            =   120
               TabIndex        =   243
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               Caption         =   "Third"
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
               Left            =   120
               TabIndex        =   242
               Top             =   1320
               Width           =   615
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               Caption         =   "Second"
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
               Left            =   120
               TabIndex        =   241
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               Caption         =   "First"
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
               Left            =   360
               TabIndex        =   240
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Frame fraPaymentDate 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Index           =   2
            Left            =   -68700
            TabIndex        =   218
            Top             =   420
            Width           =   3015
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   11
               Left            =   1440
               TabIndex        =   226
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   11
               Left            =   720
               TabIndex        =   225
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   1440
               TabIndex        =   224
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   720
               TabIndex        =   223
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   1440
               TabIndex        =   222
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   720
               TabIndex        =   221
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   1440
               TabIndex        =   220
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   720
               TabIndex        =   219
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "12th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   46
               Left            =   315
               TabIndex        =   230
               Top             =   1860
               Width           =   300
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "11th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   45
               Left            =   315
               TabIndex        =   229
               Top             =   1380
               Width           =   300
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "10th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   44
               Left            =   315
               TabIndex        =   228
               Top             =   900
               Width           =   300
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "9th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   43
               Left            =   390
               TabIndex        =   227
               Top             =   420
               Width           =   225
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Half Yearly Payment Dates"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   -71580
            TabIndex        =   211
            Top             =   1260
            Width           =   3135
            Begin VB.ComboBox cboM6 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   215
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboD6 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   214
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboM5 
               Enabled         =   0   'False
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
               Left            =   1560
               TabIndex        =   213
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboD5 
               Enabled         =   0   'False
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
               Left            =   840
               TabIndex        =   212
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "First"
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
               Left            =   240
               TabIndex        =   217
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Second"
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
               Left            =   120
               TabIndex        =   216
               Top             =   840
               Width           =   615
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Yearly Payment Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   -71317
            TabIndex        =   208
            Top             =   1440
            Width           =   2535
            Begin VB.ComboBox cboD7 
               Enabled         =   0   'False
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
               Left            =   240
               TabIndex        =   210
               Top             =   360
               Width           =   615
            End
            Begin VB.ComboBox cboM7 
               Enabled         =   0   'False
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
               Left            =   960
               TabIndex        =   209
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame fraPaymentDate 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Index           =   1
            Left            =   -71700
            TabIndex        =   195
            Top             =   420
            Width           =   3015
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   720
               TabIndex        =   203
               Top             =   360
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   1440
               TabIndex        =   202
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   720
               TabIndex        =   201
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   1440
               TabIndex        =   200
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   720
               TabIndex        =   199
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   1440
               TabIndex        =   198
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   720
               TabIndex        =   197
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   1440
               TabIndex        =   196
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "7th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   41
               Left            =   390
               TabIndex        =   207
               Top             =   1380
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "5th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   39
               Left            =   390
               TabIndex        =   206
               Top             =   420
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "6th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   40
               Left            =   390
               TabIndex        =   205
               Top             =   900
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "8th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   42
               Left            =   390
               TabIndex        =   204
               Top             =   1860
               Width           =   225
            End
         End
         Begin VB.Frame fraPaymentDate 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Index           =   0
            Left            =   -74640
            TabIndex        =   182
            Top             =   420
            Width           =   3015
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1440
               TabIndex        =   190
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   720
               TabIndex        =   189
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   1440
               TabIndex        =   188
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   720
               TabIndex        =   187
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1440
               TabIndex        =   186
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   720
               TabIndex        =   185
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboMonth 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1440
               TabIndex        =   184
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboDay 
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   720
               TabIndex        =   183
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "4th"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   38
               Left            =   390
               TabIndex        =   194
               Top             =   1860
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "3rd"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   37
               Left            =   390
               TabIndex        =   193
               Top             =   1380
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "2nd"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   36
               Left            =   360
               TabIndex        =   192
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "1st"
               BeginProperty Font 
                  Name            =   "Myriad Condensed Web"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   240
               TabIndex        =   191
               Top             =   420
               Width           =   375
            End
         End
      End
      Begin MSForms.TextBox txtEND_DATE 
         Height          =   285
         Left            =   8580
         TabIndex        =   272
         Top             =   1620
         Width           =   860
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1517;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSTART_DATE 
         Height          =   285
         Left            =   7695
         TabIndex        =   271
         Top             =   1620
         Width           =   860
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1517;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAnnualCharge 
         Height          =   285
         Left            =   7020
         TabIndex        =   270
         Top             =   1620
         Width           =   660
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "1164;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "Issue Payable:                                  (days)"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   -67200
         TabIndex        =   269
         Top             =   840
         Width           =   2400
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "Issue Fee/Charges:                                    (days)"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -74160
         TabIndex        =   268
         Top             =   840
         Width           =   2775
      End
      Begin MSForms.ComboBox cboFrequency 
         Height          =   285
         Left            =   9450
         TabIndex        =   55
         Top             =   1620
         Width           =   1440
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2540;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNtDueDate 
         Height          =   285
         Left            =   10915
         TabIndex        =   56
         Top             =   1620
         Width           =   860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1517;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   10
         Left            =   10915
         TabIndex        =   175
         Top             =   1395
         Width           =   885
         VariousPropertyBits=   276824083
         Caption         =   "Next Due Dt:"
         Size            =   "1561;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   9
         Left            =   9450
         TabIndex        =   174
         Top             =   1395
         Width           =   750
         VariousPropertyBits=   276824083
         Caption         =   "Frequency"
         Size            =   "1323;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboHandling 
         Height          =   285
         Left            =   3915
         TabIndex        =   52
         Top             =   1620
         Width           =   1095
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "1940;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   300
         Index           =   3
         Left            =   3915
         TabIndex        =   173
         Top             =   1395
         Width           =   1290
         VariousPropertyBits=   8388627
         Caption         =   "Manual/Auto"
         Size            =   "2275;529"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   2
         Left            =   2820
         TabIndex        =   172
         Top             =   1395
         Width           =   600
         VariousPropertyBits=   276824083
         Caption         =   "Fund"
         Size            =   "1058;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFund 
         Height          =   285
         Left            =   2820
         TabIndex        =   51
         Top             =   1620
         Width           =   1095
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1931;503"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;5000"
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   12
         Left            =   6240
         TabIndex        =   171
         Top             =   960
         Width           =   840
         VariousPropertyBits=   276824083
         Caption         =   "Notice Days"
         Size            =   "1482;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   11
         Left            =   240
         TabIndex        =   170
         Top             =   945
         Width           =   915
         VariousPropertyBits=   276824083
         Caption         =   "Review Date:"
         Size            =   "1614;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   8
         Left            =   8580
         TabIndex        =   169
         Top             =   1395
         Width           =   645
         VariousPropertyBits=   276824083
         Caption         =   "End Date"
         Size            =   "1138;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   7
         Left            =   7695
         TabIndex        =   168
         Top             =   1395
         Width           =   735
         VariousPropertyBits=   276824083
         Caption         =   "Start Date"
         Size            =   "1296;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   420
         Index           =   6
         Left            =   7020
         TabIndex        =   167
         Top             =   1395
         Width           =   750
         VariousPropertyBits=   276824083
         Caption         =   "Amount"
         Size            =   "1323;741"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   5
         Left            =   6060
         TabIndex        =   166
         Top             =   1395
         Width           =   900
         VariousPropertyBits=   276824083
         Caption         =   "Charge Basis"
         Size            =   "1587;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   300
         Index           =   4
         Left            =   5010
         TabIndex        =   165
         Top             =   1395
         Width           =   1035
         VariousPropertyBits=   276824083
         Caption         =   "Chrg Method"
         Size            =   "1826;529"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   1
         Left            =   1755
         TabIndex        =   164
         Top             =   1395
         Width           =   1215
         VariousPropertyBits=   276824083
         Caption         =   "Demand Type"
         Size            =   "2143;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   163
         Top             =   1395
         Width           =   885
         VariousPropertyBits=   276824083
         Caption         =   "Charge Type"
         Size            =   "1561;370"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDEMAND_TYPE 
         Height          =   285
         Left            =   1755
         TabIndex        =   50
         Top             =   1620
         Width           =   1050
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1852;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboCHARGE_TYPE 
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   1620
         Width           =   1635
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2884;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   30
         Left            =   0
         TabIndex        =   160
         Top             =   1365
         Width           =   11895
         BackColor       =   -2147483629
         Size            =   "20981;53"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   4845
         Index           =   5
         Left            =   -74925
         Top             =   360
         Width           =   11745
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   4845
         Index           =   4
         Left            =   -74925
         Top             =   360
         Width           =   11745
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   4125
         Index           =   3
         Left            =   -74325
         Top             =   720
         Width           =   10305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   4845
         Index           =   2
         Left            =   -74920
         Top             =   360
         Width           =   11745
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   4845
         Index           =   1
         Left            =   -74920
         Top             =   360
         Width           =   11745
      End
      Begin MSForms.ComboBox cboCHARGE_METHOD 
         Height          =   285
         Left            =   5010
         TabIndex        =   53
         Top             =   1620
         Width           =   1035
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1826;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboCHARGE_BASIS 
         Height          =   285
         Left            =   6060
         TabIndex        =   54
         Top             =   1620
         Width           =   950
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1676;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Ownership Percentage:             %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   11760
         TabIndex        =   159
         Top             =   450
         Visible         =   0   'False
         Width           =   2370
      End
      Begin MSForms.ComboBox cboRecharge_NONEED 
         Height          =   315
         Left            =   12840
         TabIndex        =   47
         Top             =   855
         Visible         =   0   'False
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   11640
         TabIndex        =   158
         Top             =   855
         Visible         =   0   'False
         Width           =   915
         BackColor       =   -2147483645
         VariousPropertyBits=   276824091
         Caption         =   "Recharges:"
         Size            =   "1614;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent Receivable (p.a.):"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   6240
         TabIndex        =   157
         Top             =   450
         Width           =   1530
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Left            =   6240
         TabIndex        =   156
         Top             =   740
         Width           =   3645
         ForeColor       =   128
         BackColor       =   -2147483645
         VariousPropertyBits=   8388627
         Caption         =   "Rent payable is based on rent received from lease holders."
         Size            =   "6429;344"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image imgPremises 
         BorderStyle     =   1  'Fixed Single
         Height          =   3090
         Left            =   -66360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3090
      End
      Begin MSForms.CommandButton cmdImgLeftMove 
         Height          =   315
         Left            =   -65040
         TabIndex        =   154
         Top             =   3960
         Width           =   555
         Size            =   "979;556"
         Picture         =   "frmClientNew3.frx":2AB0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblImageName 
         Height          =   195
         Left            =   -66360
         TabIndex        =   153
         Top             =   480
         Width           =   3120
         Caption         =   "Image Name:"
         Size            =   "5503;344"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   152
         Top             =   450
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   975
         Left            =   120
         Top             =   360
         Width           =   11655
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   0
         Top             =   1395
         Width           =   11895
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   380
      Left            =   10760
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveClient 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   380
      Left            =   4352
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteClient 
      Caption         =   "&Delete"
      Height          =   380
      Left            =   8624
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditClient 
      Caption         =   "&Edit"
      Height          =   380
      Left            =   2216
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelChange 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6488
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNewClient 
      Caption         =   "&New"
      Height          =   380
      Left            =   80
      TabIndex        =   8
      Top             =   1920
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
      Height          =   1695
      Left            =   80
      ScaleHeight     =   1665
      ScaleWidth      =   11865
      TabIndex        =   82
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton cmdClient 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   2355
      End
      Begin VB.TextBox txtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2620
      End
      Begin VB.TextBox txtYearEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   2000
      End
      Begin VB.TextBox txtVATReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Width           =   2000
      End
      Begin VB.ListBox lstResidency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         ItemData        =   "frmClientNew3.frx":2F02
         Left            =   9285
         List            =   "frmClientNew3.frx":2F0C
         TabIndex        =   81
         Top             =   120
         Width           =   2000
      End
      Begin MSForms.ComboBox cboLandLordSageCustAC 
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   840
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
         FontName        =   "Myriad Condensed Web"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         Object.Width           =   "1763;4233"
      End
      Begin MSForms.ComboBox cboLandLordSageSuppAC 
         Height          =   285
         Left            =   1605
         TabIndex        =   4
         Top             =   1200
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
         FontName        =   "Myriad Condensed Web"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         Object.Width           =   "1762;4233"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year End:"
         Height          =   225
         Index           =   7
         Left            =   7800
         TabIndex        =   90
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX/VAT Number:"
         Height          =   225
         Index           =   6
         Left            =   7800
         TabIndex        =   89
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
         Height          =   225
         Index           =   5
         Left            =   7800
         TabIndex        =   88
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Residency:"
         Height          =   225
         Index           =   4
         Left            =   7800
         TabIndex        =   87
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Supplier A/C:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   86
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Customer A/C:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   85
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   645
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
      TabIndex        =   93
      Top             =   3787
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
         TabIndex        =   94
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
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   12195
      TabIndex        =   95
      Top             =   2400
      Width           =   12255
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
   Begin MSAdodcLib.Adodc adoFeeTypes 
      Height          =   375
      Left            =   1800
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
   Begin MSAdodcLib.Adodc adoFreq 
      Height          =   375
      Left            =   3600
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
      Caption         =   "FrequencyType"
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
   Begin MSAdodcLib.Adodc adoChargeType 
      Height          =   375
      Left            =   5520
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
      Caption         =   "FrequencyType"
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
   Begin VB.PictureBox picClientList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1320
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   91
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
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
         TabIndex        =   92
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2175
         Left            =   45
         TabIndex        =   155
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmClientNew3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LOAD_CLINT_CLIENTID As String

Private bDefaultAccount As Boolean
Private szPropertyID As String
Private iRecharge As Integer, iSlectedRow As Integer
Private bGlobalData As Boolean
Private bNewEdit As Boolean, bBankNewEdit As Boolean
Private IMAGE_FILE_NAME_ As String
Private szaPremisisIDType() As String

Private ADD_NEW_CLIENT As Boolean
Private AGREEMENT_EDIT_MODE As Boolean
Private AGREEMENT_ADDNEW_MODE As Boolean

Private Sub cboBank_ID_Click()
   txtBANK_NAME.text = cboBank_ID.Column(1)
   txtBANK_ADDRESS1.text = cboBank_ID.Column(3)
   txtBANK_ADDRESS2.text = cboBank_ID.Column(5)
   txtBANK_ADDRESS3.text = cboBank_ID.Column(6)
   txtBANK_POST_CODE.text = cboBank_ID.Column(4)
   txtBANK_SC.text = cboBank_ID.Column(2)
End Sub
'
'Private Sub cboCHARGE_TYPE_Click()
'   If cboCHARGE_TYPE.text = "RENT PAYABLE" Then
'      cboCHARGE_BASIS.Locked = True
'      txtAnnualCharge.Locked = True
'      cboCHARGE_BASIS.text = "%"
'      txtAnnualCharge.text = "100"
'   Else
'      cboCHARGE_BASIS.Locked = False
'      txtAnnualCharge.Locked = False
'   End If
'End Sub

Private Sub cboCHARGE_BASIS_Click()
   If cboCHARGE_BASIS.text = "%" And cboCHARGE_METHOD.text = "FIXED" Then
      MsgBox "Charge Method is ''Fixed'' in this case % is not valid choice."
      cboCHARGE_BASIS.text = ""
   End If
End Sub

Private Sub cboCHARGE_METHOD_Click()
'   If cboCHARGE_BASIS.text = "%" And cboCHARGE_METHOD.text = "FIXED" Then
'      MsgBox "Charge Basis is ''%'' in this case ''Fixed'' is not valid choice."
'      cboCHARGE_METHOD.text = ""
'   End If
End Sub

Private Sub cboFrequency_Change()
   SetNextDueDt
End Sub

Private Sub cboFrequency_Click()
   If Not cmdAgmntCancel.Enabled Then Exit Sub

   If txtSTART_DATE.text = "" Or txtEND_DATE.text = "" Then
      MsgBox "Before selecting the frequency you have to enter start date and end date.", vbCritical + vbOKOnly, "Frequency"
      If txtSTART_DATE.text = "" Then txtSTART_DATE.SetFocus
      If txtEND_DATE.text = "" Then txtEND_DATE.SetFocus
      Exit Sub
   End If
End Sub

Private Sub cboLandLordSageCustAC_LostFocus()
   If ADD_NEW_CLIENT Then
      txtClientID.text = cboLandLordSageCustAC.text
   End If
End Sub

Private Sub cboLandLordSageSuppAC_GotFocus()
   If cboLandLordSageCustAC.text = "" Then
      MsgBox "Please enter the Client's SAGE customer account number", vbInformation + vbOKOnly, "SAGE Customer Account Number"
      cboLandLordSageCustAC.SetFocus
   End If
End Sub

Private Sub cboProperty_Click()
   Dim sSQLQuery_ As String, sFilter As String
   Dim szaPropertyID() As String
   Dim rdoConn As New RDO.rdoConnection
   Dim rstAgreement As rdoResultset

   MousePointer = vbHourglass

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   szaPropertyID = Split(cboProperty.text, " / ")
   szPropertyID = szaPropertyID(0)

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT * " & _
                "FROM ClientProAgr " & _
                "WHERE " & _
                  "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
                  "ClientProAgr.PropertyID = '" & szPropertyID & "';"
   Set rstAgreement = rdoConn.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)

   With rstAgreement
      If Not .EOF Then
         txtREVIEW_DATE.text = !REVIEW_DATE
         txtNOTICE_DAYS.text = !NOTICE_DAYS
         txtRRPA.text = !RRPA
      End If
      .Close
   End With
   Set rstAgreement = Nothing

   sSQLQuery_ = "SELECT ChargeTypes.FeeType AS CHARGE_TYPE, DemandTypes.Type as DEMAND_TYPE, " & _
                  "Fund, CHARGE_BASIS, " & _
                  "AnnualCharge, START_DATE, END_DATE, Handling, " & _
                  "AGREEMENT_ID, CHARGE_METHOD, Frequency, NtDueDate " & _
                "FROM tlbAgreement, ClientProAgr, DemandTypes, ChargeTypes " & _
                "WHERE tlbAgreement.CPA_ID = ClientProAgr.CPA_ID And " & _
                  "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
                  "DemandTypes.ID = tlbAgreement.DEMAND_TYPE And " & _
                  "ChargeTypes.ID = tlbAgreement.CHARGE_TYPE And " & _
                  "ClientProAgr.PropertyID = '" & szPropertyID & "';"

   Set rstAgreement = rdoConn.OpenResultset(sSQLQuery_, rdOpenDynamic, rdConcurReadOnly)

   If adoMain.Recordset.EOF Then
      MsgBox "There is no agreement record setup for this property. Please enter agreement details.", vbCritical + vbOKOnly, "No Agreement"

      cboProperty.Locked = True
      cmdAgmntSave.Enabled = True
      MousePointer = vbDefault
      Exit Sub
   Else
      SetFlxAgreementHeader flxAgreement, rstAgreement
   End If

   rstAgreement.Close
   rdoConn.Close
   Set rstAgreement = Nothing
   Set rdoConn = Nothing

   MousePointer = vbDefault
End Sub

Private Sub cboProperty_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboProperty.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub chkLettingFee_LostFocus()
'   If chkLettingFee.Value = True And cboLettingChrgType.ListCount = 0 Then
'      MsgBox "There is no 'Fee Charge Type' defined. Please input charge type in the Global Form.", vbCritical + vbOKOnly, "No Types"
'      chkLettingFee.Value = False
'   End If
End Sub

Private Sub chkMngFee_LostFocus()
'   If chkMngFee.Value = True And cboMngChrgType.ListCount = 0 Then
'      MsgBox "There is no 'Fee Charge Type' defined. Please input charge type in the Global Form.", vbCritical + vbOKOnly, "No Types"
'      chkMngFee.Value = False
'   End If
End Sub

Private Sub chkRentPayble_LostFocus()
'   If chkRentPayble.Value = True And cboRentChrgType.ListCount = 0 Then
'      MsgBox "There is no 'Payable Type' defined. Please input payable type in the Global Form.", vbCritical + vbOKOnly, "No Types"
'      chkRentPayble.Value = False
'   End If
End Sub

Private Sub cmdAgmntAddNew_Click()
   If MsgBox("Do you want to add new agreement?", vbQuestion + vbYesNo, "Add New Agreement") = vbNo Then Exit Sub
   AgreementButtonMode NewEntryMode
   AgreementClearMode ClearOnlyTextBoxes
   AGREEMENT_ADDNEW_MODE = True
   cboCHARGE_TYPE.SetFocus
End Sub

Private Sub cmdAgmntCancel_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   AgreementButtonMode DefaultMode
   AgreementClearMode ClearOnlyTextBoxes
End Sub

Private Sub cmdAgrTopEdit_Click()
   AgreementButtonMode EditMode
   cmdAgrTopEdit.Enabled = False
   cmdAgrTopSave.Enabled = True
End Sub

Private Sub cmdAgrTopSave_Click()
   Dim nChoice As Integer

   nChoice = MsgBox("Are you sure to save?", vbQuestion + vbYesNoCancel, "Data Saving")
   If nChoice = vbNo Then
      AgreementButtonMode DefaultMode
      AgreementClearMode ClearOnlyTextBoxes
      Exit Sub
   End If
   If nChoice = vbCancel Then
      Exit Sub
   End If

   MousePointer = vbHourglass

   Dim conAgr As New RDO.rdoConnection
   Dim rstAgr As rdoResultset, rstCPA_ID As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conAgr.CursorDriver = rdUseIfNeeded
   conAgr.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
        "FROM ClientProAgr " & _
        "WHERE " & _
            "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
            "ClientProAgr.PropertyID = '" & szPropertyID & "';"
   Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   With rstAgr
      If .EOF Then
         .AddNew
         !ClientID = txtClientID.text
         !PROPERTYID = szPropertyID
      Else
         .Edit
      End If

'      !RECHARGES = cboRecharge.Value
      !REVIEW_DATE = txtREVIEW_DATE.text
      !NOTICE_DAYS = txtNOTICE_DAYS.text

      .Update
      .Close
   End With
   
   conAgr.Close
   Set rstAgr = Nothing
   Set conAgr = Nothing

   AGREEMENT_ADDNEW_MODE = False
   
   AgreementButtonMode DefaultMode
   cmdAgrTopEdit.Enabled = True
   cmdAgrTopSave.Enabled = False
   MousePointer = vbDefault

   MsgBox "Data has been updated successfully", vbInformation, "Saved"
   Exit Sub
ErrorHandler:
   MsgBox ERR.Number & ERR.description & " ", vbCritical + vbOK, "PCM Error: 125"
   MousePointer = vbDefault
End Sub

Private Sub cmdCancelBank_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   CommandButtonEnabled True
   LockingAcText True
   NewBankText True, True
   flxOtherBankDetails_RowColChange
'   cmdNewBank.Visible = False
   cmdNewBank.Enabled = True
End Sub

Private Sub cmdClientDetailsCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText True
   CommandButtonEnable True
End Sub

Private Sub CommandButtonEnable(bEnable As Boolean)
   cmdClientDetailsEdit.Enabled = bEnable
   cmdClientDetailsSave.Enabled = Not bEnable
   cmdClientDetailsCancel.Enabled = Not bEnable
End Sub

Private Sub cmdClientDetailsEdit_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText False
   CommandButtonEnable False
End Sub

Private Sub cmdClientDetailsSave_Click()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM Client " & _
           "WHERE ClientID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   
   With rstClient
      .Edit
      !ClientAddressLine1 = txtClientAddressLine1.text
      !ClientAddressLine2 = txtClientAddressLine2.text
      !ClientAddressLine3 = txtClientAddressLine3.text
      !ClientPostCode = txtClientPostCode.text
      !ClientOfficeEmail = txtClientOfficeEmail.text
      !ClientPersonalEmail = txtClientPersonalEmail.text
      !ClientHomeTel = txtClientHomeTel.text
      !ClientMobile = txtClientMobile.text
      !ClientOfficeAddressLine1 = txtClientOfficeAddressLine1.text
      !ClientOfficeAddressLine2 = txtClientOfficeAddressLine2.text
      !ClientOfficeAddressLine3 = txtClientOfficeAddressLine3.text
      !ClientOfficePostCode = txtClientOfficePostCode.text
      !ClientOfficeTel = txtClientOfficeTel.text

      .Update
      .Close
   End With
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
   MsgBox "Data has been updated successfully", vbInformation + vbOKOnly, "Data Update"
   CommandButtonEnable True
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachment cmbFiles, "Client", txtClientID.text
   MsgBox "File has been saved successfull, Thanks"
End Sub

Private Sub cmdDeleteBank_Click()
   If MsgBox("Do you want to delete current account details?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
   
   flxOtherBankDetails.RemoveItem (flxOtherBankDetails.Row)

   flxOtherBankDetails_RowColChange
   NewBankText True, False
   cmdNewBank.Caption = "New"
   LockingAcText True
   MsgBox "Record has been deleted successfully.", vbInformation + vbOKOnly, "Delete"
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachment cmbFiles, cmbFiles.Column(2), txtClientID.text, "Client"
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdEditBank_Click()
   MousePointer = vbHourglass

   cmdNewBank.Caption = "Edit"
   bBankNewEdit = False

'   cmdNewBank.Visible = True
   LockingAcText False

   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
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
'   cmdNewBank.Visible = True
   cboBank_ID.SetFocus

   LockingAcText False
   NewBankText True, True
   cboBank_ID.Locked = False
   
   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
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

   adoBank.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT BANK_ID, BANK_NAME, SORT_CODE, " & _
                     "BANK_ADDRESS1, BANK_POST_CODE, " & _
                     "BANK_ADDRESS2, BANK_ADDRESS3 " & _
                "FROM tlbBank"

   adoBank.RecordSource = sSQLQuery_
   adoBank.CommandType = adCmdText
   adoBank.Refresh

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoBank.Recordset.RecordCount
   TotalCol = adoBank.Recordset.Fields.Count

   Dim data() As String

   ReDim data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoBank.Recordset.RecordCount - 1
       For j = 0 To adoBank.Recordset.Fields.Count - 1
           data(j, i) = IIf(IsNull(adoBank.Recordset.Fields(j)), "", adoBank.Recordset.Fields(j))
       Next j
       adoBank.Recordset.MoveNext
   Next i

   cboBank_ID.Column() = data()
End Function

Private Sub cmdAddNewClient_Click()
   If MsgBox("Do you want to add new client?", vbYesNo + vbQuestion, "Add New Client") = vbNo Then Exit Sub
   bNewEdit = True
   
   MousePointer = vbHourglass

   SageCustomerAccCombo
   SageSupplierAccCombo

   ADD_NEW_CLIENT = True
   UnlockMainClientText True
   MainCommandButtonEnable True

   txtClientName.SetFocus
   
   MousePointer = vbDefault
End Sub

Private Sub SageCustomerAccCombo()
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
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then
   
      Set oSalesRecord = oWS.CreateObject("SalesRecord")
   
      Dim TotalRow, TotalCol As Long
      Dim data() As String
      Dim i As Integer
          
      TotalRow = oSalesRecord.Count
      TotalCol = 2
      cboLandLordSageCustAC.Clear
      
      ReDim data(TotalCol, TotalRow) As String
      
      oSalesRecord.MoveFirst
      For i = 0 To TotalRow - 1
         'cboTest.AddItem adoClient.Recordset.Fields(1)
         data(0, i) = CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value)
         data(1, i) = CStr(oSalesRecord.Fields.Item("NAME").Value)
         oSalesRecord.MoveNext
      Next i
      '
      cboLandLordSageCustAC.Column() = data()
      cboLandLordSageCustAC.ColumnCount = TotalCol
      cboLandLordSageCustAC.BoundColumn = 1
   '         cboLandLordSageCustAC.TextColumn = 2
   
      'Disconnect
      oWS.Disconnect
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

Private Sub UnlockMainClientText(bUnlock As Boolean)
'   txtClientID.Locked = Not bUnlock
   txtClientName.Locked = Not bUnlock
   cboLandLordSageCustAC.Locked = Not bUnlock
   cboLandLordSageSuppAC.Locked = Not bUnlock
'   txtAcBalance.Locked = Not bUnlock
   txtVATReg.Locked = Not bUnlock
   txtYearEndDate.Locked = Not bUnlock
End Sub

Private Sub cmdAgmntEdit_Click()
   If MsgBox("Do you want to edit the agreement?", vbQuestion + vbYesNo, "Edit Agreement") = vbNo Then Exit Sub

   AgreementButtonMode EditMode
   AGREEMENT_EDIT_MODE = True
   cboCHARGE_TYPE.SetFocus
End Sub

Private Sub cmdAgmntSave_Click()
   If cboCHARGE_TYPE.text = "" Or _
         cboDEMAND_TYPE.text = "" Or _
         cboFund.text = "" Or _
         cboHandling.text = "" Or _
         cboCHARGE_METHOD.text = "" Or _
         cboCHARGE_BASIS.text = "" Or _
         txtAnnualCharge.text = "" Or _
         txtSTART_DATE.text = "" Or _
         txtEND_DATE.text = "" Or _
         cboFrequency.text = "" Then
      MsgBox "You can not leave any field blank.", vbCritical + vbOKOnly, "Saving Data"
      If cboDEMAND_TYPE.text = "" Then cboDEMAND_TYPE.SetFocus
      If cboFund.text = "" Then cboFund.SetFocus
      If cboHandling.text = "" Then cboHandling.SetFocus
      If cboCHARGE_METHOD.text = "" Then cboCHARGE_METHOD.SetFocus
      If cboCHARGE_BASIS.text = "" Then cboCHARGE_BASIS.SetFocus
      If txtAnnualCharge.text = "" Then txtAnnualCharge.SetFocus
      If txtSTART_DATE.text = "" Then txtSTART_DATE.SetFocus
      If cboFrequency.text = "" Then cboFrequency.SetFocus
      Exit Sub
   End If

   If cboCHARGE_BASIS.text = "%" And cboCHARGE_METHOD.text = "FIXED" Then
      MsgBox "In the case of ''Fixed'' charge method, charge basis can not be ''%''."
      cboCHARGE_METHOD.text = ""
      Exit Sub
   End If

   Dim nChoice As Integer

   nChoice = MsgBox("Are you sure to save?", vbQuestion + vbYesNoCancel, "Data Saving")
   If nChoice = vbNo Then
      AgreementButtonMode DefaultMode
      AgreementClearMode ClearOnlyTextBoxes
      Exit Sub
   End If
   If nChoice = vbCancel Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass

   Dim conAgr As New RDO.rdoConnection
   Dim rstAgr As rdoResultset, rstCPA_ID As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conAgr.CursorDriver = rdUseIfNeeded
   conAgr.EstablishConnection rdDriverNoPrompt

   If AGREEMENT_ADDNEW_MODE Then
      szSQL = "SELECT * " & _
              "FROM tlbAgreement;"
      Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
      rstAgr.AddNew
      AGREEMENT_ADDNEW_MODE = False

      szSQL = "SELECT CPA_ID " & _
              "FROM ClientProAgr " & _
              "WHERE " & _
                  "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
                  "ClientProAgr.PropertyID = '" & szPropertyID & "';"
      Set rstCPA_ID = conAgr.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

      rstAgr!CPA_ID = rstCPA_ID!CPA_ID
      rstAgr!AGR_DATE = Format(Now, "dd mmmm yyyy")
      rstCPA_ID.Close
      Set rstCPA_ID = Nothing
   End If

   If AGREEMENT_EDIT_MODE Then
      szSQL = "SELECT * " & _
           "FROM tlbAgreement " & _
           "WHERE AGREEMENT_ID = " & txtAGREEMENT_ID.text & ";"

      Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
      rstAgr.Edit
      AGREEMENT_EDIT_MODE = False
   End If

   With rstAgr
      !CHARGE_TYPE = cboCHARGE_TYPE.Value
      !DEMAND_TYPE = cboDEMAND_TYPE.Value
      !Fund = cboFund.Value
      !Handling = cboHandling.Value
      !CHARGE_METHOD = cboCHARGE_METHOD.text
      !CHARGE_BASIS = cboCHARGE_BASIS.text
      !AnnualCharge = txtAnnualCharge.text
      !START_DATE = Format(txtSTART_DATE.text, "dd mmmm yyyy")
      !END_DATE = Format(txtEND_DATE.text, "dd mmmm yyyy")
      !Frequency = cboFrequency.Value
      !NtDueDate = txtNtDueDate.text
      .Update
      .Close
   End With

   conAgr.Close
   Set rstAgr = Nothing
   Set conAgr = Nothing

   Call cboProperty_Click        'refresh the grid

   MousePointer = vbDefault

   cboProperty.Locked = False
   MsgBox "Agreement has been updated successfully.", vbInformation + vbOKOnly, "Agreement"

   AgreementButtonMode DefaultMode
   Exit Sub

ErrorHandler:

   rstAgr.Close
   Set rstAgr = Nothing
   conAgr.Close
   Set conAgr = Nothing

   MsgBox ERR.Number & ERR.description & " " & "Do not leave any field as blank.", vbCritical + vbOK, "PCM Error: 125"
   AgreementButtonMode DefaultMode
   MousePointer = vbDefault
End Sub

Private Sub cmdClient_Click()
   Call PrepareList

   picClientList.Top = picMain.Top + txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = picMain.Left + txtClientID.Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxClientList
   LoadAllClientFlxGrd
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDeleteClient_Click()
'===========================================================================================
'This button should be visible, because user should not get facility to delete any record.
'we should give user a facility to see or remove the recode from the current list.
'===========================================================================================
   If txtClientID.text = "" Then
      MsgBox "Please select a client to delete.", vbInformation, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Are you sure to delete current client?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If Not rstClient.EOF Then
      MsgBox "This client could not be deleted. This client has property in the database." & _
            (Chr$(13) + Chr$(10)) & "To delete this client delete all properties of this client.", vbCritical + vbOKOnly, "Delete Not Possible"
      rstClient.Close
      Set rstClient = Nothing
      conClient.Close
      Set conClient = Nothing
      Exit Sub
   End If

   rstClient.Close

   MousePointer = vbHourglass

   szSQL = "DELETE * FROM CLIENT WHERE CLIENTID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   rstClient.Close
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing

   MsgBox "Client has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"

   MousePointer = vbDefault
End Sub

Private Sub cmdEditClient_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to make change to the current client?", vbYesNo + vbQuestion, "Edit Client") = vbNo Then Exit Sub
   bNewEdit = False

   MousePointer = vbHourglass

   MainCommandButtonEnable True

   Dim szTemp As String
   
   If cboLandLordSageSuppAC.ListCount = 0 Then
      szTemp = cboLandLordSageSuppAC.text
      SageSupplierAccCombo
      cboLandLordSageSuppAC.text = szTemp
   End If
   If cboLandLordSageCustAC.ListCount = 0 Then
      szTemp = cboLandLordSageCustAC.text
      SageCustomerAccCombo
      cboLandLordSageCustAC.text = szTemp
   End If
   
   ADD_NEW_CLIENT = False
   LockingAllText False
   UnlockMainClientText True
   
   MousePointer = vbDefault
End Sub

Private Sub MainCommandButtonEnable(bEnabled As Boolean)
   cmdAddNewClient.Enabled = Not bEnabled
   cmdEditClient.Enabled = Not bEnabled
   cmdSaveClient.Enabled = bEnabled
   cmdDeleteClient.Enabled = Not bEnabled
   cmdCancelChange.Enabled = bEnabled
   
   cmdClient.Enabled = Not bEnabled
   lstResidency.Enabled = bEnabled
End Sub

Private Sub cmdFeeType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "CFT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoFeeTypes.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT CODE, VALUE " & _
              "FROM SECONDARYCODE " & _
              "WHERE SECONDARYCODE.PRIMARYCODE = 'CFT' " & _
              "ORDER BY VALUE;"

   adoFeeTypes.RecordSource = sSQLQuery_
   adoFeeTypes.CommandType = adCmdText
   adoFeeTypes.Refresh
End Sub
'
'Private Sub cmdFeeTypesCancel_Click()
'   Dim i As Integer
'
'   If MsgBox("Do you want to discard all modified/new fees?", vbQuestion + vbYesNo, "New Fees") = vbYes Then
'      i = 1
'      While i < flxFeeType.Rows
'         If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "1" Or flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "2" Then
'            flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "0"
'         End If
'         If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "3" Then
'            flxFeeType.RemoveItem i
'            i = i - 1
'         End If
'         i = i + 1
'      Wend
''      ClientGlobalSetting
''      ConfigureFlxFeeType
''      ComponentInFrameEnableMode Me, imgFeeTypes, DefaultMode
'   End If
'End Sub
'
'Private Sub cmdFeeTypesEdit_Click()
'   If flxFeeType.TextMatrix(flxFeeType.Row, flxFeeType.Cols - 1) = "0" Then
'      bNewEdit = False
''      ComponentInFrameEnableMode Me, imgFeeTypes, EditMode
'      flxFeeType.TextMatrix(flxFeeType.Row, flxFeeType.Cols - 1) = "1"
'   End If
'End Sub

Private Sub cmdFeeTypesNew_Click()
   MousePointer = vbHourglass
   
   bNewEdit = True
'   ComponentInFrameEnableMode Me, imgFeeTypes, NewEntryMode

   MousePointer = vbDefault
End Sub
'
'Private Sub cmdFeeTypesSave_Click()
'   If cmdUpdate.Enabled Then
'      MsgBox "Please update the data into grid by pressing the >> button.", vbInformation + vbOKOnly, "Save"
'      cmdUpdate.SetFocus
'      Exit Sub
'   End If
'   If MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save") = vbNo Then Exit Sub
'
'   Dim szSQL As String, i As Integer
'   Dim adoConn As New ADODB.Connection
'   Dim oResultSet As ADODB.Recordset
'
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'   Set oResultSet = New ADODB.Recordset
'
'   i = 1
'   While i < flxFeeType.Rows
'      If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "2" Then
'         szSQL = "SELECT * " & _
'              "FROM ClientGDFees " & _
'              "WHERE ThisID = " & flxFeeType.TextMatrix(i, flxFeeType.Cols - 3) & ";"  'flxFeeType.TextMatrix(i, flxFeeType.Cols - 3) -> ThisID of the record
'
'         oResultSet.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
'
'         oResultSet.Fields("FeeType").Value = flxFeeType.TextMatrix(i, 0)
'         oResultSet.Fields("Handling").Value = flxFeeType.TextMatrix(i, 1)
'         oResultSet.Fields("Frequency").Value = flxFeeType.TextMatrix(i, 2)
'         oResultSet.Fields("NtDueDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 3)), "dd mmmm yyyy")
'         oResultSet.Fields("StDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 4)), "dd mmmm yyyy")
'         oResultSet.Fields("ChargeType").Value = CByte(flxFeeType.TextMatrix(i, 5))
'         oResultSet.Update
'         oResultSet.Close
'      End If
'      If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "3" Then
'         szSQL = "SELECT * " & _
'              "FROM ClientGDFees;"
'
'         oResultSet.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
'
'         oResultSet.AddNew
'         oResultSet.Fields("FeeType").Value = flxFeeType.TextMatrix(i, 0)
'         oResultSet.Fields("Handling").Value = flxFeeType.TextMatrix(i, 1)
'         oResultSet.Fields("Frequency").Value = flxFeeType.TextMatrix(i, 2)
'         oResultSet.Fields("NtDueDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 3)), "dd mmmm yyyy")
'         oResultSet.Fields("StDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 4)), "dd mmmm yyyy")
'         oResultSet.Fields("ChargeType").Value = CByte(flxFeeType.TextMatrix(i, 5))
'         oResultSet.Fields("ParentID").Value = CByte(flxFeeType.TextMatrix(i, flxFeeType.Cols - 2))
'         oResultSet.Update
'         oResultSet.Close
'      End If
'      i = i + 1
'   Wend
'   MsgBox "Data has been saved successfully", vbInformation + vbOKOnly, "Saved"
''   ClientGlobalSetting
''   ConfigureFlxFeeType
''   ComponentInFrameEnableMode Me, imgFeeTypes, DefaultMode
'End Sub

Private Sub cmdFreq_Click()
   frmSecondaryCode.PRIMARY_CODE_SHOW = "FREQ"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1
   Dim sSQLQuery_ As String
   
   adoFreq.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT CODE, VALUE " & _
              "FROM SECONDARYCODE " & _
              "WHERE SECONDARYCODE.PRIMARYCODE = 'FREQ' " & _
              "ORDER BY VALUE;"

   adoFreq.RecordSource = sSQLQuery_
   adoFreq.CommandType = adCmdText
   adoFreq.Refresh
End Sub

Private Sub cmdGridUnitLookup_Click()
   picClientList.Visible = False
End Sub

Private Sub cmdGSCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   Dim i As Integer
   
   On Error Resume Next
   For i = 0 To 67
      Label1(i).ForeColor = vbBlack
   Next i
   
   EnableGlobalControl False
End Sub

Private Sub cmdGSEdit_Click()
   MousePointer = vbHourglass

   EnableGlobalControl True

   MousePointer = vbDefault
End Sub

Private Sub EnableGlobalControl(bEnable As Boolean)
   Dim i As Integer

'   chkLettingFee.Enabled = bEnable
'   chkMngFee.Enabled = bEnable
'   chkRentPayble.Enabled = bEnable
   For i = 0 To 5
'      If i < 3 Then fraFee(i).Enabled = bEnable
      If i < 6 Then fraPaymentDate(i).Enabled = bEnable
   Next i

   cmdGSEdit.Enabled = Not bEnable
   cmdGSSave.Enabled = bEnable
   cmdGSCancel.Enabled = bEnable
End Sub

Private Function ControlFilled() As Boolean
   ControlFilled = True
   
   If txtFeeIsuDays.text = "" Then
      Label83(0).ForeColor = vbRed
      ControlFilled = False
   End If
   If txtPayIsuDays.text = "" Then
      Label83(13).ForeColor = vbRed
      ControlFilled = False
   End If
'   If chkLettingFee.Value Then
'      If cboLettingAM.text = "" Then
'         Label1(23).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboLettingFreq.text = "" Then
'         Label1(25).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtLettingNtDueDt.text = "" Then
'         Label1(24).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtLettingStDt.text = "" Then
'         Label1(26).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboLettingChrgType.text = "" Then
'         Label1(64).ForeColor = vbRed
'         ControlFilled = False
'      End If
'   End If

'   If chkMngFee.Value Then
'      If cboMngAM.text = "" Then
'         Label1(27).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboMngFreq.text = "" Then
'         Label1(29).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtMngNtDueDt.text = "" Then
'         Label1(28).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtMngStDt.text = "" Then
'         Label1(30).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboMngChrgType.text = "" Then
'         Label1(65).ForeColor = vbRed
'         ControlFilled = False
'      End If
'   End If

'   If chkRentPayble.Value Then
'      If cboRentAM.text = "" Then
'         Label1(31).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboRentFreq.text = "" Then
'         Label1(32).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtRentNtDueDt.text = "" Then
'         Label1(33).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If txtRentStDt.text = "" Then
'         Label1(34).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If cboRentChrgType.text = "" Then
'         Label1(66).ForeColor = vbRed
'         ControlFilled = False
'      End If
'   End If

   Dim i As Integer
   
   For i = 0 To 11         'MONTHLY
      If cboDay(i).text = "" Or cboMonth(i).text = "" Then
         Label1(35 + i).ForeColor = vbRed
         ControlFilled = False
      End If
      If i < 4 Then        'QUARTERLY
         If cboQDay(i).text = "" Or cboQMth(i).text = "" Then
            Label1(47 + i).ForeColor = vbRed
            ControlFilled = False
         End If
      End If
      If i < 2 Then        'HALF YEARLY
         If cboHDay(i).text = "" Or cboHMth(i).text = "" Then
            Label1(51 + i).ForeColor = vbRed
            ControlFilled = False
         End If
      End If
   Next i
   If cboYDay.text = "" Or cboYMth.text = "" Then        'YEARLY
      Label1(67).ForeColor = vbRed
      ControlFilled = False
   End If
End Function

Private Sub cmdGSSave_Click()
   Dim sSQLQuery_ As String, i As Integer

   If Not ControlFilled Then
      MsgBox "Please type/select red marked fields.", vbCritical + vbOKOnly, "Blank Fields"
      Exit Sub
   End If

   MousePointer = vbHourglass

'Saving client's global data
   If cmdGSEdit.Caption = "Edit" Then
      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"
   Else
      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData;"
   End If

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If cmdGSEdit.Caption = "Add New" Then
      adoMain.Recordset.AddNew
   End If

   adoMain.Recordset.Fields("ClientID").Value = txtClientID.text
   
   For i = 0 To 11
      adoMain.Recordset.Fields("MonthlyDueDate" & (i + 1) & "").Value = _
                        cboDay(i).text & " " & cboMonth(i).text
      If i < 4 Then _
         adoMain.Recordset.Fields("QuarterlyDueDate" & (i + 1) & "").Value = _
                        cboQDay(i).text & " " & cboQMth(i).text
      If i < 2 Then _
         adoMain.Recordset.Fields("HalfYearlyDueDate" & (i + 1) & "").Value = _
                        cboHDay(i).text & " " & cboHMth(i).text
   Next i
   adoMain.Recordset.Fields("YearlyDueDate").Value = _
                     cboYDay.text & " " & cboYMth.text

   adoMain.Recordset.Fields("FeeIsuDays").Value = CInt(txtFeeIsuDays.text)
   adoMain.Recordset.Fields("PayIsuDays").Value = CInt(txtPayIsuDays.text)
   
   adoMain.Recordset.Update
   adoMain.Recordset.Close

   MousePointer = vbDefault

   If cmdGSEdit.Caption = "Edit" Then
      MsgBox "Data has been updated successfully", vbInformation + vbOKOnly, "Update Data"
   Else
      MsgBox "Data has been saved successfully", vbInformation + vbOKOnly, "Save Data"
   End If

   EnableGlobalControl False
End Sub

Private Sub cmdImgDelete_Click()
   If imgPremises.Picture = 0 Then Exit Sub
   If MsgBox("Are you sure to delete the image?", vbQuestion + vbYesNo, "Delete Image") = vbNo Then Exit Sub
   DeleteImage imgPremises, IMAGE_FILE_NAME_, szaPremisisIDType(0), szaPremisisIDType(1)
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdImgLeftMove_Click()
   IMAGE_FILE_NAME_ = MoveNextImage(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), IMAGE_FILE_NAME_, lblImageName)
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
      MsgBox "File has been moved from original location.", vbExclamation
      
   MousePointer = vbDefault
End Sub
'
'Private Sub cmdResidency_Click()
'   lstResidency.Top = txtResidency.Top
'   lstResidency.Left = txtResidency.Left
'   lstResidency.Visible = True
'   lstResidency.ZOrder 0
'   lstResidency.SetFocus
'End Sub

Private Sub cmdSaveBank_Click()
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String, szWhere As String

   On Error GoTo ErrorHandler

   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "New" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank;"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

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
'      cmdNewBank.Visible = False
   End If

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "Edit" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank " & _
              "WHERE BANK_ID = '" & cboBank_ID.text & "';"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

      rstBank.Edit
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update
'
      rstBank.Close

      NewBankText True, False
'      cmdNewBank.Visible = False
   End If

   If bDefaultAccount And cmdNewBank.Caption = "New" Then
      szSQL = "SELECT * " & _
              "FROM CLIENT " & _
              "WHERE CLIENTID = '" & txtClientID.text & "'"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
      With rstBank
         .Edit
         !BANK_ID = cboBank_ID.text
         .Update
         .Close
      End With
   End If

   If cmdNewBank.Caption = "Edit" Then
      szWhere = " Where BANK_AC_NUM = '" & flxOtherBankDetails.TextMatrix(iSlectedRow, 5) & "' And " & _
                     "BANK_SC = '" & flxOtherBankDetails.TextMatrix(iSlectedRow, 6) & "';"
   Else
      szWhere = ";"
   End If

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks" & szWhere
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   With rstBank
      If cmdNewBank.Caption = "New" Then
         .AddNew
      Else
         .Edit
      End If
      !CLIENT_ID = txtClientID.text
      !BANK_ID = cboBank_ID.text
      !Bank_AC_Name = txtBank_AC_Name.text
      !BANK_AC_NUM = txtBANK_AC_NUM.text
      !BANK_SC = txtBANK_SC.text
      !DEFAULT_AC = bDefaultAccount
      !PaymentMethod = cboPaymentMethod.text
      !BacsRef = txtBacsRef.text
      .Update
   End With

   If cmdNewBank.Caption = "New" Then
      MsgBox "Data has been saved successfully.", vbInformation + vbOKOnly, "Add New"
   Else
      MsgBox "Data has been updated successfully.", vbInformation + vbOKOnly, "Edit"
   End If

   LoadAllBankAC
   
'   If cmdNewBank.Caption = "New" Then
'      flxOtherBankDetails.AddItem ""
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 1) = cboBank_ID.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 2) = txtBANK_NAME.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 3) = txtBANK_POST_CODE.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 4) = txtBank_AC_Name.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 5) = txtBANK_AC_NUM.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 6) = txtBANK_SC.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 7) = IIf(bDefaultAccount, "YES", "NO")
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 8) = txtBANK_ADDRESS1.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 9) = txtBANK_ADDRESS2.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 10) = txtBANK_ADDRESS3.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 11) = cboPaymentMethod.text
'      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 12) = txtBACSRef.text
'   Else
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 1) = cboBank_ID.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 2) = txtBANK_NAME.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 3) = txtBANK_POST_CODE.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 4) = txtBank_AC_Name.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 5) = txtBANK_AC_NUM.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 6) = txtBANK_SC.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 7) = IIf(bDefaultAccount, "YES", "NO")
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 8) = txtBANK_ADDRESS1.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 9) = txtBANK_ADDRESS2.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 10) = txtBANK_ADDRESS3.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 11) = cboPaymentMethod.text
'      flxOtherBankDetails.TextMatrix(iSlectedRow, 12) = txtBACSRef.text
'   End If
   LockingAcText True
   CommandButtonEnabled True
   cmdNewBank.Enabled = True
'   cmdNewBank.Visible = False

NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub cmdSaveClient_Click()
   If MsgBox("Are you sure to save/update?", vbQuestion + vbYesNo, "Saving Data") = vbNo Then Exit Sub
   If txtClientID.text = "" Then
      MsgBox "Please type client id.", vbCritical + vbOKOnly, "Client"
      txtClientID.SetFocus
      Exit Sub
   End If
   If txtClientName.text = "" Then
      MsgBox "Please type client's name.", vbCritical + vbOKOnly, "Client"
      txtClientName.SetFocus
      Exit Sub
   End If
   If cboLandLordSageCustAC.text = "" Then
      MsgBox "Please select client's Sage Customer Account.", vbCritical + vbOKOnly, "Client"
      cboLandLordSageCustAC.SetFocus
      Exit Sub
   End If
   If cboLandLordSageSuppAC.text = "" Then
      MsgBox "Please select client's Sage Supplier Account.", vbCritical + vbOKOnly, "Client"
      cboLandLordSageSuppAC.SetFocus
      Exit Sub
   End If

   If lstResidency.ListIndex < 0 Then lstResidency.ListIndex = 0

'   If txtResidency.text = "" Then
'      MsgBox "Please select client's residency.", vbCritical + vbOKOnly, "Client"
'      txtResidency.SetFocus
'      Exit Sub
'   End If
'   If txtVATReg.text = "" Then
'      MsgBox "Please type client's VAT Registration number.", vbCritical + vbOKOnly, "Client"
'      txtVATReg.SetFocus
'      Exit Sub
'   End If
   If txtYearEndDate.text = "" Then
      MsgBox "Please type year end date.", vbCritical + vbOKOnly, "Client"
      txtYearEndDate.SetFocus
      Exit Sub
   End If
   
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   szSQL = "SELECT ClientID, ClientName, LandLordSageCustAC, " & _
                  "LandLordSageSuppAC, Residency, AcBalance, VATReg, " & _
                  "YearEndDate " & _
           "FROM Client " & _
           "WHERE ClientID = '" & txtClientID.text & "';"

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   If PostToDBUsingADODB(Me, picMain, adoConn, szSQL, bNewEdit) Then
      MsgBox "Data has been saved succfully.", vbOKOnly, "Data Save"
   Else
      MsgBox "Data has not been saved.", vbOKOnly, "Data Save"
   End If
   UnlockMainClientText False
   MainCommandButtonEnable False
End Sub
'
'Private Sub cmdSecondaryCode_Click()
'   Dim sSQLQuery As String
'   Dim adoConn As New ADODB.Connection
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "RECRG"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'
'   ' Recharge
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'RECRG'"
'   populateCombo adoConn, sSQLQuery, cboRecharge
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveMemo("Client", "ClientMemo", txtClientID.text, "ClientID", txtNote) Then
      MsgBox "Memo has been saved successfully.", vbInformation + vbOKOnly, "Memo"
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
'Private Sub cmdUpdate_Click()
'   Dim i As Integer
'
'   If Not bNewEdit Then
'      If MsgBox("Do you want to update?", vbQuestion + vbYesNo, "Update Fees") = vbNo Then Exit Sub
'
'      For i = 1 To flxFeeType.Rows - 1
'         If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "1" Then
'            flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "2"
'            Exit For
'         End If
'      Next i
'
'      flxFeeType.TextMatrix(i, 0) = cboFeeType.text
'      flxFeeType.TextMatrix(i, 1) = cboHandling.text
'      flxFeeType.TextMatrix(i, 2) = cboFrequency.text
'      flxFeeType.TextMatrix(i, 3) = txtNtDueDate.text
'      flxFeeType.TextMatrix(i, 4) = txtStartDate.text
'      flxFeeType.TextMatrix(i, 5) = cboChargeType.text
'   Else
'      If MsgBox("Do you want to add new?", vbQuestion + vbYesNo, "New Fees") = vbNo Then Exit Sub
'
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 0) = cboFeeType.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 1) = cboHandling.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 2) = cboFrequency.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 3) = txtNtDueDate.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 4) = txtStartDate.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, 5) = cboChargeType.text
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, flxFeeType.Cols - 2) = GetParentID(flxClientList.TextMatrix(flxClientList.Row, 1)) 'parent record id
'      flxFeeType.TextMatrix(flxFeeType.Rows - 1, flxFeeType.Cols - 1) = "3"           'Newly added
'      flxFeeType.AddItem ""
'   End If
'   cmdUpdate.Enabled = False
'End Sub

Private Function GetParentID(szParent As String) As String
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   szSQL = "SELECT Record_ID " & _
           "FROM ClientGlobalData " & _
           "WHERE ClientGlobalData.ClientID = '" & flxClientList.TextMatrix(flxClientList.Row, 1) & "';"
   
   Dim adoRst As ADODB.Recordset
   Set adoRst = New ADODB.Recordset

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
   
   GetParentID = adoRst.Fields("Record_ID").Value

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Function

Private Sub cmdUploadImageAdd_Click()
   If MsgBox("Do you want to add new image?", vbQuestion + vbYesNo, "Image Attachment") = vbNo Then Exit Sub
   IMAGE_FILE_NAME_ = AddNewImage(imgPremises, szaPremisisIDType(1), szaPremisisIDType(0), lblImageName)
   MsgBox "Image has been uploaded successfull."
End Sub

Private Sub flxAgreement_RowColChange()
   populateControl Me, flxAgreement
   cmdAgmntEdit.Enabled = True
End Sub

Private Sub flxClientList_Click()
   If flxClientList.TextMatrix(flxClientList.Row, 1) = "" Then Exit Sub

   Dim sSQLQuery_ As String, sFilter As String

   txtClientID.text = flxClientList.TextMatrix(flxClientList.Row, 1)

   MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   sSQLQuery_ = "SELECT * " & _
                "FROM CLIENT " & _
                "WHERE CLIENT.ClientID = '" & flxClientList.TextMatrix(flxClientList.Row, 1) & "';"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      MsgBox "Error in Database.", vbExclamation
   Else
      LoadClientProperty

      lblLoading.Caption = "Please wait, tree is building..."
      fmeLoading.Refresh
      DrawLandLordTree tvwLandLord, imgList, txtClientID.text, True

      lblLoading.Caption = "Please wait, global data is loading..."

      fmeLoading.Refresh

      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"
      adoMain.RecordSource = sSQLQuery_
      adoMain.CommandType = adCmdText
      adoMain.Refresh
      If adoMain.Recordset.RecordCount > 0 Then
         bGlobalData = True
      Else
         bGlobalData = False
      End If
      RetrieveMemo "Client", "ClientMemo", txtClientID.text, "ClientID", txtNote
   End If

   fmeLoading.Visible = False
   MousePointer = vbDefault

   picClientList.Visible = False
End Sub

Private Sub LoadFormByClient()
   Dim sSQLQuery_ As String, sFilter As String

   txtClientID.text = LOAD_CLINT_CLIENTID

   MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   sSQLQuery_ = "SELECT * " & _
                "FROM CLIENT " & _
                "WHERE CLIENT.ClientID = '" & LOAD_CLINT_CLIENTID & "';"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      MsgBox "Error in Database.", vbExclamation
   Else
      LoadClientProperty

      lblLoading.Caption = "Please wait, tree is building..."
      fmeLoading.Refresh
      DrawLandLordTree tvwLandLord, imgList, txtClientID.text, True

      lblLoading.Caption = "Please wait, global data is loading..."

      fmeLoading.Refresh

      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"
      adoMain.RecordSource = sSQLQuery_
      adoMain.CommandType = adCmdText
      adoMain.Refresh
      If adoMain.Recordset.RecordCount > 0 Then
         bGlobalData = True
      Else
         bGlobalData = False
      End If
      RetrieveMemo "Client", "ClientMemo", txtClientID.text, "ClientID", txtNote
   End If

'   ClientGlobalSetting     'Load client`s global settings

   fmeLoading.Visible = False
   MousePointer = vbDefault

   picClientList.Visible = False
End Sub
'
'Private Sub ClientGlobalSetting()
'   Dim szSQL As String
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'
'   szSQL = "SELECT FeeType, Handling, Frequency, NtDueDate, StDate, " & _
'               "ChargeType, ThisID, ParentID, '0' AS RECSTATUS " & _
'           "FROM ClientGDFees, ClientGlobalData " & _
'           "WHERE ClientGlobalData.Record_ID = ClientGDFees.ParentID AND " & _
'               "ClientGlobalData.ClientID = '" & flxClientList.TextMatrix(flxClientList.Row, 1) & "';"
'   populateGrid adoConn, szSQL, flxFeeType
'
'   adoConn.Close
'End Sub

Private Sub LoadGlobalData()
   Dim szTemp() As String, i As Integer

   For i = 0 To 11
      szTemp = Split(adoMain.Recordset.Fields("MonthlyDueDate" & (i + 1) & "").Value)
      cboDay(i).text = szTemp(0)
      cboMonth(i).text = szTemp(1)
      If i < 4 Then
         szTemp = Split(adoMain.Recordset.Fields("QuarterlyDueDate" & (i + 1) & "").Value)
         cboQDay(i).text = szTemp(0)
         cboQMth(i).text = szTemp(1)
      End If
      If i < 2 Then
         szTemp = Split(adoMain.Recordset.Fields("HalfYearlyDueDate" & (i + 1) & "").Value)
         cboHDay(i).text = szTemp(0)
         cboHMth(i).text = szTemp(1)
      End If
   Next i
   szTemp = Split(adoMain.Recordset.Fields("YearlyDueDate").Value)
   cboYDay.text = szTemp(0)
   cboYMth.text = szTemp(1)

   txtFeeIsuDays.text = adoMain.Recordset.Fields("FeeIsuDays").Value
   txtPayIsuDays.text = adoMain.Recordset.Fields("PayIsuDays").Value
End Sub

Private Sub LoadClientProperty()
   Dim conClient As New RDO.rdoConnection
   Dim rstProperty As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   cboProperty.Clear

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT PropertyID, PropertyName  " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyName;"

   Set rstProperty = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstProperty.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstProperty.EOF
      cboProperty.AddItem rstProperty!PROPERTYID & " / " & rstProperty!PropertyName
      rstProperty.MoveNext
   Wend
   
NoRes:
   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
End Sub
'
'Private Sub flxFeeType_RowColChange()
'   If cmdFeeTypesNew.Enabled = False Or cmdFeeTypesEdit.Enabled = False Then
'      flxFeeType.Row = iSlectedRow
'      Exit Sub
'   End If
''   displayRowInControl Me, imgFeeTypes, flxFeeType
'   iSlectedRow = flxFeeType.Row
'End Sub

Private Sub flxOtherBankDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxOtherBankDetails.ToolTipText = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.MouseRow, flxOtherBankDetails.MouseCol)
End Sub

Private Sub flxOtherBankDetails_RowColChange()
   Dim iCol As Integer
   
   MousePointer = vbHourglass

   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 10)
   cboPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 11)
   txtBacsRef.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 12)
   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   
   iSlectedRow = flxOtherBankDetails.Row
   
   MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   If LOAD_CLINT_CLIENTID <> "" Then
      LoadFormByClient
   End If
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50

   MousePointer = vbHourglass
   tabMain.Tab = 0
   cboPaymentMethod.AddItem "CHEQUE"
   cboPaymentMethod.AddItem "BACS"
   cboPaymentMethod.AddItem "DIRECT DEBIT"
   cboPaymentMethod.AddItem "Bank TRANSFER"
   cboPaymentMethod.AddItem "TT"
   cboPaymentMethod.AddItem "CHAPS"

   FillDaysMonths

   AGREEMENT_ADDNEW_MODE = False
   AGREEMENT_EDIT_MODE = False

   MousePointer = vbDefault
End Sub

Public Sub FillDaysMonths()

    Dim i As Integer, j As Integer
    
    For i = 0 To 11
      For j = 1 To 31
         cboDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboMonth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
   
   For i = 0 To 3
      For j = 1 To 31
         cboQDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboQMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
    
   For i = 0 To 1
      For j = 1 To 31
         cboHDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboHMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
    
   For j = 1 To 31
      cboYDay.AddItem Format(j, "00")
   Next j
   For j = 1 To 12
      cboYMth.AddItem Format("1/" & j & "/2000", "MMMM")
   Next j
   
   cboHandling.AddItem "Automatic"
   cboHandling.AddItem "Manual"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<ClientID|<ClientName|<ClientPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 300        'Solid column
   conFlxGrid.ColWidth(1) = 900        'Client ID
   conFlxGrid.ColWidth(2) = 3000       'Client Name
   conFlxGrid.ColWidth(3) = 800        'Post Code
   conFlxGrid.Rows = 2
'
   conFlxGrid.RowHeightMin = 300
End Sub

Private Sub imgClose_Click()
   picClientList.Visible = False
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
   txtTVInfoAdd(0).text = Rst!ProAddressLine1
   txtTVInfoAdd(1).text = Rst!ProAddressLine2
   txtTVInfoAdd(2).text = Rst!ProAddressLine3
   txtTVInfoPC.text = Rst!PROPOSTCODE
   
   Rst.Close
   Set Rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Function UnitDetails(szID As String) As Boolean
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
      If Rst!UnitName <> "" Then txtTVInfoName.text = Rst!UnitName
      txtTVInfoAdd(0).text = IIf(Rst!UnitAddressLine1 <> "", Rst!UnitAddressLine1, "")
      txtTVInfoAdd(1).text = IIf(Rst!UnitAddressLine2 <> "", Rst!UnitAddressLine2, "")
      txtTVInfoAdd(2).text = IIf(Rst!UnitAddressLine3 <> "", Rst!UnitAddressLine3, "")
      txtTVInfoPC.text = IIf(Rst!UnitPostCode <> "", Rst!UnitPostCode, "")
      If Rst!OCCUPIED = "N" Then
         UnitDetails = False
         Exit Function
      Else
         lblTenantIDLink.Caption = Rst!SageAccountNumber
         lblTenantNameLink.Caption = IIf(IsNull(Rst!TenantCompanyName), "", Rst!TenantCompanyName)
         Rst.Close
         Conn.Close
         Set Rst = Nothing
         Set Conn = Nothing
'
         szStr = LeaseDetails(szID)
         If szStr = "NULL" Then
            MsgBox "Please update lease information of this unit.", vbInformation + vbOKOnly, "Error"
            UnitDetails = False
         Else
            UnitDetails = True
            szaTemp = Split(szStr, " # ")
'
            txtPreOccupiedFr.text = szaTemp(0)
            txtPreOccupiedTo.text = szaTemp(1)
            txtPreTenancyType.text = szaTemp(2)
            txtPreRentRvw.text = szaTemp(3)
         End If
      End If
   End If
End Function

Private Sub LoadAllClientFlxGrd()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!ClientID
      flxClientList.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(iRow, 3) = IIf(IsNull(rstClient!ClientPostCode), "", rstClient!ClientPostCode)
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub
'
'Private Sub lstResidency_DblClick()
'   txtResidency.text = lstResidency.text
'   lstResidency.Visible = False
'End Sub
'
'Private Sub lstResidency_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then lstResidency_DblClick
'End Sub

Private Sub optLetting_Fees_Click(Index As Integer)
'   txtLETTING_FEES_VALUE(Index).SetFocus
End Sub

Private Sub optManagement_Fees_Click(Index As Integer)
'   txtMGT_FEES_VALUE(Index).SetFocus
End Sub

Private Sub optRecharges_Click(Index As Integer)
'   iRecharge = Index
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
   MousePointer = vbHourglass

   Select Case tabMain.Tab
   Case 1:                    'Property
      tvwLandLord.SetFocus
   Case 2:                    'Agreement
      If cboProperty.ListCount < 1 Then
         MsgBox "No property has been entered for this client.", vbCritical + vbOKOnly, "No Property"
      Else
         cboProperty.SetFocus
         AgreementButtonMode DefaultMode
         PopulateCodes
      End If
   Case 3:                    'Bank Payment details
      If cboBank_ID.text = "" Or flxOtherBankDetails.TextMatrix(1, 1) = "" Then
         LoadAllBankAC
         flxOtherBankDetails.Row = 0
         flxOtherBankDetails.Col = 0
      End If
   Case 5:                    'Global setting
'      tabFee.Tab = 0
      If bGlobalData Then
         cmdGSEdit.Caption = "Edit"
      Else
         cmdGSEdit.Caption = "Add New"
      End If

'      ConfigureFlxFeeType
'      ComponentInFrameEnableMode frmClientNew3, imgFeeTypes, DefaultMode
      PopulateDataCombo

   Case 6:                    'Memo and File attachment
      If txtClientID.text <> "" Then _
            Call LoadAttachmentFiles(cmbFiles, txtClientID.text, "Client")
   End Select
   MousePointer = vbDefault
End Sub

Private Sub PopulateDataCombo()
   Dim sSQLQuery_ As String
   
   adoFreq.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT CODE, VALUE " & _
              "FROM SECONDARYCODE " & _
              "WHERE SECONDARYCODE.PRIMARYCODE = 'FREQ' " & _
              "ORDER BY VALUE;"

   adoFreq.RecordSource = sSQLQuery_
   adoFreq.CommandType = adCmdText
   adoFreq.Refresh
   
   adoFeeTypes.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   
   sSQLQuery_ = "SELECT CODE, VALUE " & _
              "FROM SECONDARYCODE " & _
              "WHERE SECONDARYCODE.PRIMARYCODE = 'CFT' " & _
              "ORDER BY VALUE;"

   adoFeeTypes.RecordSource = sSQLQuery_
   adoFeeTypes.CommandType = adCmdText
   adoFeeTypes.Refresh

   adoChargeType.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT ID, FeeType " & _
              "FROM ChargeTypes;"

   adoChargeType.RecordSource = sSQLQuery_
   adoChargeType.CommandType = adCmdText
   adoChargeType.Refresh
End Sub

Public Sub SetFlxAgreementHeader(flxControl As MSHFlexGrid, ByVal rstAgreement As rdoResultset)
   Dim sSQLQuery_ As String, szHeader As String, iCol As Integer

   flxControl.Clear
   flxControl.Rows = 2
   flxControl.Cols = 12

   szHeader$ = "<CHARGE_TYPE|<DEMAND_TYPE|<Fund|<Handling" & _
               "|<CHARGE_METHOD|<CHARGE_BASIS|<AnnualCharge" & _
               "|<START_DATE|<END_DATE|<Frequency|<NtDueDate|<AGREEMENT_ID"
   flxControl.FormatString = szHeader$

   For iCol = 0 To flxControl.Cols - 3
      flxControl.ColWidth(iCol) = Label7(iCol + 1).Left - Label7(iCol).Left
   Next iCol
   flxControl.ColWidth(10) = txtNtDueDate.Width - 55
   flxControl.ColWidth(11) = 0

   Dim iRecCol As Integer, iRow As Integer

   SetControlStyle flxControl
   rstAgreement.MoveFirst

   For iRow = 1 To RDORecordCount(rstAgreement)
      For iRecCol = 0 To flxControl.Cols - 1
         For iCol = 0 To flxControl.Cols - 1
            If flxControl.TextMatrix(0, iCol) = rstAgreement.rdoColumns(iRecCol).Name Then Exit For
         Next iCol
         flxControl.TextMatrix(iRow, iCol) = IIf(IsNull(rstAgreement.rdoColumns(iRecCol).Value), "", rstAgreement.rdoColumns(iRecCol).Value)
      Next iRecCol
      rstAgreement.MoveNext
      If Not rstAgreement.EOF Then flxControl.AddItem ""
   Next iRow
   
   flxControl.Row = 0
   flxControl.Col = 0
End Sub

Private Sub LoadTypes()
   Dim sSQLQuery_ As String

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT * FROM ChargeTypes ORDER BY ID"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh
End Sub

Private Sub LoadAllBankAC()
   ConfigureFlxOtherBank

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT tlbClientBanks.*, tlbBank.* " & _
           "FROM tlbClientBanks, tlbBank, Client " & _
           "WHERE Client.ClientID = '" & txtClientID.text & "' And " & _
               "Client.BANK_ID = tlbBank.BANK_ID And " & _
               "tlbBank.BANK_ID = tlbClientBanks.BANK_ID"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

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
           "WHERE CLIENT_ID = '" & txtClientID.text & "' And " & _
               "tlbBank.BANK_ID = tlbClientBanks.BANK_ID " & _
           "ORDER BY Bank_AC_Name;"

   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

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
   flxOtherBankDetails.Row = 0
   flxOtherBankDetails.Col = 0
NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub ConfigureFlxOtherBank()
   Dim szHeader As String, i As Integer

   flxOtherBankDetails.Clear
   flxOtherBankDetails.Cols = 13

   szHeader = "|<BANK_ID|<BANK_NAME|<BANK_POST_CODE|<BANK_AC_NAME|<BANK_AC_NUM|<BANK_SC|<DEFAULT_AC"
   flxOtherBankDetails.FormatString = szHeader

   flxOtherBankDetails.ColWidth(0) = 400

   For i = 1 To flxOtherBankDetails.Cols - 1
      flxOtherBankDetails.ColWidth(i) = (flxOtherBankDetails.Width - 600) / 7
   Next i
   flxOtherBankDetails.ColWidth(8) = 0
   flxOtherBankDetails.ColWidth(9) = 0
   flxOtherBankDetails.ColWidth(10) = 0
   flxOtherBankDetails.ColWidth(11) = 0      'PaymentMethod
   flxOtherBankDetails.ColWidth(12) = 0      'BacsRef

   flxOtherBankDetails.RowHeightMin = 315
End Sub

Private Sub tvwLandLord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If tvwLandLord.Nodes.Count = 0 Then Exit Sub

   If Button = 1 Then
      szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
      fraType.Caption = szaPremisisIDType(1)

      IMAGE_FILE_NAME_ = ImageLoader(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), lblImageName)

      txtTVInfoName.text = tvwLandLord.SelectedItem.text
      txtTVInfoAdd(0).text = ""
      txtTVInfoAdd(1).text = ""
      txtTVInfoAdd(2).text = ""
      txtTVInfoPC.text = ""

      If szaPremisisIDType(1) = "CLIENT" Then
         fraOccupied.Visible = False
      End If
      If szaPremisisIDType(1) = "PROPERTY" Then
         fraOccupied.Visible = False
         PropertyDetails szaPremisisIDType(0)
      End If
      If szaPremisisIDType(1) = "UNITS" Then
         If UnitDetails(szaPremisisIDType(0)) Then
            fraOccupied.Visible = True
         Else
            fraOccupied.Visible = False
         End If
      End If
   End If
End Sub

Private Sub LockingAcText(bLock As Boolean)
   cboPaymentMethod.Locked = bLock
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
'   txtClientID.Locked = bLock
   txtClientName.Locked = bLock
   lstResidency.Enabled = Not bLock
'   txtAcBalance.Locked = bLock
   txtVATReg.Locked = bLock
   txtYearEndDate.Locked = bLock
   txtClientAddressLine1.Locked = bLock
   txtClientAddressLine2.Locked = bLock
   txtClientAddressLine3.Locked = bLock
   txtClientPostCode.Locked = bLock
   txtClientHomeTel.Locked = bLock
   txtClientOfficeTel.Locked = bLock
   txtClientMobile.Locked = bLock
   txtClientPersonalEmail.Locked = bLock
   txtClientOfficeEmail.Locked = bLock
   txtClientOfficeAddressLine1.Locked = bLock
   txtClientOfficeAddressLine2.Locked = bLock
   txtClientOfficeAddressLine3.Locked = bLock
   txtClientOfficePostCode.Locked = bLock
End Sub

Private Sub SageSupplierAccCombo()
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
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then
   
      Set oPurchaseRecord = oWS.CreateObject("PurchaseRecord")
   
      Dim TotalRow, TotalCol As Long
      Dim data() As String
      Dim i As Integer
   
      TotalRow = oPurchaseRecord.Count
      TotalCol = 2
      cboLandLordSageSuppAC.Clear
   
      ReDim data(TotalCol, TotalRow) As String
   
      oPurchaseRecord.MoveFirst
      For i = 0 To TotalRow - 1
         data(0, i) = CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value)
         data(1, i) = CStr(oPurchaseRecord.Fields.Item("NAME").Value)
         oPurchaseRecord.MoveNext
      Next i
   
      cboLandLordSageSuppAC.Column() = data()
      cboLandLordSageSuppAC.ColumnCount = TotalCol
      cboLandLordSageSuppAC.BoundColumn = 1
   
      'Disconnect
      oWS.Disconnect
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

Private Sub txtBANK_AC_NUM_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtBANK_ADDRESS1_DblClick()
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_ADDRESS2_DblClick()
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_ADDRESS3_DblClick()
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_NAME_DblClick()
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_POST_CODE_DblClick()
   MsgBox "To edit the bank details, please go to Bank screen through Global secreen!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_SC_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtEND_DATE_Change()
   TextBoxChangeDate txtEND_DATE
End Sub

Private Sub txtEND_DATE_GotFocus()
   If txtSTART_DATE.text = "" Then
      MsgBox "Please enter start date before the end date.", vbCritical + vbOKOnly, "End Date"
      txtSTART_DATE.SetFocus
   End If
End Sub

Private Sub txtEND_DATE_KeyPress(KeyAscii As MSForms.ReturnInteger)
   TextBoxKeyPrsDate txtEND_DATE, KeyAscii
End Sub

Private Sub txtEND_DATE_LostFocus()
   TextBoxFormatDate txtEND_DATE
End Sub

Private Sub txtNOTICE_DAYS_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

'Private Sub txtNtDueDate_Change()
'   TextBoxChangeDate txtNtDueDate
'End Sub
'
'Private Sub txtNtDueDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   TextBoxKeyPrsDate txtNtDueDate, KeyAscii
'End Sub
'
'Private Sub txtNtDueDate_LostFocus()
'   If txtNtDueDate.text <> "" Then TextBoxFormatDate txtNtDueDate
'End Sub
'
'Private Sub txtResidency_GotFocus()
'   If cmdResidency.Enabled Then cmdResidency.SetFocus
'End Sub

Private Sub txtREVIEW_DATE_Change()
   TextBoxChangeDate txtREVIEW_DATE
End Sub

Private Sub txtREVIEW_DATE_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtREVIEW_DATE, KeyAscii
End Sub

Private Sub txtREVIEW_DATE_LostFocus()
   TextBoxFormatDate txtREVIEW_DATE
End Sub

Private Sub txtRRPA_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSTART_DATE_Change()
   TextBoxChangeDate txtSTART_DATE
End Sub

Private Sub txtSTART_DATE_KeyPress(KeyAscii As MSForms.ReturnInteger)
   TextBoxKeyPrsDate txtSTART_DATE, KeyAscii
End Sub

Private Sub txtSTART_DATE_LostFocus()
   TextBoxFormatDate txtSTART_DATE
End Sub
'
'Private Sub txtStartDate_Change()
'   TextBoxChangeDate txtStartDate
'End Sub
'
'Private Sub txtStartDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   TextBoxKeyPrsDate txtStartDate, KeyAscii
'End Sub
'
'Private Sub txtStartDate_LostFocus()
'   If txtStartDate.text <> "" Then TextBoxFormatDate txtStartDate
'End Sub

Private Sub txtYearEndDate_Change()
   TextBoxChangeDate txtYearEndDate
End Sub

Private Sub txtYearEndDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtYearEndDate, KeyAscii
End Sub

Private Sub txtYearEndDate_LostFocus()
   TextBoxFormatDate txtYearEndDate
End Sub

Private Sub AgreementButtonMode(ByVal mode As ComponentMode)
   Select Case mode

   Case ComponentMode.DefaultMode
      cboCHARGE_TYPE.Locked = True
      cboDEMAND_TYPE.Locked = True
      cboCHARGE_BASIS.Locked = True
      cboCHARGE_METHOD.Locked = True
      txtAnnualCharge.Locked = True
      txtSTART_DATE.Locked = True
      txtEND_DATE.Locked = True
      cboFund.Locked = True
      cboHandling.Locked = True
      cboFrequency.Locked = True
'      txtNtDueDate.Locked = True
      
      flxAgreement.Enabled = True
      
      cmdAgmntAddNew.Enabled = True
      cmdAgmntEdit.Enabled = False
      cmdAgmntSave.Enabled = False
      cmdAgmntCancel.Enabled = False
   
   Case ComponentMode.NewEntryMode
      cboCHARGE_TYPE.Locked = False
      cboDEMAND_TYPE.Locked = False
      cboCHARGE_BASIS.Locked = False
      cboCHARGE_METHOD.Locked = False
      txtAnnualCharge.Locked = False
      txtSTART_DATE.Locked = False
      txtEND_DATE.Locked = False
      cboFund.Locked = False
      cboHandling.Locked = False
      cboFrequency.Locked = False
'      txtNtDueDate.Locked = False
      
      flxAgreement.Enabled = False
      
      cmdAgmntAddNew.Enabled = False
      cmdAgmntEdit.Enabled = False
      cmdAgmntSave.Enabled = True
      cmdAgmntCancel.Enabled = True
   
   Case ComponentMode.EditMode
      cboCHARGE_TYPE.Locked = False
      cboDEMAND_TYPE.Locked = False
      cboCHARGE_BASIS.Locked = False
      cboCHARGE_METHOD.Locked = False
      txtAnnualCharge.Locked = False
      txtSTART_DATE.Locked = False
      txtEND_DATE.Locked = False
      cboFund.Locked = False
      cboHandling.Locked = False
      cboFrequency.Locked = False
'      txtNtDueDate.Locked = False

      flxAgreement.Enabled = False

      cmdAgmntAddNew.Enabled = False
      cmdAgmntEdit.Enabled = False
      cmdAgmntSave.Enabled = True
      cmdAgmntCancel.Enabled = True
   End Select
End Sub

Public Sub SetNextDueDt()
   Select Case cboFrequency.Value
      Case cboFrequency.List(0):                               'Weekly in advance
         txtNtDueDate.text = txtSTART_DATE.text
      Case cboFrequency.List(1):                               'Weekly in arrears
         txtNtDueDate.text = DateAdd("d", 7, txtSTART_DATE.text)
      Case cboFrequency.List(2):                               'Fortnightly in advance
         txtNtDueDate.text = txtSTART_DATE.text
      Case cboFrequency.List(3):                               'Fortnightly in arrears
         txtNtDueDate.text = DateAdd("d", 14, txtSTART_DATE.text)
      Case cboFrequency.List(4):                               'Monthly in advance
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InAdv, Pay_Monthly)
      Case cboFrequency.List(5):                               'Monthly in arrears
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InArr, Pay_Monthly)
      Case cboFrequency.List(6):                               'Quarterly in advance
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InAdv, Pay_Quarterly)
      Case cboFrequency.List(7):                               'Quarterly in arrears
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InArr, Pay_Quarterly)
      Case cboFrequency.List(8):                               'Half yearly in advance
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InAdv, Pay_Half_Yearly)
      Case cboFrequency.List(9):                               'Half yearly in arrears
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InArr, Pay_Half_Yearly)
      Case cboFrequency.List(10):                              'yearly in advance
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InAdv, Pay_Yearly)
      Case cboFrequency.List(11):                              'yearly in arrears
         txtNtDueDate.text = NextPayingDate(txtSTART_DATE.text, InArr, Pay_Yearly)
   End Select
End Sub

Private Sub AgreementClearMode(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
      cboCHARGE_TYPE.text = ""
      cboDEMAND_TYPE.text = ""
      cboCHARGE_BASIS.text = ""
      cboCHARGE_METHOD.text = ""
      txtAnnualCharge.text = ""
      txtSTART_DATE.text = ""
      txtEND_DATE.text = ""
      cboFund.text = ""
      cboHandling.text = ""
      cboFrequency.text = ""
      txtNtDueDate.text = ""
   
   Case CearEntryComponents.ClearOnlyComboBoxes
      cboCHARGE_BASIS.Clear
      cboCHARGE_METHOD.Clear
      cboFund.Clear
      cboHandling.Clear
      cboFrequency.Clear
   
   Case CearEntryComponents.ClearBoth
      AgreementClearMode ClearOnlyTextBoxes
      AgreementClearMode ClearOnlyComboBoxes
   End Select
End Sub

Public Sub PopulateCodes()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   ' Charge Type
   sSQLQuery = "SELECT ID, FeeType " & _
                 "FROM ChargeTypes;"

   populateCombo adoConn, sSQLQuery, cboCHARGE_TYPE

   ' Demand Type
   sSQLQuery = "SELECT ID, TYPE " & _
                 "FROM DemandTypes;"

   populateCombo adoConn, sSQLQuery, cboDEMAND_TYPE

   ' Fund
   LoadFund
   
   ' Charge Basis
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'AMTP'"

   populateCombo adoConn, sSQLQuery, cboCHARGE_BASIS

   ' Amt/%
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'CRGBS'"

   populateCombo adoConn, sSQLQuery, cboCHARGE_METHOD

   ' Frequency
   sSQLQuery = "SELECT ID, Frequency " & _
                 "FROM Frequencies;"

   populateCombo adoConn, sSQLQuery, cboFrequency

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadFund()
   Dim data() As String
   Dim rRow As Integer

   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   '  Set oSDO = New SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   '  Set oWS = oSDO.Workspaces.Add("WkpsSupplier")
   Dim oDepartmentData As SageDataObject120.DepartmentData

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'   szDataPath = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & Sdsn & "", "DataPathname")
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
   ' Try to Connect - Will Throw an Exception if it Fails
    If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

       Set oDepartmentData = oWS.CreateObject("DepartmentData")
       
       ReDim data(2, oDepartmentData.Count) As String
       
       For rRow = 2 To oDepartmentData.Count
          oDepartmentData.Read (rRow)
          data(0, rRow - 2) = CStr(rRow - 1)
          data(1, rRow - 2) = CStr(oDepartmentData.Fields.Item("NAME").Value)
       Next rRow
       'Disconnect
       oWS.Disconnect
    End If
    cboFund.Clear
    cboFund.Column() = data()

   ' Destroy Objects
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

      MsgBox "(frmLease4 LoadDept) The SDO generated the following error: " & oSDO.LastError.text
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub
