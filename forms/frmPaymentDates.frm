VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPaymentDates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Dates"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaymentDates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10980
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   9420
      TabIndex        =   56
      Top             =   3840
      Width           =   1455
      Begin VB.CommandButton cmdAutoSetup 
         Caption         =   "C&lose"
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   3840
      Width           =   1935
      Begin VB.CommandButton cmdAutoSetup 
         Caption         =   "A&uto Date Fill"
         Enabled         =   0   'False
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraButton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2220
      TabIndex        =   26
      Top             =   3840
      Width           =   7035
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit "
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   1965
         TabIndex        =   30
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel "
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   5400
         TabIndex        =   28
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New "
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save "
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   3675
         TabIndex        =   27
         Top             =   240
         Width           =   1395
      End
   End
   Begin TabDlg.SSTab tabPaymentDates 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Monthly Payment Dates"
      TabPicture(0)   =   "frmPaymentDates.frx":17002
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMthPayDt(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMthPayDt(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMthPayDt(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMthPayDt(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblMthPayDt(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblMthPayDt(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMthPayDt(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblMthPayDt(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblMthPayDt(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblMthPayDt(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblMthPayDt(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblMthPayDt(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblMthPayDt(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblMthPayDt(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "flxMthPayDt(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtMthPayDt(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMthPayDt(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMthPayDt(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtMthPayDt(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMthPayDt(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMthPayDt(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtMthPayDt(6)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtMthPayDt(7)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtMthPayDt(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtMthPayDt(9)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtMthPayDt(10)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtMthPayDt(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtMthPayDt(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Quarterly Payment Dates"
      TabPicture(1)   =   "frmPaymentDates.frx":1701E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMthPayDt(16)"
      Tab(1).Control(1)=   "txtMthPayDt(13)"
      Tab(1).Control(2)=   "txtMthPayDt(14)"
      Tab(1).Control(3)=   "txtMthPayDt(15)"
      Tab(1).Control(4)=   "flxMthPayDt(1)"
      Tab(1).Control(5)=   "Label1(0)"
      Tab(1).Control(6)=   "txtDescription(0)"
      Tab(1).Control(7)=   "lblMthPayDt(19)"
      Tab(1).Control(8)=   "lblMthPayDt(17)"
      Tab(1).Control(9)=   "lblMthPayDt(16)"
      Tab(1).Control(10)=   "lblMthPayDt(15)"
      Tab(1).Control(11)=   "lblMthPayDt(13)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Half Yearly payments"
      TabPicture(2)   =   "frmPaymentDates.frx":1703A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtMthPayDt(17)"
      Tab(2).Control(1)=   "txtMthPayDt(18)"
      Tab(2).Control(2)=   "flxMthPayDt(2)"
      Tab(2).Control(3)=   "Label1(1)"
      Tab(2).Control(4)=   "txtDescription(1)"
      Tab(2).Control(5)=   "lblMthPayDt(22)"
      Tab(2).Control(6)=   "lblMthPayDt(20)"
      Tab(2).Control(7)=   "lblMthPayDt(18)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Yearly payments"
      TabPicture(3)   =   "frmPaymentDates.frx":17056
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtMthPayDt(19)"
      Tab(3).Control(1)=   "flxMthPayDt(3)"
      Tab(3).Control(2)=   "Label1(2)"
      Tab(3).Control(3)=   "txtDescription(2)"
      Tab(3).Control(4)=   "lblMthPayDt(24)"
      Tab(3).Control(5)=   "lblMthPayDt(21)"
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   -69780
         MaxLength       =   5
         TabIndex        =   53
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   -70800
         MaxLength       =   5
         TabIndex        =   51
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   -68700
         MaxLength       =   5
         TabIndex        =   49
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
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
         Index           =   12
         Left            =   9020
         MaxLength       =   25
         TabIndex        =   47
         Top             =   675
         Width           =   1675
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   18
         Left            =   -69180
         MaxLength       =   5
         TabIndex        =   42
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   -73425
         MaxLength       =   5
         TabIndex        =   35
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   -71850
         MaxLength       =   5
         TabIndex        =   36
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   -70275
         MaxLength       =   5
         TabIndex        =   37
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   8400
         MaxLength       =   5
         TabIndex        =   25
         Top             =   680
         Width           =   600
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   24
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   6960
         MaxLength       =   5
         TabIndex        =   23
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   22
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   21
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   20
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   19
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   18
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   17
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   16
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   15
         Top             =   680
         Width           =   680
      End
      Begin VB.TextBox txtMthPayDt 
         Alignment       =   1  'Right Justify
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
         Left            =   480
         MaxLength       =   5
         TabIndex        =   14
         Top             =   680
         Width           =   680
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMthPayDt 
         Height          =   2610
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4604
         _Version        =   393216
         Cols            =   13
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
         _Band(0).Cols   =   13
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMthPayDt 
         Height          =   2610
         Index           =   1
         Left            =   -74400
         TabIndex        =   32
         Top             =   960
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4604
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
            Name            =   "Myriad Web"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMthPayDt 
         Height          =   2610
         Index           =   2
         Left            =   -71760
         TabIndex        =   33
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4604
         _Version        =   393216
         Cols            =   3
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMthPayDt 
         Height          =   2610
         Index           =   3
         Left            =   -70680
         TabIndex        =   34
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4604
         _Version        =   393216
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
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   2
         Left            =   -68160
         TabIndex        =   63
         Top             =   480
         Width           =   1800
         Caption         =   "Description"
         Size            =   "3175;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDescription 
         Height          =   285
         Index           =   2
         Left            =   -68160
         TabIndex        =   62
         Top             =   675
         Width           =   1675
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2955;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   1
         Left            =   -67560
         TabIndex        =   61
         Top             =   480
         Width           =   1800
         Caption         =   "Description"
         Size            =   "3175;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDescription 
         Height          =   285
         Index           =   1
         Left            =   -67560
         TabIndex        =   60
         Top             =   675
         Width           =   1675
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2955;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   -67140
         TabIndex        =   59
         Top             =   480
         Width           =   1800
         Caption         =   "Description"
         Size            =   "3175;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDescription 
         Height          =   285
         Index           =   0
         Left            =   -67140
         TabIndex        =   58
         Top             =   675
         Width           =   1675
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2955;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   22
         Left            =   -70800
         TabIndex        =   52
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   19
         Left            =   -73425
         TabIndex        =   50
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   14
         Left            =   9015
         TabIndex        =   48
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   24
         Left            =   -69780
         TabIndex        =   46
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "No."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   21
         Left            =   -70680
         TabIndex        =   45
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   20
         Left            =   -69180
         TabIndex        =   44
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "No."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   18
         Left            =   -71760
         TabIndex        =   43
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "4th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   17
         Left            =   -68700
         TabIndex        =   41
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "3rd"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   16
         Left            =   -70275
         TabIndex        =   40
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   15
         Left            =   -71850
         TabIndex        =   39
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "No."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   13
         Left            =   -74400
         TabIndex        =   38
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "No."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "12th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   11
         Left            =   8400
         TabIndex        =   13
         Top             =   480
         Width           =   330
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "11th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   10
         Left            =   7680
         TabIndex        =   12
         Top             =   480
         Width           =   330
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "10th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   9
         Left            =   6960
         TabIndex        =   11
         Top             =   480
         Width           =   330
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "9th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   8
         Left            =   6240
         TabIndex        =   10
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "8th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   9
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "6th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "5th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "7th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "3rd"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblMthPayDt 
         AutoSize        =   -1  'True
         Caption         =   "4th"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPaymentDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SELECTED_ROW As Integer
Private bEDIT As Boolean
Private bFS As Boolean 'FORWARD SLASH - PRINTING

Private Sub cmdAddNew_Click()
   If MsgBox("Do you want to add new Set of payment Dates?", vbQuestion + vbYesNo, "Payment Dates") = vbNo Then Exit Sub

   ComponentInFrameEnableMode Me, fraButton, NewEntryMode
   ComponentEnableModePaymentDates Me, NewEntryMode
   tabPaymentDates.Tab = 0
   txtMthPayDt(0).SetFocus
   SELECTED_ROW = 0
   cmdAutoSetup(0).Enabled = True
End Sub

Private Sub cmdAutoSetup_Click(Index As Integer)
   If Index = 1 Then
      Unload Me
      Exit Sub
   End If

   Dim dtDate As Date, var

   On Error GoTo ErrorHandler

   var = InputBox("Please type the first payment date of the year. (dd/mm/yyyy)", "Frist Payment Date", "01/01/" & Year(Date))
   If var = "" Then Exit Sub

   dtDate = Format(var, "dd mmmm yyyy")

   SetAddDates dtDate

   Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
      cmdAutoSetup_Click (0)
   End If
End Sub

Private Sub SetAddDates(dtDate As Date)
   Dim i As Integer

'  Monthly Payment Dates
   txtMthPayDt(0).text = Format(dtDate, "dd/mm")
   For i = 1 To 11
      txtMthPayDt(i).text = Format(DateAdd("m", i, dtDate), "dd/mm")
   Next i

'  Quarterly Payment Dates
   txtMthPayDt(13).text = Format(dtDate, "dd/mm")
   txtMthPayDt(14).text = Format(DateAdd("m", 3, dtDate), "dd/mm")
   txtMthPayDt(15).text = Format(DateAdd("m", 6, dtDate), "dd/mm")
   txtMthPayDt(16).text = Format(DateAdd("m", 9, dtDate), "dd/mm")

'  Halfyearly Payment Dates
   txtMthPayDt(17).text = Format(dtDate, "dd/mm")
   txtMthPayDt(18).text = Format(DateAdd("m", 6, dtDate), "dd/mm")

'  Yearly Payment Date
   txtMthPayDt(19).text = Format(dtDate, "dd/mm")
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Do you want to cancel the editing?", vbQuestion + vbYesNo, "Payment Date") = vbNo Then Exit Sub
   ComponentInFrameEnableMode Me, fraButton, DefaultMode
   ComponentEnableModePaymentDates Me, DefaultMode
   cmdAutoSetup(0).Enabled = False
End Sub

Private Sub cmdEdit_Click()
   If SELECTED_ROW = 0 Then Exit Sub

   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Payment Dates") = vbNo Then Exit Sub
   
   ComponentInFrameEnableMode Me, fraButton, EditMode
   ComponentEnableModePaymentDates Me, EditMode
   cmdAutoSetup(0).Enabled = True
End Sub

Private Sub cmdSave_Click()
   Dim i As Integer

   bEDIT = False
   For i = 0 To 19
      If i <> 12 Then If txtMthPayDt(i).text = "" Then Exit For
   Next i
   If i <> 20 Then
      ShowMsgInTaskBar "All Monthly, Quarterly and Half-Yearly & Yearly dates should be entered"
      Exit Sub
   End If

   RearrangeDates

   If txtMthPayDt(12).text = "" Then
      ShowMsgInTaskBar "Please type a description."
      txtMthPayDt(12).SetFocus
      Exit Sub
   End If

   If SELECTED_ROW = 0 Then
      AddNewDates
   Else
      EditDates
   End If

   ComponentInFrameEnableMode Me, fraButton, DefaultMode
   ComponentEnableModePaymentDates Me, DefaultMode
   LoadFlxMthPayDt

   SELECTED_ROW = 0
   cmdAutoSetup(0).Enabled = False
End Sub

Private Sub RearrangeDates()
   Dim szTemp As String, i As Integer, j As Integer

   For i = 0 To 10
      For j = i + 1 To 11
         If DateDiff("d", CDate(txtMthPayDt(i).text & "/" & Year(Date)), CDate(txtMthPayDt(j).text & "/" & Year(Date))) < 0 Then
            szTemp = txtMthPayDt(i).text
            txtMthPayDt(i).text = txtMthPayDt(j).text
            txtMthPayDt(j).text = szTemp
         End If
      Next j
   Next i
   For i = 13 To 15
      For j = i + 1 To 16
         If DateDiff("d", CDate(txtMthPayDt(i).text & "/" & Year(Date)), CDate(txtMthPayDt(j).text & "/" & Year(Date))) < 0 Then
            szTemp = txtMthPayDt(i).text
            txtMthPayDt(i).text = txtMthPayDt(j).text
            txtMthPayDt(j).text = szTemp
         End If
      Next j
   Next i
   For i = 17 To 17
      For j = i + 1 To 18
         If DateDiff("d", CDate(txtMthPayDt(i).text & "/" & Year(Date)), CDate(txtMthPayDt(j).text & "/" & Year(Date))) < 0 Then
            szTemp = txtMthPayDt(i).text
            txtMthPayDt(i).text = txtMthPayDt(j).text
            txtMthPayDt(j).text = szTemp
         End If
      Next j
   Next i
End Sub

Private Sub EditDates()
   Dim sSQLQuery_ As String, szHeader As String, iRecCount As Integer, j As Integer
   Dim conPayDates As New ADODB.Connection
   Dim adoRst As ADODB.Recordset

   conPayDates.Open getConnectionString
   Set adoRst = New ADODB.Recordset

   sSQLQuery_ = "SELECT DateSetID, NameOfSet, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, " & _
                  "MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, " & _
                  "QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate " & _
                "FROM PaymentDates WHERE DateSetID = " & flxMthPayDt(0).TextMatrix(SELECTED_ROW, 0) & ";"
   adoRst.Open sSQLQuery_, conPayDates, adOpenDynamic, adLockOptimistic

   adoRst.Fields(1).Value = txtMthPayDt(12).text
   For j = 2 To 20
      If j < 14 Then adoRst.Fields(j).Value = Format(txtMthPayDt(j - 2).text, "dd mmmm")
      If j > 13 Then adoRst.Fields(j).Value = Format(txtMthPayDt(j - 1).text, "dd mmmm")
   Next j
   adoRst.Update
 
   Set adoRst = Nothing
   conPayDates.Close
   Set conPayDates = Nothing
End Sub

Private Sub AddNewDates()
   Dim sSQLQuery_ As String, szHeader As String, iRecCount As Integer, j As Integer
   Dim conPayDates As New ADODB.Connection
   Dim adoRst As ADODB.Recordset

   conPayDates.Open getConnectionString
   Set adoRst = New ADODB.Recordset

   sSQLQuery_ = "SELECT COUNT(DateSetID) AS DATEID FROM PaymentDates;"
   adoRst.Open sSQLQuery_, conPayDates, adOpenStatic, adLockReadOnly

   iRecCount = adoRst.Fields("DATEID").Value

   adoRst.Close

   sSQLQuery_ = "SELECT DateSetID, NameOfSet, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, " & _
                  "MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, " & _
                  "QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate " & _
                "FROM PaymentDates;"
   adoRst.Open sSQLQuery_, conPayDates, adOpenDynamic, adLockOptimistic

   adoRst.AddNew

   adoRst.Fields(0).Value = iRecCount + 1
   adoRst.Fields(1).Value = txtMthPayDt(12).text
   For j = 2 To 20
      If j < 14 Then adoRst.Fields(j).Value = Format(txtMthPayDt(j - 2).text, "dd mmmm")
      If j > 13 Then adoRst.Fields(j).Value = Format(txtMthPayDt(j - 1).text, "dd mmmm")
   Next j

   adoRst.Update

   Set adoRst = Nothing
   conPayDates.Close
   Set conPayDates = Nothing
End Sub

Private Sub flxMthPayDt_Click(Index As Integer)
   Dim i As Integer

   SELECTED_ROW = flxMthPayDt(Index).row

   For i = 0 To 3
      If i <> Index Then flxMthPayDt(i).row = SELECTED_ROW
   Next i

   If flxMthPayDt(0).TextMatrix(flxMthPayDt(0).row, 0) = "" Then Exit Sub

   For i = 0 To 19
      If i < 13 Then _
         txtMthPayDt(i).text = flxMthPayDt(0).TextMatrix(SELECTED_ROW, i + 1)
      If i > 12 And i < 17 Then _
         txtMthPayDt(i).text = flxMthPayDt(1).TextMatrix(SELECTED_ROW, i - 12)
      If i > 16 And i < 19 Then _
         txtMthPayDt(i).text = flxMthPayDt(2).TextMatrix(SELECTED_ROW, i - 16)
      If i = 19 Then _
         txtMthPayDt(i).text = flxMthPayDt(3).TextMatrix(SELECTED_ROW, i - 18)
   Next i
End Sub

Private Sub flxMthPayDt_RowColChange1(Index As Integer)
   Dim i As Integer

   SELECTED_ROW = flxMthPayDt(0).row

   If flxMthPayDt(0).TextMatrix(flxMthPayDt(0).row, 0) = "" Then Exit Sub

   For i = 0 To 19
      If i < 13 Then _
         txtMthPayDt(i).text = flxMthPayDt(0).TextMatrix(SELECTED_ROW, i + 1)
      If i > 12 And i < 17 Then _
         txtMthPayDt(i).text = flxMthPayDt(1).TextMatrix(SELECTED_ROW, i - 12)
      If i > 16 And i < 19 Then _
         txtMthPayDt(i).text = flxMthPayDt(2).TextMatrix(SELECTED_ROW, i - 16)
      If i = 19 Then _
         txtMthPayDt(i).text = flxMthPayDt(3).TextMatrix(SELECTED_ROW, i - 18)
   Next i
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   tabPaymentDates.BackColor = Me.BackColor
   Frame1(0).BackColor = Me.BackColor
   fraButton.BackColor = Me.BackColor

   tabPaymentDates.Tab = 0
   ConfigureFlxGrids

   LoadFlxMthPayDt

   ComponentInFrameEnableMode Me, fraButton, DefaultMode
   ComponentEnableModePaymentDates Me, DefaultMode

   SELECTED_ROW = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If frmMMain.IsRibbonVersion Then
      frmGlobalx.Enabled = True
   Else
'      frmGlobal1.Enabled = True
   End If
End Sub

Private Sub LoadFlxMthPayDt()
   Dim sSQLQuery_ As String, szHeader As String
   Dim conPayDates As New ADODB.Connection
   Dim adoRst As ADODB.Recordset
   Dim i As Integer, j As Integer

   conPayDates.Open getConnectionString
   sSQLQuery_ = "SELECT DateSetID, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, MonthlyDueDate4" & _
                  ", MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8" & _
                  ", MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, NameOfSet " & _
                  ", QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4 " & _
                  ", HalfYearlyDueDate1, HalfYearlyDueDate2 " & _
                  ", YearlyDueDate " & _
                "FROM PaymentDates " & _
                "ORDER BY DateSetID ASC;"

   Set adoRst = New ADODB.Recordset
   adoRst.Open sSQLQuery_, conPayDates, adOpenStatic, adLockOptimistic

   flxMthPayDt(0).Rows = 2

   For i = 0 To adoRst.RecordCount - 1
      flxMthPayDt(0).TextMatrix(i + 1, 0) = adoRst.Fields(0)
      flxMthPayDt(1).TextMatrix(i + 1, 0) = adoRst.Fields(0)
      flxMthPayDt(2).TextMatrix(i + 1, 0) = adoRst.Fields(0)
      flxMthPayDt(3).TextMatrix(i + 1, 0) = adoRst.Fields(0)

      For j = 1 To 12
          If IsNull(adoRst.Fields(j)) Then
              flxMthPayDt(0).TextMatrix(i + 1, j) = ""
          Else
              flxMthPayDt(0).TextMatrix(i + 1, j) = Format(adoRst.Fields(j), "dd/mm")
          End If
      Next j
      If IsNull(adoRst.Fields(j)) Then
          flxMthPayDt(0).TextMatrix(i + 1, j) = ""
      Else
          flxMthPayDt(0).TextMatrix(i + 1, j) = adoRst.Fields(j)
      End If

      For j = 14 To 17
          If IsNull(adoRst.Fields(j)) Then
              flxMthPayDt(1).TextMatrix(i + 1, j - 13) = ""
          Else
              flxMthPayDt(1).TextMatrix(i + 1, j - 13) = Format(adoRst.Fields(j), "dd/mm")
          End If
      Next j
      
      For j = 18 To 19
          If IsNull(adoRst.Fields(j)) Then
              flxMthPayDt(2).TextMatrix(i + 1, j - 17) = ""
          Else
              flxMthPayDt(2).TextMatrix(i + 1, j - 17) = Format(adoRst.Fields(j), "dd/mm")
          End If
      Next j
      
      If IsNull(adoRst.Fields(20)) Then
          flxMthPayDt(3).TextMatrix(i + 1, 1) = ""
      Else
          flxMthPayDt(3).TextMatrix(i + 1, 1) = Format(adoRst.Fields(20), "dd/mm")
      End If
      adoRst.MoveNext

      If Not adoRst.EOF Then flxMthPayDt(0).AddItem ""
      If Not adoRst.EOF Then flxMthPayDt(1).AddItem ""
      If Not adoRst.EOF Then flxMthPayDt(2).AddItem ""
      If Not adoRst.EOF Then flxMthPayDt(3).AddItem ""
   Next i

   adoRst.Close
   conPayDates.Close
   Set conPayDates = Nothing
End Sub

Private Sub ConfigureFlxGrids()
   Dim i As Integer, iFlxCol As Integer, szHeader As String

   flxMthPayDt(0).Cols = 14

   szHeader$ = "<DateSetID|<MonthlyDueDate1|<MonthlyDueDate2|<MonthlyDueDate3|<MonthlyDueDate4" & _
               "|<MonthlyDueDate5|<MonthlyDueDate6|<MonthlyDueDate7|<MonthlyDueDate8" & _
               "|<MonthlyDueDate9|<MonthlyDueDate10|<MonthlyDueDate11|<MonthlyDueDate12|<NameOfSet"
   flxMthPayDt(0).FormatString = szHeader$
   flxMthPayDt(0).RowHeight(0) = 0
   flxMthPayDt(0).ColWidth(0) = txtMthPayDt(0).Left - flxMthPayDt(0).Left
   iFlxCol = 1
   For i = 0 To 11
      flxMthPayDt(0).ColWidth(iFlxCol) = txtMthPayDt(i + 1).Left - txtMthPayDt(i).Left
      iFlxCol = iFlxCol + 1
   Next i
   flxMthPayDt(0).ColWidth(iFlxCol) = flxMthPayDt(0).Width + flxMthPayDt(0).Left - txtMthPayDt(i).Left - 300
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
   szHeader$ = "<DateSetID|<QuarterlyDueDate1|<QuarterlyDueDate2|<QuarterlyDueDate3|<QuarterlyDueDate4"
   flxMthPayDt(1).FormatString = szHeader$
   flxMthPayDt(1).RowHeight(0) = 0
   flxMthPayDt(1).ColWidth(0) = txtMthPayDt(13).Left - flxMthPayDt(1).Left
   iFlxCol = 1
   For i = 13 To 15
      flxMthPayDt(1).ColWidth(iFlxCol) = txtMthPayDt(i + 1).Left - txtMthPayDt(i).Left
      iFlxCol = iFlxCol + 1
   Next i
   flxMthPayDt(1).ColWidth(iFlxCol) = flxMthPayDt(1).Width + flxMthPayDt(1).Left - txtMthPayDt(i).Left - 300
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
   szHeader$ = "<DateSetID|<HalfYearlyDueDate1|<HalfYearlyDueDate2"
   flxMthPayDt(2).FormatString = szHeader$
   flxMthPayDt(2).RowHeight(0) = 0
   flxMthPayDt(2).ColWidth(0) = txtMthPayDt(17).Left - flxMthPayDt(2).Left
   flxMthPayDt(2).ColWidth(1) = txtMthPayDt(18).Left - txtMthPayDt(17).Left
   flxMthPayDt(2).ColWidth(2) = flxMthPayDt(2).Width + flxMthPayDt(2).Left - txtMthPayDt(18).Left - 300
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
   szHeader$ = "<DateSetID|<YearlyDueDate"
   flxMthPayDt(3).FormatString = szHeader$
   flxMthPayDt(3).RowHeight(0) = 0
   flxMthPayDt(3).ColWidth(0) = txtMthPayDt(19).Left - flxMthPayDt(3).Left
   flxMthPayDt(3).ColWidth(1) = txtMthPayDt(19).Width - 300
End Sub

Private Sub txtMthPayDt_Change(Index As Integer)
   If Index = 12 Then
      txtDescription(0).text = txtMthPayDt(Index).text
      txtDescription(1).text = txtMthPayDt(Index).text
      txtDescription(2).text = txtMthPayDt(Index).text
   End If
End Sub

Private Sub txtMthPayDt_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 12 Then Exit Sub

   bFS = True
   If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then KeyAscii = 0
   If KeyAscii = 8 And Len(txtMthPayDt(Index).text) = 3 Then bFS = False
End Sub

Private Sub txtMthPayDt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 12 Then Exit Sub

   If Len(txtMthPayDt(Index).text) = 2 And bFS Then
      txtMthPayDt(Index).text = txtMthPayDt(Index).text & "/"
      bFS = True
   End If
   txtMthPayDt(Index).SelStart = Len(txtMthPayDt(Index).text)
   If Len(txtMthPayDt(Index).text) = 5 And Index < 19 Then
      txtMthPayDt(Index + 1).SetFocus
      SelTxtInCtrl txtMthPayDt(Index + 1)
   End If
End Sub

Private Sub txtMthPayDt_LostFocus(Index As Integer)
   If Index = 12 Then Exit Sub

   If txtMthPayDt(Index).text = "" Then Exit Sub
   If Not IsDate(txtMthPayDt(Index).text & " " & Year(Now)) Then
      ShowMsgInTaskBar "Please enter a valid date (dd/mm).", , "N"
      txtMthPayDt(Index).SetFocus
   Else
      txtMthPayDt(Index).text = Format(CDate(txtMthPayDt(Index).text & " " & Year(Now)), "dd/mm")
   End If

   Dim i As Integer
   If Index > 0 And Index < 12 Then
      For i = 0 To 11
         If txtMthPayDt(i).text = txtMthPayDt(Index).text And i <> Index Then
            ShowMsgInTaskBar "Two payment dates cannot be same.", , "N"
            txtMthPayDt(Index).text = ""
            txtMthPayDt(Index).SetFocus
            Exit For
         End If
      Next i
   End If
   If Index > 12 And Index < 17 Then
      For i = 13 To 16
         If txtMthPayDt(i).text = txtMthPayDt(Index).text And i <> Index Then
            ShowMsgInTaskBar "Two payment dates cannot be same.", , "N"
            txtMthPayDt(Index).text = ""
            txtMthPayDt(Index).SetFocus
            Exit For
         End If
      Next i
   End If
   If Index = 17 Or Index = 18 Then
      If txtMthPayDt(17).text = txtMthPayDt(18).text Then
         ShowMsgInTaskBar "Two payment dates cannot be same.", , "N"
         txtMthPayDt(Index).text = ""
         txtMthPayDt(Index).SetFocus
      End If
   End If
End Sub
